"""
비상장 외감법인 감사보고서(OpenDART)에서 감사시간 / 자산총계 / 매출액을 뽑아
'규모 대비 감사시간이 많은 회사' 순위를 만드는 스크립트.

사용 예:
  python dart_audit_hours.py --start 20260101 --end 20260630 --period 2025.12 --limit 200 --max-assets 1500
  (--limit 없이 돌리면 전체. 일일 호출 한도에 걸리면 저장 후 종료, 다음날 같은 명령으로 이어서 실행)
  파싱 결과는 parse_progress.jsonl에 쌓인다. 중간에 멈춰도 같은 명령으로 다시 실행하면 남은 건만 처리한다.
  파싱 로직을 고친 뒤에는 --fresh로 처음부터 다시 돌릴 것.

필요 패키지:
  pip install requests beautifulsoup4 lxml pandas numpy
"""
import os
import re
import sys
import json
import time
import zipfile
import argparse
from collections import Counter
from datetime import datetime, timedelta

import numpy as np
import pandas as pd
import requests
from bs4 import BeautifulSoup
import warnings
try:
    from bs4 import XMLParsedAsHTMLWarning
    warnings.filterwarnings("ignore", category=XMLParsedAsHTMLWarning)
except ImportError:
    pass

API = "https://opendart.fss.or.kr/api"
KEY = os.environ.get("DART_API_KEY")
CACHE_DIR = "dart_cache"
FILINGS_CSV = "filings.csv"
PROGRESS_FILE = "parse_progress.jsonl"


class LimitReached(Exception):
    pass


# ---------------------------------------------------------------- 1. 공시 목록
def month_ranges(start, end):
    s = datetime.strptime(start, "%Y%m%d")
    e = datetime.strptime(end, "%Y%m%d")
    while s <= e:
        nxt = (s.replace(day=1) + timedelta(days=32)).replace(day=1)
        chunk_end = min(nxt - timedelta(days=1), e)
        yield s.strftime("%Y%m%d"), chunk_end.strftime("%Y%m%d")
        s = nxt


def fetch_filings(start, end):
    """외부감사관련(F) > 감사보고서(F001), 기타법인(E=비상장) 목록"""
    rows = []
    for bgn, end_ in month_ranges(start, end):
        page = 1
        while True:
            r = requests.get(f"{API}/list.json", params=dict(
                crtfc_key=KEY, bgn_de=bgn, end_de=end_,
                pblntf_ty="F", pblntf_detail_ty="F001", corp_cls="E",
                page_no=page, page_count=100), timeout=30)
            d = r.json()
            if d["status"] == "013":          # 조회 데이터 없음
                break
            if d["status"] == "020":
                raise LimitReached("목록 조회 중 일일 한도 초과")
            if d["status"] != "000":
                raise RuntimeError(d)
            rows += d["list"]
            print(f"  목록 {bgn}~{end_} {page}/{d['total_page']}페이지")
            if page >= d["total_page"]:
                break
            page += 1
            time.sleep(0.3)
    return pd.DataFrame(rows)


# ---------------------------------------------------------------- 2. 원문 다운로드
def decode(raw):
    for enc in ("utf-8", "cp949"):
        try:
            return raw.decode(enc)
        except UnicodeDecodeError:
            pass
    return raw.decode("utf-8", errors="ignore")


def get_doc_texts(rcept_no):
    os.makedirs(CACHE_DIR, exist_ok=True)
    path = os.path.join(CACHE_DIR, f"{rcept_no}.zip")
    if not os.path.exists(path) or os.path.getsize(path) == 0:   # 0바이트 = 끊긴 다운로드, 다시 받는다
        r = requests.get(f"{API}/document.xml",
                         params=dict(crtfc_key=KEY, rcept_no=rcept_no), timeout=60)
        if r.content[:2] != b"PK":            # zip이 아니면 오류 응답
            msg = decode(r.content[:500])
            if "020" in msg:
                raise LimitReached("원문 다운로드 중 일일 한도 초과")
            raise ValueError(msg)
        tmp = path + ".part"
        with open(tmp, "wb") as f:
            f.write(r.content)
        os.replace(tmp, path)
        time.sleep(0.3)
    with zipfile.ZipFile(path) as z:
        return [decode(z.read(n)) for n in z.namelist()]


# ---------------------------------------------------------------- 3. 파싱
NUM_RE = re.compile(r"^\(?-?[\d,]+(\.\d+)?\)?$")
CELL_TAGS = ["td", "th", "te", "tu"]       # DART XML은 TE/TU 셀 태그도 사용


def to_num(s):
    s = s.replace(" ", "")
    if not s or not NUM_RE.match(s):
        return None
    neg = s.startswith("(") or s.startswith("-")
    v = float(s.strip("()").replace(",", "").lstrip("-"))
    return -v if neg else v


def table_rows(tbl):
    out = []
    for tr in tbl.find_all("tr"):
        cells = [c.get_text(" ", strip=True) for c in tr.find_all(CELL_TAGS)]
        if cells:
            out.append(cells)
    return out


def row_nums(cells):
    return [n for n in (to_num(c) for c in cells) if n is not None]


def norm(s):
    s = re.sub(r"\s+", "", s)
    s = re.sub(r"^[ⅠⅡⅢⅣⅤⅥⅦⅧⅨⅩIVX\d\.\(\)]+", "", s)   # 'Ⅰ.' '1.' 같은 앞번호 제거
    return s


def cell_num(s):
    """'-' 또는 빈칸은 0, 숫자면 float, 그 외는 None"""
    t = s.replace(" ", "")
    if t in ("-", "−", "–", ""):
        return 0.0
    return to_num(t)


def parse_hours(soup):
    """외부감사 실시내용 > 감사참여자 구분별 인원수 및 감사시간 표.
    당기/전기가 '열'로 배치되어 있고, 투입시간 '합계' 행의 맨 끝 두 칸이 합계 당기·전기."""
    for tbl in soup.find_all("table"):
        flat = tbl.get_text("").replace(" ", "")
        if "투입" not in flat or "인원수" not in flat:
            continue
        rows = table_rows(tbl)
        for cells in reversed(rows):              # 투입시간 합계 행은 표의 마지막 쪽
            if len(cells) < 3:
                continue
            if "합계" not in [c.replace(" ", "") for c in cells[:2]]:
                continue
            cur, prev = cell_num(cells[-2]), cell_num(cells[-1])
            if cur is None or prev is None:       # 머리글 행(합계/당기/전기 글자) 건너뜀
                continue
            if cur <= 0:
                continue
            return cur, (prev if prev > 0 else None)   # 전기 '-' = 초도감사 가능성
    return None, None


def unit_of(tbl):
    txt = tbl.get_text(" ")
    m = re.search(r"단\s*위\s*[:：]?\s*(백만원|천원|원)", txt)
    if not m:
        prev = tbl.find_previous(string=re.compile(r"단\s*위"))
        if prev:
            m = re.search(r"(백만원|천원|원)", str(prev))
    unit = m.group(1) if m else "원"
    return {"백만원": 1e6, "천원": 1e3, "원": 1.0}[unit]


def first_big(cells):
    """주석번호 같은 작은 숫자는 건너뛰고 첫 금액(당기)"""
    for n in row_nums(cells[1:]):
        if abs(n) >= 1000:
            return n
    return None


REV_LABELS = {"매출액", "매출", "영업수익", "수익(매출액)", "매출수익"}


def parse_fin(soup):
    assets = revenue = None
    for tbl in soup.find_all("table"):
        for cells in table_rows(tbl):
            label = norm(cells[0])
            if assets is None and label == "자산총계":
                v = first_big(cells)
                if v is not None:
                    assets = v * unit_of(tbl) / 1e8          # 억원
            if revenue is None and (label in REV_LABELS or label.startswith("매출액")):
                v = first_big(cells)
                if v is not None:
                    revenue = v * unit_of(tbl) / 1e8
        if assets is not None and revenue is not None:
            break
    return assets, revenue


def parse_auditor(full_text):
    names = re.findall(r"([가-힣A-Za-z]{1,15}회계법인)", full_text)
    names = [n for n in names if n != "회계법인"]
    return Counter(names).most_common(1)[0][0] if names else None


def parse_gaap(full_text):
    """감사의견/경영진 책임 문단에서 '...기준에 따라 ... 공정하게 표시' 앞의 회계기준 판별"""
    t = re.sub(r"\s+", "", full_text)
    for m in re.finditer("공정하게표시", t):
        window = t[max(0, m.start() - 400):m.start()]
        if "한국채택국제회계기준" in window:
            return "K-IFRS"
        if "일반기업회계기준" in window:
            return "일반기업회계기준"
        if "중소기업회계기준" in window:
            return "중소기업회계기준"
    return None


def parse_report(rcept_no):
    res = dict(hours=None, hours_prev=None, assets_eok=None, revenue_eok=None, auditor=None)
    full = ""
    for txt in get_doc_texts(rcept_no):
        soup = BeautifulSoup(txt, "lxml")
        full += soup.get_text(" ")
        if res["hours"] is None:
            res["hours"], res["hours_prev"] = parse_hours(soup)
        if res["assets_eok"] is None or res["revenue_eok"] is None:
            a, r = parse_fin(soup)
            res["assets_eok"] = res["assets_eok"] or a
            res["revenue_eok"] = res["revenue_eok"] or r
    res["auditor"] = parse_auditor(full)
    res["gaap"] = parse_gaap(full)
    return res


# ---------------------------------------------------------------- 3-1. 이어받기
def load_progress():
    """이전 실행에서 파싱을 마친 건. 오류 건은 다시 시도하도록 넣지 않는다."""
    done = {}
    if not os.path.exists(PROGRESS_FILE):
        return done
    with open(PROGRESS_FILE, encoding="utf-8") as f:
        for line in f:
            try:
                r = json.loads(line)
            except ValueError:                # 강제 종료로 잘린 줄
                continue
            if isinstance(r, dict) and r.get("rcept_no") and not r.get("error"):
                done[r["rcept_no"]] = r
    return done


def open_progress(fresh):
    if fresh or not os.path.exists(PROGRESS_FILE):
        return open(PROGRESS_FILE, "w", encoding="utf-8")
    with open(PROGRESS_FILE, "rb") as f:      # 잘린 마지막 줄 뒤에 바로 이어 쓰지 않도록
        f.seek(0, os.SEEK_END)
        cut = f.tell() > 0 and (f.seek(-1, os.SEEK_END), f.read(1))[1] != b"\n"
    pf = open(PROGRESS_FILE, "a", encoding="utf-8")
    if cut:
        pf.write("\n")
    return pf


# ---------------------------------------------------------------- 4. 순위
def rank(df, max_assets=None):
    df = df.copy()
    for c in ["hours", "hours_prev", "assets_eok", "revenue_eok"]:
        df[c] = pd.to_numeric(df[c], errors="coerce").astype(float)
    ok = df.dropna(subset=["hours", "assets_eok"]).copy()
    if max_assets is not None:
        before = len(ok)
        ok = ok[ok["assets_eok"] <= max_assets]
        print(f"자산 {max_assets:,.0f}억 이하 필터: {before}건 → {len(ok)}건")
    if len(ok) < 10:
        print(f"\n[주의] 감사시간·자산이 모두 추출된 건이 {len(ok)}건뿐이라 순위를 계산하지 않습니다. parse_failed.csv를 확인하세요.")
        return None
    ok = ok[(ok["hours"] > 0) & (ok["assets_eok"] > 0)]
    ok["revenue_eok"] = ok["revenue_eok"].clip(lower=0)
    # log(감사시간) = a + b*log(자산) + c*log(매출+1) 회귀 → 잔차가 클수록 규모 대비 시간 과다
    # 매출을 못 읽은 회사를 매출 0으로 넣으면 예측시간이 낮아져 배수가 부풀려진다.
    # → 매출이 있는 회사만으로 적합하고, 매출 없는 회사는 자산만 쓴 회귀로 예측한다.
    has_rev = ok["revenue_eok"].notna().to_numpy()
    y = np.log(ok["hours"]).to_numpy()
    la = np.log(ok["assets_eok"]).to_numpy()
    X = np.column_stack([np.ones(len(ok)), la, np.log1p(ok["revenue_eok"].fillna(0))])
    coef, *_ = np.linalg.lstsq(X[has_rev], y[has_rev], rcond=None)
    coef_a, *_ = np.linalg.lstsq(X[:, :2], y, rcond=None)
    pred = np.where(has_rev, X @ coef, X[:, :2] @ coef_a)
    print(f"매출 미확인 {int((~has_rev).sum())}건은 자산만으로 예측 (상수 {coef_a[0]:.3f}, log자산 {coef_a[1]:.3f})")
    ok["예측시간"] = np.exp(pred).round(0)
    ok["예측대비배수"] = (ok["hours"] / ok["예측시간"]).round(2)
    ok["시간_per_자산억"] = (ok["hours"] / ok["assets_eok"]).round(2)
    ok["시간_per_매출억"] = (ok["hours"] / ok["revenue_eok"].replace(0, np.nan)).round(2)
    ok["전기대비증가율"] = ((ok["hours"] / ok["hours_prev"]) - 1).round(3)
    print(f"\n회귀계수: 상수 {coef[0]:.3f}, log자산 {coef[1]:.3f}, log매출 {coef[2]:.3f} (n={int(has_rev.sum())})")
    return ok.sort_values("예측대비배수", ascending=False)


# ---------------------------------------------------------------- main
def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--start", required=True, help="접수일 시작 yyyymmdd")
    ap.add_argument("--end", required=True, help="접수일 끝 yyyymmdd")
    ap.add_argument("--period", default=None, help="보고서명 결산기 필터 예) 2025.12")
    ap.add_argument("--limit", type=int, default=None, help="테스트용 처리 건수 제한")
    ap.add_argument("--include-kifrs", action="store_true", help="K-IFRS 적용 회사도 포함(기본은 제외)")
    ap.add_argument("--max-assets", type=float, default=1500, help="자산총계 상한(억원), 기본 1500")
    ap.add_argument("--fresh", action="store_true", help=f"{PROGRESS_FILE}을 무시하고 처음부터 다시 파싱")
    args = ap.parse_args()

    if not KEY:
        sys.exit("DART_API_KEY 환경변수가 없습니다.")

    if os.path.exists(FILINGS_CSV):
        filings = pd.read_csv(FILINGS_CSV, dtype=str)
        print(f"기존 목록 사용: {len(filings)}건 (새로 받으려면 {FILINGS_CSV} 삭제)")
    else:
        print("공시 목록 조회 중...")
        filings = fetch_filings(args.start, args.end)
        filings.to_csv(FILINGS_CSV, index=False, encoding="utf-8-sig")
        print(f"목록 {len(filings)}건 저장")

    if args.period:
        filings = filings[filings["report_nm"].str.contains(re.escape(args.period), na=False)]
    n0 = len(filings)
    # 연결감사보고서 제외 (중복 제거 전에 먼저 빼야 별도 감사보고서가 남음)
    filings = filings[~filings["report_nm"].str.contains("연결", na=False)]
    n1 = len(filings)
    # 상장회사 제외: 법인구분 E(기타)만, 종목코드 있는 회사 제외
    if "corp_cls" in filings:
        filings = filings[filings["corp_cls"] == "E"]
    if "stock_code" in filings:
        filings = filings[filings["stock_code"].fillna("").str.strip() == ""]
    n2 = len(filings)
    print(f"필터: 결산기 해당 {n0}건 → 연결 제외 {n1}건 → 상장사 제외 {n2}건")
    # 정정 공시가 있으면 회사별 최신 접수번호만 사용
    filings = filings.sort_values("rcept_no").drop_duplicates("corp_code", keep="last")
    if args.limit:
        filings = filings.head(args.limit)
    print(f"처리 대상 {len(filings)}건")

    done = {} if args.fresh else load_progress()
    results = [done[r] for r in filings["rcept_no"] if r in done]
    todo = filings[~filings["rcept_no"].isin(done)]
    if results:
        print(f"이어받기: {len(results)}건은 {PROGRESS_FILE}에서 불러옴 → 남은 {len(todo)}건 "
              f"(처음부터 다시 하려면 --fresh)")

    with open_progress(args.fresh) as pf:
        try:
            for i, row in enumerate(todo.itertuples(), 1):
                try:
                    r = parse_report(row.rcept_no)
                except LimitReached:
                    raise
                except Exception as e:
                    r = dict(error=str(e)[:200])
                r.update(corp_name=row.corp_name, corp_code=row.corp_code,
                         rcept_no=row.rcept_no, report_nm=row.report_nm)
                results.append(r)
                pf.write(json.dumps(r, ensure_ascii=False) + "\n")
                if i % 50 == 0:
                    pf.flush()
                    print(f"  {len(results)}/{len(filings)} 처리")
        except LimitReached as e:
            print(f"\n[중단] {e}. 받은 원문은 {CACHE_DIR}에, 파싱 결과는 {PROGRESS_FILE}에 저장됨 "
                  f"→ 내일 같은 명령으로 이어서 실행")
        except KeyboardInterrupt:
            pf.flush()
            sys.exit(f"\n[중단] {len(results)}/{len(filings)}건까지 {PROGRESS_FILE}에 저장됨 "
                     f"→ 같은 명령으로 이어서 실행 (순위 파일은 만들지 않음)")

    df = pd.DataFrame(results)
    for col in ["hours", "hours_prev", "assets_eok", "revenue_eok", "auditor", "gaap", "error"]:
        if col not in df:
            df[col] = None

    failed = df[df["hours"].isna() | df["assets_eok"].isna()]
    failed.to_csv("parse_failed.csv", index=False, encoding="utf-8-sig")

    n = len(df)
    print(f"\n추출 성공률: 감사시간 {df['hours'].notna().sum()}/{n}, "
          f"자산총계 {df['assets_eok'].notna().sum()}/{n}, "
          f"매출액 {df['revenue_eok'].notna().sum()}/{n}, "
          f"다운로드·기타오류 {df['error'].notna().sum()}/{n}")

    g = df["gaap"].fillna("판별불가")
    print("회계기준: " + ", ".join(f"{k} {v}건" for k, v in g.value_counts().items()))
    if not args.include_kifrs:
        df = df[df["gaap"] != "K-IFRS"]
        print(f"K-IFRS 제외 후 {len(df)}건")

    ranked = rank(df, args.max_assets)
    if ranked is None:
        return
    cols = ["corp_name", "auditor", "gaap", "hours", "hours_prev", "전기대비증가율",
            "assets_eok", "revenue_eok", "예측시간", "예측대비배수",
            "시간_per_자산억", "시간_per_매출억", "report_nm", "rcept_no"]
    ranked[cols].to_csv("audit_hours_rank.csv", index=False, encoding="utf-8-sig")
    print(f"\n완료: 순위 {len(ranked)}건 → audit_hours_rank.csv / 파싱실패 {len(failed)}건 → parse_failed.csv")


if __name__ == "__main__":
    main()
