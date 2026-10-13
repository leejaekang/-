# -*- coding: utf-8 -*-
"""
감사인 교체 가능성 신호 점수 (지인회계법인 외부감사 영업용, 내부 타깃팅 전용)

입력
  - mailing_list.xlsx (있으면 우선) 또는 audit_hours_rank.csv
  - contacts_cache.csv (주소·대표자 보강, 있으면)
  - filings.csv (2026 접수분) + filings_hist/*.csv (2024·2025 접수분, --fetch-history로 수집)
  - exclude_clients.csv (기장 거래처 등 제외 목록, 있으면. corp_code 또는 회사명 컬럼)

출력
  - signals_all.csv       전체 후보와 신호별 점수
  - signals_excluded.csv  제외된 회사와 사유
  - signals_top.xlsx      상위 N곳 (기본 150)
  - signals_stage2_cache.csv  2단계 원문 판독 캐시

사용 (PowerShell, 실행 전 결과 엑셀 파일은 닫아 둘 것)
  python dart_signals.py --write-config               # signals_config.json 생성 → 가중치 수정
  python dart_signals.py --fetch-history              # 2024·2025 목록 수집 (월별 캐시, 끊겨도 이어받음)
  python dart_signals.py                              # 1단계 점수
  python dart_signals.py --stage2 --stage2-n 400      # 상위 400곳 원문으로 대표이사 변경·의견 확인

API 키는 환경변수 DART_API_KEY로만 읽는다.
"""
import argparse
import copy
import html
import json
import math
import os
import re
import sys
import time
import zipfile
from collections import Counter
from datetime import date
from pathlib import Path

import numpy as np
import pandas as pd
import requests

BASE = Path(__file__).resolve().parent

# ----------------------------------------------------------------------------
# 설정 (signals_config.json이 있으면 같은 키를 덮어쓴다)
# ----------------------------------------------------------------------------
DEFAULT_CONFIG = {
    "period": "2025.12",           # 대상 결산기
    "top_n": 150,                  # signals_top.xlsx에 넣을 개수
    "max_assets_eok": 1000,        # 자산 상한(억). 주 타깃 1,000억 미만
    "asset_unit": "auto",          # 입력 자산·매출 단위: auto / 원 / 천원 / 백만원 / 억
    "our_firm": ["지인"],           # 우리 법인 (감사인이면 이미 고객 → 제외)

    "weights": {
        "multiple": 30,            # 예측대비배수: 1.0배 → 0점, multiple_full배 이상 → 만점
        "std_hours": 10,           # 자산 200억 이상만: 최소 표준감사시간 대비 초과
        "hours_yoy": 5,            # 전기 대비 감사시간 증가율이 hours_yoy_min 이상
        "big4_small": 15,          # 자산 small_asset_eok 미만인데 Big4
        "mid_small": 8,            # 자산 small_asset_eok 미만인데 준대형
        "tenure_long": 8,          # 3개년 같은 감사인 (교체 검토 시점 가능성)
        "just_changed": -15,       # 2025.12 결산에서 막 교체 (초년도, 재교체 가능성 낮음)
        "correction": 10,          # [기재정정] 공시
        "late": 8,                 # 전년보다 late_days일 이상 늦게 제출
        "ceo_change": 15,          # 2단계: 대표이사 변경
    },
    "multiple_full": 2.0,
    "std_exempt_eok": 200,         # 2025~2027 표준감사시간 적용 면제 구간(자산 200억 미만)
    "std_full_ratio": 1.5,         # 최소 표준감사시간의 1.5배 이상이면 std_hours 만점
    "hours_yoy_min": 20.0,         # %
    "small_asset_eok": 300,
    "late_days": 30,

    # 감사인 이름은 '회계법인/유한/(주)/공백'을 지운 짧은 이름으로 비교한다.
    # 준대형 목록은 예시다. 실제 기준(매출 순위 등)으로 고쳐 쓸 것.
    "big4": ["삼일", "삼정", "안진", "한영"],
    "mid_tier": ["삼덕", "대주", "신한", "한울", "BDO성현", "성현", "이촌", "우리", "정동", "대현", "삼화", "안경"],

    # 주소 근접 가점: 위에서부터 처음 맞는 규칙 하나만 적용
    "proximity": [
        ["강남구", 15], ["서초구", 15],
        ["송파구", 10], ["성남시", 10],
        ["강동구", 6], ["광진구", 6], ["용산구", 6], ["성동구", 6], ["하남시", 6], ["과천시", 6],
        ["서울", 4],
    ],

    # 제외 규칙
    "exclude_frequent_change": 2,  # 3개 결산기 사이 감사인 교체 횟수가 이 값 이상이면 제외
    "exclude_nonclean_opinion": True,
    "exclude_going_concern": True,

    # 입력 컬럼명 수동 지정 (자동 인식이 틀릴 때만). 예: {"multiple": "예측대비배수"}
    "column_map": {},

    # 2024·2025 목록 수집 범위 (접수연도, 월)
    "history_years": [2024, 2025],
    "history_months": [1, 2, 3, 4, 5, 6],
}

CONFIG_FILE = BASE / "signals_config.json"
FILINGS_FILE = BASE / "filings.csv"
HIST_DIR = BASE / "filings_hist"
CACHE_DIR = BASE / "dart_cache"
STAGE2_CACHE = BASE / "signals_stage2_cache.csv"
EXCLUDE_FILE = BASE / "exclude_clients.csv"

OUT_ALL = BASE / "signals_all.csv"
OUT_EXCL = BASE / "signals_excluded.csv"
OUT_TOP = BASE / "signals_top.xlsx"

API = "https://opendart.fss.or.kr/api"

# 입력 컬럼 후보 (정확히 같은 이름 우선, 없으면 포함 여부)
COL_CANDIDATES = {
    "corp_code": ["corp_code", "고유번호", "회사코드"],
    "corp_name": ["corp_name", "회사명", "기업명", "법인명"],
    "rcept_no": ["rcept_no", "접수번호"],
    "assets": ["자산총계", "자산총액", "자산", "assets", "total_assets"],
    "sales": ["매출액", "매출", "sales", "revenue"],
    "multiple": ["예측대비배수", "배수", "multiple"],
    "hours": ["감사시간_당기", "당기감사시간", "투입시간_당기", "감사시간(당기)", "당기_감사시간", "감사시간"],
    "hours_yoy": ["전기대비증가율", "증가율"],
    "std_hours": ["최소표준감사시간", "최소 표준감사시간", "표준감사시간"],
    "address": ["주소", "adres", "address", "소재지"],
    "ceo": ["대표자", "대표이사", "ceo_nm"],
}


class DartLimit(Exception):
    """status 020: 하루 한도 초과"""


# ----------------------------------------------------------------------------
# 공통 유틸
# ----------------------------------------------------------------------------
def load_config():
    cfg = copy.deepcopy(DEFAULT_CONFIG)
    if CONFIG_FILE.exists():
        user = json.loads(CONFIG_FILE.read_text(encoding="utf-8"))
        for k, v in user.items():
            if isinstance(v, dict) and isinstance(cfg.get(k), dict):
                cfg[k].update(v)
            else:
                cfg[k] = v
        print(f"[설정] {CONFIG_FILE.name} 적용")
    return cfg


def read_table(path):
    path = Path(path)
    if path.suffix.lower() in (".xlsx", ".xls"):
        return pd.read_excel(path, dtype=str)
    for enc in ("utf-8-sig", "cp949"):
        try:
            return pd.read_csv(path, dtype=str, encoding=enc)
        except UnicodeDecodeError:
            continue
    raise RuntimeError(f"인코딩을 알 수 없음: {path}")


def safe_write(write_fn, path):
    """엑셀로 열려 있으면 PermissionError → 안내 후 다른 이름으로 저장"""
    try:
        write_fn(path)
        return path
    except PermissionError:
        alt = path.with_name(f"{path.stem}_{time.strftime('%H%M%S')}{path.suffix}")
        print(f"  ! {path.name} 이(가) 열려 있어 저장 실패 → {alt.name} 으로 저장 (다음엔 엑셀을 닫고 실행)")
        write_fn(alt)
        return alt


def to_num(s):
    return pd.to_numeric(
        s.astype(str).str.replace(",", "", regex=False).str.replace("%", "", regex=False).str.strip()
        .replace({"": np.nan, "nan": np.nan, "None": np.nan, "-": np.nan}),
        errors="coerce",
    )


def find_col(df, key, cfg):
    forced = cfg["column_map"].get(key)
    if forced:
        if forced not in df.columns:
            sys.exit(f"[오류] column_map의 '{forced}' 컬럼이 입력에 없습니다. 실제 컬럼: {list(df.columns)}")
        return forced
    cands = COL_CANDIDATES[key]
    norm = {c: re.sub(r"\s+", "", str(c)) for c in df.columns}
    for cand in cands:
        for c, n in norm.items():
            if n == re.sub(r"\s+", "", cand):
                return c
    for cand in cands:
        cn = re.sub(r"\s+", "", cand)
        for c, n in norm.items():
            if cn in n and not (key == "hours" and re.search("전기|표준|증가", n)):
                return c
    return None


def short_auditor(name):
    if not isinstance(name, str):
        return ""
    s = re.sub(r"\(주\)|주식회사|유한|회계법인|감사반|\s+", "", name)
    return s.strip()


def is_correction(report_nm):
    return isinstance(report_nm, str) and "정정" in report_nm


def period_of(report_nm):
    m = re.search(r"\((\d{4}\.\d{2})\)", str(report_nm))
    return m.group(1) if m else None


def shift_period(period, years):
    y, m = period.split(".")
    return f"{int(y) - years}.{m}"


def api_key():
    key = os.environ.get("DART_API_KEY", "").strip()
    if not key:
        sys.exit("[오류] 환경변수 DART_API_KEY가 없습니다. PowerShell: $env:DART_API_KEY = \"발급키\"")
    return key


# ----------------------------------------------------------------------------
# 2024·2025 공시 목록 수집
# ----------------------------------------------------------------------------
LIST_FIELDS = ["corp_code", "corp_name", "stock_code", "corp_cls", "report_nm", "rcept_no", "flr_nm", "rcept_dt", "rm"]


def fetch_month(key, year, month, session):
    start = date(year, month, 1)
    end = date(year + (month == 12), month % 12 + 1, 1)
    end = date.fromordinal(end.toordinal() - 1)
    rows, page, total_page = [], 1, 1
    while page <= total_page:
        params = {
            "crtfc_key": key, "bgn_de": start.strftime("%Y%m%d"), "end_de": end.strftime("%Y%m%d"),
            "pblntf_ty": "F", "pblntf_detail_ty": "F001", "corp_cls": "E",
            "page_no": page, "page_count": 100,
        }
        for attempt in range(4):
            try:
                r = session.get(f"{API}/list.json", params=params, timeout=30)
                data = r.json()
                break
            except (requests.RequestException, ValueError):
                if attempt == 3:
                    raise
                time.sleep(2 ** (attempt + 1))
        st = data.get("status")
        if st == "013":
            break
        if st == "020":
            raise DartLimit()
        if st != "000":
            raise RuntimeError(f"list.json 오류 {st}: {data.get('message')}")
        total_page = int(data.get("total_page", 1))
        rows.extend({f: it.get(f, "") for f in LIST_FIELDS} for it in data.get("list", []))
        page += 1
        time.sleep(0.15)
    return pd.DataFrame(rows, columns=LIST_FIELDS)


def fetch_history(cfg, refresh=False):
    key = api_key()
    HIST_DIR.mkdir(exist_ok=True)
    s = requests.Session()
    for y in cfg["history_years"]:
        for m in cfg["history_months"]:
            out = HIST_DIR / f"F001_E_{y}{m:02d}.csv"
            if out.exists() and not refresh:
                print(f"  {out.name} 있음 → 건너뜀")
                continue
            try:
                df = fetch_month(key, y, m, s)
            except DartLimit:
                print("  ! 하루 호출 한도 초과(020). 내일 같은 명령으로 이어받으면 됩니다.")
                return
            df.to_csv(out, index=False, encoding="utf-8-sig")
            print(f"  {out.name}: {len(df):,}건")
    print("[완료] 목록 수집")


def load_filings():
    frames = []
    if FILINGS_FILE.exists():
        frames.append(read_table(FILINGS_FILE))
    else:
        print(f"  ! {FILINGS_FILE.name} 없음")
    if HIST_DIR.exists():
        frames += [read_table(p) for p in sorted(HIST_DIR.glob("*.csv"))]
    if not frames:
        return pd.DataFrame(columns=LIST_FIELDS)
    f = pd.concat(frames, ignore_index=True)
    f = f.drop_duplicates("rcept_no")
    f["corp_code"] = f["corp_code"].astype(str).str.zfill(8)
    f["period"] = f["report_nm"].map(period_of)
    f["is_corr"] = f["report_nm"].map(is_correction)
    f["is_gijae"] = f["report_nm"].astype(str).str.contains("기재정정")
    f["rcept_dt"] = pd.to_datetime(f["rcept_dt"].astype(str), format="%Y%m%d", errors="coerce")
    return f


def filing_summary(f, period):
    """회사별: 원본 최초 접수일, 감사인(최신 접수 기준), 기재정정 여부, 원본·최신 접수번호"""
    g = f[f["period"] == period].sort_values(["corp_code", "rcept_dt", "rcept_no"])
    if g.empty:
        return pd.DataFrame(columns=["corp_code", "auditor", "orig_dt", "gijae", "rcept_no_latest", "rcept_no_orig"])
    latest = g.groupby("corp_code").tail(1).set_index("corp_code")
    orig = g[~g["is_corr"]].groupby("corp_code").head(1).set_index("corp_code")
    out = pd.DataFrame({
        "auditor": latest["flr_nm"],
        "rcept_no_latest": latest["rcept_no"],
    })
    out["orig_dt"] = orig["rcept_dt"]
    out["rcept_no_orig"] = orig["rcept_no"]
    out["gijae"] = g.groupby("corp_code")["is_gijae"].any()
    return out.reset_index()


# ----------------------------------------------------------------------------
# 입력 로드
# ----------------------------------------------------------------------------
def detect_unit_factor(assets, unit):
    """입력 단위 → 억 원 환산 계수"""
    factors = {"원": 1e-8, "천원": 1e-5, "백만원": 1e-2, "억": 1.0}
    if unit != "auto":
        return factors[unit], unit
    med = assets.dropna().median()
    if pd.isna(med):
        return 1e-8, "원"
    if med > 1e8:
        u = "원"
    elif med > 1e5:
        u = "천원"
    elif med > 1e3:
        u = "백만원"
    else:
        u = "억"
    return factors[u], u


def load_base(cfg, filings, input_path=None):
    if input_path:
        src = Path(input_path)
    else:
        src = next((p for p in (BASE / "mailing_list.xlsx", BASE / "audit_hours_rank.csv") if p.exists()), None)
    if src is None or not src.exists():
        sys.exit("[오류] mailing_list.xlsx / audit_hours_rank.csv 가 없습니다. 전체 파싱이 끝났는지 확인하세요.")
    df = read_table(src)
    print(f"[입력] {src.name}: {len(df):,}행, 수정 {time.strftime('%Y-%m-%d %H:%M', time.localtime(src.stat().st_mtime))}")

    cols = {k: find_col(df, k, cfg) for k in COL_CANDIDATES}
    print("[컬럼 인식] " + ", ".join(f"{k}={v}" for k, v in cols.items() if v))
    missing = [k for k in ("assets", "multiple") if not cols[k]]
    if missing:
        sys.exit(f"[오류] 필수 컬럼 {missing} 인식 실패. 실제 컬럼: {list(df.columns)}\n"
                 f"       signals_config.json의 column_map에 지정하세요.")

    out = pd.DataFrame(index=df.index)
    if cols["corp_code"]:
        out["corp_code"] = df[cols["corp_code"]].astype(str).str.replace(r"\.0$", "", regex=True).str.zfill(8)
    elif cols["rcept_no"]:
        m = filings.set_index("rcept_no")["corp_code"]
        out["corp_code"] = df[cols["rcept_no"]].map(m)
    else:
        sys.exit("[오류] corp_code·rcept_no 둘 다 없어 공시 목록과 연결할 수 없습니다.")

    out["회사명"] = df[cols["corp_name"]] if cols["corp_name"] else ""
    raw_assets = to_num(df[cols["assets"]])
    factor, unit = detect_unit_factor(raw_assets, cfg["asset_unit"])
    print(f"[단위] 자산·매출 입력 단위 = {unit} (틀리면 signals_config.json의 asset_unit 지정)")
    out["자산_억"] = (raw_assets * factor).round(1)
    out["매출_억"] = (to_num(df[cols["sales"]]) * factor).round(1) if cols["sales"] else np.nan
    out["예측대비배수"] = to_num(df[cols["multiple"]])
    out["감사시간"] = to_num(df[cols["hours"]]) if cols["hours"] else np.nan
    if cols["hours_yoy"]:
        yoy = to_num(df[cols["hours_yoy"]])
        if yoy.dropna().abs().median() < 2:   # 0.25 같은 비율이면 %로
            yoy = yoy * 100
        out["감사시간증가율%"] = yoy.round(1)
    else:
        out["감사시간증가율%"] = np.nan
    out["최소표준감사시간"] = to_num(df[cols["std_hours"]]) if cols["std_hours"] else np.nan
    out["주소"] = df[cols["address"]] if cols["address"] else np.nan
    out["대표자"] = df[cols["ceo"]] if cols["ceo"] else np.nan

    # 원래 입력 컬럼(연락처 등)을 뒤에 보존
    extra = df.drop(columns=[c for c in cols.values() if c], errors="ignore")
    extra.columns = [f"원본_{c}" for c in extra.columns]
    out = pd.concat([out, extra], axis=1)

    contacts = BASE / "contacts_cache.csv"
    if contacts.exists() and (out["주소"].isna().all() or out["대표자"].isna().all()):
        c = read_table(contacts)
        cc = find_col(c, "corp_code", cfg)
        if cc:
            c[cc] = c[cc].astype(str).str.zfill(8)
            c = c.drop_duplicates(cc).set_index(cc)
            for key, col in (("address", "주소"), ("ceo", "대표자")):
                src_col = find_col(c, key, cfg)
                if src_col and out[col].isna().all():
                    out[col] = out["corp_code"].map(c[src_col])
            print("[보강] contacts_cache.csv 에서 주소·대표자")
    return out.dropna(subset=["corp_code"]).drop_duplicates("corp_code")


# ----------------------------------------------------------------------------
# 1단계 신호
# ----------------------------------------------------------------------------
def add_stage1(df, filings, cfg):
    p0 = cfg["period"]
    p1, p2 = shift_period(p0, 1), shift_period(p0, 2)
    s0, s1, s2 = (filing_summary(filings, p).set_index("corp_code") for p in (p0, p1, p2))
    print(f"[목록] {p0}: {len(s0):,}사, {p1}: {len(s1):,}사, {p2}: {len(s2):,}사")
    if len(s1) == 0:
        print("  ! 전년 목록이 없습니다. 감사인 변경·제출 지연 신호를 쓰려면 --fetch-history 먼저 실행")

    cc = df["corp_code"]
    df["감사인"] = cc.map(s0["auditor"])
    df["감사인_전년"] = cc.map(s1["auditor"])
    df["감사인_전전년"] = cc.map(s2["auditor"])
    df["접수번호"] = cc.map(s0["rcept_no_latest"])
    df["접수번호_전년"] = cc.map(s1["rcept_no_latest"])
    df["제출일"] = cc.map(s0["orig_dt"])
    df["제출일_전년"] = cc.map(s1["orig_dt"])
    df["기재정정"] = cc.map(s0["gijae"]).fillna(False).astype(bool)

    a0 = df["감사인"].map(short_auditor)
    a1 = df["감사인_전년"].map(short_auditor)
    a2 = df["감사인_전전년"].map(short_auditor)

    def n_changes(x, y, z):
        seq = [v for v in (z, y, x) if v]
        return sum(1 for i in range(1, len(seq)) if seq[i] != seq[i - 1])

    df["감사인교체횟수"] = [n_changes(x, y, z) for x, y, z in zip(a0, a1, a2)]
    df["올해교체"] = (a0 != "") & (a1 != "") & (a0 != a1)
    df["3년동일"] = (a0 != "") & (a0 == a1) & (a1 == a2)
    df["전년보고서없음"] = a1 == ""

    def dayshift(d0, d1):
        if pd.isna(d0) or pd.isna(d1):
            return np.nan
        return (d0 - d1).days - (date(d0.year, 1, 1) - date(d1.year, 1, 1)).days

    df["제출지연일"] = [dayshift(x, y) for x, y in zip(df["제출일"], df["제출일_전년"])]

    w = cfg["weights"]
    pts = pd.DataFrame(index=df.index)
    mult = df["예측대비배수"]
    pts["p_배수"] = (w["multiple"] * ((mult - 1) / (cfg["multiple_full"] - 1)).clip(0, 1)).fillna(0)

    std_ok = (df["자산_억"] >= cfg["std_exempt_eok"]) & df["최소표준감사시간"].gt(0) & df["감사시간"].gt(0)
    ratio = df["감사시간"] / df["최소표준감사시간"]
    pts["p_표준시간"] = np.where(
        std_ok, w["std_hours"] * ((ratio - 1) / (cfg["std_full_ratio"] - 1)).clip(0, 1), 0)
    pts["p_표준시간"] = pts["p_표준시간"].fillna(0)

    pts["p_시간증가"] = np.where(df["감사시간증가율%"] >= cfg["hours_yoy_min"], w["hours_yoy"], 0)

    small = df["자산_억"] < cfg["small_asset_eok"]
    big4 = a0.isin(cfg["big4"])
    mid = a0.isin(cfg["mid_tier"])
    df["감사인구분"] = np.where(big4, "Big4", np.where(mid, "준대형", np.where(a0 != "", "기타", "")))
    pts["p_대형감사인"] = np.where(small & big4, w["big4_small"], np.where(small & mid, w["mid_small"], 0))

    pts["p_교체이력"] = np.where(df["올해교체"], w["just_changed"], np.where(df["3년동일"], w["tenure_long"], 0))
    pts["p_정정"] = np.where(df["기재정정"], w["correction"], 0)
    pts["p_지연"] = np.where(df["제출지연일"] >= cfg["late_days"], w["late"], 0)

    def prox(addr):
        if not isinstance(addr, str):
            return 0, ""
        for pat, p in cfg["proximity"]:
            if pat in addr:
                return p, pat
        return 0, ""

    pr = df["주소"].map(prox)
    pts["p_근접"] = [p for p, _ in pr]
    df["근접지역"] = [lab for _, lab in pr]

    for c in pts.columns:
        df[c] = pts[c].astype(float).round(1)
    return df


# ----------------------------------------------------------------------------
# 2단계: 원문 판독
# ----------------------------------------------------------------------------
NAME_STOP = {"이사회", "귀중", "귀하", "변경", "선임", "사임", "취임", "회장", "사장", "부회장", "부사장",
             "외", "및", "등", "공동", "각자", "직무대행", "명의", "서명", "날인", "주식회사"}
# '대표이사' 뒤에 공백이나 콜론이 있어야 한다 ('대표이사의', '대표이사는' 배제).
# 이름: 붙여 쓴 2~4자, 또는 '홍 길 동'처럼 한 칸씩 띄운 형태
CEO_RE = re.compile(r"대\s*표\s*이\s*사(?:\s+(?:사장|회장|부회장))?(?:\s*[:：]\s*|\s+)"
                    r"([가-힣]{2,4}(?![가-힣])|[가-힣](?:\s[가-힣]){1,3}(?![가-힣]))")


def zip_texts(path, rcept_no):
    main, allt = "", []
    with zipfile.ZipFile(path) as z:
        names = [n for n in z.namelist() if n.lower().endswith(".xml")]
        for n in names:
            b = z.read(n)
            try:
                t = b.decode("utf-8")
            except UnicodeDecodeError:
                t = b.decode("cp949", errors="ignore")
            t = html.unescape(re.sub(r"<[^>]+>", " ", t))
            t = re.sub(r"\s+", " ", t)
            allt.append(t)
            if Path(n).stem == rcept_no:
                main = t
    if not main and allt:
        main = max(allt, key=len)
    return main, " ".join(allt)


def extract_ceo(text):
    names = []
    for m in CEO_RE.finditer(text):
        n = re.sub(r"\s", "", m.group(1))
        if n in NAME_STOP or n[0] in "및등외의는께":
            continue
        if any(s in n for s in ("이사회", "귀중", "회사")):
            continue
        names.append(n)
    return Counter(names).most_common(1)[0][0] if names else ""


def read_opinion(text):
    head = text[:20000]   # 감사보고서 본문 범위 (재무제표 주석의 같은 표현 오인 방지)
    if re.search(r"의견\s*거절\s*근거|의견을\s*표명하지\s*않", head):
        op = "의견거절"
    elif re.search(r"부적정\s*의견\s*근거", head):
        op = "부적정"
    elif re.search(r"한정\s*의견\s*근거", head):
        op = "한정"
    elif re.search(r"적정하게\s*표시하고\s*있습니다", head):
        op = "적정"
    else:
        op = "미확인"
    gc = bool(re.search(r"계속\s*기업\s*(?:가정에\s*)?(?:관련|관한)\s*(?:된\s*)?중요한\s*불확실성", head))
    return op, gc


def download_doc(key, rcept_no, session):
    path = CACHE_DIR / f"{rcept_no}.zip"
    if path.exists():
        return path, "cached"
    r = session.get(f"{API}/document.xml", params={"crtfc_key": key, "rcept_no": rcept_no}, timeout=60)
    if r.content[:2] == b"PK":
        tmp = path.with_suffix(".part")
        tmp.write_bytes(r.content)
        tmp.replace(path)
        time.sleep(0.15)
        return path, "ok"
    m = re.search(r"<status>(\d+)</status>|\"status\"\s*:\s*\"(\d+)\"", r.text)
    st = (m.group(1) or m.group(2)) if m else "?"
    if st == "020":
        raise DartLimit()
    return None, st


def run_stage2(df, cfg, n):
    CACHE_DIR.mkdir(exist_ok=True)
    cache = read_table(STAGE2_CACHE).set_index("rcept_no") if STAGE2_CACHE.exists() else pd.DataFrame()
    cache = cache.to_dict("index") if len(cache) else {}

    targets = df[~df["제외"]].head(n)
    need = [r for r in pd.concat([targets["접수번호"], targets["접수번호_전년"]]).dropna().unique()
            if r not in cache]
    print(f"[2단계] 상위 {len(targets)}곳, 새로 판독할 원문 {len(need)}건")

    key, session, limited = None, requests.Session(), False
    for i, rno in enumerate(need, 1):
        path = CACHE_DIR / f"{rno}.zip"
        if not path.exists():
            if limited:
                continue
            key = key or api_key()
            try:
                path, st = download_doc(key, rno, session)
            except DartLimit:
                print("  ! 하루 호출 한도 초과(020). 받은 것까지만 판독하고 저장합니다.")
                limited = True
                continue
            except requests.RequestException as e:
                print(f"  ! {rno} 다운로드 실패: {e}")
                continue
            if path is None:
                cache[rno] = {"opinion": f"원문없음({st})", "going_concern": "", "ceo": ""}
                continue
        try:
            main, _ = zip_texts(path, rno)
            op, gc = read_opinion(main)
            cache[rno] = {"opinion": op, "going_concern": str(gc), "ceo": extract_ceo(main)}
        except (zipfile.BadZipFile, OSError) as e:
            cache[rno] = {"opinion": f"판독실패({type(e).__name__})", "going_concern": "", "ceo": ""}
        if i % 50 == 0:
            print(f"  {i}/{len(need)}")

    c = pd.DataFrame.from_dict(cache, orient="index")
    c.index.name = "rcept_no"
    safe_write(lambda p: c.to_csv(p, encoding="utf-8-sig"), STAGE2_CACHE)
    return c


def apply_stage2(df, cache, cfg):
    get = lambda r, k: cache.at[r, k] if isinstance(r, str) and r in cache.index else ""
    df["감사의견"] = [get(r, "opinion") for r in df["접수번호"]]
    df["계속기업불확실성"] = [str(get(r, "going_concern")) == "True" for r in df["접수번호"]]
    df["대표이사_원문"] = [get(r, "ceo") for r in df["접수번호"]]
    df["대표이사_전년원문"] = [get(r, "ceo") for r in df["접수번호_전년"]]

    def norm(n):
        return re.sub(r"[^가-힣A-Za-z]", "", str(n)) if isinstance(n, str) else ""

    changed = []
    for cur, ceo_now, prev in zip(df["대표이사_원문"], df["대표자"], df["대표이사_전년원문"]):
        now = norm(cur) or norm(ceo_now)      # 올해 원문 → 없으면 기업개황 현재 대표자
        prev = norm(prev)
        if not now or not prev:
            changed.append(np.nan)
        else:
            changed.append(prev not in now and now not in prev)
    df["대표이사변경"] = changed
    df["p_대표이사"] = np.where(df["대표이사변경"] == True, cfg["weights"]["ceo_change"], 0).astype(float)  # noqa: E712
    return df


# ----------------------------------------------------------------------------
# 제외·점수·출력
# ----------------------------------------------------------------------------
def load_exclusions(cfg):
    codes, names = set(), set()
    if EXCLUDE_FILE.exists():
        e = read_table(EXCLUDE_FILE)
        cc, cn = find_col(e, "corp_code", cfg), find_col(e, "corp_name", cfg)
        if cc:
            codes = set(e[cc].dropna().astype(str).str.zfill(8))
        if cn:
            names = set(e[cn].dropna().map(lambda s: re.sub(r"\(주\)|주식회사|\s", "", s)))
        print(f"[제외목록] {EXCLUDE_FILE.name}: 코드 {len(codes)}, 회사명 {len(names)}")
    return codes, names


def finalize(df, cfg):
    codes, names = load_exclusions(cfg)
    reasons = [[] for _ in range(len(df))]

    def flag(mask, why):
        for i in np.flatnonzero(np.asarray(mask, dtype=bool)):
            reasons[i].append(why)

    flag(df["자산_억"] > cfg["max_assets_eok"], f"자산 {cfg['max_assets_eok']}억 초과")
    flag(df["감사인"].map(short_auditor).isin(cfg["our_firm"]), "현 감사인이 우리 법인")
    flag(df["감사인교체횟수"] >= cfg["exclude_frequent_change"], "감사인 잦은 교체")
    flag(df["corp_code"].isin(codes), "제외목록(기장 등)")
    flag(df["회사명"].fillna("").map(lambda s: re.sub(r"\(주\)|주식회사|\s", "", s)).isin(names) & (len(names) > 0),
         "제외목록(기장 등)")
    if "감사의견" in df:
        if cfg["exclude_nonclean_opinion"]:
            flag(df["감사의견"].isin(["한정", "부적정", "의견거절"]), "비적정 의견")
        if cfg["exclude_going_concern"]:
            flag(df["계속기업불확실성"], "계속기업 불확실성")
    df["제외사유"] = [", ".join(dict.fromkeys(r)) for r in reasons]
    df["제외"] = df["제외사유"] != ""

    pcols = [c for c in df.columns if c.startswith("p_")]
    df["신호점수"] = df[pcols].sum(axis=1).round(1)

    labels = {"p_배수": "감사시간 많음", "p_표준시간": "표준시간 초과", "p_시간증가": "시간 증가",
              "p_대형감사인": "규모 대비 대형감사인", "p_교체이력": "감사인 이력", "p_정정": "기재정정",
              "p_지연": "제출 지연", "p_근접": "근접", "p_대표이사": "대표이사 변경"}
    df["신호요약"] = df[pcols].apply(
        lambda r: " / ".join(f"{labels.get(c, c)}({r[c]:+g})" for c in pcols if r[c] != 0), axis=1)
    df = df.sort_values(["제외", "신호점수", "예측대비배수"], ascending=[True, False, False]).reset_index(drop=True)
    df.insert(0, "순위", pd.array(np.where(df["제외"], np.nan, np.arange(1, len(df) + 1)), dtype="Int64"))
    return df


TOP_COLS = ["순위", "회사명", "corp_code", "신호점수", "신호요약", "자산_억", "매출_억", "감사시간", "예측대비배수",
            "감사시간증가율%", "최소표준감사시간", "감사인", "감사인구분", "감사인_전년", "감사인_전전년",
            "감사인교체횟수", "기재정정", "제출일", "제출일_전년", "제출지연일", "근접지역", "주소", "대표자",
            "대표이사_원문", "대표이사_전년원문", "대표이사변경", "감사의견", "계속기업불확실성", "접수번호"]


def write_outputs(df, cfg):
    keep = df[~df["제외"]]
    excl = df[df["제외"]]
    safe_write(lambda p: df.to_csv(p, index=False, encoding="utf-8-sig"), OUT_ALL)
    safe_write(lambda p: excl[["회사명", "corp_code", "제외사유", "자산_억", "감사인", "감사인_전년", "감사인_전전년"]]
               .to_csv(p, index=False, encoding="utf-8-sig"), OUT_EXCL)

    top = keep.head(cfg["top_n"])
    cols = [c for c in TOP_COLS if c in top.columns] + [c for c in top.columns if c.startswith("원본_")]
    top = top[cols].copy()
    for c in ("제출일", "제출일_전년"):
        if c in top:
            top[c] = pd.to_datetime(top[c]).dt.strftime("%Y-%m-%d")

    def write_xlsx(p):
        with pd.ExcelWriter(p, engine="openpyxl") as xw:
            top.to_excel(xw, sheet_name="상위", index=False)
            pd.DataFrame({"주의": [
                "내부 타깃팅·통화 준비용. 서면(커버레터)에 현 감사인 비교, 감사시간 과다, 보수 절감액을 쓰지 말 것.",
                "감사시간·배수는 외부 추정치. 대표이사 변경은 원문 인식 결과라 오인식 가능.",
            ]}).to_excel(xw, sheet_name="주의", index=False)
            ws = xw.sheets["상위"]
            ws.freeze_panes = "C2"
            for col in ws.columns:
                width = max(len(str(c.value)) if c.value is not None else 0 for c in col[:50])
                ws.column_dimensions[col[0].column_letter].width = min(max(8, width * 1.6), 50)

    out = safe_write(write_xlsx, OUT_TOP)

    print(f"\n[결과] 후보 {len(keep):,}사, 제외 {len(excl):,}사 → {Path(out).name} 상위 {len(top)}곳")
    if len(excl):
        rc = Counter(r for s in excl["제외사유"] for r in s.split(", "))
        print("  제외 사유: " + ", ".join(f"{k} {v}" for k, v in rc.most_common()))
    print("  감사인 구분(후보): " + ", ".join(f"{k or '미확인'} {v}" for k, v in keep["감사인구분"].value_counts().items()))
    unmatched = keep["감사인"].isna().sum()
    if unmatched:
        print(f"  ! 올해 목록에서 감사인을 못 찾은 회사 {unmatched}곳 (corp_code·결산기 확인)")
    print(top[["순위", "회사명", "신호점수", "신호요약"]].head(15).to_string(index=False))


def main():
    ap = argparse.ArgumentParser(description="감사인 교체 가능성 신호 점수")
    ap.add_argument("--input", help="기본: mailing_list.xlsx, 없으면 audit_hours_rank.csv")
    ap.add_argument("--period", help="대상 결산기 (기본 2025.12)")
    ap.add_argument("--top", type=int, help="상위 N곳 (기본 150)")
    ap.add_argument("--fetch-history", action="store_true", help="2024·2025 접수 목록 수집 후 종료")
    ap.add_argument("--refresh", action="store_true", help="--fetch-history 시 이미 받은 달도 다시 받기")
    ap.add_argument("--stage2", action="store_true", help="상위 후보 원문으로 대표이사 변경·감사의견 확인")
    ap.add_argument("--stage2-n", type=int, default=400)
    ap.add_argument("--write-config", action="store_true", help="signals_config.json 기본값 생성 후 종료")
    a = ap.parse_args()

    if a.write_config:
        if CONFIG_FILE.exists():
            sys.exit(f"{CONFIG_FILE.name} 이미 있음. 지우고 다시 실행하거나 직접 수정하세요.")
        CONFIG_FILE.write_text(json.dumps(DEFAULT_CONFIG, ensure_ascii=False, indent=2), encoding="utf-8")
        print(f"[생성] {CONFIG_FILE.name}")
        return

    cfg = load_config()
    if a.period:
        cfg["period"] = a.period
    if a.top:
        cfg["top_n"] = a.top

    if a.fetch_history:
        fetch_history(cfg, refresh=a.refresh)
        return

    filings = load_filings()
    df = load_base(cfg, filings, a.input)
    df = add_stage1(df, filings, cfg)
    df = finalize(df, cfg)

    if a.stage2:
        cache = run_stage2(df, cfg, a.stage2_n)
        df = apply_stage2(df.drop(columns=["순위", "제외", "제외사유", "신호점수", "신호요약"]), cache, cfg)
        df = finalize(df, cfg)
    elif STAGE2_CACHE.exists():
        # 이전 2단계 판독 결과가 있으면 재사용 (API 호출 없음)
        cache = read_table(STAGE2_CACHE).set_index("rcept_no")
        df = apply_stage2(df.drop(columns=["순위", "제외", "제외사유", "신호점수", "신호요약"]), cache, cfg)
        df = finalize(df, cfg)
        print("[2단계] 기존 판독 캐시 반영")

    write_outputs(df, cfg)


if __name__ == "__main__":
    main()
