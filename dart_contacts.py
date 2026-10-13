"""
audit_hours_rank.csv 의 회사들에 OpenDART 기업개황(대표자, 주소, 전화, 팩스, 홈페이지)을 붙여
우편 발송 + 전화 후속용 명단(mailing_list.xlsx)을 만든다.

사용:
  python -m pip install openpyxl        (최초 1회)
  python dart_contacts.py               (전체)
  python dart_contacts.py --top 100     (예측대비배수 상위 100개만)
"""
import os
import sys
import time
import argparse
import requests
import pandas as pd

API = "https://opendart.fss.or.kr/api"
KEY = os.environ.get("DART_API_KEY")
RANK_CSV = "audit_hours_rank.csv"
FILINGS_CSV = "filings.csv"
CACHE_CSV = "contacts_cache.csv"
OUT_XLSX = "mailing_list.xlsx"


def fetch_company(corp_code):
    r = requests.get(f"{API}/company.json",
                     params=dict(crtfc_key=KEY, corp_code=corp_code), timeout=30)
    d = r.json()
    if d.get("status") == "020":
        raise RuntimeError("일일 호출 한도 초과. 내일 다시 실행하면 이어서 진행됩니다.")
    if d.get("status") != "000":
        return dict(corp_code=corp_code, 조회오류=d.get("message", str(d)))
    return dict(corp_code=corp_code,
                대표자=d.get("ceo_nm", ""),
                주소=d.get("adres", ""),
                전화=d.get("phn_no", ""),
                팩스=d.get("fax_no", ""),
                홈페이지=d.get("hm_url", ""),
                업종코드=d.get("induty_code", ""),
                설립일=d.get("est_dt", ""),
                결산월=d.get("acc_mt", ""))


# ---------------------------------------------------------------- 표준감사시간 (한공회 제1호, 2025.12.9 개정)
# 열 순서: 제조업, 서비스업, 건설업, 금융업, 도소매업, 기타 / 행: 기업규모(억원)
IND = ["제조업", "서비스업", "건설업", "금융업", "도소매업", "기타"]
STD_TABLE = {
    8: [(500, [690, 480, 600, 630, 620, 190]), (1000, [750, 530, 660, 690, 680, 210]),
        (2000, [830, 590, 730, 760, 750, 230]), (5000, [940, 660, 820, 860, 850, 260]),
        (10000, [1030, 730, 910, 950, 940, 290]), (20000, [1140, 800, 1000, 1040, 1030, 320])],
    9: [(200, [450, 350, 380, 390, 400, 380]), (500, [550, 440, 470, 480, 500, 470]),
        (600, [580, 460, 490, 500, 520, 490]), (700, [600, 470, 510, 510, 540, 510]),
        (800, [620, 490, 530, 530, 560, 530]), (900, [630, 500, 540, 550, 570, 540]),
        (1000, [650, 510, 560, 560, 590, 560]), (2000, [760, 600, 650, 660, 690, 650])],
    10: [(100, [420, 340, 340, 350, 390, 310]), (200, [480, 390, 390, 390, 440, 350]),
         (300, [510, 420, 420, 420, 480, 380]), (400, [540, 440, 440, 450, 510, 400]),
         (500, [570, 460, 460, 470, 530, 420]), (1000, [650, 520, 520, 530, 600, 480])],
}
APPLY_RATE_2026 = {8: 0.90, 9: 0.80, 10: 0.70}   # 부칙 2025.12.9 적용률 (2027년분은 2026 하반기 재심의)


def std_industry(induty_code):
    """KSIC 중분류(앞 2자리) → 표준감사시간 업종 (별표 12)"""
    try:
        d = int(str(induty_code).strip()[:2])
    except (ValueError, TypeError):
        return None
    if 10 <= d <= 34:
        return "제조업"
    if d in (41, 42):
        return "건설업"
    if 45 <= d <= 47:
        return "도소매업"
    if 64 <= d <= 66:
        return "금융업"
    if 1 <= d <= 8 or 35 <= d <= 39:
        return "기타"
    if 49 <= d <= 99:
        return "서비스업"
    return None


def std_group(assets):
    if assets >= 1000:
        return 8
    if assets >= 500:
        return 9
    if assets >= 200:
        return 10
    return 11


def table_hours(group, industry, size):
    """기업규모(억원)로 표 조회. 표 사이는 직선 보간, 표 범위 밖은 끝 값 사용."""
    rows = STD_TABLE[group]
    k = IND.index(industry)
    xs = [r[0] for r in rows]
    ys = [r[1][k] for r in rows]
    note = ""
    if size <= xs[0]:
        return ys[0], f"기업규모 {size:,.0f}억 < 표 하한 {xs[0]:,}억 → 하한값 사용"
    if size >= xs[-1]:
        return ys[-1], f"표 상한 {xs[-1]:,}억 값 사용"
    for (x0, y0), (x1, y1) in zip(zip(xs, ys), zip(xs[1:], ys[1:])):
        if x0 <= size <= x1:
            return y0 + (y1 - y0) * (size - x0) / (x1 - x0), note


def lower_bound_std(assets, rev, skill_coef):
    """업종을 모를 때 쓰는 하한: 해당 그룹·규모에서 가장 낮은 업종의 최소표준시간"""
    if pd.isna(assets):
        return None
    g = std_group(assets)
    if g == 11:
        return None
    size = (assets + (0 if pd.isna(rev) else rev)) / 2
    return min(table_hours(g, ind, size)[0] for ind in IND) * APPLY_RATE_2026[g] * skill_coef


def add_std_hours(out, skill_coef, rate):
    res = []
    for _, r in out.iterrows():
        d = dict(표준업종=None, 표준감사그룹=None, 기업규모_억=None, 표준시간표값=None,
                 적용률=None, 최소표준시간=None, 초도감사반영시간=None, 최소표준시간_실제대비=None,
                 표준시간비고="")
        assets, rev = r.get("자산(억)"), r.get("매출(억)")
        ind = std_industry(r.get("업종코드"))
        d["표준업종"] = ind
        if pd.isna(assets):
            d["표준시간비고"] = "자산 없음"
            res.append(d); continue
        rev = 0 if pd.isna(rev) else rev
        g = std_group(assets)
        d["표준감사그룹"] = g
        size = (assets + rev) / 2                        # 기업규모 = (자산+매출)/2
        d["기업규모_억"] = round(size, 1)
        if g == 11:
            d["표준시간비고"] = "그룹11(자산 200억 미만): 표준감사시간 적용 유예"
            res.append(d); continue
        if ind is None:
            d["표준시간비고"] = "업종코드 판별 불가"
            res.append(d); continue
        base, note = table_hours(g, ind, size)
        rate_g = APPLY_RATE_2026[g]
        mn = base * rate_g * skill_coef
        d.update(표준시간표값=round(base), 적용률=rate_g,
                 최소표준시간=round(mn), 초도감사반영시간=round(mn * 1.10),
                 표준시간비고=note)
        if not pd.isna(r.get("감사시간")) and r.get("감사시간"):
            d["최소표준시간_실제대비"] = round(mn / r["감사시간"], 2)
        if rate:
            d["예상보수_최소(원)"] = round(mn * rate, -4)
            d["예상보수_초도(원)"] = round(mn * 1.10 * rate, -4)
        res.append(d)
    return pd.concat([out.reset_index(drop=True), pd.DataFrame(res)], axis=1)


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--top", type=int, default=None, help="예측대비배수 상위 N개만")
    ap.add_argument("--skill-coef", type=float, default=1.0,
                    help="숙련도조정계수(기준숙련도/감사팀숙련도). 기본 1.0")
    ap.add_argument("--over-std", action="store_true",
                    help="작년 실제 감사시간이 최저 표준감사시간보다 많은 회사만 (그룹11 제외)")
    ap.add_argument("--rate", type=float, default=None, help="시간당 보수(원). 넣으면 예상보수 계산")
    args = ap.parse_args()

    if not KEY:
        sys.exit("DART_API_KEY 환경변수가 없습니다.")
    for f in (RANK_CSV, FILINGS_CSV):
        if not os.path.exists(f):
            sys.exit(f"{f} 파일이 없습니다. dart_audit_hours.py를 먼저 실행하세요.")

    rank = pd.read_csv(RANK_CSV, dtype={"rcept_no": str})
    filings = pd.read_csv(FILINGS_CSV, dtype=str)[["rcept_no", "corp_code"]]
    rank = rank.merge(filings, on="rcept_no", how="left")
    if args.over_std:
        # 1차: 업종 무관 하한보다 시간이 많은 회사만 남겨 기업개황 호출 수를 줄임
        lb = [lower_bound_std(a, r, args.skill_coef)
              for a, r in zip(rank["assets_eok"], rank["revenue_eok"])]
        before = len(rank)
        rank = rank[[l is not None and h > l for l, h in zip(lb, rank["hours"])]]
        print(f"1차 선별(업종 무관 하한 초과): {before}개사 → {len(rank)}개사")
    if args.top:
        rank = rank.head(args.top)
    rank.insert(0, "예측배수순위", range(1, len(rank) + 1))

    cache = pd.read_csv(CACHE_CSV, dtype=str) if os.path.exists(CACHE_CSV) else pd.DataFrame(columns=["corp_code"])
    done = set(cache["corp_code"])
    todo = [c for c in rank["corp_code"].dropna().unique() if c not in done]
    print(f"대상 {len(rank)}개사, 새로 조회 {len(todo)}건 (기존 캐시 {len(done)}건)")

    new_rows = []
    try:
        for i, code in enumerate(todo, 1):
            new_rows.append(fetch_company(code))
            if i % 20 == 0:
                print(f"  {i}/{len(todo)} 조회")
            time.sleep(0.3)
    finally:
        if new_rows:
            cache = pd.concat([cache, pd.DataFrame(new_rows)], ignore_index=True)
            cache.to_csv(CACHE_CSV, index=False, encoding="utf-8-sig")

    out = rank.merge(cache, on="corp_code", how="left")
    out = out.rename(columns={"corp_name": "회사명", "auditor": "현감사인(추정)",
                              "hours": "감사시간", "hours_prev": "전기감사시간",
                              "assets_eok": "자산(억)", "revenue_eok": "매출(억)",
                              "rcept_no": "접수번호"})
    cols = ["예측배수순위", "회사명", "대표자", "주소", "전화", "팩스", "홈페이지",
            "업종코드", "결산월", "현감사인(추정)", "감사시간", "전기감사시간",
            "자산(억)", "매출(억)", "예측대비배수", "접수번호"]
    if "조회오류" in out:
        cols.append("조회오류")
    out = out[[c for c in cols if c in out.columns]]
    out = add_std_hours(out, args.skill_coef, args.rate)
    if args.over_std:
        before = len(out)
        out = out[pd.to_numeric(out["감사시간"], errors="coerce")
                  > pd.to_numeric(out["최소표준시간"], errors="coerce")]
        print(f"2차 선별(업종 반영 최저 표준시간 초과): {before}개사 → {len(out)}개사")
    # 작년 실제 감사시간 - 최저 표준감사시간 차이가 큰 순서로 정렬
    out["차이시간(실제-최소표준)"] = (pd.to_numeric(out["감사시간"], errors="coerce")
                                   - pd.to_numeric(out["최소표준시간"], errors="coerce")).round(0)
    out = out.sort_values("차이시간(실제-최소표준)", ascending=False, na_position="last").reset_index(drop=True)
    out.insert(0, "차이순위", [i + 1 if pd.notna(v) else None
                               for i, v in enumerate(out["차이시간(실제-최소표준)"])])
    lead = ["차이순위", "회사명", "감사시간", "최소표준시간", "차이시간(실제-최소표준)",
            "최소표준시간_실제대비", "초도감사반영시간"]
    out = out[lead + [c for c in out.columns if c not in lead]]
    # 발송·전화 후속 관리용 빈 칸
    for c in ["발송일", "통화일", "통화상대", "통화결과", "다음연락일", "비고"]:
        out[c] = ""

    out.to_excel(OUT_XLSX, index=False)
    no_addr = out["주소"].isna().sum() + (out["주소"] == "").sum() if "주소" in out else len(out)
    print(f"\n완료: {len(out)}개사 → {OUT_XLSX} (주소 없음 {no_addr}건)")


if __name__ == "__main__":
    main()
