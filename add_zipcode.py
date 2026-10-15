# -*- coding: utf-8 -*-
"""
mailing_list.xlsx 주소 → 우편번호(5자리)  (행정안전부 도로명주소 검색 API)

결과는 zipcode_cache.csv(주소별)에 쌓고, build_mail.py가 주소 라벨에 "(우) 06134"로 넣는다.
명단을 다시 뽑아도 캐시에 있는 주소는 다시 조회하지 않는다.

준비 (최초 1회)
  1. business.juso.go.kr → API 신청 → 검색API(도로명주소) → 승인키 발급
  2. PowerShell:  setx JUSO_API_KEY "발급받은키"   → PowerShell 창을 새로 열 것
사용
  python add_zipcode.py            # 캐시에 없는 주소만 조회
  python add_zipcode.py --retry    # 못 찾은 주소도 다시 조회
"""
import argparse
import os
import re
import sys
import time
from pathlib import Path

import pandas as pd
import requests

BASE = Path(__file__).resolve().parent
MAIL = BASE / "mailing_list.xlsx"
CACHE = BASE / "zipcode_cache.csv"
API = "https://business.juso.go.kr/addrlink/addrLinkApi.do"

ROAD = re.compile(r"^(.*?\S(?:로\d+번?길|로|길)\s*(?:지하\s*)?\d+(?:-\d+)?)(?=\D|$)")
JIBUN = re.compile(r"^(.*?\S(?:동|가|리)\s+(?:산\s*)?\d+(?:-\d+)?)(?=\D|$)")


def keyword(addr):
    """검색어: 괄호·층·호수를 버리고 '도로명 + 건물번호'(없으면 지번)까지만"""
    a = re.sub(r"\(.*?\)", " ", str(addr))
    a = re.sub(r"\s+", " ", a).strip()
    a = re.sub(r"(로|길)\s+(\d+)\s*(번?길)", r"\1\2\3", a)   # '양재대로 11길' → '양재대로11길'
    a = re.sub(r"(\d+)\s+(번?길)", r"\1\2", a)                 # '언주로85 길' → '언주로85길'
    for pat in (ROAD, JIBUN):
        m = pat.match(a)
        if m:
            return m.group(1).strip()
    return a


def variants(kw):
    """못 찾았을 때 바꿔 볼 검색어: '로N길'↔'로N번길', 읍·면 빼기"""
    out = [kw]
    if re.search(r"로\d+길", kw):
        out.append(re.sub(r"(로\d+)길", r"\1번길", kw))
    elif re.search(r"로\d+번길", kw):
        out.append(re.sub(r"(로\d+)번길", r"\1길", kw))
    out += [re.sub(r"\s\S+[읍면]\s", " ", v) for v in list(out) if re.search(r"\s\S+[읍면]\s", v)]
    return out


def lookup(key, kw):
    r = requests.get(API, params={"confmKey": key, "currentPage": 1, "countPerPage": 5,
                                  "keyword": kw, "resultType": "json"}, timeout=15)
    r.raise_for_status()
    res = r.json()["results"]
    code = res["common"]["errorCode"]
    if code != "0":
        raise RuntimeError(f"{code} {res['common']['errorMessage']}")
    rows = res.get("juso") or []
    zips = {j["zipNo"] for j in rows}
    if not rows:
        return "", "", "없음"
    note = "" if len(zips) == 1 else f"후보 {len(rows)}건·우편번호 {len(zips)}개, 첫 번째 사용"
    return rows[0]["zipNo"], rows[0]["roadAddr"], note


def main():
    ap = argparse.ArgumentParser(description="주소 라벨용 우편번호 조회")
    ap.add_argument("--retry", action="store_true", help="못 찾은 주소도 다시 조회")
    a = ap.parse_args()
    key = os.environ.get("JUSO_API_KEY")
    if not key:
        sys.exit("[오류] 환경변수 JUSO_API_KEY가 없습니다. 파일 맨 위 '준비'를 보세요.")
    m = pd.read_excel(MAIL, sheet_name="발송명단", dtype=str)
    addrs = [x for x in m["주소"].dropna().unique()]

    cache = pd.read_csv(CACHE, dtype=str, encoding="utf-8-sig").fillna("") if CACHE.exists() else \
        pd.DataFrame(columns=["주소", "검색어", "우편번호", "도로명주소", "비고"])
    # 검색어 규칙이 바뀌어 저장된 검색어와 다르면 다시 조회 (손으로 넣은 줄은 비고가 '수동'으로 시작)
    stale = (cache["검색어"] != cache["주소"].map(keyword)) & ~cache["비고"].str.startswith("수동")
    cache = cache[~stale]
    done = set(cache.loc[cache["우편번호"] != "", "주소"]) if a.retry else set(cache["주소"])
    todo = [x for x in addrs if x not in done]
    print(f"주소 {len(addrs)}개 중 조회 {len(todo)}개 (캐시 {len(addrs) - len(todo)}개)")

    new = []
    for i, addr in enumerate(todo, 1):
        kw = keyword(addr)
        try:
            for v in variants(kw):
                z, road, note = lookup(key, v)
                if z:
                    if v != kw:
                        note = (note + " " if note else "") + f"검색어 바꿔 찾음: {v}"
                    break
        except Exception as e:                       # 키 오류 등은 바로 멈춘다
            if new:
                break
            sys.exit(f"[오류] {e}")
        new.append({"주소": addr, "검색어": kw, "우편번호": z, "도로명주소": road, "비고": note})
        if i % 20 == 0:
            print(f"  {i}/{len(todo)}")
        time.sleep(0.1)

    if new:
        cache = cache[~cache["주소"].isin([n["주소"] for n in new])]
        cache = pd.concat([cache, pd.DataFrame(new)], ignore_index=True)
        cache.to_csv(CACHE, index=False, encoding="utf-8-sig")
    got = cache[cache["주소"].isin(addrs)]
    miss = got[got["우편번호"] == ""]
    multi = got[got["비고"].str.startswith("후보")]
    print(f"[완료] {CACHE.name}: 우편번호 {len(got) - len(miss)}/{len(addrs)}  "
          f"(못 찾음 {len(miss)}, 후보 여러 개 {len(multi)} → 비고 확인)")
    for _, r in miss.iterrows():
        print(f"  못 찾음: {r['주소']}  [검색어: {r['검색어']}]")


if __name__ == "__main__":
    main()
