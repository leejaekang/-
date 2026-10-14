import os, sys, time, argparse, re
import pandas as pd
import requests

API = "https://opendart.fss.or.kr/api"
KEY = os.environ.get("DART_API_KEY")
CACHE_DIR = "dart_cache"
FILINGS_CSV = "filings.csv"


def decode(raw):
    for enc in ("utf-8", "cp949"):
        try:
            return raw.decode(enc)
        except UnicodeDecodeError:
            pass
    return raw.decode("utf-8", errors="ignore")


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--period", default="2025.12")
    args = ap.parse_args()
    if not KEY:
        sys.exit("DART_API_KEY 환경변수가 없습니다.")
    if not os.path.exists(FILINGS_CSV):
        sys.exit(f"{FILINGS_CSV}가 이 폴더에 없습니다. 현재 폴더: {os.getcwd()}")
    os.makedirs(CACHE_DIR, exist_ok=True)

    f = pd.read_csv(FILINGS_CSV, dtype=str)
    n0 = len(f)
    f = f[f["report_nm"].str.contains(re.escape(args.period), na=False)]
    n1 = len(f)
    f = f[~f["report_nm"].str.contains("연결", na=False)]
    n2 = len(f)
    if "stock_code" in f.columns:
        f = f[f["stock_code"].fillna("").str.strip() == ""]
    n3 = len(f)
    f = f.sort_values("rcept_no").drop_duplicates("corp_code", keep="last")
    print(f"목록 {n0}건 → 결산기 {n1}건 → 연결 제외 {n2}건 → 상장사 제외 {n3}건 → 처리 대상 {len(f)}건")

    have = set(x[:-4] for x in os.listdir(CACHE_DIR)          # 0바이트 = 끊긴 다운로드, 다시 받는다
               if x.endswith(".zip") and os.path.getsize(os.path.join(CACHE_DIR, x)) > 0)
    todo = [r for r in f["rcept_no"] if r not in have]
    done_n = len(f) - len(todo)
    print(f"이미 받음 {done_n}건 / 남은 건수 {len(todo)}건")
    if not todo:
        print("모두 받았습니다. 이제 dart_audit_hours.py를 실행하면 됩니다.")
        return

    ok = err = 0
    for i, rcept_no in enumerate(todo, 1):
        path = os.path.join(CACHE_DIR, f"{rcept_no}.zip")
        try:
            r = requests.get(f"{API}/document.xml",
                             params=dict(crtfc_key=KEY, rcept_no=rcept_no), timeout=60)
        except Exception as e:
            err += 1
            print(f"  [오류] {rcept_no}: {str(e)[:80]}")
            time.sleep(2)
            continue
        if r.content[:2] != b"PK":
            msg = decode(r.content[:500])
            if "020" in msg:
                print(f"\n[중단] 오늘 일일 호출 한도에 도달했습니다. 이번 실행에서 {ok}건 받음.")
                print(f"남은 건수 {len(todo) - i + 1}건. 내일 같은 명령으로 이어서 받으세요.")
                return
            err += 1
            print(f"  [건너뜀] {rcept_no}: {msg[:80]}")
        else:
            tmp = path + ".part"
            with open(tmp, "wb") as fh:
                fh.write(r.content)
            os.replace(tmp, path)
            ok += 1
        if i % 50 == 0:
            print(f"  {i}/{len(todo)} 다운로드 (전체 {done_n + ok}/{len(f)})")
        time.sleep(0.3)

    print(f"\n완료: 이번 실행 {ok}건 받음, 오류·건너뜀 {err}건.")
    print("이제 dart_audit_hours.py를 실행하면 순위 파일이 만들어집니다.")


if __name__ == "__main__":
    main()