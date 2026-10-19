# -*- coding: utf-8 -*-
"""
mailing_list.xlsx(발송명단) → 커버레터와 주소 라벨 Word 파일

출력
  - 커버레터_전체.docx   회사별 1장씩 (수신: 회사명·대표이사)
  - 주소라벨.docx        A4 라벨지 (기본: 2열×7행 = 14칸, 99.1×38.1mm. 처음엔 일반 용지에 시험 인쇄할 것)

사용 (PowerShell, 실행 전 mailing_list.xlsx와 결과 docx는 닫아 둘 것)
  python build_mail.py                          # 날짜는 오늘
  python build_mail.py --date "2026년 11월 16일"

영업 원칙: 서면에 현 감사인 비교, 감사시간 과다, 보수 절감액을 쓰지 않는다.
아래 문안은 '선임 전 비교 견적 기회 + 맞춤 감사계획·견적 제시'에서 멈춘다. 고칠 때도 이 선을 지킬 것.
필요 패키지: python-docx qrcode[pil] (python -m pip install python-docx "qrcode[pil]")
커버레터_전체.docx는 회사마다 [커버레터, 회신서] 2쪽. 단면 인쇄하면 회사별로 순서대로 나온다
"""
import argparse
import io
import sys
import time
from datetime import date
from pathlib import Path
from urllib.parse import quote

import pandas as pd
from docx import Document
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_CELL_VERTICAL_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_BREAK
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
from docx.shared import Mm, Pt

BASE = Path(__file__).resolve().parent
MAIL = BASE / "mailing_list.xlsx"
OUT_LETTER = BASE / "커버레터_전체.docx"
OUT_LABEL = BASE / "주소라벨.docx"
ZIP_CACHE = BASE / "zipcode_cache.csv"     # python add_zipcode.py 로 생성

# ----------------------------------------------------------------------------
# 문안 (여기만 고치면 된다)
# ----------------------------------------------------------------------------
FONT = "맑은 고딕"
TITLE = "2027 사업연도 외부감사인 선임 관련 안내"
BODY = [
    "안녕하십니까. 서울 강남구 삼성동에 있는 지인회계법인입니다.",
    "12월 결산 회사가 2027 사업연도 외부감사인을 새로 선임하시는 경우, 선임 기한은 2027년 2월 15일입니다"
    "(사업연도 개시일부터 45일 이내). 감사인을 정하시기 전에 비교 견적을 받아 보실 수 있도록 안내드립니다.",
    "저희는 귀사의 업종과 규모, 회계 처리 현황을 먼저 살펴본 뒤 귀사에 맞춘 감사계획과 견적을 제시해 드리겠습니다. "
    "검토를 원하시면 아래 연락처로 편하게 연락 주십시오. 일정을 맞춰 방문 설명드리겠습니다.",
    "뒷장의 회신서에 표시해 보내 주시거나 회신서의 QR코드로 신청하셔도 됩니다. "
    "법인 소개와 외부감사 준비 일정표도 함께 넣었습니다.",
]
FIRM_POINTS = [
    "2024년 4월 설립, 상근 공인회계사 18명 (경력 15년 이상 13명)",
    "회계감사 12명 · 세무자문 3명 · 경영자문 3명",
    "비상장 회사 외부감사에 집중합니다",
]
CLOSING = [
    "안내문이 잘 전달되었는지 확인하기 위해 담당자가 한 차례 전화드릴 수 있습니다. "
    "안내를 원하지 않으시면 뒷장 회신서의 QR코드로 '앞으로 안내를 원하지 않음'을 선택해 주십시오. "
    "이후 연락드리지 않겠습니다.",
    "감사합니다.",
]
CONTACT = [
    "지인회계법인",
    "서울특별시 강남구 테헤란로103길 6, 5층 (삼성동)",
    "담당 이재강 공인회계사 (직통) 010-6624-2061",
    "이메일 ljkdlworkd@gmail.com",
]
SIGNATURE = "지인회계법인  대표이사  송 재 철"

# 회신서 (회사별 1장, 커버레터 바로 뒤 쪽). 회신번호 = DART corp_code
REPLY_FAX = ""          # 법인 팩스번호. 비워 두면 팩스 안내 줄을 넣지 않는다
REPLY_MOBILE = "010-6624-2061"
REPLY_EMAIL = "ljkdlworkd@gmail.com"
# QR 연결 주소. 구글 설문 '미리 채워진 링크'를 만들고 회신번호 값은 {code}, 회사명 값은 {company}로 바꿔 넣는다
#   예: "https://docs.google.com/forms/d/e/.../viewform?usp=pp_url&entry.111={code}&entry.222={company}"
# 비워 두면 회신번호가 제목에 들어간 이메일 작성 화면(mailto)으로 연결한다
REPLY_URL = ("https://docs.google.com/forms/d/e/1FAIpQLSdzERh9J_VlZoPHk4JWDY732KO1ZJ5iONG22ygSHStNa9rIgQ/viewform"
             "?usp=pp_url&entry.1830759413={code}&entry.1697068679={company}")
REPLY_ITEMS = [
    "2027 사업연도 외부감사 견적·감사계획 요청",
    "세액공제 누락·세무 현안 무료 진단 신청 (1회)",
    "담당 회계사 방문 상담 희망",
    "자료만 이메일로 받아 보고 싶음",
    "앞으로 안내를 원하지 않음",
]

# 봉투(주소 라벨) 맨 위 한 줄. 개봉률을 위한 기한 안내. 비우면 넣지 않는다
LABEL_NOTICE = "2027 외부감사인 선임기한(2027. 2. 15.) 안내"

# 라벨지 규격 (mm). 기본값은 2열×7행 14칸(99.1×38.1). 쓰는 라벨지 포장에 적힌 규격으로 고칠 것
LABEL = {"cols": 2, "rows": 7, "w": 99.1, "h": 38.1, "gap": 2.5, "top": 15.1, "left": 4.65}


# ----------------------------------------------------------------------------
def set_font(doc):
    st = doc.styles["Normal"]
    st.font.name = FONT
    st.font.size = Pt(10.5)
    st.element.rPr.rFonts.set(qn("w:eastAsia"), FONT)      # 한글 글꼴
    st.paragraph_format.space_after = Pt(0)
    st.paragraph_format.line_spacing = 1.25


def para(doc_or_cell, text="", size=None, bold=False, align=None, before=0, after=0):
    p = doc_or_cell.add_paragraph()
    if text:
        r = p.add_run(text)
        r.bold = bold
        if size:
            r.font.size = Pt(size)
    if align is not None:
        p.alignment = align
    p.paragraph_format.space_before = Pt(before)
    p.paragraph_format.space_after = Pt(after)
    return p


def bottom_border(p):
    pPr = p._p.get_or_add_pPr()
    bdr = OxmlElement("w:pBdr")
    b = OxmlElement("w:bottom")
    for k, v in (("w:val", "single"), ("w:sz", "6"), ("w:space", "4"), ("w:color", "808080")):
        b.set(qn(k), v)
    bdr.append(b)
    pPr.append(bdr)


def recipient(row):
    ceo = str(row["대표자"]).strip() if pd.notna(row["대표자"]) else ""
    return f"대표이사 {ceo} 귀하" if ceo else "대표이사 귀하"


def build_letters(m, when):
    doc = Document()
    set_font(doc)
    sec = doc.sections[0]
    sec.page_width, sec.page_height = Mm(210), Mm(297)
    sec.top_margin, sec.bottom_margin = Mm(22), Mm(18)
    sec.left_margin = sec.right_margin = Mm(25)

    for i, (_, row) in enumerate(m.iterrows()):
        first = para(doc, when, align=WD_ALIGN_PARAGRAPH.RIGHT)
        if i:
            first.paragraph_format.page_break_before = True
        para(doc, str(row["회사명"]), size=12, bold=True, before=12)
        para(doc, recipient(row), size=11, after=14)
        t = para(doc, TITLE, size=14, bold=True, align=WD_ALIGN_PARAGRAPH.CENTER, after=6)
        bottom_border(t)
        for text in BODY:
            para(doc, text, before=10)
        para(doc, "연락처", bold=True, before=14, after=2)
        for text in CONTACT:
            para(doc, f"  {text}")
        for text in CLOSING:
            para(doc, text, before=12)
        para(doc, SIGNATURE, size=11, bold=True, align=WD_ALIGN_PARAGRAPH.RIGHT, before=16)
        build_reply(doc, row)
    return doc


def reply_link(code, company):
    if REPLY_URL:
        return REPLY_URL.replace("{code}", quote(code)).replace("{company}", quote(company))
    subject = f"[회신 {code}] {company} 견적·진단 문의"
    return f"mailto:{REPLY_EMAIL}?subject={quote(subject)}"


def qr_png(text):
    import qrcode
    img = qrcode.make(text, box_size=10, border=1, error_correction=qrcode.constants.ERROR_CORRECT_M)
    buf = io.BytesIO()
    img.save(buf, format="PNG")
    buf.seek(0)
    return buf


def cell_border(cell, sz="6", color="808080"):
    tcPr = cell._tc.get_or_add_tcPr()
    b = OxmlElement("w:tcBorders")
    for side in ("top", "left", "bottom", "right"):
        e = OxmlElement(f"w:{side}")
        for k, v in (("w:val", "single"), ("w:sz", sz), ("w:color", color)):
            e.set(qn(k), v)
        b.append(e)
    tcPr.append(b)


def build_reply(doc, row):
    """회사별 회신서 1쪽: 회사명·회신번호 미리 인쇄, 체크 항목, 담당자 칸, QR"""
    company = str(row["회사명"])
    code = str(row["corp_code"]).strip() if pd.notna(row.get("corp_code")) else ""
    t = para(doc, "회 신 서", size=18, bold=True, align=WD_ALIGN_PARAGRAPH.CENTER)
    t.paragraph_format.page_break_before = True
    sub = "해당 항목에 표시하시어 이 종이를 사진 찍어 문자로 보내 주시거나"
    sub += ", 팩스로 보내 주십시오." if REPLY_FAX else " 이메일로 보내 주십시오."
    para(doc, sub, size=9.5, align=WD_ALIGN_PARAGRAPH.CENTER, before=4)
    para(doc, "오른쪽 QR코드를 휴대폰 카메라로 비추면 바로 신청하실 수 있습니다.", size=9.5,
         align=WD_ALIGN_PARAGRAPH.CENTER, after=10)

    # 회사명·회신번호 + QR
    top = doc.add_table(rows=1, cols=2)
    top.autofit = False
    left, right = top.rows[0].cells
    left.width, right.width = Mm(118), Mm(42)
    left.vertical_alignment = right.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
    left.paragraphs[0].text = ""
    para(left, "회사명", size=9).paragraph_format.space_before = Pt(0)
    para(left, company, size=15, bold=True, after=6)
    para(left, "회신번호", size=9)
    para(left, code or "-", size=13, bold=True)
    rp = right.paragraphs[0]
    rp.alignment = WD_ALIGN_PARAGRAPH.CENTER
    rp.add_run().add_picture(qr_png(reply_link(code, company)), width=Mm(34))
    para(right, "QR로 신청", size=8.5, align=WD_ALIGN_PARAGRAPH.CENTER)
    cell_border(left)
    cell_border(right)

    para(doc, "요청 사항 (해당하는 곳에 ✓ 표시, 여러 개 가능)", size=10.5, bold=True, before=10, after=2)
    for item in REPLY_ITEMS:
        para(doc, f"☐  {item}", size=11, before=2)

    # '담당자'만 쓰면 지인회계법인 담당자로 읽혀 우리 연락처를 적어 보낸 회신이 있었다 (10.19) → '귀사'를 명시
    para(doc, "회신하시는 분 (귀사 담당자) 정보", size=10.5, bold=True, before=10, after=3)
    fields = [("귀사 담당자\n성함 · 직위", ""), ("귀사 담당자 전화", ""), ("이메일", ""),
              ("연락 희망 시간", "☐ 오전     ☐ 오후     ☐ 언제든"), ("메모", "")]
    tb = doc.add_table(rows=len(fields), cols=2)
    tb.autofit = False
    for (label, val), tr in zip(fields, tb.rows):
        tr.height = Mm(9.5 if label != "메모" else 16)
        a, b = tr.cells
        a.width, b.width = Mm(35), Mm(125)
        for c in (a, b):
            c.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
            cell_border(c)
        a.paragraphs[0].add_run(label).bold = True
        if val:
            b.paragraphs[0].add_run(val)

    para(doc, "회신 받는 곳 (지인회계법인)", size=10.5, bold=True, before=10, after=2)
    if REPLY_FAX:
        para(doc, f"  팩스  {REPLY_FAX}", size=11)
    para(doc, f"  전화·문자(사진)  {REPLY_MOBILE}  ·  이메일  {REPLY_EMAIL}", size=11)
    para(doc, "  지인회계법인 담당 이재강 공인회계사", size=10)
    para(doc, "적어 주신 귀사 담당자 정보는 회신 응대와 상담 연락에만 사용하며, 상담이 끝나면 파기합니다. "
              "'앞으로 안내를 원하지 않음'에 표시하시면 이후 연락드리지 않습니다.", size=8.5, before=8)


def load_zipcodes():
    """add_zipcode.py가 만든 zipcode_cache.csv (주소 → 우편번호). 없으면 우편번호 없이 인쇄"""
    if not ZIP_CACHE.exists():
        return {}
    z = pd.read_csv(ZIP_CACHE, dtype=str, encoding="utf-8-sig").fillna("")
    return dict(zip(z["주소"], z["우편번호"]))


def build_labels(m):
    L = LABEL
    doc = Document()
    set_font(doc)
    sec = doc.sections[0]
    sec.page_width, sec.page_height = Mm(210), Mm(297)
    sec.top_margin, sec.bottom_margin = Mm(L["top"]), Mm(0)
    sec.left_margin, sec.right_margin = Mm(L["left"]), Mm(0)
    sec.header_distance = sec.footer_distance = Mm(0)

    zipmap = load_zipcodes()
    per_page = L["cols"] * L["rows"]
    rows = [r for _, r in m.iterrows()]
    for start in range(0, len(rows), per_page):
        chunk = rows[start:start + per_page]
        ncols = L["cols"] * 2 - 1                         # 라벨 사이 간격 열
        tbl = doc.add_table(rows=L["rows"], cols=ncols)
        tbl.alignment = WD_TABLE_ALIGNMENT.LEFT
        tbl.autofit = False
        for r_i, tr in enumerate(tbl.rows):
            tr.height = Mm(L["h"])
            trPr = tr._tr.get_or_add_trPr()
            h = trPr.find(qn("w:trHeight"))
            h.set(qn("w:hRule"), "exact")
            for c_i, cell in enumerate(tr.cells):
                cell.width = Mm(L["gap"] if c_i % 2 else L["w"])
                if c_i % 2:
                    continue
                idx = r_i * L["cols"] + c_i // 2
                cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
                cell.paragraphs[0].text = ""
                if idx >= len(chunk):
                    continue
                row = chunk[idx]
                p = cell.paragraphs[0]
                p.paragraph_format.left_indent = Mm(6)
                if LABEL_NOTICE:
                    run = p.add_run(f"■ {LABEL_NOTICE}")
                    run.font.size, run.bold = Pt(7.5), True
                    p = cell.add_paragraph()
                    p.paragraph_format.left_indent = Mm(6)
                    p.paragraph_format.space_before = Pt(2)
                addr = str(row["주소"]) if pd.notna(row["주소"]) else ""
                if zipmap.get(addr):
                    run = p.add_run(f"(우) {zipmap[addr]}   ")
                    run.font.size, run.bold = Pt(9), True
                p.add_run(addr).font.size = Pt(9)
                for text, size, bold in ((str(row["회사명"]), 11, True), (recipient(row), 10, False)):
                    q = cell.add_paragraph()
                    q.paragraph_format.left_indent = Mm(6)
                    q.paragraph_format.space_before = Pt(2)
                    run = q.add_run(text)
                    run.font.size, run.bold = Pt(size), bold
        for col_i, col in enumerate(tbl.columns):
            col.width = Mm(L["gap"] if col_i % 2 else L["w"])
        if start + per_page < len(rows):
            doc.add_paragraph().add_run().add_break(WD_BREAK.PAGE)
    return doc


def save(doc, path):
    try:
        doc.save(path)
        return path
    except PermissionError:
        alt = path.with_name(f"{path.stem}_{time.strftime('%H%M%S')}{path.suffix}")
        print(f"  ! {path.name} 이(가) 열려 있어 {alt.name} 으로 저장 (다음엔 Word를 닫고 실행)")
        doc.save(alt)
        return alt


def main():
    ap = argparse.ArgumentParser(description="커버레터·주소 라벨 Word 파일 생성")
    ap.add_argument("--date", help="편지 날짜 (기본: 오늘)")
    a = ap.parse_args()
    if not MAIL.exists():
        sys.exit(f"[오류] {MAIL.name} 이(가) 없습니다. 먼저 python dart_signals.py --write-mailing")
    m = pd.read_excel(MAIL, sheet_name="발송명단", dtype=str)
    m = m[m["회사명"].notna()]
    # 커버레터·회신서와 주소 라벨을 같은 순서(회사명 가나다순)로 인쇄해 봉투 작업 때 맞춰 넣기 쉽게 한다
    m = m.sort_values("회사명", key=lambda s: s.str.replace(r"\s|\(주\)|주식회사", "", regex=True)).reset_index(drop=True)
    no_addr = m["주소"].isna().sum()
    if no_addr:
        print(f"  ! 주소 없는 {no_addr}곳은 라벨에 회사명만 들어갑니다")
    today = date.today()
    when = a.date or f"{today.year}년 {today.month}월 {today.day}일"

    p1 = save(build_letters(m, when), OUT_LETTER)
    p2 = save(build_labels(m), OUT_LABEL)
    pages = -(-len(m) // (LABEL["cols"] * LABEL["rows"]))
    print(f"[완료] {p1.name}: {len(m)}곳 × 2쪽(커버레터+회신서) / {p2.name}: {len(m)}칸, {pages}쪽 (날짜 {when})")


if __name__ == "__main__":
    main()
