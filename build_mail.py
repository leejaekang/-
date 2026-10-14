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
필요 패키지: python-docx (python -m pip install python-docx)
"""
import argparse
import sys
import time
from datetime import date
from pathlib import Path

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
]
FIRM_POINTS = [
    "2024년 4월 설립, 상근 공인회계사 18명 (경력 15년 이상 13명)",
    "회계감사 12명 · 세무자문 3명 · 경영자문 3명",
    "비상장 회사 외부감사에 집중합니다",
]
CLOSING = [
    "안내문이 잘 전달되었는지 확인하기 위해 담당자가 한 차례 전화드릴 수 있습니다. "
    "안내를 원하지 않으시면 위 연락처로 알려 주십시오. 이후 연락드리지 않겠습니다.",
    "감사합니다.",
]
CONTACT = [
    "지인회계법인",
    "서울특별시 강남구 테헤란로103길 6, 5층 (삼성동)",
    "대표전화 070-8765-1910",
    "담당 이재강 공인회계사 · 010-6624-2061 · ljkdlworkd@gmail.com",
]
SIGNATURE = "지인회계법인  공인회계사  이 재 강"

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
        para(doc, "지인회계법인 소개", bold=True, before=14, after=2)
        for text in FIRM_POINTS:
            para(doc, f"  ·  {text}")
        para(doc, "연락처", bold=True, before=14, after=2)
        for text in CONTACT:
            para(doc, f"  {text}")
        for text in CLOSING:
            para(doc, text, before=12)
        para(doc, SIGNATURE, size=11, bold=True, align=WD_ALIGN_PARAGRAPH.RIGHT, before=16)
    return doc


def build_labels(m):
    L = LABEL
    doc = Document()
    set_font(doc)
    sec = doc.sections[0]
    sec.page_width, sec.page_height = Mm(210), Mm(297)
    sec.top_margin, sec.bottom_margin = Mm(L["top"]), Mm(0)
    sec.left_margin, sec.right_margin = Mm(L["left"]), Mm(0)
    sec.header_distance = sec.footer_distance = Mm(0)

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
                p.add_run(str(row["주소"]) if pd.notna(row["주소"]) else "").font.size = Pt(9)
                for text, size, bold in ((str(row["회사명"]), 11, True), (recipient(row), 10, False)):
                    q = cell.add_paragraph()
                    q.paragraph_format.left_indent = Mm(6)
                    q.paragraph_format.space_before = Pt(3)
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
    no_addr = m["주소"].isna().sum()
    if no_addr:
        print(f"  ! 주소 없는 {no_addr}곳은 라벨에 회사명만 들어갑니다")
    today = date.today()
    when = a.date or f"{today.year}년 {today.month}월 {today.day}일"

    p1 = save(build_letters(m, when), OUT_LETTER)
    p2 = save(build_labels(m), OUT_LABEL)
    pages = -(-len(m) // (LABEL["cols"] * LABEL["rows"]))
    print(f"[완료] {p1.name}: {len(m)}장 / {p2.name}: {len(m)}칸, {pages}쪽 (날짜 {when})")


if __name__ == "__main__":
    main()
