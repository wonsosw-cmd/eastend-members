# -*- coding: utf-8 -*-
"""EASTEND 오프라인 멤버십 — 매장 운영 가이드 PPT (v2.4: 웰컴 3,000P + 익일 적립)"""
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE
from pptx.oxml.ns import qn

NAVY = RGBColor(0x1B, 0x2A, 0x4A)
TEAL = RGBColor(0x2A, 0x9D, 0x8F)
AMBER = RGBColor(0xE9, 0xA1, 0x3B)
BG = RGBColor(0xF6, 0xF7, 0xF9)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
TEXT = RGBColor(0x22, 0x29, 0x3A)
MUTED = RGBColor(0x6B, 0x74, 0x86)
LINE = RGBColor(0xE4, 0xE8, 0xEE)
TEAL_T = RGBColor(0xE2, 0xF3, 0xF1)
AMBER_T = RGBColor(0xFC, 0xF0, 0xDE)
NAVY_T = RGBColor(0xE8, 0xEC, 0xF5)
RED = RGBColor(0xD6, 0x45, 0x45)

FONT = "Pretendard"
QR_SAMPLE = r"G:\공유 드라이브\영업마케팅실\★. 데일리 매출 보고\★오프라인 회원관리\QR_포스터용_PNG\C005_시티브리즈_플래그십.png"
OUT = r"G:\공유 드라이브\영업마케팅실\★. 데일리 매출 보고\★오프라인 회원관리\EASTEND_멤버십_매장가이드_v1.pptx"

prs = Presentation()
prs.slide_width = Inches(13.333)
prs.slide_height = Inches(7.5)
BLANK = prs.slide_layouts[6]


def slide(bg=BG):
    s = prs.slides.add_slide(BLANK)
    r = s.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, prs.slide_width, prs.slide_height)
    r.fill.solid(); r.fill.fore_color.rgb = bg
    r.line.fill.background()
    r.shadow.inherit = False
    return s


def box(s, x, y, w, h, fill=None, line=None, radius=0.08):
    shp = s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(x), Inches(y), Inches(w), Inches(h))
    try:
        shp.adjustments[0] = radius
    except Exception:
        pass
    if fill is None:
        shp.fill.background()
    else:
        shp.fill.solid(); shp.fill.fore_color.rgb = fill
    if line is None:
        shp.line.fill.background()
    else:
        shp.line.color.rgb = line; shp.line.width = Pt(1)
    shp.shadow.inherit = False
    return shp


def circle(s, x, y, d, fill):
    shp = s.shapes.add_shape(MSO_SHAPE.OVAL, Inches(x), Inches(y), Inches(d), Inches(d))
    shp.fill.solid(); shp.fill.fore_color.rgb = fill
    shp.line.fill.background()
    shp.shadow.inherit = False
    return shp


def text(s, x, y, w, h, runs, size=14, color=TEXT, bold=False, align=PP_ALIGN.LEFT,
         anchor=MSO_ANCHOR.TOP, line_spacing=1.15, space_after=4):
    tb = s.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    tf = tb.text_frame
    tf.word_wrap = True
    tf.vertical_anchor = anchor
    tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
    if isinstance(runs, str):
        runs = [runs]
    for i, para in enumerate(runs):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = align
        p.line_spacing = line_spacing
        p.space_after = Pt(space_after)
        if isinstance(para, str):
            para = [(para, {})]
        para = [(it, {}) if isinstance(it, str) else it for it in para]
        for t, st in para:
            r = p.add_run()
            r.text = t
            r.font.name = FONT
            r.font.size = Pt(st.get("size", size))
            r.font.bold = st.get("bold", bold)
            r.font.color.rgb = st.get("color", color)
            rPr = r._r.get_or_add_rPr()
            ea = rPr.find(qn('a:ea'))
            if ea is None:
                ea = rPr.makeelement(qn('a:ea'), {})
                rPr.append(ea)
            ea.set('typeface', FONT)
    return tb


def num_badge(s, x, y, n, color=TEAL, d=0.42):
    circle(s, x, y, d, color)
    text(s, x, y + 0.015, d, d - 0.03, str(n), size=17, color=WHITE, bold=True,
         align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)


def title_bar(s, step, title, sub=""):
    text(s, 0.65, 0.42, 1.9, 0.35, step, size=13, color=TEAL, bold=True)
    text(s, 0.65, 0.72, 10.5, 0.75, title, size=30, color=NAVY, bold=True)
    if sub:
        text(s, 0.65, 1.38, 12.0, 0.4, sub, size=13.5, color=MUTED)


# ───────────────────────── 1. 표지
s = slide(NAVY)
text(s, 0.9, 1.55, 6.0, 0.4, "CITYBREEZE  ×  ARTID", size=15, color=AMBER, bold=True)
text(s, 0.9, 2.05, 11.5, 1.9,
     [[("EASTEND 오프라인 멤버십", {})], [("매장 운영 가이드", {})]],
     size=44, color=WHITE, bold=True, line_spacing=1.1)
text(s, 0.9, 4.15, 11.0, 0.9,
     [["QR 가입  ·  구매 시 자동 적립  ·  마일리지 에누리"],
      [("전 매장 공통 절차 안내  |  2026. 07", {"color": RGBColor(0xAF, 0xB9, 0xCE), "size": 14})]],
     size=17, color=WHITE)
stats = [("신규가입 혜택", "3,000P 지급"), ("적립률", "구매액의 1% 자동"), ("사용 (1P=1원)", "5,000P부터 1,000P 단위")]
for i, (k, v) in enumerate(stats):
    x = 0.9 + i * 4.0
    box(s, x, 5.45, 3.6, 1.25, fill=RGBColor(0x25, 0x36, 0x5C))
    text(s, x + 0.3, 5.68, 3.0, 0.3, k, size=12.5, color=RGBColor(0xAF, 0xB9, 0xCE), bold=True)
    text(s, x + 0.3, 6.02, 3.2, 0.5, v, size=19, color=WHITE, bold=True)

# ───────────────────────── 2. 한눈에 보기
s = slide()
title_bar(s, "OVERVIEW", "멤버십, 이렇게 돌아갑니다",
          "매장은 QR 가입만 안내하면 됩니다. 회원 등록·적립·사용은 본사와 판매 프로그램이 처리합니다.")
flow = [
    ("1", "QR 스캔 → 가입", "매장 QR 포스터를 스캔해\n정보 입력·동의하면 가입 완료\n(매장이 할 일은 안내뿐!)", TEAL),
    ("2", "본사가 회원 등록", "가입 명단을 본사에서 받아\n판매 프로그램에 등록\n+ 웰컴 3,000P 지급", AMBER),
    ("3", "구매 시 자동 적립", "구매하면 결제금액의 1%가\n프로그램에서 자동 적립\n(아티드 당일 · 시티브리즈 익일)", NAVY),
    ("4", "포인트로 에누리", "판매 프로그램에서 조회 후\n5,000P부터 1,000P 단위 차감\n(1P = 1원)", TEAL),
]
for i, (n, t, d, c) in enumerate(flow):
    x = 0.65 + i * 3.13
    box(s, x, 2.0, 2.85, 3.3, fill=WHITE, line=LINE)
    num_badge(s, x + 0.28, 2.3, n, color=c)
    text(s, x + 0.28, 2.95, 2.35, 0.75, t, size=16.5, color=NAVY, bold=True)
    text(s, x + 0.28, 3.72, 2.35, 1.45, d, size=12, color=MUTED, line_spacing=1.3)
    if i < 3:
        text(s, x + 2.83, 3.35, 0.35, 0.5, "→", size=20, color=MUTED, bold=True, align=PP_ALIGN.CENTER)
box(s, 0.65, 5.65, 12.05, 1.15, fill=NAVY_T)
text(s, 1.0, 5.92, 11.4, 0.7,
     [[("매장에서 하는 일은 단 하나  —  ", {"bold": True, "color": NAVY}),
       ("QR 가입 안내!   (적립·조회·에누리는 판매 프로그램에서 처리)", {"color": TEXT})]],
     size=15)

# ───────────────────────── 3. STEP 1 고객 가입
s = slide()
title_bar(s, "STEP 1", "고객 가입 안내 — QR만 보여주세요",
          "가입은 고객 본인 휴대폰에서 진행됩니다. 매장은 QR 포스터만 안내하면 됩니다.")
steps = [
    ("QR 스캔", "계산대·피팅룸 옆에 비치된 매장 전용 QR을 고객 휴대폰 카메라로 스캔"),
    ("정보 입력", "이름·휴대전화번호·생년월일(필수), 성별(선택) 입력"),
    ("약관 동의", "필수/선택 항목을 구분해 동의 — ‘전체 동의’는 편의 기능(개별 해제 가능)"),
    ("가입 완료", "신규가입 3,000P 지급 — 1% 적립은 아티드 당일 / 시티브리즈 다음 날부터"),
]
for i, (t, d) in enumerate(steps):
    y = 2.0 + i * 1.15
    num_badge(s, 0.75, y + 0.08, i + 1, color=TEAL)
    text(s, 1.4, y, 6.3, 0.35, t, size=16, color=NAVY, bold=True)
    text(s, 1.4, y + 0.38, 6.5, 0.65, d, size=12.5, color=MUTED, line_spacing=1.25)
box(s, 8.6, 1.95, 4.1, 4.75, fill=WHITE, line=LINE)
try:
    s.shapes.add_picture(QR_SAMPLE, Inches(9.55), Inches(2.35), Inches(2.2), Inches(2.2))
except Exception:
    pass
text(s, 8.85, 4.7, 3.6, 0.4, "매장별 전용 QR 포스터", size=14, color=NAVY, bold=True, align=PP_ALIGN.CENTER)
text(s, 8.85, 5.12, 3.6, 1.3,
     [["매장마다 QR이 다릅니다 (매장 실적 집계)."],
      ["포스터 인쇄본은 본사에서 배포하며,"],
      ["훼손 시 본사로 재요청해주세요."]],
     size=11.5, color=MUTED, align=PP_ALIGN.CENTER, line_spacing=1.25)
box(s, 0.65, 6.35, 7.6, 0.75, fill=AMBER_T)
text(s, 0.95, 6.52, 7.1, 0.45,
     [[("멘트 예시  ", {"bold": True, "color": NAVY}),
       ("“지금 QR 찍고 가입하시면 3,000P를 바로 드려요!”", {"color": TEXT})]], size=13)

# ───────────────────────── 4. STEP 2 적립 안내 (자동)
s = slide()
title_bar(s, "STEP 2", "적립은 자동입니다 — 매장이 할 일 없음",
          "구매 시 판매 프로그램에서 결제금액의 1%가 자동 적립됩니다. 별도 등록·입력이 필요 없습니다.")
box(s, 0.65, 1.95, 5.95, 3.0, fill=WHITE, line=LINE)
text(s, 1.0, 2.2, 5.2, 0.4, "🎁 신규가입 웰컴", size=16, color=TEAL, bold=True)
text(s, 1.0, 2.72, 5.3, 2.0,
     [["QR 가입 명단을 본사가 프로그램에 등록하면"],
      [("웰컴 3,000P 지급", {"bold": True, "color": NAVY})], [""],
      ["매장은 가입 안내만 하면 되고,"],
      ["등록·지급은 본사가 처리합니다."]],
     size=13, color=TEXT, line_spacing=1.3)
box(s, 6.75, 1.95, 5.95, 3.0, fill=WHITE, line=LINE)
text(s, 7.1, 2.2, 5.2, 0.4, "🛍 구매 적립 — 브랜드별로 다름!", size=16, color=AMBER, bold=True)
text(s, 7.1, 2.72, 5.3, 2.0,
     [["구매(결제) 시 결제금액의 1%가 자동 적립"], [""],
      [("아티드", {"bold": True, "color": TEAL}), ("  →  가입 ", {}), ("당일", {"bold": True, "color": TEAL}), (" 구매분부터 바로 적립", {})],
      [("시티브리즈", {"bold": True, "color": RED}), ("  →  가입 ", {}), ("다음 날", {"bold": True, "color": RED}), (" 구매분부터 적립", {})],
      [("            (가입 당일 구매분은 적립 안 됨)", {"color": MUTED, "size": 11.5})]],
     size=13, color=TEXT, line_spacing=1.3)
box(s, 0.65, 5.2, 12.05, 1.75, fill=NAVY_T)
text(s, 1.0, 5.4, 11.5, 0.35, "적립 규칙", size=14, color=NAVY, bold=True)
text(s, 1.0, 5.78, 11.5, 1.1,
     [[("·  적립액 = 결제금액의 1% (원단위 절사, 구매 시 자동)      ·  별도 등록·영수증 입력 불필요", {})],
      [("·  적립 시작:  ", {}), ("아티드 = 가입 당일부터", {"bold": True, "color": TEAL}),
       ("   /   ", {}), ("시티브리즈 = 가입 익일부터", {"bold": True, "color": RED}), ("  (웰컴 3,000P는 두 브랜드 모두 지급)", {})],
      [("·  포인트는 ", {}), ("적립일로부터 1년", {"bold": True, "color": RED}),
       ("  경과 시 자동 소멸  ·  반품 시 적립 취소는 본사에서 처리", {})]],
     size=13, color=TEXT, line_spacing=1.3)

# ───────────────────────── 7. FAQ
s = slide()
title_bar(s, "FAQ", "자주 묻는 질문 & 꼭 지켜주세요")
faqs = [
    ("포인트는 언제 쌓이나요?", "구매 시 결제금액의 1%가 자동 적립됩니다. 매장에서 별도 등록할 것이 없습니다."),
    ("미가입 고객이에요", "QR 가입을 먼저 안내해주세요. 웰컴 3,000P는 두 브랜드 모두 지급됩니다."),
    ("가입 당일 구매도 적립되나요?", "아티드는 네(당일부터 적립), 시티브리즈는 아니요(다음 날 구매분부터). 우리 매장 브랜드 기준으로 안내해주세요."),
    ("반품 고객의 적립은?", "적립된 포인트는 본사(영업마케팅실)로 알려주시면 차감 처리합니다."),
]
for i, (q, a) in enumerate(faqs):
    x = 0.65 + (i % 2) * 6.2
    y = 1.85 + (i // 2) * 1.5
    box(s, x, y, 5.95, 1.3, fill=WHITE, line=LINE)
    text(s, x + 0.3, y + 0.18, 5.4, 0.35, [[("Q. ", {"color": TEAL, "bold": True}), (q, {"bold": True, "color": NAVY})]], size=13.5)
    text(s, x + 0.3, y + 0.58, 5.4, 0.65, a, size=12, color=MUTED, line_spacing=1.25)
box(s, 0.65, 5.05, 12.05, 1.75, fill=RGBColor(0xFB, 0xEC, 0xEC))
text(s, 1.0, 5.28, 11.5, 0.35, "개인정보 — 반드시 지켜주세요", size=14.5, color=RED, bold=True)
text(s, 1.0, 5.7, 11.5, 1.0,
     [["·  고객 정보(전화번호·이름 등)는 외부에 공유 금지"],
      ["·  고객 정보는 적립·차감 업무에만 사용하고, 화면을 고객 외 타인에게 보여주지 않기"],
      ["·  고객이 탈퇴·동의 철회를 요청하면 본사(영업마케팅실)로 즉시 전달"]],
     size=13, color=TEXT, line_spacing=1.35)

# ───────────────────────── 8. 마무리
s = slide(NAVY)
text(s, 0.9, 2.1, 11.5, 0.9, "오늘부터 시작하세요!", size=38, color=WHITE, bold=True)
text(s, 0.9, 3.05, 11.5, 0.5, "QR 포스터 비치 → 첫 고객 가입 안내 — 이게 전부입니다!", size=16, color=RGBColor(0xAF, 0xB9, 0xCE))
links = [
    ("QR 포스터 (인쇄)", "wonsosw-cmd.github.io/eastend-members/qr.html"),
    ("포인트 조회·사용", "매장 판매 프로그램에서 처리"),
    ("문의", "영업마케팅실  ·  wonsosw@eastend.co.kr"),
]
for i, (k, v) in enumerate(links):
    y = 4.0 + i * 0.85
    box(s, 0.9, y, 11.5, 0.7, fill=RGBColor(0x25, 0x36, 0x5C))
    text(s, 1.25, y + 0.17, 2.6, 0.4, k, size=13.5, color=AMBER, bold=True)
    text(s, 3.9, y + 0.15, 8.3, 0.4, v, size=14.5, color=WHITE)

prs.save(OUT)
print("saved:", OUT)
