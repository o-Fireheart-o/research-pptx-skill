# -*- coding: utf-8 -*-
"""Build the user manual deck for the research-pptx skill.

  python -I docs/build_manual.py            (run from the repository root; writes docs/research-pptx 사용 매뉴얼.pptx)

Re-run after changing the text below. Images: docs/img/ (renders of the example storyboard deck).
"""
import os

from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_CONNECTOR, MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.util import Inches, Pt

HERE = os.path.dirname(os.path.abspath(__file__))
OUT = os.path.join(HERE, "research-pptx 사용 매뉴얼.pptx")
FONT = "맑은 고딕"
NAVY, TEAL, INK, GRAY = RGBColor(0x1F, 0x3A, 0x5F), RGBColor(0x0F, 0x8B, 0x8D), RGBColor(0x22, 0x22, 0x22), RGBColor(0x6B, 0x72, 0x80)
PALE, LINE, RED, WHITE, WARM = RGBColor(0xF1, 0xF5, 0xF9), RGBColor(0xCB, 0xD5, 0xE1), RGBColor(0xC0, 0x1F, 0x2F), RGBColor(0xFF, 0xFF, 0xFF), RGBColor(0xFF, 0xF4, 0xE5)

prs = Presentation()
prs.slide_width, prs.slide_height = Inches(13.333), Inches(7.5)
BLANK = prs.slide_layouts[6]
page = [0]


def text(slide, x, y, w, h, lines, size=16, color=INK, bold=False, align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.TOP, gap=4):
    """lines: str or list of str / (str, {size, color, bold}) ."""
    tb = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    tf = tb.text_frame
    tf.word_wrap = True
    tf.vertical_anchor = anchor
    tf.margin_left = tf.margin_right = Inches(0.05)
    tf.margin_top = tf.margin_bottom = Inches(0.03)
    fill(tf, lines, size, color, bold, align, gap)
    return tb


def fill(tf, lines, size, color, bold, align, gap=4):
    if isinstance(lines, str):
        lines = [lines]
    for i, line in enumerate(lines):
        opt = {}
        if isinstance(line, tuple):
            line, opt = line
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = align
        p.space_after = Pt(gap)
        r = p.add_run()
        r.text = line
        r.font.name = FONT
        r.font.size = Pt(opt.get("size", size))
        r.font.bold = opt.get("bold", bold)
        r.font.color.rgb = opt.get("color", color)
        rpr = r._r.get_or_add_rPr()
        for tag in ("ea", "cs"):
            el = rpr.makeelement("{http://schemas.openxmlformats.org/drawingml/2006/main}" + tag, {"typeface": FONT})
            rpr.append(el)


def box(slide, x, y, w, h, lines, size=14, fillc=PALE, linec=None, color=INK, bold=False, align=PP_ALIGN.LEFT,
        anchor=MSO_ANCHOR.MIDDLE, shape=MSO_SHAPE.ROUNDED_RECTANGLE, gap=3):
    s = slide.shapes.add_shape(shape, Inches(x), Inches(y), Inches(w), Inches(h))
    if shape == MSO_SHAPE.ROUNDED_RECTANGLE:
        s.adjustments[0] = 0.08
    s.fill.solid()
    s.fill.fore_color.rgb = fillc
    if linec is None:
        s.line.fill.background()
    else:
        s.line.color.rgb = linec
        s.line.width = Pt(1)
    s.shadow.inherit = False
    tf = s.text_frame
    tf.word_wrap = True
    tf.vertical_anchor = anchor
    tf.margin_left = tf.margin_right = Inches(0.14)
    tf.margin_top = tf.margin_bottom = Inches(0.08)
    fill(tf, lines, size, color, bold, align, gap)
    return s


def arrow(slide, x1, y1, x2, y2, color=NAVY, width=1.5):
    c = slide.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, Inches(x1), Inches(y1), Inches(x2), Inches(y2))
    c.line.color.rgb = color
    c.line.width = Pt(width)
    ln = c.line._get_or_add_ln()
    ln.append(ln.makeelement("{http://schemas.openxmlformats.org/drawingml/2006/main}tailEnd", {"type": "triangle"}))
    return c


def new(title, lead=None, chapter=""):
    s = prs.slides.add_slide(BLANK)
    page[0] += 1
    text(s, 0.6, 0.38, 10.6, 0.7, title, size=28, color=NAVY, bold=True, anchor=MSO_ANCHOR.MIDDLE)
    if chapter:
        text(s, 10.2, 0.45, 2.55, 0.4, chapter, size=12, color=GRAY, align=PP_ALIGN.RIGHT)
    if lead:
        text(s, 0.6, 1.1, 12.1, 0.55, lead, size=16, color=TEAL, bold=True)
    text(s, 12.2, 6.98, 0.6, 0.3, str(page[0]), size=11, color=GRAY, align=PP_ALIGN.RIGHT)
    return s


def say(slide, x, y, w, h, words, size=15):
    """What the user types to the agent."""
    return box(slide, x, y, w, h, [("이렇게 말하기", {"size": 11, "color": TEAL, "bold": True}), words],
               size=size, fillc=WHITE, linec=TEAL, anchor=MSO_ANCHOR.MIDDLE)


# ------------------------------------------------------------------ 1 cover
s = prs.slides.add_slide(BLANK)
page[0] += 1
bg = s.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, prs.slide_width, prs.slide_height)
bg.fill.solid(); bg.fill.fore_color.rgb = NAVY; bg.line.fill.background()
text(s, 0.9, 2.3, 11.5, 1.0, "research-pptx 사용 매뉴얼", size=44, color=WHITE, bold=True)
text(s, 0.9, 3.4, 11.5, 0.6, "연구 제안서·연구 보고서 PPT를 양식에 맞춰 만드는 Codex 스킬", size=20, color=RGBColor(0xCF, 0xE3, 0xF1))
text(s, 0.9, 6.3, 11.5, 0.4, "스토리라인을 먼저 정하고, 양식대로 만들고, 근거 없는 내용은 넣지 않습니다", size=14, color=RGBColor(0xCF, 0xE3, 0xF1))

# ------------------------------------------------------------------ 2 what it does
s = new("이 스킬이 하는 일", "양식을 읽고, 흐름을 먼저 합의하고, 그 내용을 양식대로 PPT로 만듭니다", "1. 소개")
cards = [("① 양식 읽기", "기관 양식이나 지난 발표 자료에서\n슬라이드 유형, 글꼴, 색, 정해진 목차를\n읽어 그대로 따릅니다"),
         ("② 스토리라인 합의", "무엇을 어떤 순서로 말할지를\n뼈대 덱(PPT)으로 보여 줍니다\n여기서 고치고 확정합니다"),
         ("③ 양식대로 제작", "샘플 한두 장을 먼저 확인받고\n같은 방식으로 전체를 만든 뒤\n렌더링해서 검수합니다")]
for i, (h, b) in enumerate(cards):
    x = 0.6 + i * 4.1
    box(s, x, 1.95, 3.85, 0.65, h, size=18, fillc=NAVY, color=WHITE, bold=True, align=PP_ALIGN.CENTER)
    box(s, x, 2.65, 3.85, 1.75, b, size=14, fillc=PALE, align=PP_ALIGN.CENTER)
    if i < 2:
        arrow(s, x + 3.88, 3.1, x + 4.08, 3.1)
box(s, 0.6, 4.75, 12.1, 1.75,
    [("가장 중요한 원칙: 근거 없는 내용을 만들지 않습니다", {"size": 17, "bold": True, "color": RED}),
     "· 수치, 실험 결과, 참고문헌은 내가 준 자료에 있는 것만 씁니다",
     "· 자료에 없어서 채울 수 없는 자리는 [확인 필요]로 남기고 목록으로 알려 줍니다",
     "· 결과의 해석을 자료보다 강하게 바꾸지 않습니다"],
    size=14, fillc=WARM, anchor=MSO_ANCHOR.MIDDLE)

# ------------------------------------------------------------------ 3 install: prerequisites + ask Codex
REPO = "https://github.com/o-Fireheart-o/research-pptx-skill"
s = new("설치 ① 준비물과 가장 쉬운 설치", "스킬은 GitHub에 있습니다. Codex에게 설치를 시키면 됩니다", "2. 준비")
box(s, 0.6, 1.85, 5.2, 0.5, "준비물 (한 번만)", size=16, fillc=NAVY, color=WHITE, bold=True)
box(s, 0.6, 2.4, 5.2, 2.55,
    ["· Windows PC와 PowerPoint",
     "· Codex (앱 또는 CLI)",
     "· Python 3.11 이상",
     "· Python 패키지 세 개",
     ("  python -m pip install python-pptx lxml Pillow", {"size": 12, "color": TEAL, "bold": True}),
     ("  ※ 없으면 스킬을 처음 쓸 때 설치해도 되는지 물어봅니다", {"size": 11, "color": GRAY})],
    size=14, fillc=PALE, anchor=MSO_ANCHOR.TOP)
box(s, 6.1, 1.85, 6.6, 0.5, "가장 쉬운 방법: Codex에게 시키기", size=16, fillc=TEAL, color=WHITE, bold=True)
say(s, 6.1, 2.4, 6.6, 1.5,
    REPO + " 저장소의 research-pptx 폴더를 내 Codex 스킬 폴더에 설치해 줘.", size=14)
box(s, 6.1, 4.0, 6.6, 0.95,
    ["① 위 문장을 Codex에 붙여 넣습니다",
     "② 설치가 끝나면 Codex를 닫았다가 다시 엽니다"],
    size=14, fillc=PALE)
box(s, 0.6, 5.25, 5.9, 0.5, "설치 확인", size=15, fillc=NAVY, color=WHITE, bold=True)
say(s, 0.6, 5.8, 5.9, 0.75, "어떤 스킬을 쓸 수 있어?", size=14)
box(s, 6.8, 5.25, 5.9, 1.3,
    ["목록에 research-pptx가 보이면 설치된 것입니다",
     "보이지 않으면 다음 쪽의 방법으로 직접 설치합니다"],
    size=13, fillc=WHITE, linec=LINE, color=GRAY)

# ------------------------------------------------------------------ 4 install: manual from GitHub
s = new("설치 ② GitHub에서 직접 설치", REPO, "2. 준비")
box(s, 0.6, 1.85, 5.9, 0.5, "방법 A: ZIP으로 받기", size=16, fillc=NAVY, color=WHITE, bold=True)
box(s, 0.6, 2.4, 5.9, 2.65,
    ["① 위 주소를 브라우저에서 엽니다",
     "② 초록색 Code 버튼 > Download ZIP",
     "③ 받은 ZIP의 압축을 풉니다",
     "④ 안에 있는 research-pptx 폴더를 복사해 아래 위치에 붙여 넣습니다",
     ("  C:\\Users\\<사용자>\\.codex\\skills\\", {"size": 13, "color": TEAL, "bold": True}),
     "⑤ Codex를 다시 엽니다"],
    size=13, fillc=PALE, anchor=MSO_ANCHOR.TOP)
box(s, 6.8, 1.85, 5.9, 0.5, "방법 B: git 명령으로 받기 (PowerShell)", size=16, fillc=NAVY, color=WHITE, bold=True)
box(s, 6.8, 2.4, 5.9, 2.65,
    [("git clone " + REPO + ".git", {"size": 11, "color": TEAL, "bold": True}),
     ("Copy-Item -Recurse research-pptx-skill\\research-pptx", {"size": 11, "color": TEAL, "bold": True}),
     ("    \"$env:USERPROFILE\\.codex\\skills\\\"", {"size": 11, "color": TEAL, "bold": True}),
     "",
     "· 첫 줄은 저장소를 내려받고, 둘째 줄은 스킬 폴더를 Codex 스킬 폴더로 복사합니다",
     "· 끝나면 Codex를 다시 엽니다"],
    size=13, fillc=PALE, anchor=MSO_ANCHOR.TOP)
box(s, 0.6, 5.3, 5.9, 1.25,
    [("설치 후 폴더 모습", {"size": 14, "bold": True, "color": NAVY}),
     ("C:\\Users\\<사용자>\\.codex\\skills\\research-pptx\\SKILL.md", {"size": 12, "color": TEAL, "bold": True}),
     ("※ research-pptx 폴더 바로 안에 SKILL.md가 있어야 합니다", {"size": 11, "color": GRAY})],
    size=13, fillc=WHITE, linec=LINE)
box(s, 6.8, 5.3, 5.9, 1.25,
    [("새 버전으로 바꾸기", {"size": 14, "bold": True, "color": NAVY}),
     "같은 방법으로 다시 받아 research-pptx 폴더를 덮어씁니다",
     ("※ 스킬이 보이지 않으면 폴더를 .agents\\skills\\ 로 옮겨 봅니다", {"size": 11, "color": GRAY})],
    size=13, fillc=WHITE, linec=LINE)

# ------------------------------------------------------------------ 4 folder
s = new("작업 폴더 만들기", "문서 하나에 폴더 하나. 받은 자료를 세 폴더에 나눠 넣습니다", "2. 준비")
rows = [("template", "발표 양식이나 지난 발표 자료 (.pptx)", "내가 넣음", NAVY),
        ("source", "공고문, 초안, 논문, 데이터, 지난 제안서·보고서", "내가 넣음", NAVY),
        ("figures", "쓸 그림과 그래프", "내가 넣음", NAVY),
        ("storyboard.pptx", "뼈대 덱. 스토리라인을 보고 고치는 파일", "만들어짐 · 내가 고침", TEAL),
        ("output", "샘플 슬라이드와 완성본. backup 폴더에 고치기 전 파일", "만들어짐", GRAY),
        ("storyline.md · work", "글로 된 스토리라인과 작업 기록. 열어 보지 않아도 됩니다", "만들어짐", GRAY)]
for i, (name, desc, who, c) in enumerate(rows):
    y = 1.9 + i * 0.75
    box(s, 0.6, y, 2.9, 0.62, name, size=15, fillc=c, color=WHITE, bold=True)
    box(s, 3.6, y, 6.6, 0.62, desc, size=14, fillc=PALE)
    box(s, 10.3, y, 2.4, 0.62, who, size=13, fillc=WHITE, linec=c, color=c, bold=True, align=PP_ALIGN.CENTER)
text(s, 0.6, 6.5, 12.1, 0.4, "※ template 폴더의 양식 원본은 고쳐지지 않습니다. 완성본은 항상 output에 따로 만들어집니다", size=13, color=GRAY)

# ------------------------------------------------------------------ 5 flow
s = new("전체 흐름", "내가 확인해 줄 곳은 두 번입니다: 스토리라인과 샘플 슬라이드", "3. 흐름")
stages = [("0", "입력 확인", False), ("1", "양식 분석", False), ("2", "스토리라인", True), ("3", "자료 준비", False),
          ("4", "샘플 확인", True), ("5", "전체 조립", False), ("6", "검수", False), ("7", "수정 반복", False), ("8", "파생판", False)]
w = 1.28
for i, (n, name, gate) in enumerate(stages):
    x = 0.6 + i * (w + 0.075)
    c = TEAL if gate else NAVY
    box(s, x, 2.0, w, 0.5, n, size=16, fillc=c, color=WHITE, bold=True, align=PP_ALIGN.CENTER)
    box(s, x, 2.52, w, 0.85, name, size=13, fillc=PALE, linec=c if gate else None, bold=gate, align=PP_ALIGN.CENTER)
    if gate:
        box(s, x - 0.1, 3.5, w + 0.2, 0.42, "내가 확인", size=12, fillc=TEAL, color=WHITE, bold=True, align=PP_ALIGN.CENTER)
notes = [("2. 스토리라인", "뼈대 덱을 열어 흐름을 봅니다. 순서와 제목을 고치고 \"이대로 만들어 줘\"라고 하면 확정됩니다."),
         ("4. 샘플 슬라이드", "대표 한두 장을 먼저 봅니다. 글의 양, 글자 크기, 배치가 마음에 드는지 답합니다."),
         ("왜 두 번 확인하나", "다 만든 뒤에 구성이나 배치를 바꾸면 모든 장을 다시 만들어야 합니다. 앞에서 정하는 편이 빠르고 사용량도 적게 듭니다.")]
for i, (h, b) in enumerate(notes):
    y = 4.25 + i * 0.8
    box(s, 0.6, y, 2.6, 0.68, h, size=14, fillc=TEAL if i < 2 else GRAY, color=WHITE, bold=True, align=PP_ALIGN.CENTER)
    box(s, 3.3, y, 9.4, 0.68, b, size=13, fillc=PALE)

# ------------------------------------------------------------------ 6 start
s = new("시작하기", "작업 폴더에서 Codex를 열고 하고 싶은 일을 말합니다. 스킬 이름은 말하지 않아도 됩니다", "4. 사용")
say(s, 0.6, 1.95, 12.1, 1.05, "template 폴더의 양식으로 연구 제안 발표 자료를 만들 거야. source 폴더 자료를 보고 스토리라인부터 잡아 줘. 발표는 15분이야.")
say(s, 0.6, 3.15, 12.1, 1.05, "지난 연차 보고(template)와 같은 형식으로 2차년도 보고 자료를 만들어 줘. 지난 제안서는 source에 있어.")
box(s, 0.6, 4.5, 5.9, 0.5, "처음에 알려 주면 좋은 것", size=15, fillc=NAVY, color=WHITE, bold=True)
box(s, 0.6, 5.05, 5.9, 1.55,
    ["· 문서 종류: 제안 발표 / 연차 보고 / 최종 보고 / 회의 자료",
     "· 발표 시간, 분량 제한",
     "· 평가 항목이 있는 공고문의 위치",
     "· 누가 보는지 (같은 분야 심사위원인지)"],
    size=13, fillc=PALE, anchor=MSO_ANCHOR.TOP)
box(s, 6.8, 4.5, 5.9, 0.5, "시작점은 세 가지", size=15, fillc=NAVY, color=WHITE, bold=True)
box(s, 6.8, 5.05, 5.9, 1.55,
    ["· 자료만 있다 → 초안 스토리라인을 만들어 줍니다",
     "· 머릿속에 흐름이 있다 → 제목만 불러 주면 내용을 채워 줍니다",
     "· 이미 만든 개요나 PPT가 있다 → 그대로 읽어 옮깁니다"],
    size=13, fillc=PALE, anchor=MSO_ANCHOR.TOP)

# ------------------------------------------------------------------ 7 storyboard map
s = new("뼈대 덱 ① 1번 장은 전체 흐름도", "장(章)마다 기둥 하나, 그 아래 슬라이드 제목이 순서대로 놓입니다", "4. 사용")
s.shapes.add_picture(os.path.join(HERE, "img", "storyboard_map.png"), Inches(0.6), Inches(1.8), width=Inches(8.3))
pic_h = 8.3 * 9 / 16
frame = s.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0.6), Inches(1.8), Inches(8.3), Inches(pic_h))
frame.fill.background(); frame.line.color.rgb = LINE
legend = [("제목만 이어 읽기", "위에서 아래로, 왼쪽에서 오른쪽으로 읽어서 이야기가 이어지면 구성이 된 것입니다"),
          ("빨간 테두리", "아직 못 정한 것(?)이 남은 장"),
          ("회색 머리", "아직 슬라이드가 없는 장(章)"),
          ("위쪽 글", "종류, 발표 시간, 핵심 주장. 여기는 고치면 반영됩니다")]
for i, (h, b) in enumerate(legend):
    y = 1.8 + i * 1.02
    box(s, 9.15, y, 3.55, 0.36, h, size=13, fillc=NAVY, color=WHITE, bold=True)
    box(s, 9.15, y + 0.38, 3.55, 0.56, b, size=11, fillc=PALE)
box(s, 9.15, 5.92, 3.55, 0.55, "흐름도의 상자는 보기 전용입니다\n고칠 때는 2번 장부터 고칩니다", size=11, fillc=WARM, color=RED, bold=True, align=PP_ALIGN.CENTER)
text(s, 0.6, 6.55, 8.3, 0.35, "※ 그림은 가상의 과제로 만든 예시입니다", size=11, color=GRAY)

# ------------------------------------------------------------------ 8 storyboard slide
s = new("뼈대 덱 ② 2번 장부터는 기획 장", "실제 슬라이드 한 장마다 기획 장이 하나씩 있습니다", "4. 사용")
s.shapes.add_picture(os.path.join(HERE, "img", "storyboard_slide.png"), Inches(0.6), Inches(1.8), width=Inches(8.3))
frame = s.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0.6), Inches(1.8), Inches(8.3), Inches(pic_h))
frame.fill.background(); frame.line.color.rgb = LINE
legend = [("제목", "그 장에서 말하려는 결론 한 문장"),
          ("본문", "그 장에 담을 내용"),
          ("빨간 ? 줄", "아직 못 정한 것, 확인할 것"),
          ("점선 상자", "넣을 그림이나 표"),
          ("오른쪽 위 작은 글씨", "이 장이 속한 장(章) 이름"),
          ("슬라이드 노트", "만드는 쪽에 전할 말을 적는 곳")]
for i, (h, b) in enumerate(legend):
    y = 1.8 + i * 0.78
    box(s, 9.15, y, 3.55, 0.34, h, size=13, fillc=NAVY, color=WHITE, bold=True)
    box(s, 9.15, y + 0.36, 3.55, 0.36, b, size=11, fillc=PALE)
text(s, 0.6, 6.55, 12.1, 0.35, "※ 제목은 \"실험 결과\"처럼 주제어가 아니라 \"제안 소재가 오차를 절반으로 줄였다\"처럼 결론으로 씁니다", size=12, color=GRAY)

# ------------------------------------------------------------------ 9 edit storyboard
s = new("뼈대 덱 고치기", "PowerPoint에서 평소처럼 고치면 됩니다", "4. 사용")
edits = [("순서 바꾸기", "보기 > 여러 슬라이드에서\n끌어서 옮깁니다"),
         ("내용 고치기", "제목과 본문을 고쳐 씁니다\n모르는 것은 ? 로 시작해 적습니다"),
         ("장 추가", "비슷한 장을 복제해 고쳐 쓰거나\n빈 슬라이드에 제목만 적습니다"),
         ("장 삭제", "필요 없는 슬라이드를\n지웁니다"),
         ("전할 말", "\"표는 4군으로 나눠 줘\" 같은 말은\n슬라이드 노트에 적습니다"),
         ("그림 미리 넣기", "쓸 그림을 그 장에\n붙여 넣어도 됩니다")]
for i, (h, b) in enumerate(edits):
    x = 0.6 + (i % 3) * 4.1
    y = 1.9 + (i // 3) * 1.55
    box(s, x, y, 3.85, 0.45, h, size=15, fillc=NAVY, color=WHITE, bold=True, align=PP_ALIGN.CENTER)
    box(s, x, y + 0.47, 3.85, 0.92, b, size=13, fillc=PALE, align=PP_ALIGN.CENTER)
box(s, 0.6, 5.1, 5.9, 1.45,
    [("고친 뒤에는", {"size": 15, "bold": True, "color": RED}),
     "① 저장합니다",
     "② PowerPoint를 닫습니다",
     "③ 읽어 달라고 말합니다"],
    size=13, fillc=WARM, anchor=MSO_ANCHOR.MIDDLE)
say(s, 6.8, 5.1, 5.9, 0.7, "뼈대 덱 고쳤어. 읽어 줘.", size=14)
say(s, 6.8, 5.85, 5.9, 0.7, "말로 해도 됩니다: 3번과 4번 순서 바꿔 줘.", size=14)

# ------------------------------------------------------------------ 10 sample
s = new("샘플 슬라이드 확인", "\"이대로 만들어 줘\"라고 하면 대표 한두 장을 먼저 보여 줍니다", "4. 사용")
say(s, 0.6, 1.9, 5.9, 0.8, "스토리라인 이대로 만들어 줘.")
box(s, 0.6, 2.9, 5.9, 0.5, "받는 것", size=15, fillc=NAVY, color=WHITE, bold=True)
box(s, 0.6, 3.45, 5.9, 1.3,
    ["· output\\sample.pptx (양식대로 완성한 한두 장)",
     "· 글과 표가 가장 많은 장 하나",
     "· 그림이나 도식이 중심인 장 하나"],
    size=13, fillc=PALE, anchor=MSO_ANCHOR.TOP)
box(s, 6.8, 1.9, 5.9, 0.5, "보고 답할 것", size=15, fillc=TEAL, color=WHITE, bold=True)
box(s, 6.8, 2.45, 5.9, 2.3,
    ["· 한 장에 들어간 글의 양이 맞는가 (더 줄일지, 더 넣을지)",
     "· 글자 크기와 표의 모습",
     "· 그림과 글의 자리",
     "· 결론 줄의 위치와 표기",
     "· 양식의 장식과 어울리는가"],
    size=13, fillc=PALE, anchor=MSO_ANCHOR.TOP)
box(s, 0.6, 5.05, 12.1, 1.45,
    [("여기서 정한 것이 나머지 모든 장의 기준이 됩니다", {"size": 15, "bold": True, "color": NAVY}),
     "· 마음에 들면 \"좋아, 이대로 전체 만들어 줘\", 아니면 고칠 곳을 말합니다 (\"글을 조금 줄이고 표 글자를 키워 줘\")",
     "· 같은 기관 양식을 다시 쓰거나 다섯 장 이하의 짧은 문서는 이 단계를 건너뜁니다"],
    size=13, fillc=WHITE, linec=LINE, anchor=MSO_ANCHOR.MIDDLE)

# ------------------------------------------------------------------ 11 result
s = new("결과 받기", "완성본은 output 폴더에 생기고, 보고는 세 묶음으로 옵니다", "4. 사용")
three = [("결과", "완성본 파일 위치와 장수\n고치기 전 백업 위치", NAVY),
         ("확인이 필요한 것", "자료가 없어 [확인 필요]로 둔 자리\n받은 그림을 그대로 넣은 곳\n대조하지 못한 수치", RED),
         ("임의로 정한 것", "순서나 표현을 바꾼 곳과 이유\n원치 않으면 되돌립니다", TEAL)]
for i, (h, b, c) in enumerate(three):
    x = 0.6 + i * 4.1
    box(s, x, 1.9, 3.85, 0.5, h, size=16, fillc=c, color=WHITE, bold=True, align=PP_ALIGN.CENTER)
    box(s, x, 2.45, 3.85, 1.35, b, size=13, fillc=PALE, align=PP_ALIGN.CENTER)
box(s, 0.6, 4.1, 5.9, 0.5, "\"완성했다\"는 말의 뜻", size=15, fillc=NAVY, color=WHITE, bold=True)
box(s, 0.6, 4.65, 5.9, 1.85,
    ["· 스토리라인의 모든 장이 들어 있다",
     "· 모든 장을 렌더링해서 눈으로 확인했다",
     "· 양식에 있던 예전 문구와 그림이 남아 있지 않다",
     "· 남은 [확인 필요]가 보고에 모두 적혀 있다"],
    size=13, fillc=PALE, anchor=MSO_ANCHOR.TOP)
box(s, 6.8, 4.1, 5.9, 0.5, "막힌 장이 있으면", size=15, fillc=RED, color=WHITE, bold=True)
box(s, 6.8, 4.65, 5.9, 1.85,
    ["· 몇 번 장이 왜 남았는지 알려 줍니다",
     "· 없는 수치나 그래프를 그럴듯하게 채우지 않습니다",
     "· 무엇을 주면 풀리는지와 선택지를 함께 알려 줍니다",
     "· 나머지 장은 그대로 진행됩니다"],
    size=13, fillc=PALE, anchor=MSO_ANCHOR.TOP)

# ------------------------------------------------------------------ 12 after edits
s = new("완성본을 직접 고친 뒤", "내가 PowerPoint에서 고친 파일이 그때부터 원본입니다", "4. 사용")
box(s, 0.6, 1.9, 5.9, 0.5, "이렇게 됩니다", size=15, fillc=NAVY, color=WHITE, bold=True)
box(s, 0.6, 2.45, 5.9, 2.2,
    ["· 내가 다듬은 문장과 배치는 그대로 둡니다",
     "· 요청한 곳만 고칩니다",
     "· 고치기 전 파일을 output\\backup에 날짜와 함께 남깁니다",
     "· 장을 옮기거나 지웠으면 목차, 그림·표 번호를 다시 맞춥니다"],
    size=13, fillc=PALE, anchor=MSO_ANCHOR.TOP)
box(s, 6.8, 1.9, 5.9, 0.5, "지켜 줄 것", size=15, fillc=RED, color=WHITE, bold=True)
box(s, 6.8, 2.45, 5.9, 2.2,
    ["· 고친 뒤 저장하고 PowerPoint를 닫은 다음 말합니다",
     "· 어느 파일을 고쳤는지 말합니다",
     "· 지도교수와 공동연구자의 의견이 서로 다르면 어느 쪽을 따를지 정해서 알려 줍니다"],
    size=13, fillc=PALE, anchor=MSO_ANCHOR.TOP)
say(s, 0.6, 4.95, 12.1, 0.72, "output의 발표자료.pptx를 내가 고쳤어. 그 파일 기준으로 7번 장 표 글자 키우고, 12번 그림을 figures의 새 그림으로 바꿔 줘.")
say(s, 0.6, 5.82, 12.1, 0.72, "교수님 의견 받았어. source의 의견.txt 보고 반영해 줘. 반영 못 하는 건 이유를 알려 줘.")

# ------------------------------------------------------------------ 13 usage
s = new("사용량 아끼기", "요금제의 한도 안에서 끝까지 쓰려면", "5. 요령")
tips = [("두 번의 확인에 답하기", "스토리라인과 샘플에서 정해 두면\n다 만든 뒤 다시 만드는 일이 없습니다"),
        ("단계가 끝나면 새 대화", "대화가 길수록 한 번 답할 때 많이 듭니다\n스토리라인 확정 뒤, 샘플 승인 뒤가 좋습니다"),
        ("고칠 것은 모아서 한 번에", "\"7번 표 글자 키우고, 12번 그림 바꾸고,\n15번 삭제\"처럼 한 번에 말합니다"),
        ("양식 기록 저장하기", "문서를 마칠 때 \"양식 기록을 저장할까요?\"에\n그렇다고 하면 다음에 분석을 건너뜁니다")]
for i, (h, b) in enumerate(tips):
    x = 0.6 + (i % 2) * 6.2
    y = 1.9 + (i // 2) * 1.6
    box(s, x, y, 5.9, 0.48, h, size=15, fillc=NAVY, color=WHITE, bold=True)
    box(s, x, y + 0.5, 5.9, 0.95, b, size=13, fillc=PALE)
say(s, 0.6, 5.25, 12.1, 0.75, "(새 대화에서) work/NOTES.md와 work/progress.json 읽고 이어서 해 줘.")
text(s, 0.6, 6.15, 12.1, 0.5, "※ 양식 기록은 C:\\Users\\<사용자>\\.research-pptx\\profiles\\ 에 쌓입니다. 연구 내용은 저장되지 않습니다", size=12, color=GRAY)

# ------------------------------------------------------------------ 14 faq
s = new("자주 생기는 일", None, "5. 요령")
faq = [("스킬이 안 쓰이는 것 같다", "\"research-pptx 스킬로 해 줘\"라고 말합니다. 그래도 안 되면 설치 폴더를 확인합니다 (3~4쪽)"),
       ("뼈대 덱을 고쳤는데 반영이 안 됐다", "저장하고 PowerPoint를 닫았는지 봅니다. 1번 장 흐름도의 상자를 고친 것은 읽히지 않습니다"),
       ("[확인 필요]가 많이 남았다", "자료에 없는 내용입니다. 해당 자료를 source에 넣거나 값을 말해 주면 채워집니다"),
       ("글이 너무 많다 / 적다", "샘플 단계에서 말하는 것이 가장 좋습니다. 뒤에 말하면 모든 장을 다시 만듭니다"),
       ("양식의 예전 내용이 남아 있다", "몇 번 장인지 알려 주면 지웁니다. 지난 발표 자료를 양식으로 쓸 때 생길 수 있습니다"),
       ("며칠 뒤 이어서 하고 싶다", "같은 폴더에서 \"work/NOTES.md 읽고 이어서 해 줘\"라고 합니다")]
for i, (q, a) in enumerate(faq):
    y = 1.35 + i * 0.88
    box(s, 0.6, y, 4.0, 0.76, q, size=14, fillc=NAVY, color=WHITE, bold=True)
    box(s, 4.7, y, 8.0, 0.76, a, size=13, fillc=PALE)

# ------------------------------------------------------------------ 15 cheat sheet
s = new("한 장 요약", "순서대로 이렇게 말하면 됩니다", "5. 요령")
steps = [("시작", "template 양식으로 연구 제안 발표 자료를 만들 거야. source 자료를 보고 스토리라인부터 잡아 줘. 발표는 15분이야."),
         ("뼈대 덱을 고친 뒤", "뼈대 덱 고쳤어. 읽어 줘."),
         ("스토리라인 확정", "이대로 만들어 줘."),
         ("샘플을 본 뒤", "좋아, 이대로 전체 만들어 줘.   /   글을 조금 줄이고 표 글자를 키워 줘."),
         ("완성본을 고친 뒤", "output의 파일을 내가 고쳤어. 그 파일 기준으로 7번 장만 고쳐 줘."),
         ("이어서 할 때", "work/NOTES.md와 work/progress.json 읽고 이어서 해 줘.")]
for i, (h, b) in enumerate(steps):
    y = 1.85 + i * 0.78
    box(s, 0.6, y, 2.6, 0.66, h, size=14, fillc=TEAL, color=WHITE, bold=True, align=PP_ALIGN.CENTER)
    box(s, 3.3, y, 9.4, 0.66, b, size=13, fillc=WHITE, linec=TEAL)
    if i < len(steps) - 1:
        arrow(s, 1.9, y + 0.67, 1.9, y + 0.78, TEAL)

prs.save(OUT)
print(OUT, len(prs.slides))
