from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE
from pathlib import Path

ROOT = Path(__file__).parent
NAVY = RGBColor(22, 46, 112)
GOLD = RGBColor(243, 183, 63)
INK = RGBColor(24, 35, 59)
MUTED = RGBColor(100, 112, 138)
PALE = RGBColor(243, 246, 251)
WHITE = RGBColor(255, 255, 255)

prs = Presentation()
prs.slide_width = Inches(13.333)
prs.slide_height = Inches(7.5)

def textbox(slide, x, y, w, h, text, size, color, bold=False, font="Aptos"):
    shape = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    tf = shape.text_frame
    tf.clear()
    tf.word_wrap = True
    tf.margin_left = Inches(.02)
    tf.margin_right = Inches(.02)
    tf.margin_top = Inches(.02)
    p = tf.paragraphs[0]
    p.text = text
    p.font.name = font
    p.font.size = Pt(size)
    p.font.bold = bold
    p.font.color.rgb = color
    return shape

def base(title, kicker="ART205 GROUP 1"):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    bg = slide.background.fill
    bg.solid()
    bg.fore_color.rgb = WHITE
    bar = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, prs.slide_width, Inches(.16))
    bar.fill.solid(); bar.fill.fore_color.rgb = GOLD; bar.line.fill.background()
    textbox(slide, .72, .48, 5, .3, kicker, 10, NAVY, True)
    textbox(slide, .72, .88, 11.7, .72, title, 29, NAVY, True)
    textbox(slide, .72, 7.12, 11.8, .2, "DIGITAL TRAINING CENTER  |  Sample website project", 9, MUTED)
    return slide

def bullet_list(slide, x, y, w, items, size=19, gap=.73):
    for i, item in enumerate(items):
        cy = y + i * gap
        dot = slide.shapes.add_shape(MSO_SHAPE.OVAL, Inches(x), Inches(cy+.08), Inches(.13), Inches(.13))
        dot.fill.solid(); dot.fill.fore_color.rgb = GOLD; dot.line.fill.background()
        textbox(slide, x+.32, cy, w-.32, .62, item, size, INK)

def card(slide, x, y, w, h, head, body):
    shape = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(x), Inches(y), Inches(w), Inches(h))
    shape.fill.solid(); shape.fill.fore_color.rgb = PALE
    shape.line.color.rgb = RGBColor(223, 229, 240)
    textbox(slide, x+.22, y+.22, w-.44, .45, head, 17, NAVY, True)
    textbox(slide, x+.22, y+.78, w-.44, h-.94, body, 13, MUTED)

def logo_mark(slide, x, y, size):
    block = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(x), Inches(y), Inches(size), Inches(size))
    block.fill.solid(); block.fill.fore_color.rgb = NAVY; block.line.fill.background()
    screen = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(x+size*.23), Inches(y+size*.2), Inches(size*.54), Inches(size*.43))
    screen.fill.solid(); screen.fill.fore_color.rgb = WHITE; screen.line.fill.background()
    base_shape = slide.shapes.add_shape(MSO_SHAPE.TRAPEZOID, Inches(x+size*.16), Inches(y+size*.65), Inches(size*.68), Inches(size*.13))
    base_shape.fill.solid(); base_shape.fill.fore_color.rgb = GOLD; base_shape.line.fill.background()
    for sx, sy, sw, sh in [(x+size*.37,y+size*.43,size*.12,size*.12),(x+size*.49,y+size*.36,size*.13,size*.12),(x+size*.62,y+size*.25,size*.08,size*.12)]:
        dot = slide.shapes.add_shape(MSO_SHAPE.OVAL, Inches(sx), Inches(sy), Inches(sw), Inches(sh))
        dot.fill.solid(); dot.fill.fore_color.rgb = GOLD; dot.line.fill.background()

# 1. Opening
slide = prs.slides.add_slide(prs.slide_layouts[6])
slide.background.fill.solid(); slide.background.fill.fore_color.rgb = NAVY
logo_mark(slide, 9.0, 1.25, 2.65)
textbox(slide, .9, .9, 6.7, .35, "ART205 WEBSITE PROJECT  ·  GROUP 1", 12, GOLD, True)
textbox(slide, .9, 1.7, 7.4, 1.8, "DIGITAL TRAINING\nCENTER", 37, WHITE, True)
textbox(slide, .95, 4.0, 6.7, .7, "A practical digital skills learning concept", 20, WHITE)
textbox(slide, .95, 5.65, 7.4, .7, "KELVIN PADMORE  ·  11509\nSIMEON J. JOE  ·  11735  ·  [ADD MEMBERS]", 13, GOLD, True)

# 2. Purpose and planning
slide = base("The idea and the audience")
bullet_list(slide, .9, 2.0, 7.3, [
    "A sample local learning center for useful computer and creative skills",
    "Designed for learners who want clear steps and hands-on practice",
    "Four pages guide visitors from the concept to sample courses and an inquiry",
])
card(slide, 8.7, 2.05, 3.55, 2.9, "Project boundary", "Course details, schedules and contacts are examples. The site is a class demonstration, not a live service.")

# 3. Brand
slide = base("A simple visual identity")
logo_mark(slide, .95, 2.1, 2.4)
card(slide, 4.0, 2.05, 3.4, 1.45, "NAVY  #162E70", "Trust, clarity and structure")
card(slide, 7.75, 2.05, 3.4, 1.45, "GOLD  #F3B73F", "Energy and a highlight for calls to action")
bullet_list(slide, 4.05, 4.1, 7.2, ["Logo idea: open laptop with an upward arrow", "Original logo and learning illustration are local SVG files"], 17, .75)

# 4. Pages and content
slide = base("Four connected pages")
card(slide, .8, 2.05, 2.8, 2.7, "Home", "Introduction, featured courses and clear next steps")
card(slide, 3.85, 2.05, 2.8, 2.7, "About", "Purpose, teaching approach and team details")
card(slide, 6.9, 2.05, 2.8, 2.7, "Courses", "Six samples, a schedule and level comparison")
card(slide, 9.95, 2.05, 2.55, 2.7, "Contact", "Labelled demonstration-only inquiry form")
textbox(slide, .9, 5.45, 11.4, .65, "The navigation, logo and stylesheet work from local files, so the project can be shown offline.", 18, INK)

# 5. Build and responsiveness
slide = base("How the website is built")
bullet_list(slide, .95, 1.95, 7.6, [
    "Semantic HTML pages share one external CSS file",
    "Flexbox aligns navigation and action groups; Grid arranges cards",
    "A mobile media query changes the layout for narrow screens",
    "Type, class and ID selectors plus descendant and child combinators",
    "Hover and keyboard focus styles help visitors follow interactions",
], 16, .77)
card(slide, 9.0, 2.2, 3.1, 2.9, "Responsive check", "Resize the browser. The course grid moves from three columns to two and then one.")

# 6. Interaction and form
slide = base("Small interactions, clear limits")
card(slide, .9, 2.0, 3.6, 2.8, "Course requirements", "A JavaScript button reveals and hides the sample requirements panel.")
card(slide, 4.85, 2.0, 3.6, 2.8, "Mobile navigation", "A button opens and closes the menu on small screens.")
card(slide, 8.8, 2.0, 3.6, 2.8, "Inquiry form", "The labelled form validates locally and confirms that nothing is sent or saved.")
textbox(slide, .95, 5.35, 11.2, .65, "Live demo: open index.html, visit Courses, toggle the panel, then preview the form.", 17, NAVY, True)

# 7. Team, learning and next steps
slide = base("Teamwork, evidence and next steps")
bullet_list(slide, .95, 1.9, 11.3, [
    "Complete member names and IDs, then record each person’s actual contribution",
    "Add two hand-drawn page sketches, two student-created visuals and five local observations",
    "Capture three real progress screenshots and list all sources consulted",
    "Each member should explain their work and make a small live code change",
    "Submission date in the instruction: November 25, 2026. Presentation: November 28 at 10:00 AM",
], 16, .78)
textbox(slide, .95, 6.12, 11.4, .45, "Questions and offline website demonstration", 18, NAVY, True)

prs.save(ROOT / "ART205_Group1_Presentation.pptx")
