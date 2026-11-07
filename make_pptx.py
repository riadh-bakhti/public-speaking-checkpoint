# -*- coding: utf-8 -*-
"""
Generate the GOMYCODE Public Speaking checkpoint presentation:
"Do social media networks strengthen or degrade human relationships?"
Position defended: social media networks STRENGTHEN human relationships.
"""

from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE

# ----------------------------------------------------------------------------
# Theme
# ----------------------------------------------------------------------------
BG      = RGBColor(0x0E, 0x1B, 0x2A)
CARD    = RGBColor(0x16, 0x28, 0x3D)
TEAL    = RGBColor(0x00, 0xD1, 0xB2)
CORAL   = RGBColor(0xFF, 0x7A, 0x59)
BLUE    = RGBColor(0x4D, 0xA3, 0xFF)
GOLD    = RGBColor(0xFF, 0xC8, 0x57)
WHITE   = RGBColor(0xF5, 0xF8, 0xFC)
MUTED   = RGBColor(0x9F, 0xB3, 0xC8)

FONT = "Segoe UI"

prs = Presentation()
prs.slide_width = Inches(13.333)
prs.slide_height = Inches(7.5)
BLANK = prs.slide_layouts[6]


# ----------------------------------------------------------------------------
# Helpers
# ----------------------------------------------------------------------------
def new_slide():
    slide = prs.slides.add_slide(BLANK)
    fill = slide.background.fill
    fill.solid()
    fill.fore_color.rgb = BG
    return slide


def add_rect(slide, x, y, w, h, color, shape=MSO_SHAPE.RECTANGLE, radius=None):
    shp = slide.shapes.add_shape(shape, Inches(x), Inches(y), Inches(w), Inches(h))
    shp.fill.solid()
    shp.fill.fore_color.rgb = color
    shp.line.fill.background()
    if radius is not None and shape == MSO_SHAPE.ROUNDED_RECTANGLE:
        try:
            shp.adjustments[0] = radius
        except Exception:
            pass
    return shp


def add_text(slide, text, x, y, w, h, size=20, bold=False, color=WHITE,
             align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.TOP, italic=False,
             line_spacing=1.0):
    tb = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    tf = tb.text_frame
    tf.word_wrap = True
    tf.vertical_anchor = anchor
    tf.margin_left = 0
    tf.margin_right = 0
    tf.margin_top = 0
    tf.margin_bottom = 0
    p = tf.paragraphs[0]
    p.alignment = align
    p.line_spacing = line_spacing
    run = p.add_run()
    run.text = text
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.italic = italic
    run.font.color.rgb = color
    run.font.name = FONT
    return tb


def add_bullets(slide, items, x, y, w, h, size=19, color=WHITE, accent=TEAL,
                space_after=14, line_spacing=1.05):
    tb = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    tf = tb.text_frame
    tf.word_wrap = True
    tf.margin_left = 0
    tf.margin_right = 0
    tf.margin_top = 0
    tf.margin_bottom = 0
    for i, item in enumerate(items):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.space_after = Pt(space_after)
        p.line_spacing = line_spacing
        r1 = p.add_run()
        r1.text = "▸  "
        r1.font.size = Pt(size)
        r1.font.bold = True
        r1.font.color.rgb = accent
        r1.font.name = FONT
        r2 = p.add_run()
        r2.text = item
        r2.font.size = Pt(size)
        r2.font.color.rgb = color
        r2.font.name = FONT
    return tb


def add_footer(slide, page):
    add_rect(slide, 0.7, 6.95, 11.93, 0.02, RGBColor(0x24, 0x3A, 0x52))
    add_text(slide, "Riadh Bakhti  ·  Public Speaking Checkpoint",
             0.7, 7.03, 8, 0.3, size=10, color=MUTED)
    add_text(slide, str(page), 12.0, 7.03, 0.63, 0.3, size=10,
             color=MUTED, align=PP_ALIGN.RIGHT)


def kicker_label(slide, text, color):
    add_text(slide, text.upper(), 0.75, 0.62, 8, 0.35, size=13, bold=True,
             color=color)


def slide_title(slide, text, size=38, y=1.02, color=WHITE):
    add_text(slide, text, 0.75, y, 11.8, 1.2, size=size, bold=True, color=color)


def notes(slide, text):
    slide.notes_slide.notes_text_frame.text = text


# ----------------------------------------------------------------------------
# 1. Title slide
# ----------------------------------------------------------------------------
s = new_slide()
add_rect(s, 0, 0, 13.333, 0.12, TEAL)
add_rect(s, 0.75, 1.15, 1.5, 0.08, CORAL)
add_text(s, "PUBLIC SPEAKING CHECKPOINT", 0.75, 0.75, 8, 0.4, size=14,
         bold=True, color=TEAL)
add_text(s, "Do Social Media Networks\nStrengthen or Degrade\nHuman Relationships?",
         0.75, 1.5, 9.6, 3.0, size=44, bold=True, color=WHITE, line_spacing=0.95)
add_rect(s, 0.78, 4.72, 6.4, 0.06, TEAL)
add_text(s, "Position defended:  social media networks STRENGTHEN human relationships",
         0.75, 4.95, 11, 0.5, size=20, bold=True, color=TEAL)
add_text(s, "Riadh Bakhti  ·  Full-stack JavaScript Developer  ·  GOMYCODE",
         0.75, 6.25, 9, 0.4, size=16, color=MUTED)
add_rect(s, 10.6, 1.5, 2.0, 2.0, CARD, MSO_SHAPE.OVAL)
add_text(s, "9", 10.6, 2.05, 2.0, 1.0, size=64, bold=True, color=TEAL,
         align=PP_ALIGN.CENTER)
add_text(s, "arguments", 10.6, 2.95, 2.0, 0.4, size=14, color=MUTED,
         align=PP_ALIGN.CENTER)
notes(s, "Good morning everyone. Today I will defend a clear position: social media networks strengthen human relationships. I will support it with 3 factual, 3 emotional and 3 rational arguments, and I will answer the main objection.")


# ----------------------------------------------------------------------------
# 2. Structure of the message
# ----------------------------------------------------------------------------
s = new_slide()
kicker_label(s, "Structure of the message", TEAL)
slide_title(s, "A clear roadmap in four movements")
parts = [
    ("01", "Factual arguments", "What the data and observable facts show", TEAL),
    ("02", "Emotional arguments", "What people feel and live every day", CORAL),
    ("03", "Rational arguments", "What logic and evidence demonstrate", BLUE),
    ("04", "Objection & conclusion", "Risks acknowledged, position reaffirmed", GOLD),
]
for i, (num, title, desc, color) in enumerate(parts):
    x = 0.75 + (i % 2) * 6.05
    y = 2.65 + (i // 2) * 2.05
    add_rect(s, x, y, 5.7, 1.75, CARD, MSO_SHAPE.ROUNDED_RECTANGLE, 0.08)
    add_text(s, num, x + 0.3, y + 0.3, 0.9, 0.7, size=34, bold=True, color=color)
    add_text(s, title, x + 1.25, y + 0.32, 4.2, 0.5, size=21, bold=True,
             color=WHITE)
    add_text(s, desc, x + 1.25, y + 0.85, 4.2, 0.6, size=14, color=MUTED)
add_footer(s, 2)
notes(s, "My message is structured in four movements: facts, emotions, reason, then the objection and conclusion.")


# ----------------------------------------------------------------------------
# 3. Thesis
# ----------------------------------------------------------------------------
s = new_slide()
kicker_label(s, "Thesis", TEAL)
slide_title(s, "Social media do not replace relationships — they extend them")
add_rect(s, 0.75, 2.7, 11.8, 1.9, CARD, MSO_SHAPE.ROUNDED_RECTANGLE, 0.06)
add_rect(s, 0.75, 2.7, 0.14, 1.9, TEAL)
add_text(s, "“A relationship is not weakened because it is partly digital;\nit is weakened when it is neglected.”",
         1.15, 3.0, 11.0, 1.3, size=24, italic=True, color=WHITE)
add_bullets(s, [
    "Social networks are a channel, not a substitute for human contact.",
    "They remove distance, time and cost barriers to staying connected.",
    "The real question is not the tool, but the intention with which we use it.",
], 0.75, 5.0, 11.6, 1.6, size=19, accent=TEAL)
add_footer(s, 3)
notes(s, "My thesis: the tool is neutral; it is the use that matters. And used intentionally, social media extend our relationships.")


# ----------------------------------------------------------------------------
# Section divider helper
# ----------------------------------------------------------------------------
def section_slide(number, title, subtitle, color, page):
    s = new_slide()
    add_rect(s, 0, 0, 0.18, 7.5, color)
    add_text(s, number, 1.0, 1.6, 4.0, 2.4, size=120, bold=True, color=color)
    add_text(s, title, 1.1, 4.0, 11, 0.9, size=40, bold=True, color=WHITE)
    add_text(s, subtitle, 1.1, 4.95, 11, 0.6, size=19, color=MUTED)
    add_text(s, str(page), 12.0, 7.03, 0.63, 0.3, size=10, color=MUTED,
             align=PP_ALIGN.RIGHT)
    return s


# ----------------------------------------------------------------------------
# Argument slide helper
# ----------------------------------------------------------------------------
def argument_slide(kicker, title, stat, stat_label, points, accent, page,
                   note_text=""):
    s = new_slide()
    add_rect(s, 0, 0, 13.333, 0.1, accent)
    kicker_label(s, kicker, accent)
    slide_title(s, title, size=34)
    add_rect(s, 0.75, 2.45, 3.55, 3.4, CARD, MSO_SHAPE.ROUNDED_RECTANGLE, 0.08)
    add_text(s, stat, 0.75, 2.8, 3.55, 1.2, size=52, bold=True, color=accent,
             align=PP_ALIGN.CENTER)
    add_text(s, stat_label, 1.05, 4.15, 2.95, 1.4, size=15, color=MUTED,
             align=PP_ALIGN.CENTER)
    add_bullets(s, points, 4.85, 2.55, 7.7, 3.4, size=18, accent=accent,
                space_after=16)
    add_footer(s, page)
    if note_text:
        notes(s, note_text)
    return s


# ----------------------------------------------------------------------------
# 4. Section — Factual
# ----------------------------------------------------------------------------
section_slide("01", "Factual arguments", "What the data and observable facts show", TEAL, 4)

# 4.1 Factual 1 — scale
argument_slide(
    "Factual argument 1 / 3",
    "Never before have so many people been able to connect",
    "5.2 B",
    "people use social media worldwide — about 64% of the global population (DataReportal, 2025)",
    [
        "More than five billion human beings can now reach each other instantly.",
        "The average user spends about 2h20 per day on social platforms — time dedicated to interaction.",
        "Each account is a door open to family, friends, communities and opportunities.",
    ],
    TEAL, 5,
    "Fact one: the scale is unprecedented. Five billion people connected — this is not a niche phenomenon, it is the new social infrastructure.",
)

# 4.2 Factual 2 — distance
argument_slide(
    "Factual argument 2 / 3",
    "Distance is no longer a barrier to maintaining a relationship",
    "281 M",
    "international migrants in the world (UN, 2020) — most of them stay in touch with family through social media",
    [
        "WhatsApp alone connects more than 2 billion users in over 180 countries.",
        "Video calls allow families separated by borders to see each other daily, for free.",
        "Students, expatriates and travellers keep a living link with their loved ones.",
    ],
    TEAL, 6,
    "Fact two: migration used to mean separation. Today, a video call keeps the relationship alive across continents.",
)

# 4.3 Factual 3 — communities
argument_slide(
    "Factual argument 3 / 3",
    "Communities of support now exist at global scale",
    "1.8 B",
    "people use Facebook groups every month — support, health, education and mutual aid communities",
    [
        "Rare-disease patients find others who understand their reality.",
        "Disaster response and mutual aid are organised in real time on social platforms.",
        "90% of young adults in the United States use social media (Pew Research Center, 2024).",
    ],
    TEAL, 7,
    "Fact three: communities are not only local. Online groups give isolated people access to support networks they could never find in their village or neighbourhood.",
)


# ----------------------------------------------------------------------------
# 5. Section — Emotional
# ----------------------------------------------------------------------------
section_slide("02", "Emotional arguments", "What people feel and live every day", CORAL, 8)

argument_slide(
    "Emotional argument 1 / 3",
    "Presence beyond distance: grandparents and grandchildren",
    "♥",
    "A daily video call can keep a grandparent inside the life of a child who lives far away",
    [
        "Seeing a smile, hearing a voice: the emotional bond is maintained, not replaced.",
        "Loneliness decreases when a simple call becomes a daily ritual.",
        "Children grow up knowing their grandparents — not as strangers, but as faces and voices.",
    ],
    CORAL, 9,
    "Emotion one: think of a grandmother who sees her grandson every evening on a screen. That is not a fake relationship — that is a preserved one.",
)

argument_slide(
    "Emotional argument 2 / 3",
    "Finding your people when you feel alone",
    "♥",
    "Support groups give a voice and a place to people who thought they were the only ones",
    [
        "A teenager struggling with anxiety finds a community that listens without judging.",
        "Patients share experiences, courage and practical advice.",
        "Belonging to a group reduces the feeling of isolation and restores hope.",
    ],
    CORAL, 10,
    "Emotion two: social media can be the first place where a person hears: 'you are not alone'. That experience changes lives.",
)

argument_slide(
    "Emotional argument 3 / 3",
    "Shared joy, empathy and collective emotion",
    "♥",
    "Millions of people celebrate, mourn and support each other together in real time",
    [
        "Birthdays, weddings and achievements are shared instantly with those who care.",
        "Collective empathy appears after tragedies: messages, donations, solidarity.",
        "Humour and memes create complicity and shared culture between generations.",
    ],
    CORAL, 11,
    "Emotion three: emotions are contagious online too. Joy shared is joy multiplied; sorrow shared is sorrow divided.",
)


# ----------------------------------------------------------------------------
# 6. Section — Rational
# ----------------------------------------------------------------------------
section_slide("03", "Rational arguments", "What logic and evidence demonstrate", BLUE, 12)

argument_slide(
    "Rational argument 1 / 3",
    "Lower cost and higher speed mean more maintained relationships",
    "≈ 0",
    "marginal cost to send a message anywhere in the world, instantly",
    [
        "A letter took days and money; a message takes a second and costs nothing.",
        "Weak ties — old colleagues, distant friends — survive instead of disappearing.",
        "According to Granovetter, weak ties are precisely the ones that create opportunities (jobs, collaborations).",
    ],
    BLUE, 13,
    "Rationally, relationships survive when maintaining them is cheap and easy. Social media make maintenance almost free.",
)

argument_slide(
    "Rational argument 2 / 3",
    "Inclusion: communication adapted to everyone",
    "24/7",
    "asynchronous communication that respects schedules, abilities and time zones",
    [
        "People with disabilities can participate without physical barriers.",
        "Introverted or socially anxious people can express themselves at their own pace.",
        "Families spread across time zones stay connected without forcing anyone to wake up at night.",
    ],
    BLUE, 14,
    "Rationally, a channel that adapts to the person is more inclusive than a channel that forces the person to adapt to it.",
)

argument_slide(
    "Rational argument 3 / 3",
    "Active use is correlated with greater social capital",
    "＋",
    "Research links active social-media use to increased bridging and bonding social capital (Ellison et al., 2007; Pew, 2011)",
    [
        "Active use (talking, commenting, organising) builds relationships; passive scrolling does not.",
        "Bridging capital: access to new information, opportunities and diverse networks.",
        "Bonding capital: stronger ties with family and close friends.",
    ],
    BLUE, 15,
    "Rationally, the scientific literature distinguishes active from passive use. Active use builds social capital — which is exactly the point.",
)


# ----------------------------------------------------------------------------
# 7. Objection & rebuttal
# ----------------------------------------------------------------------------
s = new_slide()
kicker_label(s, "Objection & rebuttal", GOLD)
slide_title(s, "“But social media isolate, compare and polarise people…”", size=32)
add_rect(s, 0.75, 2.5, 5.6, 3.7, CARD, MSO_SHAPE.ROUNDED_RECTANGLE, 0.07)
add_text(s, "THE OBJECTION", 1.1, 2.8, 4.9, 0.4, size=13, bold=True, color=CORAL)
add_bullets(s, [
    "Screen time replaces face-to-face time.",
    "Comparison and cyberbullying damage self-esteem.",
    "Echo chambers and misinformation divide society.",
], 1.1, 3.35, 4.9, 2.5, size=17, accent=CORAL, space_after=12)
add_rect(s, 6.95, 2.5, 5.6, 3.7, CARD, MSO_SHAPE.ROUNDED_RECTANGLE, 0.07)
add_text(s, "OUR ANSWER", 7.3, 2.8, 4.9, 0.4, size=13, bold=True, color=TEAL)
add_bullets(s, [
    "The risks are real — but they depend on usage, not on the tool itself.",
    "The solution is digital literacy, moderation and platform design.",
    "Abandoning the tool would also abandon its benefits: connection, support, opportunity.",
], 7.3, 3.35, 4.9, 2.5, size=16, accent=TEAL, space_after=10)
add_text(s, "A balanced conclusion:  social media strengthen relationships when used actively and intentionally.",
         0.75, 6.35, 11.8, 0.4, size=17, bold=True, color=GOLD)
add_footer(s, 16)
notes(s, "I acknowledge the objection: yes, there are real risks. But they come from how the tool is used. The answer is education and intentional use, not rejection.")


# ----------------------------------------------------------------------------
# 8. Conclusion
# ----------------------------------------------------------------------------
s = new_slide()
kicker_label(s, "Conclusion", TEAL)
slide_title(s, "Social media strengthen human relationships — when we use them to connect")
add_rect(s, 0.75, 2.55, 11.8, 0.03, TEAL)
concl = [
    ("FACTS", "5 billion people connected, families across borders, global support communities.", TEAL),
    ("EMOTIONS", "Presence, belonging, shared joy — relationships are lived, even through a screen.", CORAL),
    ("REASON", "Lower cost, more inclusion, greater social capital — when use is active.", BLUE),
]
for i, (label, text, color) in enumerate(concl):
    x = 0.75 + i * 4.0
    add_rect(s, x, 2.9, 3.7, 2.5, CARD, MSO_SHAPE.ROUNDED_RECTANGLE, 0.08)
    add_text(s, label, x + 0.3, 3.15, 3.1, 0.4, size=13, bold=True, color=color)
    add_text(s, text, x + 0.3, 3.65, 3.1, 1.6, size=15, color=WHITE)
add_rect(s, 0.75, 5.75, 11.8, 0.9, TEAL, MSO_SHAPE.ROUNDED_RECTANGLE, 0.12)
add_text(s, "Let us build relationships online that make us better humans offline.",
         1.05, 5.98, 11.2, 0.5, size=20, bold=True, color=BG, align=PP_ALIGN.CENTER)
add_footer(s, 17)
notes(s, "To conclude: facts, emotions and reason all converge. Social media strengthen relationships when we use them intentionally. Thank you.")


# ----------------------------------------------------------------------------
# 9. References
# ----------------------------------------------------------------------------
s = new_slide()
kicker_label(s, "References", MUTED)
slide_title(s, "Sources", size=34)
refs = [
    "DataReportal — Digital 2025: Global Overview Report.",
    "United Nations — International Migrant Stock 2020.",
    "Pew Research Center — Social Media Use in 2024.",
    "Ellison, Steinfield & Lampe (2007) — The Benefits of Facebook “Friends”: Social Capital and College Students.",
    "Granovetter, M. (1973) — The Strength of Weak Ties.",
    "Statista — WhatsApp: number of users and countries.",
]
add_bullets(s, refs, 0.75, 2.5, 11.6, 3.8, size=17, accent=TEAL, space_after=14)
add_footer(s, 18)
notes(s, "These are the main sources used to build the factual and rational arguments of this presentation.")

prs.save("Social-Media-and-Human-Relationships.pptx")
print("PPTX generated:", len(prs.slides._sldIdLst), "slides")
