
import os
import re
import time
from io import BytesIO
from pathlib import Path
from xml.sax.saxutils import escape

import streamlit as st
from dotenv import load_dotenv
from google import genai
from google.genai import errors

import arabic_reshaper
from bidi.algorithm import get_display

from PIL import Image as PILImage

from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.enums import TA_CENTER, TA_LEFT, TA_RIGHT
from reportlab.lib.units import mm
from reportlab.platypus import (
    SimpleDocTemplate, Paragraph, Spacer,
    HRFlowable, Table, TableStyle, Image as RLImage,
    KeepTogether
)

from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.section import WD_SECTION
from docx.oxml import OxmlElement
from docx.oxml.ns import qn


# =====================================================
# CONFIGURATION
# =====================================================

load_dotenv()

st.set_page_config(
    page_title="Saso | Build a Career That Stands Out",
    page_icon="✨",
    layout="wide",
    initial_sidebar_state="collapsed"
)

GOLD = "#E5B642"
DARK = "#050505"


# =====================================================
# SASO DESIGN
# =====================================================

st.markdown("""
<style>

.stApp {
    background:
        radial-gradient(ellipse at 50% 5%, #211600, transparent 42%),
        #030303 !important;
    color: white;
}

header[data-testid="stHeader"] {
    background: transparent;
}

.block-container {
    max-width: 1150px;
    padding: 2rem 3rem 3rem;
}

/* FULL PAGE FRAME */

.saso-frame {
    position: fixed;
    inset: 13px;
    z-index: 999990;
    pointer-events: none;
    border: 2px solid #e8b947;
    border-radius: 30px;
    box-shadow:
        0 0 8px #ffd76a,
        0 0 18px #e9a900,
        0 0 35px rgba(255,183,0,.5),
        inset 0 0 15px rgba(255,193,44,.2);
}

.saso-frame::before {
    content: "";
    position: absolute;
    inset: 7px;
    border: 1px solid #f5d47d;
    border-radius: 23px;
    box-shadow:
        0 0 10px rgba(255,203,74,.7),
        inset 0 0 8px rgba(255,203,74,.2);
}

.frame-corner {
    position: fixed;
    z-index: 999999;
    pointer-events: none;
    color: #ffd66b;
    font-family: Georgia, serif;
    font-size: 45px;
    text-shadow:
        0 0 7px white,
        0 0 15px #ffbf26,
        0 0 30px #e7a300;
}

.corner-tl {top: 17px;left: 20px;}
.corner-tr {top: 17px;right: 20px;transform:scaleX(-1);}
.corner-bl {bottom: 17px;left: 20px;transform:scaleY(-1);}
.corner-br {bottom: 17px;right: 20px;transform:scale(-1);}

/* LOGO */

.saso-header {
    text-align: center;
    padding: 30px 10px;
}

.saso-logo {
    font-family: Georgia, serif;
    font-size: clamp(85px, 13vw, 145px);
    font-weight: 900;
    letter-spacing: 5px;
    line-height: 1.1;
    background: linear-gradient(
        180deg,
        #fff8d8,
        #ffe58b,
        #ffbf24,
        #fff3b1,
        #d88b00,
        #fff0a6
    );
    -webkit-background-clip: text;
    background-clip: text;
    color: transparent;
    filter:
        drop-shadow(0 0 8px #ffcf53)
        drop-shadow(0 0 20px #ffb000)
        drop-shadow(0 0 35px rgba(255,170,0,.8));
    animation: logoGlow 3s ease-in-out infinite alternate;
}

.saso-tagline {
    color: #fff8e7;
    font-family: Georgia, serif;
    font-size: clamp(13px, 2vw, 20px);
    letter-spacing: 2px;
    margin-top: 12px;
    text-shadow: 0 0 12px rgba(255,197,54,.5);
}

@keyframes logoGlow {
    from {
        filter: drop-shadow(0 0 8px #ffcf53)
                drop-shadow(0 0 18px #ffb000);
    }
    to {
        filter: drop-shadow(0 0 14px white)
                drop-shadow(0 0 30px #ffc400)
                drop-shadow(0 0 48px #ff9d00);
    }
}

/* HEADINGS */

h1, h2, h3, h4 {
    color: #ffdc80 !important;
    text-shadow: 0 0 10px rgba(255,183,0,.6);
}

.hero-title {
    text-align: center;
    color: #ffdb78;
    font-family: Georgia, serif;
    font-size: 34px;
    font-weight: bold;
    text-shadow: 0 0 15px #ffbd35;
    margin-bottom: 8px;
}

.hero-subtitle {
    text-align: center;
    color: #e8dcc2;
    margin-bottom: 30px;
}

/* FORM */

[data-testid="stForm"] {
    background:
        radial-gradient(ellipse at top, rgba(255,183,0,.08), transparent 60%),
        linear-gradient(145deg,#100d07,#070707);
    border: 2px solid #dca52f;
    border-radius: 26px;
    padding: 30px;
    box-shadow:
        0 0 10px #f1b832,
        0 0 25px rgba(255,181,0,.4),
        inset 0 0 22px rgba(255,189,44,.08);
}

/* INPUTS */

.stTextInput label,
.stTextArea label,
.stSelectbox label,
.stFileUploader label {
    color: #fff0c4 !important;
    font-weight: bold;
    text-shadow: 0 0 7px rgba(255,183,0,.4);
}

.stTextInput input,
.stTextArea textarea,
.stSelectbox div[data-baseweb="select"] {
    background: #100d08 !important;
    color: white !important;
    border: 1px solid #d59c24 !important;
    border-radius: 12px !important;
    box-shadow: 0 0 7px rgba(255,181,0,.3);
    transition: .25s ease;
}

.stTextInput input:hover,
.stTextArea textarea:hover,
.stSelectbox div[data-baseweb="select"]:hover {
    border-color: #ffda70 !important;
    box-shadow: 0 0 12px #ffbd35;
}

.stTextInput input:focus,
.stTextArea textarea:focus {
    border-color: #fff0a8 !important;
    box-shadow:
        0 0 10px #fff0a8,
        0 0 22px #ffbd35 !important;
}

/* SELECT LANGUAGE */

.stSelectbox div[data-baseweb="select"] {
    border: 2px solid #e5b642 !important;
    box-shadow:
        0 0 8px rgba(255,193,44,.6),
        0 0 18px rgba(255,193,44,.25);
}

/* FILE UPLOADER */

[data-testid="stFileUploader"] section {
    background: #100d08;
    border: 2px dashed #dca52f;
    border-radius: 16px;
    box-shadow: 0 0 12px rgba(255,181,0,.3);
}

[data-testid="stFileUploader"] section:hover {
    border-color: #ffe08a;
    box-shadow: 0 0 20px #e5b642;
}

/* BUTTONS */

.stFormSubmitButton > button,
.stDownloadButton > button {
    background: linear-gradient(
        180deg,#fff0a9,#ffd45b,#e7a000,#bd7900
    ) !important;
    color: #120b00 !important;
    border: 2px solid #fff0a8 !important;
    border-radius: 15px !important;
    min-height: 55px;
    font-weight: 900 !important;
    letter-spacing: .5px;
    box-shadow:
        0 0 9px #ffdf79,
        0 0 20px #ffb700,
        0 0 35px rgba(255,169,0,.6);
    transition: all .3s ease !important;
}

.stFormSubmitButton > button:hover,
.stDownloadButton > button:hover {
    transform: translateY(-3px);
    box-shadow:
        0 0 15px white,
        0 0 30px #ffd24f,
        0 0 50px #ff9d00;
}

/* RESULT */

.result-title {
    text-align: center;
    color: #ffe18a;
    font-size: 28px;
    font-weight: bold;
    margin: 25px 0;
    text-shadow: 0 0 15px #e5a000;
}

hr {
    border-color: rgba(255,195,55,.5) !important;
    box-shadow: 0 0 8px rgba(255,183,0,.4);
}

[data-testid="stAlert"] {
    border: 1px solid #c99528;
    border-radius: 13px;
    box-shadow: 0 0 13px rgba(255,183,0,.25);
}

.saso-footer {
    text-align: center;
    color: #e8bd57;
    letter-spacing: 3px;
    font-size: 12px;
    padding: 35px 5px;
    text-shadow: 0 0 12px #e7a000;
}

@media(max-width:768px) {
    .block-container {
        padding: 1.5rem 1rem;
    }
    [data-testid="stForm"] {
        padding: 17px;
    }
    .saso-frame {
        inset: 7px;
        border-radius: 18px;
    }
    .frame-corner {
        font-size: 30px;
    }
    .corner-tl {top:10px;left:11px;}
    .corner-tr {top:10px;right:11px;}
    .corner-bl {bottom:10px;left:11px;}
    .corner-br {bottom:10px;right:11px;}
    .saso-tagline {
        letter-spacing: 1px;
    }
}

</style>
""", unsafe_allow_html=True)


st.markdown("""
<div class="saso-frame"></div>
<div class="frame-corner corner-tl">❦</div>
<div class="frame-corner corner-tr">❦</div>
<div class="frame-corner corner-bl">❦</div>
<div class="frame-corner corner-br">❦</div>

<div class="saso-header">
    <div class="saso-logo">Saso</div>
    <div class="saso-tagline">
        Build a Career That Stands Out
    </div>
    <div style="
        color:#ffcf53;
        font-size:22px;
        margin-top:20px;
        text-shadow:0 0 15px #ffb700;">
        ━━━━━━━ ✦ ━━━━━━━
    </div>
</div>
""", unsafe_allow_html=True)


# =====================================================
# ARABIC TEXT HELPERS
# =====================================================

def is_arabic(text):
    return bool(re.search(r"[\u0600-\u06FF]", text))


def shape_arabic(text):
    if is_arabic(text):
        return get_display(arabic_reshaper.reshape(text))
    return text


def get_arabic_font():
    candidates = [
        Path("C:/Windows/Fonts/arial.ttf"),
        Path("C:/Windows/Fonts/tahoma.ttf"),
        Path("C:/Windows/Fonts/Arial.ttf"),
        Path("/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf"),
    ]

    for font in candidates:
        if font.exists():
            return str(font)

    return None


# =====================================================
# TEXT PROCESSING
# =====================================================

def split_documents(result):

    parts = re.split(
        r"(?im)^\s*COVER LETTER\s*$",
        result,
        maxsplit=1
    )

    if len(parts) == 2:
        cv = re.sub(
            r"(?im)^\s*PROFESSIONAL CV\s*$",
            "",
            parts[0]
        ).strip()

        letter = parts[1].strip()

    else:
        cv = result.strip()
        letter = ""

    return cv, letter


def get_sections(content):

    headings = [
        "PROFESSIONAL SUMMARY",
        "CAREER OBJECTIVE",
        "EDUCATION",
        "WORK EXPERIENCE",
        "SKILLS",
        "CERTIFICATIONS",
        "PROJECTS",
        "ACHIEVEMENTS",
        "CONTACT",
        "SUMMARY",
        "الملخص المهني",
        "الهدف الوظيفي",
        "التعليم",
        "الخبرات العملية",
        "الخبرة العملية",
        "المهارات",
        "الشهادات",
        "المشروعات",
        "الإنجازات",
        "معلومات التواصل"
    ]

    sections = []
    current_title = "PROFILE"
    current_lines = []

    for line in content.splitlines():

        clean = line.strip().strip("*# ").rstrip(":").strip()

        if not clean:
            continue

        if any(clean.upper() == h.upper() for h in headings):

            if current_lines:
                sections.append((current_title, current_lines))

            current_title = clean
            current_lines = []

        else:
            current_lines.append(clean)

    if current_lines:
        sections.append((current_title, current_lines))

    return sections


# =====================================================
# PDF GENERATOR
# =====================================================

def create_pdf(content, document_type="CV", photo=None, candidate_name=""):

    buffer = BytesIO()

    doc = SimpleDocTemplate(
        buffer,
        pagesize=A4,
        rightMargin=17 * mm,
        leftMargin=17 * mm,
        topMargin=17 * mm,
        bottomMargin=19 * mm
    )

    styles = getSampleStyleSheet()

    gold = colors.HexColor("#B38B2E")
    dark = colors.HexColor("#252525")
    grey = colors.HexColor("#666666")

    font_path = get_arabic_font()

    title_style = ParagraphStyle(
        "TitleSaso",
        parent=styles["Title"],
        fontName="Helvetica-Bold",
        fontSize=23,
        leading=29,
        textColor=gold,
        alignment=TA_CENTER,
        spaceAfter=7
    )

    heading_style = ParagraphStyle(
        "HeadingSaso",
        parent=styles["Heading2"],
        fontName="Helvetica-Bold",
        fontSize=11,
        leading=16,
        textColor=gold,
        spaceBefore=10,
        spaceAfter=5
    )

    body_style = ParagraphStyle(
        "BodySaso",
        parent=styles["BodyText"],
        fontName="Helvetica",
        fontSize=9,
        leading=14,
        textColor=dark,
        spaceAfter=4
    )

    if font_path:
        from reportlab.pdfbase import pdfmetrics
        from reportlab.pdfbase.ttfonts import TTFont

        try:
            if "SasoArabic" not in pdfmetrics.getRegisteredFontNames():
                pdfmetrics.registerFont(
                    TTFont("SasoArabic", font_path)
                )

            title_style.fontName = "SasoArabic"
            heading_style.fontName = "SasoArabic"
            body_style.fontName = "SasoArabic"

        except Exception:
            pass

    story = []

    title = (
        "PROFESSIONAL CV"
        if document_type == "CV"
        else "COVER LETTER"
    )

    story.append(Paragraph(title, title_style))

    if candidate_name:
        story.append(
            Paragraph(
                escape(shape_arabic(candidate_name)),
                ParagraphStyle(
                    "Name",
                    parent=body_style,
                    alignment=TA_CENTER,
                    fontSize=14,
                    textColor=dark,
                    spaceAfter=8
                )
            )
        )

    story.append(HRFlowable(
        width="100%",
        thickness=2,
        color=gold,
        spaceAfter=12
    ))

    if photo and document_type == "CV":

        try:
            image_buffer = BytesIO(photo)

            portrait = RLImage(
                image_buffer,
                width=30 * mm,
                height=35 * mm
            )

            portrait.hAlign = "RIGHT"

            story.append(portrait)
            story.append(Spacer(1, 8))

        except Exception:
            pass

    sections = get_sections(content)

    if not sections:
        sections = [("PROFILE", content.splitlines())]

    for section_title, lines in sections:

        shaped_title = shape_arabic(section_title)

        story.append(KeepTogether([
            Paragraph(
                escape(shaped_title),
                heading_style
            ),
            HRFlowable(
                width="100%",
                thickness=0.8,
                color=gold,
                spaceAfter=6
            )
        ]))

        for line in lines:

            if not line.strip():
                continue

            shaped = shape_arabic(line)

            story.append(
                Paragraph(
                    escape(shaped),
                    body_style
                )
            )

        story.append(Spacer(1, 5))

    def add_footer(canvas, document):

        canvas.saveState()

        canvas.setStrokeColor(gold)
        canvas.setLineWidth(1)

        canvas.line(
            17 * mm,
            14 * mm,
            193 * mm,
            14 * mm
        )

        canvas.setFont("Helvetica", 8)
        canvas.setFillColor(grey)

        canvas.drawCentredString(
            A4[0] / 2,
            9 * mm,
            f"SASO | Page {document.page}"
        )

        canvas.restoreState()

    doc.build(
        story,
        onFirstPage=add_footer,
        onLaterPages=add_footer
    )

    buffer.seek(0)

    return buffer.getvalue()


# =====================================================
# WORD GENERATOR
# =====================================================

def set_paragraph_rtl(paragraph):

    pPr = paragraph._p.get_or_add_pPr()

    bidi = OxmlElement("w:bidi")
    pPr.append(bidi)

    paragraph.alignment = WD_ALIGN_PARAGRAPH.RIGHT


def set_cell_shading(cell, fill):

    tcPr = cell._tc.get_or_add_tcPr()

    shading = OxmlElement("w:shd")
    shading.set(qn("w:fill"), fill)

    tcPr.append(shading)


def create_word(content, document_type="CV", photo=None, candidate_name=""):

    document = Document()

    section = document.sections[0]

    section.top_margin = Inches(0.65)
    section.bottom_margin = Inches(0.65)
    section.left_margin = Inches(0.7)
    section.right_margin = Inches(0.7)

    gold_rgb = RGBColor(184, 137, 37)

    title = document.add_paragraph()
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER

    run = title.add_run(
        "PROFESSIONAL CV"
        if document_type == "CV"
        else "COVER LETTER"
    )

    run.bold = True
    run.font.size = Pt(23)
    run.font.color.rgb = gold_rgb

    if candidate_name:
        name_p = document.add_paragraph()
        name_p.alignment = WD_ALIGN_PARAGRAPH.CENTER

        name_run = name_p.add_run(candidate_name)
        name_run.bold = True
        name_run.font.size = Pt(16)

    if photo and document_type == "CV":

        try:
            document.add_picture(
                BytesIO(photo),
                width=Inches(1.15),
                height=Inches(1.3)
            )
        except Exception:
            pass

    sections = get_sections(content)

    for section_title, lines in sections:

        heading = document.add_paragraph()

        if any("\u0600" <= c <= "\u06FF" for c in section_title):
            set_paragraph_rtl(heading)

        run = heading.add_run(section_title)
        run.bold = True
        run.font.size = Pt(13)
        run.font.color.rgb = gold_rgb

        for line in lines:

            p = document.add_paragraph()

            if is_arabic(line):
                set_paragraph_rtl(p)
                p.alignment = WD_ALIGN_PARAGRAPH.RIGHT

            run = p.add_run(line)
            run.font.size = Pt(10)

    buffer = BytesIO()

    document.save(buffer)
    buffer.seek(0)

    return buffer.getvalue()


# =====================================================
# GEMINI GENERATOR
# =====================================================

def generate_documents(prompt, model_name):

    api_key = os.getenv("GEMINI_API_KEY")

    client = genai.Client(api_key=api_key)

    for attempt in range(3):

        try:

            response = client.models.generate_content(
                model=model_name,
                contents=prompt,
                config={
                    "temperature": 0.2,
                    "max_output_tokens": 3000
                }
            )

            if not response.text:
                raise ValueError("Empty Gemini response.")

            return response.text

        except errors.ServerError:

            if attempt == 2:
                raise

            time.sleep(3 * (attempt + 1))


# =====================================================
# MAIN PAGE
# =====================================================

st.markdown("""
<div class="hero-title">✦ CREATE YOUR FUTURE ✦</div>
<div class="hero-subtitle">
Create your professional CV and cover letter with AI
</div>
""", unsafe_allow_html=True)


with st.form("resume_form"):

    st.subheader("✦ Personal Information")

    col1, col2 = st.columns(2)

    with col1:
        name = st.text_input(
            "Full Name / الاسم بالكامل",
            placeholder="Enter your name..."
        )

        phone = st.text_input(
            "Phone Number / رقم الهاتف",
            placeholder="Enter your phone..."
        )

    with col2:
        email = st.text_input(
            "Email Address / البريد الإلكتروني",
            placeholder="Enter your email..."
        )

        location = st.text_input(
            "Location / العنوان",
            placeholder="Enter your location..."
        )

    st.divider()

    st.subheader("✦ Career Information")

    job = st.text_input(
        "Target Job / الوظيفة المستهدفة",
        placeholder="Enter your target job..."
    )

    experience = st.text_area(
        "Work Experience / الخبرات العملية",
        placeholder="Describe your previous jobs...",
        height=110
    )

    education = st.text_area(
        "Education / التعليم",
        placeholder="Your education and qualifications...",
        height=100
    )

    skills = st.text_area(
        "Skills / المهارات",
        placeholder="Your professional skills...",
        height=100
    )

    language = st.selectbox(
        "Output Language / لغة المستند",
        ["English", "Arabic"]
    )

    st.divider()

    st.subheader("✦ Personal Photo / الصورة الشخصية")

    photo_file = st.file_uploader(
        "Upload your professional photo",
        type=["png", "jpg", "jpeg"],
        help="Your photo will be included in the CV PDF and Word file."
    )

    submitted = st.form_submit_button(
        "✦ GENERATE MY CV & COVER LETTER ✦",
        use_container_width=True
    )


# =====================================================
# PHOTO PREVIEW
# =====================================================

if photo_file:

    st.image(
        photo_file,
        caption="Your CV Photo Preview",
        width=160
    )


# =====================================================
# GENERATION
# =====================================================

if submitted:

    api_key = os.getenv("GEMINI_API_KEY")

    model_name = os.getenv(
        "GEMINI_MODEL",
        "gemini-3.8-flash"
    )

    if not api_key:

        st.error("Gemini API key is missing.")

    elif not name.strip() or not job.strip():

        st.warning("Please enter your name and target job.")

    else:

        photo_bytes = None

        if photo_file:

            try:

                image = PILImage.open(photo_file)
                image = image.convert("RGB")

                image.thumbnail((900, 1100))

                image_buffer = BytesIO()

                image.save(
                    image_buffer,
                    format="JPEG",
                    quality=90
                )

                photo_bytes = image_buffer.getvalue()

            except Exception as e:

                st.error(f"Invalid photo: {e}")
                st.stop()

        prompt = f"""
Create a professional CV and tailored cover letter.

Output language: {language}

Candidate:
Name: {name}
Email: {email}
Phone: {phone}
Location: {location}
Target Job: {job}

Work Experience:
{experience}

Education:
{education}

Skills:
{skills}

Important rules:
- Use only the information provided.
- Never invent qualifications, experience or employers.
- Do not assume the candidate has a degree.
- If a section has no information, omit it.
- Create a professional summary based on available facts.
- Tailor the cover letter to the target job.
- Keep the writing professional and concise.

Return exactly two sections:

PROFESSIONAL CV

Include suitable headings such as:
PROFESSIONAL SUMMARY
WORK EXPERIENCE
EDUCATION
SKILLS
CERTIFICATIONS
PROJECTS

COVER LETTER

Return clean text with clear section headings.
"""

        try:

            with st.spinner("✨ Creating your professional documents..."):

                result = generate_documents(
                    prompt,
                    model_name
                )

                st.session_state["result"] = result
                st.session_state["photo_bytes"] = photo_bytes
                st.session_state["candidate_name"] = name

            st.success("✨ Your documents have been generated successfully!")

        except errors.ServerError as e:

            st.error(
                "Gemini server is currently unavailable. Please try again later."
            )

            st.caption(str(e))

        except Exception as e:

            st.error(f"Generation failed: {e}")


# =====================================================
# RESULTS
# =====================================================

if "result" in st.session_state:

    result = st.session_state["result"]

    photo_bytes = st.session_state.get("photo_bytes")
    candidate_name = st.session_state.get("candidate_name", "")

    st.divider()

    st.markdown(
        '<div class="result-title">✦ YOUR DOCUMENTS ✦</div>',
        unsafe_allow_html=True
    )

    st.text_area(
        "Document Preview",
        value=result,
        height=400
    )

    cv_content, letter_content = split_documents(result)

    st.markdown(
        '<div class="result-title">✦ DOWNLOAD YOUR DOCUMENTS ✦</div>',
        unsafe_allow_html=True
    )

    col1, col2 = st.columns(2)

    with col1:

        st.download_button(
            "⬇ DOWNLOAD CV PDF",
            data=create_pdf(
                cv_content,
                "CV",
                photo_bytes,
                candidate_name
            ),
            file_name="Saso_Professional_CV.pdf",
            mime="application/pdf",
            use_container_width=True
        )

    with col2:

        if letter_content:

            st.download_button(
                "⬇ DOWNLOAD COVER LETTER PDF",
                data=create_pdf(
                    letter_content,
                    "COVER LETTER",
                    None,
                    candidate_name
                ),
                file_name="Saso_Cover_Letter.pdf",
                mime="application/pdf",
                use_container_width=True
            )

        else:

            st.info("Cover letter section was not detected.")

    st.divider()

    st.markdown(
        '<div class="result-title">✦ EDITABLE WORD FILES ✦</div>',
        unsafe_allow_html=True
    )

    word1, word2 = st.columns(2)

    with word1:

        st.download_button(
            "⬇ DOWNLOAD CV WORD",
            data=create_word(
                cv_content,
                "CV",
                photo_bytes,
                candidate_name
            ),
            file_name="Saso_Professional_CV.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            use_container_width=True
        )

    with word2:

        if letter_content:

            st.download_button(
                "⬇ DOWNLOAD COVER LETTER WORD",
                data=create_word(
                    letter_content,
                    "COVER LETTER",
                    None,
                    candidate_name
                ),
                file_name="Saso_Cover_Letter.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                use_container_width=True
            )


# =====================================================
# FOOTER
# =====================================================

st.markdown("""
<div style="
    text-align:center;
    color:#e8bd57;
    font-size:19px;
    margin-top:35px;
    text-shadow:0 0 12px #e7a000;">
    ━━━━━━━ ✦ ━━━━━━━
</div>

<div class="saso-footer">
SASO ✦ BUILD A CAREER THAT STANDS OUT
</div>
""", unsafe_allow_html=True)
