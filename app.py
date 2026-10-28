
import io
import re
import html
from datetime import date

import pandas as pd
import streamlit as st
from pypdf import PdfReader
from docx import Document
from sklearn.feature_extraction.text import TfidfVectorizer
from sklearn.metrics.pairwise import cosine_similarity
from fpdf import FPDF


# =========================================================
# PAGE CONFIGURATION
# =========================================================

st.set_page_config(
    page_title="SkillBridge AI | Career Intelligence",
    page_icon="✦",
    layout="wide",
    initial_sidebar_state="collapsed",
)


# =========================================================
# DESIGN SYSTEM — IVORY, FOREST GREEN & CHAMPAGNE
# =========================================================

st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=DM+Sans:wght@400;500;600;700&family=Playfair+Display:wght@500;600;700&display=swap');

:root {
    --ivory: #F7F6F0;
    --paper: #FFFFFF;
    --forest: #193D32;
    --forest-dark: #102C24;
    --sage: #DCE7DD;
    --sage-light: #EDF2EB;
    --gold: #B49A62;
    --gold-light: #E7D9B5;
    --ink: #202B26;
    --muted: #748078;
    --line: #E4E6DD;
    --danger: #A7473D;
}

html, body, [class*="css"] {
    font-family: 'DM Sans', sans-serif;
}

.stApp {
    background: var(--ivory);
    color: var(--ink);
}

header[data-testid="stHeader"] {
    background: transparent;
}

#MainMenu, footer {
    visibility: hidden;
}

.block-container {
    max-width: 1180px;
    padding: 2rem 2.4rem 5rem;
}

/* Remove default divider clutter */
[data-testid="stVerticalBlock"] > div {
    gap: 0.8rem;
}

/* Brand header */
.brand-row {
    display: flex;
    align-items: center;
    justify-content: space-between;
    gap: 20px;
    padding: 4px 0 24px;
    border-bottom: 1px solid var(--line);
    margin-bottom: 26px;
}

.brand-lockup {
    display: flex;
    align-items: center;
    gap: 12px;
}

.brand-mark {
    width: 43px;
    height: 43px;
    display: flex;
    align-items: center;
    justify-content: center;
    border-radius: 14px;
    background: var(--forest);
    color: var(--gold-light);
    font-size: 25px;
    box-shadow: 0 5px 14px rgba(25, 61, 50, 0.12);
}

.brand-name {
    color: var(--forest-dark);
    font-size: 17px;
    font-weight: 700;
    letter-spacing: -0.5px;
}

.brand-subtitle {
    color: var(--muted);
    font-size: 11px;
    letter-spacing: 1.4px;
    text-transform: uppercase;
    margin-top: 3px;
}

.brand-status {
    border: 1px solid #D9E2D8;
    background: #EEF3EB;
    color: var(--forest);
    border-radius: 50px;
    padding: 8px 13px;
    font-size: 11px;
    font-weight: 700;
    letter-spacing: 0.7px;
}

/* Hero */
.hero {
    position: relative;
    overflow: hidden;
    border-radius: 25px;
    background:
        radial-gradient(ellipse at 90% 5%,
            rgba(180, 154, 98, 0.16), transparent 35%),
        linear-gradient(125deg, #193D32 0%, #214B3C 65%, #16382D 100%);
    padding: 54px 54px 46px;
    margin: 8px 0 34px;
    box-shadow: 0 18px 40px rgba(25, 61, 50, 0.10);
}

.hero::after {
    content: "✦";
    position: absolute;
    right: 7%;
    top: 10%;
    font-size: 180px;
    color: rgba(231, 217, 181, 0.055);
    line-height: 1;
    pointer-events: none;
}

.hero-eyebrow {
    color: #E2D2A7;
    font-size: 11px;
    font-weight: 700;
    letter-spacing: 2.3px;
    text-transform: uppercase;
    margin-bottom: 22px;
}

.hero h1 {
    position: relative;
    z-index: 1;
    max-width: 760px;
    font-family: 'Playfair Display', Georgia, serif;
    color: #FFFDF7;
    font-size: clamp(39px, 5.3vw, 66px);
    line-height: 1.12;
    font-weight: 500;
    letter-spacing: -1.8px;
    margin: 0 0 20px;
}

.hero h1 em {
    color: #E2D2A7;
    font-weight: 500;
}

.hero-copy {
    position: relative;
    z-index: 1;
    max-width: 610px;
    color: #D9E4DB;
    font-size: 14px;
    line-height: 1.9;
    margin-bottom: 30px;
}

.hero-features {
    position: relative;
    z-index: 1;
    display: flex;
    flex-wrap: wrap;
    align-items: center;
    gap: 10px;
}

.hero-feature {
    border: 1px solid rgba(231, 217, 181, 0.23);
    border-radius: 50px;
    padding: 9px 13px;
    color: #F2EEE0;
    font-size: 11px;
    background: rgba(255, 255, 255, 0.035);
}

/* Section headings */
.section-heading {
    display: flex;
    align-items: flex-start;
    gap: 16px;
    margin: 36px 0 16px;
}

.section-number {
    flex: 0 0 40px;
    width: 40px;
    height: 40px;
    display: flex;
    align-items: center;
    justify-content: center;
    background: var(--sage-light);
    color: var(--forest);
    border: 1px solid #DCE5D9;
    border-radius: 13px;
    font-size: 12px;
    font-weight: 700;
}

.section-title {
    font-family: 'Playfair Display', Georgia, serif;
    color: var(--forest-dark);
    font-size: 25px;
    line-height: 1.25;
    font-weight: 600;
    margin: 0;
}

.section-description {
    color: var(--muted);
    font-size: 12px;
    line-height: 1.7;
    margin-top: 5px;
}

/* Cards */
.panel {
    background: var(--paper);
    border: 1px solid var(--line);
    border-radius: 19px;
    padding: 24px;
    margin-bottom: 16px;
    box-shadow: 0 4px 18px rgba(30, 48, 37, 0.025);
}

.panel-title {
    color: var(--forest-dark);
    font-size: 14px;
    font-weight: 700;
    margin-bottom: 5px;
}

.panel-copy {
    color: var(--muted);
    font-size: 12px;
    line-height: 1.7;
    margin-bottom: 14px;
}

.micro-label {
    color: var(--gold);
    font-size: 10px;
    font-weight: 700;
    letter-spacing: 1.5px;
    text-transform: uppercase;
}

/* Form controls */
.stTextInput label,
.stTextArea label,
.stSelectbox label,
.stMultiSelect label,
.stSlider label,
.stFileUploader label,
.stNumberInput label,
.stRadio label,
.stCheckbox label {
    color: #34453A !important;
    font-size: 12px !important;
    font-weight: 600 !important;
}

.stTextInput input,
.stTextArea textarea,
.stNumberInput input {
    background: #FDFDF9 !important;
    border: 1px solid #DDE2D8 !important;
    border-radius: 10px !important;
    color: var(--ink) !important;
    font-size: 13px !important;
}

.stTextInput input:focus,
.stTextArea textarea:focus,
.stNumberInput input:focus {
    border-color: #73917B !important;
    box-shadow: 0 0 0 2px rgba(115, 145, 123, 0.12) !important;
}

.stSelectbox [data-baseweb="select"] > div,
.stMultiSelect [data-baseweb="select"] > div {
    background: #FDFDF9 !important;
    border-color: #DDE2D8 !important;
    border-radius: 10px !important;
}

.stFileUploader [data-testid="stFileUploaderDropzone"] {
    background: #F8FAF5;
    border: 1.5px dashed #C8D5C8;
    border-radius: 14px;
}

.stFileUploader [data-testid="stFileUploaderDropzone"] button {
    border: 1px solid #CAD8CA !important;
    border-radius: 9px !important;
    color: var(--forest) !important;
    background: white !important;
}

div[data-testid="stForm"] {
    background: transparent;
    border: none;
    padding: 0;
}

.stSlider [data-baseweb="slider"] {
    padding-top: 4px;
}

/* Buttons */
.stButton > button,
.stFormSubmitButton > button {
    background: var(--forest) !important;
    color: #FFFDF7 !important;
    border: 1px solid var(--forest) !important;
    border-radius: 11px !important;
    padding: 12px 21px !important;
    min-height: 46px;
    font-size: 13px !important;
    font-weight: 700 !important;
    transition: all 0.18s ease;
}

.stButton > button:hover,
.stFormSubmitButton > button:hover {
    background: #285542 !important;
    border-color: #285542 !important;
    box-shadow: 0 5px 15px rgba(25, 61, 50, 0.12);
}

.stDownloadButton > button {
    background: #FDFDF9 !important;
    color: var(--forest) !important;
    border: 1px solid #C9D5C9 !important;
    border-radius: 10px !important;
    font-size: 12px !important;
    font-weight: 700 !important;
    min-height: 42px;
}

/* Small badges and result panels */
.skill-chip {
    display: inline-block;
    background: #EEF3EB;
    color: #244D39;
    border: 1px solid #D9E4D6;
    padding: 6px 10px;
    border-radius: 7px;
    margin: 3px 4px 3px 0;
    font-size: 11px;
    font-weight: 600;
}

.skill-chip.missing {
    background: #F7F0E2;
    color: #80663B;
    border-color: #E9DDBF;
}

.metric-card {
    background: white;
    border: 1px solid var(--line);
    border-radius: 16px;
    padding: 20px;
    min-height: 126px;
}

.metric-label {
    color: var(--muted);
    font-size: 11px;
    font-weight: 600;
    margin-bottom: 13px;
}

.metric-value {
    color: var(--forest-dark);
    font-family: 'Playfair Display', Georgia, serif;
    font-size: 34px;
    line-height: 1.1;
}

.metric-note {
    color: var(--muted);
    font-size: 10px;
    margin-top: 8px;
}

.result-banner {
    background: #EAF0E7;
    border: 1px solid #D5E1D2;
    border-radius: 15px;
    padding: 18px 21px;
    margin: 16px 0 20px;
}

.result-banner strong {
    color: var(--forest-dark);
    font-size: 14px;
}

.result-banner p {
    color: #5F7164;
    font-size: 12px;
    line-height: 1.7;
    margin: 6px 0 0;
}

.roadmap-card {
    display: flex;
    gap: 15px;
    background: white;
    border: 1px solid var(--line);
    border-radius: 14px;
    padding: 18px;
    margin-bottom: 11px;
}

.roadmap-number {
    flex: 0 0 34px;
    height: 34px;
    border-radius: 11px;
    background: var(--forest);
    color: #F2E6C5;
    display: flex;
    align-items: center;
    justify-content: center;
    font-size: 12px;
    font-weight: 700;
}

.roadmap-title {
    color: var(--forest-dark);
    font-size: 13px;
    font-weight: 700;
    margin: 0 0 6px;
}

.roadmap-copy {
    color: var(--muted);
    font-size: 12px;
    line-height: 1.7;
}

.footer {
    border-top: 1px solid var(--line);
    padding-top: 20px;
    margin-top: 50px;
    display: flex;
    justify-content: space-between;
    gap: 15px;
    flex-wrap: wrap;
    color: var(--muted);
    font-size: 10px;
}

.stAlert {
    border-radius: 12px;
}

[data-testid="stDataFrame"] {
    border: 1px solid var(--line);
    border-radius: 12px;
    overflow: hidden;
}

@media (max-width: 700px) {
    .block-container {
        padding: 1rem 1rem 3rem;
    }

    .hero {
        padding: 32px 23px;
        border-radius: 19px;
    }

    .hero h1 {
        font-size: 40px;
        letter-spacing: -1px;
    }

    .section-title {
        font-size: 22px;
    }

    .panel {
        padding: 17px;
    }

    .brand-status {
        display: none;
    }
}
</style>
""", unsafe_allow_html=True)


# =========================================================
# CAREER ROLE LIBRARY
# =========================================================

ROLE_PROFILES = {
    "Data Analyst": {
        "skills": [
            "SQL", "Excel", "Python", "Pandas", "Statistics",
            "Power BI", "Data Analysis", "Data Visualization",
            "Communication", "Problem Solving"
        ],
        "projects": [
            "Build a sales dashboard using Power BI.",
            "Analyse a public dataset using Python and Pandas.",
            "Create SQL queries for a business reporting database."
        ],
        "courses": [
            "SQL fundamentals and joins",
            "Excel and Power BI dashboard development",
            "Python for data analysis"
        ]
    },
    "Machine Learning Engineer": {
        "skills": [
            "Python", "Machine Learning", "Statistics", "NumPy",
            "Pandas", "Scikit-learn", "Deep Learning",
            "SQL", "Data Preprocessing", "Model Evaluation"
        ],
        "projects": [
            "Build and evaluate a classification model.",
            "Create a complete data preprocessing pipeline.",
            "Deploy a machine learning model using Streamlit."
        ],
        "courses": [
            "Python and scientific computing",
            "Supervised and unsupervised learning",
            "Model evaluation and deployment"
        ]
    },
    "Python Developer": {
        "skills": [
            "Python", "OOP", "SQL", "Git", "REST API",
            "Debugging", "Data Structures", "Testing",
            "Problem Solving", "FastAPI"
        ],
        "projects": [
            "Build a REST API with FastAPI.",
            "Create a database-backed task management app.",
            "Develop an automated file-processing tool."
        ],
        "courses": [
            "Python programming and OOP",
            "REST API development",
            "Testing, Git and deployment"
        ]
    },
    "Full Stack Developer": {
        "skills": [
            "HTML", "CSS", "JavaScript", "React", "Node.js",
            "SQL", "Git", "REST API", "Authentication",
            "Responsive Design"
        ],
        "projects": [
            "Build a responsive portfolio with a contact form.",
            "Develop a full-stack task management application.",
            "Create a login system with database integration."
        ],
        "courses": [
            "HTML, CSS and modern JavaScript",
            "React and frontend architecture",
            "Backend APIs and database integration"
        ]
    },
    "Cloud Engineer": {
        "skills": [
            "AWS", "Azure", "Google Cloud", "Linux",
            "Networking", "Docker", "CI/CD", "Python",
            "Security", "Monitoring"
        ],
        "projects": [
            "Deploy a web application to a cloud platform.",
            "Containerise an application using Docker.",
            "Create a basic CI/CD deployment pipeline."
        ],
        "courses": [
            "Cloud fundamentals and IAM",
            "Linux, networking and Docker",
            "Cloud deployment and monitoring"
        ]
    },
    "Cybersecurity Analyst": {
        "skills": [
            "Networking", "Linux", "Python", "SIEM",
            "Incident Response", "Threat Analysis",
            "Security Fundamentals", "Log Analysis",
            "Vulnerability Assessment", "Risk Management"
        ],
        "projects": [
            "Analyse sample security logs for suspicious activity.",
            "Create a basic network traffic analysis report.",
            "Build a security incident response checklist."
        ],
        "courses": [
            "Networking and security fundamentals",
            "Log analysis and incident response",
            "Vulnerability assessment in authorised labs"
        ]
    },
    "Software Engineer": {
        "skills": [
            "Python", "Java", "Data Structures", "Algorithms",
            "OOP", "SQL", "Git", "Testing",
            "System Design", "Problem Solving"
        ],
        "projects": [
            "Build a well-tested application with clear architecture.",
            "Implement common data structures and algorithms.",
            "Create a database-backed application with automated tests."
        ],
        "courses": [
            "Data structures and algorithms",
            "Object-oriented programming and testing",
            "Databases and software design"
        ]
    },
    "UI/UX Designer": {
        "skills": [
            "Figma", "Wireframing", "Prototyping", "User Research",
            "Usability Testing", "Visual Design", "Typography",
            "Design Systems", "Accessibility", "Communication"
        ],
        "projects": [
            "Redesign a mobile app and document your design decisions.",
            "Create a clickable prototype for a booking experience.",
            "Conduct a small usability test and report findings."
        ],
        "courses": [
            "User research and information architecture",
            "Wireframing and interactive prototyping",
            "Accessibility and design systems"
        ]
    }
}


SKILL_LIBRARY = sorted(set(
    skill
    for profile in ROLE_PROFILES.values()
    for skill in profile["skills"]
) | {
    "C", "C++", "C#", "Java", "JavaScript", "TypeScript",
    "PHP", "R", "Go", "Rust", "Kotlin", "Swift",
    "Django", "Flask", "Spring Boot", "Angular", "Vue",
    "PostgreSQL", "MySQL", "MongoDB", "SQLite",
    "Tableau", "TensorFlow", "PyTorch", "Keras",
    "NLP", "Computer Vision", "Generative AI", "LLMs",
    "GitHub", "Kubernetes", "Terraform", "Jenkins",
    "Excel", "Power BI", "Pandas", "NumPy", "Matplotlib",
    "Data Analysis", "Machine Learning", "Deep Learning",
    "Communication", "Teamwork", "Leadership",
    "Critical Thinking", "Agile", "Project Management",
    "Figma", "HTML", "CSS", "React", "Node.js",
    "REST API", "FastAPI", "Scikit-learn"
))


# =========================================================
# HELPERS
# =========================================================

def safe_html(value):
    """Escape user-provided text before rendering it in HTML."""
    return html.escape(str(value))


def extract_resume_text(uploaded_file):
    """Extract text from PDF, DOCX or TXT files."""
    if uploaded_file is None:
        return ""

    filename = uploaded_file.name.lower()

    try:
        if filename.endswith(".pdf"):
            reader = PdfReader(uploaded_file)
            pages = []
            for page in reader.pages:
                pages.append(page.extract_text() or "")
            return "\n".join(pages)

        if filename.endswith(".docx"):
            document = Document(uploaded_file)
            paragraphs = [
                paragraph.text for paragraph in document.paragraphs
            ]
            for table in document.tables:
                for row in table.rows:
                    paragraphs.append(
                        " ".join(cell.text for cell in row.cells)
                    )
            return "\n".join(paragraphs)

        if filename.endswith(".txt"):
            return uploaded_file.getvalue().decode(
                "utf-8", errors="ignore"
            )

    except Exception as error:
        st.error(f"Could not read the resume: {error}")
        return ""

    return ""


def detect_skills(text):
    """Find known skills mentioned in resume or job text."""
    text_lower = text.lower()
    found = []

    for skill in SKILL_LIBRARY:
        pattern = (
            r"(?<![a-z0-9+#])"
            + re.escape(skill.lower())
            + r"(?![a-z0-9+#])"
        )
        if re.search(pattern, text_lower):
            found.append(skill)

    return sorted(set(found), key=str.lower)


def calculate_text_similarity(profile_text, target_text):
    """Estimate textual similarity using TF-IDF."""
    if not profile_text.strip() or not target_text.strip():
        return 0

    try:
        vectorizer = TfidfVectorizer(
            stop_words="english",
            ngram_range=(1, 2),
        )
        matrix = vectorizer.fit_transform(
            [profile_text, target_text]
        )
        similarity = cosine_similarity(
            matrix[0:1], matrix[1:2]
        )[0][0]
        return round(float(similarity) * 100)
    except ValueError:
        return 0


def build_pdf(report):
    """Create a simple downloadable PDF report."""
    pdf = FPDF()
    pdf.set_auto_page_break(auto=True, margin=15)
    pdf.add_page()

    pdf.set_font("Helvetica", "B", 20)
    pdf.cell(0, 12, "SkillBridge AI", new_x="LMARGIN", new_y="NEXT")

    pdf.set_font("Helvetica", "", 11)
    pdf.cell(
        0, 8, "Career Intelligence Report",
        new_x="LMARGIN", new_y="NEXT"
    )
    pdf.ln(4)

    def clean(value):
        value = str(value)
        return value.encode("latin-1", errors="replace").decode("latin-1")

    for heading, value in report.items():
        pdf.set_font("Helvetica", "B", 12)
        pdf.multi_cell(0, 8, clean(heading))

        pdf.set_font("Helvetica", "", 10)
        pdf.multi_cell(0, 6, clean(value))
        pdf.ln(3)

    return bytes(pdf.output())


def section_heading(number, title, description):
    st.markdown(f"""
    <div class="section-heading">
        <div class="section-number">{number}</div>
        <div>
            <div class="section-title">{title}</div>
            <div class="section-description">{description}</div>
        </div>
    </div>
    """, unsafe_allow_html=True)


def render_chips(items, missing=False):
    if not items:
        st.caption("No skills identified in this category.")
        return

    chip_class = "skill-chip missing" if missing else "skill-chip"
    chips = "".join(
        f'<span class="{chip_class}">{safe_html(skill)}</span>'
        for skill in items
    )
    st.markdown(chips, unsafe_allow_html=True)


# =========================================================
# HEADER AND HERO
# =========================================================

st.markdown("""
<div class="brand-row">
    <div class="brand-lockup">
        <div class="brand-mark">✦</div>
        <div>
            <div class="brand-name">SkillBridge AI</div>
            <div class="brand-subtitle">Career Intelligence</div>
        </div>
    </div>
    <div class="brand-status">YOUR CAREER WORKSPACE</div>
</div>

<div class="hero">
    <div class="hero-eyebrow">A more considered approach to your career</div>
    <h1>Build your next career move<br><em>with clarity.</em></h1>
    <div class="hero-copy">
        Your skills, experience, and ambitions — brought together
        in one personal career workspace. Discover your role alignment,
        identify skill gaps, and find your next steps.
    </div>
    <div class="hero-features">
        <span class="hero-feature">✦ Resume insights</span>
        <span class="hero-feature">Skill gap mapping</span>
        <span class="hero-feature">Personal learning roadmap</span>
    </div>
</div>
""", unsafe_allow_html=True)


# =========================================================
# PROFILE INPUTS
# =========================================================

section_heading(
    "01",
    "Tell us about yourself",
    "Build your profile so your career recommendations have context."
)

with st.container(border=True):
    st.markdown(
        '<div class="micro-label">YOUR PROFILE</div>',
        unsafe_allow_html=True
    )
    st.markdown("")

    with st.form("career_profile_form", clear_on_submit=False):

        col1, col2 = st.columns(2)

        with col1:
            full_name = st.text_input(
                "Full name",
                placeholder="How should we address you?",
            )
            education = st.selectbox(
                "Education",
                [
                    "Select education",
                    "Diploma",
                    "Bachelor's degree",
                    "Master's degree",
                    "PhD",
                    "Other",
                ],
            )
            experience_level = st.selectbox(
                "Experience level",
                [
                    "Fresher / Student",
                    "Entry level",
                    "Mid-level professional",
                    "Senior professional",
                    "Career switcher",
                ],
            )

        with col2:
            current_status = st.selectbox(
                "Current status",
                [
                    "Student",
                    "Looking for a job",
                    "Currently employed",
                    "Freelancer",
                    "Career switcher",
                ],
            )
            years_experience = st.number_input(
                "Years of experience",
                min_value=0,
                max_value=50,
                value=0,
                step=1,
            )
            target_role = st.selectbox(
                "Career you want to explore",
                list(ROLE_PROFILES.keys()) + ["Custom job description"],
            )

        st.markdown("---")

        section_heading(
            "02",
            "Your resume and skills",
            "Add your experience and tell us what you already know."
        )

        uploaded_resume = st.file_uploader(
            "Upload your resume",
            type=["pdf", "docx", "txt"],
            help="Supported formats: PDF, DOCX and TXT. "
                 "The maximum upload size is controlled by Streamlit.",
        )

        resume_text_input = st.text_area(
            "Resume text / professional summary",
            placeholder=(
                "Your resume text will be extracted after you submit. "
                "You can also paste your experience, projects and achievements here."
            ),
            height=140,
            help="You can paste your resume text here if you don't have a file.",
        )

        known_skills = st.multiselect(
            "Skills you already know",
            options=SKILL_LIBRARY,
            placeholder="Search and select your skills...",
        )

        custom_skills = st.text_input(
            "Other skills (optional)",
            placeholder="e.g. Salesforce, SAP, prompt engineering",
        )

        skill_confidence = st.slider(
            "How confident are you in your selected skills?",
            min_value=1,
            max_value=5,
            value=3,
            help="This helps contextualise the recommendations. "
                 "It is not a validated proficiency test.",
        )

        interests = st.text_input(
            "Interests and preferred work",
            placeholder="e.g. solving business problems, building apps, research",
        )

        section_heading(
            "03",
            "Choose your target opportunity",
            "Compare your profile with a career role or a specific job posting."
        )

        custom_job_description = ""
        if target_role == "Custom job description":
            custom_job_description = st.text_area(
                "Paste the job description",
                placeholder=(
                    "Paste the job responsibilities, requirements "
                    "and qualifications from the job posting..."
                ),
                height=180,
            )

        submitted = st.form_submit_button(
            "✦  Generate my career analysis",
            use_container_width=True,
        )


# =========================================================
# PROCESS SUBMISSION
# =========================================================

if submitted:

    if not full_name.strip():
        st.error("Please enter your name to continue.")

    elif target_role == "Custom job description" and not custom_job_description.strip():
        st.error("Please paste a job description for a custom comparison.")

    else:
        resume_text = extract_resume_text(uploaded_resume)

        if resume_text_input.strip():
            resume_text = (
                resume_text + "\n" + resume_text_input.strip()
            ).strip()

        custom_skill_list = [
            skill.strip()
            for skill in custom_skills.split(",")
            if skill.strip()
        ]

        manually_selected_skills = sorted(
            set(known_skills + custom_skill_list),
            key=str.lower,
        )

        resume_detected_skills = detect_skills(resume_text)

        all_user_skills = sorted(
            set(manually_selected_skills + resume_detected_skills),
            key=str.lower,
        )

        if target_role == "Custom job description":
            job_text = custom_job_description
            target_skills = detect_skills(job_text)

            if not target_skills:
                st.warning(
                    "No recognised skills were detected in the job description. "
                    "Try including the required technical skills explicitly."
                )
        else:
            profile = ROLE_PROFILES[target_role]
            target_skills = profile["skills"]
            job_text = " ".join(target_skills)

        if not all_user_skills and not resume_text.strip():
            st.warning(
                "For a more useful analysis, add some known skills "
                "or provide your resume text."
            )

        if not target_skills:
            target_skills = []

        matched_skills = [
            skill for skill in target_skills
            if any(
                skill.lower() == user_skill.lower()
                for user_skill in all_user_skills
            )
        ]

        missing_skills = [
            skill for skill in target_skills
            if not any(
                skill.lower() == user_skill.lower()
                for user_skill in all_user_skills
            )
        ]

        coverage = (
            round(len(matched_skills) / len(target_skills) * 100)
            if target_skills else 0
        )

        profile_text = " ".join([
            resume_text,
            " ".join(all_user_skills),
            interests,
            education,
            experience_level,
        ])

        similarity = calculate_text_similarity(
            profile_text, job_text
        )

        # This is a heuristic indicator, not a validated hiring prediction.
        career_fit = round(coverage * 0.65 + similarity * 0.35)

        if not all_user_skills and not resume_text.strip():
            career_fit = 0

        if career_fit >= 75:
            fit_label = "Strong alignment"
            fit_description = (
                "Your provided profile has substantial overlap with the "
                "selected requirements. Strengthen your evidence with projects "
                "and measurable achievements."
            )
        elif career_fit >= 45:
            fit_label = "Developing alignment"
            fit_description = (
                "You have some relevant overlap. Prioritise the missing skills "
                "and demonstrate them through focused projects."
            )
        else:
            fit_label = "Early-stage alignment"
            fit_description = (
                "There are gaps between your current profile and the selected "
                "requirements. Start with the fundamentals and build evidence "
                "through practical work."
            )

        if target_role == "Custom job description":
            suggested_projects = [
                "Build a small project demonstrating one important requirement "
                "from the job description.",
                "Create a portfolio case study explaining your approach, "
                "decisions and results.",
                "Practise a realistic task related to the role and document "
                "your solution.",
            ]
            suggested_courses = [
                "Study the fundamentals of your highest-priority missing skill.",
                "Complete guided exercises using the tools named in the job post.",
                "Practise communicating your project decisions and outcomes.",
            ]
        else:
            suggested_projects = ROLE_PROFILES[target_role]["projects"]
            suggested_courses = ROLE_PROFILES[target_role]["courses"]

        st.session_state["career_results"] = {
            "full_name": full_name.strip(),
            "education": education,
            "experience_level": experience_level,
            "current_status": current_status,
            "years_experience": int(years_experience),
            "target_role": target_role,
            "interests": interests,
            "skill_confidence": skill_confidence,
            "resume_text": resume_text,
            "user_skills": all_user_skills,
            "target_skills": target_skills,
            "matched_skills": matched_skills,
            "missing_skills": missing_skills,
            "coverage": coverage,
            "similarity": similarity,
            "career_fit": career_fit,
            "fit_label": fit_label,
            "fit_description": fit_description,
            "projects": suggested_projects,
            "courses": suggested_courses,
            "job_text": job_text,
            "generated_date": date.today().isoformat(),
        }

        st.success("Your career analysis has been generated.")


# =========================================================
# RESULTS
# =========================================================

results = st.session_state.get("career_results")

if results:

    st.markdown("---")

    section_heading(
        "04",
        "Your career snapshot",
        f"Personalised analysis for {results['full_name']}."
    )

    st.markdown(f"""
    <div class="result-banner">
        <strong>{safe_html(results['fit_label'])}</strong>
        <p>{safe_html(results['fit_description'])}</p>
    </div>
    """, unsafe_allow_html=True)

    m1, m2, m3 = st.columns(3)

    with m1:
        st.markdown(f"""
        <div class="metric-card">
            <div class="metric-label">CAREER ALIGNMENT INDICATOR</div>
            <div class="metric-value">{results['career_fit']}%</div>
            <div class="metric-note">Heuristic estimate, not a hiring probability</div>
        </div>
        """, unsafe_allow_html=True)

    with m2:
        st.markdown(f"""
        <div class="metric-card">
            <div class="metric-label">REQUIRED SKILLS COVERED</div>
            <div class="metric-value">{results['coverage']}%</div>
            <div class="metric-note">
                {len(results['matched_skills'])} of
                {len(results['target_skills'])} recognised target skills
            </div>
        </div>
        """, unsafe_allow_html=True)

    with m3:
        st.markdown(f"""
        <div class="metric-card">
            <div class="metric-label">SKILLS TO PRIORITISE</div>
            <div class="metric-value">{len(results['missing_skills'])}</div>
            <div class="metric-note">Recognised requirements not in your provided skill list</div>
        </div>
        """, unsafe_allow_html=True)

    st.markdown("")

    left, right = st.columns([1, 1])

    with left:
        with st.container(border=True):
            st.markdown(
                '<div class="panel-title">Your profile</div>',
                unsafe_allow_html=True
            )
            st.markdown(
                f"**Name:** {results['full_name']}"
            )
            st.markdown(
                f"**Education:** {results['education']}"
            )
            st.markdown(
                f"**Experience:** {results['experience_level']}"
            )
            st.markdown(
                f"**Current status:** {results['current_status']}"
            )
            st.markdown(
                f"**Years of experience:** {results['years_experience']}"
            )
            st.markdown(
                f"**Target opportunity:** {results['target_role']}"
            )

    with right:
        with st.container(border=True):
            st.markdown(
                '<div class="panel-title">Your strongest evidence</div>',
                unsafe_allow_html=True
            )
            if results["user_skills"]:
                render_chips(results["user_skills"])
            else:
                st.info(
                    "No skills have been identified yet. Add your skills "
                    "or resume details and generate the analysis again."
                )

    # -----------------------------------------------------
    # Skill gap analysis
    # -----------------------------------------------------

    section_heading(
        "05",
        "Skill gap analysis",
        "See where your provided skills overlap with the target requirements."
    )

    gap_left, gap_right = st.columns(2)

    with gap_left:
        with st.container(border=True):
            st.markdown(
                '<div class="panel-title">Skills you already have</div>',
                unsafe_allow_html=True
            )
            st.markdown(
                '<div class="panel-copy">Recognised target skills found in your selected skills.</div>',
                unsafe_allow_html=True
            )
            render_chips(results["matched_skills"])

    with gap_right:
        with st.container(border=True):
            st.markdown(
                '<div class="panel-title">Skills to develop</div>',
                unsafe_allow_html=True
            )
            st.markdown(
                '<div class="panel-copy">Requirements not found in your provided skill list.</div>',
                unsafe_allow_html=True
            )
            render_chips(results["missing_skills"], missing=True)

    gap_rows = []

    for skill in results["target_skills"]:
        is_matched = skill in results["matched_skills"]
        gap_rows.append({
            "Required skill": skill,
            "Status": "Covered" if is_matched else "Priority to learn",
            "Suggested priority": "Maintain" if is_matched else "High",
        })

    if gap_rows:
        gap_df = pd.DataFrame(gap_rows)
        st.dataframe(
            gap_df,
            use_container_width=True,
            hide_index=True,
        )
    else:
        st.info(
            "The job description did not contain recognised skills from "
            "the built-in library. Add explicit skill requirements to "
            "improve the comparison."
        )

    # -----------------------------------------------------
    # Learning roadmap
    # -----------------------------------------------------

    section_heading(
        "06",
        "Your learning roadmap",
        "A practical sequence for turning skill gaps into demonstrable ability."
    )

    roadmap_items = []

    if results["missing_skills"]:
        roadmap_items.append({
            "title": "Build the foundations",
            "copy": (
                "Start with " + ", ".join(results["missing_skills"][:3]) +
                ". Learn the core concepts and complete small guided exercises."
            ),
        })
        roadmap_items.append({
            "title": "Practise with a focused project",
            "copy": (
                "Apply the new skills to a small project. Document your "
                "approach, the tools used and the result."
            ),
        })
        roadmap_items.append({
            "title": "Demonstrate your progress",
            "copy": (
                "Publish your project, explain your decisions and practise "
                "answering questions about your implementation."
            ),
        })
    else:
        roadmap_items = [
            {
                "title": "Strengthen your evidence",
                "copy": (
                    "Build a portfolio project that demonstrates your "
                    "existing skills in a realistic setting."
                ),
            },
            {
                "title": "Develop depth",
                "copy": (
                    "Practise more advanced tasks and explain the trade-offs "
                    "behind your decisions."
                ),
            },
            {
                "title": "Prepare for applications",
                "copy": (
                    "Tailor your resume to the role and prepare examples "
                    "that show the impact of your work."
                ),
            },
        ]

    for index, item in enumerate(roadmap_items, start=1):
        st.markdown(f"""
        <div class="roadmap-card">
            <div class="roadmap-number">{index:02d}</div>
            <div>
                <div class="roadmap-title">{safe_html(item['title'])}</div>
                <div class="roadmap-copy">{safe_html(item['copy'])}</div>
            </div>
        </div>
        """, unsafe_allow_html=True)

    project_col, course_col = st.columns(2)

    with project_col:
        with st.container(border=True):
            st.markdown(
                '<div class="panel-title">Suggested projects</div>',
                unsafe_allow_html=True
            )
            st.markdown("")
            for i, project in enumerate(results["projects"], start=1):
                st.markdown(f"**{i:02d}.** {project}")

    with course_col:
        with st.container(border=True):
            st.markdown(
                '<div class="panel-title">Suggested learning topics</div>',
                unsafe_allow_html=True
            )
            st.markdown("")
            for i, course in enumerate(results["courses"], start=1):
                st.markdown(f"**{i:02d}.** {course}")

    # -----------------------------------------------------
    # Export report
    # -----------------------------------------------------

    section_heading(
        "07",
        "Take your analysis with you",
        "Export your results for reference or further editing."
    )

    report = {
        "Candidate": results["full_name"],
        "Education": results["education"],
        "Experience level": results["experience_level"],
        "Current status": results["current_status"],
        "Years of experience": results["years_experience"],
        "Target opportunity": results["target_role"],
        "Generated date": results["generated_date"],
        "Career alignment indicator": (
            f"{results['career_fit']}% — {results['fit_label']}. "
            "This is a heuristic indicator, not a validated hiring prediction."
        ),
        "Skill coverage": (
            f"{results['coverage']}% "
            f"({len(results['matched_skills'])} of "
            f"{len(results['target_skills'])} recognised skills)"
        ),
        "Text similarity": (
            f"{results['similarity']}% TF-IDF similarity. "
            "This is a text comparison, not a measure of actual competence."
        ),
        "Skills identified": ", ".join(results["user_skills"]) or "None identified",
        "Matched skills": ", ".join(results["matched_skills"]) or "None",
        "Skills to develop": ", ".join(results["missing_skills"]) or "None identified",
        "Suggested projects": "\n".join(
            f"- {item}" for item in results["projects"]
        ),
        "Suggested learning topics": "\n".join(
            f"- {item}" for item in results["courses"]
        ),
    }

    report_df = pd.DataFrame(
        [{"Section": key, "Details": value} for key, value in report.items()]
    )

    export_left, export_right = st.columns(2)

    with export_left:
        csv_data = report_df.to_csv(index=False).encode("utf-8-sig")
        st.download_button(
            "Download CSV report",
            data=csv_data,
            file_name="skillbridge_career_report.csv",
            mime="text/csv",
            use_container_width=True,
        )

    with export_right:
        try:
            pdf_data = build_pdf(report)
            st.download_button(
                "Download PDF report",
                data=pdf_data,
                file_name="skillbridge_career_report.pdf",
                mime="application/pdf",
                use_container_width=True,
            )
        except Exception as error:
            st.warning(
                f"PDF generation failed. You can still download the CSV. "
                f"Details: {error}"
            )

    st.caption(
        "This report is generated from the information you provide. "
        "It is an educational aid, not a validated recruitment assessment "
        "or guarantee of employment."
    )


# =========================================================
# FOOTER
# =========================================================

st.markdown("""
<div class="footer">
    <span>✦ SkillBridge AI · Career Intelligence</span>
    <span>Understand your skills. Plan your next move.</span>
</div>
""", unsafe_allow_html=True)
