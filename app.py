
import html
import re
from datetime import date

import pandas as pd
import streamlit as st
from docx import Document
from fpdf import FPDF
from pypdf import PdfReader
from sklearn.feature_extraction.text import TfidfVectorizer
from sklearn.metrics.pairwise import cosine_similarity


# ============================================================
# CONFIGURATION
# ============================================================

st.set_page_config(
    page_title="SkillBridge AI | Career Intelligence",
    page_icon="✦",
    layout="wide",
    initial_sidebar_state="collapsed",
)


# ============================================================
# LUXURY DESIGN SYSTEM
# Midnight Navy | Ivory | Champagne | Sage
# ============================================================

st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=DM+Sans:wght@400;500;600;700&family=Playfair+Display:wght@500;600;700&display=swap');

:root {
    --navy: #17232D;
    --navy-light: #263B46;
    --ivory: #F6F3EB;
    --white: #FFFEFA;
    --gold: #C7A96B;
    --gold-light: #E7D7B2;
    --sage: #DDE7DC;
    --sage-dark: #345848;
    --ink: #202B31;
    --muted: #59666C;
    --border: #DEDCD3;
    --soft: #EFEEE6;
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
    max-width: 1200px;
    padding: 1.7rem 2rem 4rem;
}

/* General readable text */

p, li, label, .stMarkdown {
    color: var(--ink);
}

[data-testid="stCaptionContainer"],
.stCaption {
    color: var(--muted) !important;
}

/* Brand header */

.brand-row {
    display: flex;
    justify-content: space-between;
    align-items: center;
    gap: 18px;
    padding: 4px 0 22px;
    border-bottom: 1px solid var(--border);
    margin-bottom: 25px;
}

.brand-lockup {
    display: flex;
    align-items: center;
    gap: 12px;
}

.brand-mark {
    width: 44px;
    height: 44px;
    border-radius: 13px;
    background: var(--navy);
    display: flex;
    align-items: center;
    justify-content: center;
    color: var(--gold-light);
    font-size: 25px;
}

.brand-name {
    color: var(--navy);
    font-size: 17px;
    font-weight: 700;
    letter-spacing: -0.4px;
}

.brand-subtitle {
    color: var(--muted);
    font-size: 10px;
    letter-spacing: 1.5px;
    text-transform: uppercase;
    margin-top: 3px;
}

.brand-status {
    background: #E9E6DC;
    border: 1px solid #D8D0BD;
    color: #34453D;
    border-radius: 50px;
    padding: 8px 12px;
    font-size: 10px;
    font-weight: 700;
    letter-spacing: .8px;
}

/* Hero */

.hero {
    position: relative;
    overflow: hidden;
    background:
        radial-gradient(
            ellipse at 90% 5%,
            rgba(199, 169, 107, .18),
            transparent 35%
        ),
        linear-gradient(
            125deg,
            #17232D 0%,
            #263B46 62%,
            #1A302F 100%
        );
    border-radius: 24px;
    padding: 50px 48px 42px;
    margin: 8px 0 32px;
    box-shadow: 0 16px 38px rgba(23, 35, 45, .12);
}

.hero-eyebrow {
    position: relative;
    z-index: 1;
    color: var(--gold-light);
    font-size: 10px;
    font-weight: 700;
    letter-spacing: 2px;
    text-transform: uppercase;
    margin-bottom: 20px;
}

.hero h1 {
    position: relative;
    z-index: 1;
    font-family: 'Playfair Display', Georgia, serif;
    color: #FFFDF7;
    font-size: clamp(38px, 5vw, 64px);
    line-height: 1.12;
    font-weight: 500;
    letter-spacing: -1.3px;
    margin: 0 0 20px;
}

.hero h1 em {
    color: var(--gold-light);
    font-style: normal;
}

.hero-copy {
    position: relative;
    z-index: 1;
    color: #E0E6E4;
    max-width: 650px;
    font-size: 14px;
    line-height: 1.9;
    margin-bottom: 27px;
}

.hero-features {
    position: relative;
    z-index: 1;
    display: flex;
    flex-wrap: wrap;
    gap: 9px;
}

.hero-feature {
    border: 1px solid rgba(231, 215, 178, .3);
    border-radius: 50px;
    padding: 8px 12px;
    color: #F5F0E5;
    font-size: 11px;
    background: rgba(255, 255, 255, .035);
}

/* Section headings */

.section-heading {
    display: flex;
    align-items: flex-start;
    gap: 14px;
    margin: 34px 0 15px;
}

.section-number {
    flex: 0 0 39px;
    height: 39px;
    border-radius: 12px;
    background: #E7E9DF;
    border: 1px solid #D6DDD0;
    display: flex;
    align-items: center;
    justify-content: center;
    color: #254638;
    font-size: 11px;
    font-weight: 700;
}

.section-title {
    font-family: 'Playfair Display', Georgia, serif;
    color: var(--navy);
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

div[data-testid="stVerticalBlockBorderWrapper"] {
    background: var(--white);
    border-color: var(--border) !important;
    border-radius: 17px;
}

.panel-title {
    color: var(--navy);
    font-size: 14px;
    font-weight: 700;
    margin-bottom: 6px;
}

/* Input labels */

.stTextInput label,
.stTextArea label,
.stSelectbox label,
.stMultiSelect label,
.stSlider label,
.stFileUploader label,
.stNumberInput label,
.stRadio label,
.stCheckbox label {
    color: #263A36 !important;
    font-size: 12px !important;
    font-weight: 700 !important;
}

/* Inputs */

.stTextInput input,
.stTextArea textarea,
.stNumberInput input {
    background: #FFFEFA !important;
    color: #202B31 !important;
    border: 1px solid #CFCFC5 !important;
    border-radius: 10px !important;
    font-size: 13px !important;
}

.stTextInput input::placeholder,
.stTextArea textarea::placeholder {
    color: #737D7A !important;
    opacity: 1 !important;
}

.stTextInput input:focus,
.stTextArea textarea:focus {
    border-color: #8A9E8B !important;
    box-shadow: 0 0 0 2px rgba(52, 88, 72, .12) !important;
}

.stSelectbox [data-baseweb="select"] > div,
.stMultiSelect [data-baseweb="select"] > div {
    background: #FFFEFA !important;
    border-color: #CFCFC5 !important;
    border-radius: 10px !important;
}

.stSelectbox [data-baseweb="select"] *,
.stMultiSelect [data-baseweb="select"] * {
    color: #202B31 !important;
}

.stMultiSelect [data-baseweb="tag"] {
    background: #DDE7DC !important;
    border: 1px solid #B9CDBA !important;
}

.stMultiSelect [data-baseweb="tag"] * {
    color: #244B39 !important;
}

/* Resume uploader */

.stFileUploader [data-testid="stFileUploaderDropzone"] {
    background: #F8F8F2;
    border: 1.5px dashed #B7C4B5;
    border-radius: 13px;
}

.stFileUploader [data-testid="stFileUploaderDropzone"] * {
    color: #34453D !important;
}

.stFileUploader [data-testid="stFileUploaderDropzone"] button {
    background: #FFFEFA !important;
    border: 1px solid #B7C4B5 !important;
    color: #244B39 !important;
}

/* Buttons */

.stButton > button,
.stFormSubmitButton > button {
    background: var(--navy) !important;
    color: #FFFDF7 !important;
    border: 1px solid var(--navy) !important;
    border-radius: 10px !important;
    min-height: 45px;
    padding: 10px 18px !important;
    font-size: 12px !important;
    font-weight: 700 !important;
}

.stButton > button:hover,
.stFormSubmitButton > button:hover {
    background: #2B454F !important;
    border-color: #2B454F !important;
}

.stDownloadButton > button {
    background: #FFFEFA !important;
    color: #213F34 !important;
    border: 1px solid #B7C8B7 !important;
    border-radius: 10px !important;
    font-size: 12px !important;
    font-weight: 700 !important;
    min-height: 42px;
}

/* Results */

.skill-chip {
    display: inline-block;
    background: #EAF0E7;
    color: #244B39;
    border: 1px solid #D3E0D0;
    padding: 6px 10px;
    border-radius: 7px;
    margin: 3px 4px 3px 0;
    font-size: 11px;
    font-weight: 600;
}

.skill-chip.missing {
    background: #F4EBDD;
    color: #765B2D;
    border-color: #E4D3B4;
}

.metric-card {
    background: #FFFEFA;
    border: 1px solid var(--border);
    border-radius: 15px;
    padding: 19px;
    min-height: 125px;
}

.metric-label {
    color: var(--muted);
    font-size: 10px;
    font-weight: 700;
    letter-spacing: .6px;
    margin-bottom: 12px;
}

.metric-value {
    color: #172F2A;
    font-family: 'Playfair Display', Georgia, serif;
    font-size: 34px;
    line-height: 1.1;
}

.metric-note {
    color: var(--muted);
    font-size: 10px;
    line-height: 1.5;
    margin-top: 8px;
}

.result-banner {
    background: #E7EEE5;
    border: 1px solid #CEDCCB;
    border-radius: 14px;
    padding: 17px 20px;
    margin: 14px 0 19px;
}

.result-banner strong {
    color: #203F32;
    font-size: 14px;
}

.result-banner p {
    color: #40594A;
    font-size: 12px;
    line-height: 1.7;
    margin: 6px 0 0;
}

.roadmap-card {
    display: flex;
    gap: 14px;
    background: #FFFEFA;
    border: 1px solid var(--border);
    border-radius: 13px;
    padding: 16px;
    margin-bottom: 10px;
}

.roadmap-number {
    flex: 0 0 33px;
    height: 33px;
    border-radius: 10px;
    background: var(--navy);
    color: var(--gold-light);
    display: flex;
    align-items: center;
    justify-content: center;
    font-size: 11px;
    font-weight: 700;
}

.roadmap-title {
    color: #172F2A;
    font-size: 13px;
    font-weight: 700;
    margin: 0 0 5px;
}

.roadmap-copy {
    color: var(--muted);
    font-size: 12px;
    line-height: 1.7;
}

.stTabs [data-baseweb="tab-list"] {
    gap: 8px;
    border-bottom: 1px solid var(--border);
}

.stTabs [data-baseweb="tab"] {
    background: transparent;
    color: #4B5A5D !important;
    font-weight: 600;
}

.stTabs [aria-selected="true"] {
    color: #193D32 !important;
    border-bottom: 3px solid var(--gold) !important;
}

div[data-testid="stDataFrame"] {
    border: 1px solid var(--border);
    border-radius: 12px;
    overflow: hidden;
}

.stAlert {
    border-radius: 11px;
}

.footer {
    border-top: 1px solid var(--border);
    padding-top: 18px;
    margin-top: 45px;
    display: flex;
    justify-content: space-between;
    gap: 12px;
    flex-wrap: wrap;
    color: var(--muted);
    font-size: 10px;
}

@media (max-width: 700px) {
    .block-container {
        padding: 1rem 1rem 3rem;
    }

    .hero {
        padding: 31px 22px;
        border-radius: 18px;
    }

    .hero h1 {
        font-size: 39px;
        letter-spacing: -.6px;
    }

    .section-title {
        font-size: 22px;
    }

    .brand-status {
        display: none;
    }
}
</style>
""", unsafe_allow_html=True)


# ============================================================
# ROLE LIBRARY
# ============================================================

ROLE_PROFILES = {
    "Data Analyst": {
        "skills": [
            "SQL", "Excel", "Python", "Pandas", "Statistics",
            "Power BI", "Data Analysis", "Data Visualization",
            "Communication", "Problem Solving"
        ],
        "projects": [
            "Create a sales dashboard using Power BI.",
            "Analyse a public dataset using Python and Pandas.",
            "Build SQL reports and document business insights."
        ],
        "courses": [
            ("SQL for Data Analysis", "https://www.kaggle.com/learn/intro-to-sql"),
            ("Python", "https://www.kaggle.com/learn/python"),
            ("Power BI learning paths", "https://learn.microsoft.com/training/powerplatform/power-bi/")
        ],
        "certifications": [
            ("Microsoft Power BI Data Analyst", "https://learn.microsoft.com/credentials/certifications/power-bi-data-analyst-associate/")
        ],
        "market": "Analytics roles commonly value SQL, spreadsheet fluency, dashboarding, data storytelling and the ability to turn business questions into measurable insights."
    },

    "Machine Learning Engineer": {
        "skills": [
            "Python", "Machine Learning", "Statistics", "NumPy",
            "Pandas", "Scikit-learn", "Deep Learning", "SQL",
            "Data Preprocessing", "Model Evaluation"
        ],
        "projects": [
            "Train and evaluate a classification model.",
            "Build a reusable data preprocessing pipeline.",
            "Deploy a machine learning model in Streamlit."
        ],
        "courses": [
            ("Machine Learning Specialization", "https://www.coursera.org/specializations/machine-learning-introduction"),
            ("Scikit-learn tutorials", "https://scikit-learn.org/stable/tutorial/index.html"),
            ("Kaggle Learn", "https://www.kaggle.com/learn")
        ],
        "certifications": [
            ("AWS Machine Learning resources", "https://aws.amazon.com/machine-learning/")
        ],
        "market": "ML roles often combine Python, data preparation, evaluation, deployment and statistical reasoning. Production skills can matter as much as model selection."
    },

    "Python Developer": {
        "skills": [
            "Python", "OOP", "SQL", "Git", "REST API",
            "Debugging", "Data Structures", "Testing",
            "Problem Solving", "FastAPI"
        ],
        "projects": [
            "Build a REST API with FastAPI.",
            "Create a database-backed task manager.",
            "Develop an automated file-processing utility with tests."
        ],
        "courses": [
            ("Python tutorial", "https://docs.python.org/3/tutorial/"),
            ("FastAPI tutorial", "https://fastapi.tiangolo.com/tutorial/"),
            ("Git handbook", "https://guides.github.com/introduction/git-handbook/")
        ],
        "certifications": [
            ("Python Institute certifications", "https://pythoninstitute.org/certification-tracks")
        ],
        "market": "Python roles span backend APIs, automation, data tooling and AI integrations. Testing, version control and API design are useful differentiators."
    },

    "Full Stack Developer": {
        "skills": [
            "HTML", "CSS", "JavaScript", "React", "Node.js",
            "SQL", "Git", "REST API", "Authentication",
            "Responsive Design"
        ],
        "projects": [
            "Build a responsive portfolio website.",
            "Develop a full-stack task manager.",
            "Create authentication and database-backed CRUD features."
        ],
        "courses": [
            ("MDN web development", "https://developer.mozilla.org/en-US/docs/Learn"),
            ("React documentation", "https://react.dev/learn"),
            ("Node.js learning", "https://nodejs.org/en/learn")
        ],
        "certifications": [
            ("freeCodeCamp certifications", "https://www.freecodecamp.org/learn/")
        ],
        "market": "Full-stack work typically rewards frontend fundamentals, API integration, database knowledge, secure authentication and deployment."
    },

    "Cloud Engineer": {
        "skills": [
            "AWS", "Azure", "Google Cloud", "Linux", "Networking",
            "Docker", "CI/CD", "Python", "Security", "Monitoring"
        ],
        "projects": [
            "Deploy an application to a cloud platform.",
            "Containerise a service using Docker.",
            "Create a CI/CD pipeline with monitoring."
        ],
        "courses": [
            ("AWS Skill Builder", "https://skillbuilder.aws/"),
            ("Microsoft Learn Azure", "https://learn.microsoft.com/training/azure/"),
            ("Google Cloud training", "https://cloud.google.com/learn/training")
        ],
        "certifications": [
            ("AWS certifications", "https://aws.amazon.com/certification/"),
            ("Azure certifications", "https://learn.microsoft.com/credentials/")
        ],
        "market": "Cloud roles often require networking, Linux, access management, automation, monitoring, cost awareness and secure deployment."
    },

    "Cybersecurity Analyst": {
        "skills": [
            "Networking", "Linux", "Python", "SIEM",
            "Incident Response", "Threat Analysis",
            "Security Fundamentals", "Log Analysis",
            "Vulnerability Assessment", "Risk Management"
        ],
        "projects": [
            "Analyse sample security logs for suspicious events.",
            "Write a network traffic analysis report in a legal lab.",
            "Create an incident response playbook."
        ],
        "courses": [
            ("Cisco Networking Academy", "https://www.netacad.com/"),
            ("Microsoft security learning", "https://learn.microsoft.com/training/security/"),
            ("TryHackMe learning paths", "https://tryhackme.com/paths")
        ],
        "certifications": [
            ("CompTIA Security+", "https://www.comptia.org/certifications/security")
        ],
        "market": "Security work values networking fundamentals, log analysis, incident handling, risk communication and safe, authorised lab practice."
    },

    "Software Engineer": {
        "skills": [
            "Python", "Java", "Data Structures", "Algorithms",
            "OOP", "SQL", "Git", "Testing", "System Design",
            "Problem Solving"
        ],
        "projects": [
            "Build a tested application with clear architecture.",
            "Implement common data structures and algorithms.",
            "Create a database-backed service with automated tests."
        ],
        "courses": [
            ("CS50", "https://cs50.harvard.edu/x/"),
            ("LeetCode study plans", "https://leetcode.com/studyplan/"),
            ("Git handbook", "https://guides.github.com/introduction/git-handbook/")
        ],
        "certifications": [
            ("Oracle Java certifications", "https://education.oracle.com/java-se")
        ],
        "market": "Software engineering roles often evaluate problem-solving, code quality, data structures, testing and the ability to explain technical decisions."
    },

    "UI/UX Designer": {
        "skills": [
            "Figma", "Wireframing", "Prototyping", "User Research",
            "Usability Testing", "Visual Design", "Typography",
            "Design Systems", "Accessibility", "Communication"
        ],
        "projects": [
            "Redesign a mobile app and explain design decisions.",
            "Prototype a booking flow and test it with users.",
            "Build a small accessible design system."
        ],
        "courses": [
            ("Figma learning resources", "https://help.figma.com/hc/en-us/categories/360002051613-Learn-design"),
            ("Interaction Design Foundation", "https://www.interaction-design.org/"),
            ("Google UX Design", "https://grow.google/certificates/ux-design/")
        ],
        "certifications": [
            ("Google UX Design Certificate", "https://grow.google/certificates/ux-design/")
        ],
        "market": "UX roles commonly value research, clear interaction flows, accessible design, prototyping and evidence that design choices solve user problems."
    }
}


SKILL_LIBRARY = sorted(set(
    skill
    for profile in ROLE_PROFILES.values()
    for skill in profile["skills"]
) | {
    "C", "C++", "C#", "JavaScript", "TypeScript", "PHP",
    "R", "Go", "Rust", "Kotlin", "Swift", "Django",
    "Flask", "Spring Boot", "Angular", "Vue", "PostgreSQL",
    "MySQL", "MongoDB", "SQLite", "Tableau", "TensorFlow",
    "PyTorch", "Keras", "NLP", "Computer Vision",
    "Generative AI", "LLMs", "GitHub", "Kubernetes",
    "Terraform", "Jenkins", "Matplotlib", "Teamwork",
    "Leadership", "Critical Thinking", "Agile",
    "Project Management", "FastAPI", "Scikit-learn"
}))


# ============================================================
# RESUME PARSING
# ============================================================

def extract_resume_text(uploaded_file):
    if uploaded_file is None:
        return ""

    filename = uploaded_file.name.lower()

    try:
        if filename.endswith(".pdf"):
            reader = PdfReader(uploaded_file)
            return "\n".join(
                page.extract_text() or ""
                for page in reader.pages
            )

        if filename.endswith(".docx"):
            document = Document(uploaded_file)
            parts = [p.text for p in document.paragraphs]

            for table in document.tables:
                for row in table.rows:
                    parts.append(
                        " ".join(cell.text for cell in row.cells)
                    )

            return "\n".join(parts)

        if filename.endswith(".txt"):
            return uploaded_file.getvalue().decode(
                "utf-8", errors="ignore"
            )

    except Exception as error:
        st.error(f"Resume parsing failed: {error}")

    return ""


# ============================================================
# SKILL DETECTION
# ============================================================

def detect_skills(text):
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


# ============================================================
# OPTIONAL SBERT SEMANTIC MATCHING
# ============================================================

@st.cache_resource(show_spinner=False)
def load_sbert_model():
    from sentence_transformers import SentenceTransformer

    return SentenceTransformer(
        "sentence-transformers/all-MiniLM-L6-v2"
    )


def calculate_similarity(profile_text, job_text, use_sbert=False):
    if not profile_text.strip() or not job_text.strip():
        return 0, "Insufficient text"

    if use_sbert:
        try:
            model = load_sbert_model()

            embeddings = model.encode(
                [profile_text, job_text],
                normalize_embeddings=True
            )

            score = float(embeddings[0] @ embeddings[1]) * 100
            score = max(0, min(100, round(score)))

            return score, "SBERT semantic similarity"

        except Exception as error:
            st.warning(
                "SBERT could not be loaded. Using TF-IDF instead. "
                f"Reason: {type(error).__name__}"
            )

    try:
        vectorizer = TfidfVectorizer(
            stop_words="english",
            ngram_range=(1, 2)
        )

        matrix = vectorizer.fit_transform(
            [profile_text, job_text]
        )

        similarity = cosine_similarity(
            matrix[0:1], matrix[1:2]
        )[0][0]

        return round(float(similarity) * 100), "TF-IDF text similarity"

    except ValueError:
        return 0, "Insufficient text vocabulary"


# ============================================================
# MATCHING ENGINE
# ============================================================

def analyse_match(profile_text, user_skills, job_text, role, required_skills):
    user_skill_set = {skill.lower() for skill in user_skills}

    matched = [
        skill for skill in required_skills
        if skill.lower() in user_skill_set
    ]

    missing = [
        skill for skill in required_skills
        if skill.lower() not in user_skill_set
    ]

    coverage = (
        round(len(matched) / len(required_skills) * 100)
        if required_skills else 0
    )

    use_sbert = st.session_state.get("use_sbert", False)

    similarity, method = calculate_similarity(
        profile_text,
        job_text,
        use_sbert
    )

    score = round(0.65 * coverage + 0.35 * similarity)

    if not user_skills and not profile_text.strip():
        score = 0

    return {
        "role": role,
        "score": score,
        "coverage": coverage,
        "similarity": similarity,
        "method": method,
        "matched": matched,
        "missing": missing,
        "required": required_skills,
        "job_text": job_text
    }


# ============================================================
# PDF REPORT
# ============================================================

def build_pdf(report):
    pdf = FPDF()
    pdf.set_auto_page_break(auto=True, margin=15)
    pdf.add_page()

    pdf.set_font("Helvetica", "B", 20)
    pdf.set_text_color(23, 35, 45)
    pdf.cell(
        0, 12, "SkillBridge AI",
        new_x="LMARGIN", new_y="NEXT"
    )

    pdf.set_font("Helvetica", "", 11)
    pdf.cell(
        0, 8, "Career Intelligence Report",
        new_x="LMARGIN", new_y="NEXT"
    )
    pdf.ln(4)

    def clean(value):
        return str(value).encode(
            "latin-1", errors="replace"
        ).decode("latin-1")

    for heading, value in report.items():
        pdf.set_font("Helvetica", "B", 12)
        pdf.multi_cell(0, 8, clean(heading))

        pdf.set_font("Helvetica", "", 10)
        pdf.multi_cell(0, 6, clean(value))
        pdf.ln(3)

    return bytes(pdf.output())


# ============================================================
# DISPLAY HELPERS
# ============================================================

def section_heading(number, title, description):
    st.markdown(f"""
    <div class="section-heading">
      <div class="section-number">{number}</div>
      <div>
        <div class="section-title">{html.escape(title)}</div>
        <div class="section-description">{html.escape(description)}</div>
      </div>
    </div>
    """, unsafe_allow_html=True)


def render_chips(items, missing=False):
    if not items:
        st.caption("No recognised skills in this category yet.")
        return

    css_class = "skill-chip missing" if missing else "skill-chip"

    chips = "".join(
        f'<span class="{css_class}">{html.escape(skill)}</span>'
        for skill in items
    )

    st.markdown(chips, unsafe_allow_html=True)


def render_metric(label, value, note):
    st.markdown(f"""
    <div class="metric-card">
      <div class="metric-label">{html.escape(label)}</div>
      <div class="metric-value">{html.escape(str(value))}</div>
      <div class="metric-note">{html.escape(note)}</div>
    </div>
    """, unsafe_allow_html=True)


# ============================================================
# HEADER AND HERO
# ============================================================

st.markdown("""
<div class="brand-row">
  <div class="brand-lockup">
    <div class="brand-mark">✦</div>
    <div>
      <div class="brand-name">SkillBridge AI</div>
      <div class="brand-subtitle">Career Intelligence</div>
    </div>
  </div>
  <div class="brand-status">YOUR PERSONAL CAREER WORKSPACE</div>
</div>

<div class="hero">
  <div class="hero-eyebrow">A more considered approach to your career</div>
  <h1>Build your next career move<br><em>with clarity.</em></h1>
  <div class="hero-copy">
    Your skills, experience, and ambitions — brought together in one
    personal career workspace. Discover role alignment, identify skill
    gaps, and turn your next steps into a practical plan.
  </div>
  <div class="hero-features">
    <span class="hero-feature">✦ Resume intelligence</span>
    <span class="hero-feature">Skill gap mapping</span>
    <span class="hero-feature">Semantic job matching</span>
    <span class="hero-feature">Personal learning roadmap</span>
  </div>
</div>
""", unsafe_allow_html=True)


# ============================================================
# INPUT 01 — PROFILE
# No separate experience-years input.
# ============================================================

section_heading(
    "01",
    "Tell us about yourself",
    "Build your profile so your recommendations have context."
)

with st.container(border=True):
    col1, col2 = st.columns(2)

    with col1:
        full_name = st.text_input(
            "Full name",
            placeholder="Enter your name"
        )

        education = st.selectbox(
            "Education",
            [
                "Select education",
                "Diploma",
                "Bachelor's degree",
                "Master's degree",
                "PhD",
                "Other"
            ]
        )

    with col2:
        experience_level = st.selectbox(
            "Experience level",
            [
                "Fresher / Student",
                "Entry-level professional",
                "Mid-level professional",
                "Senior professional",
                "Career switcher"
            ]
        )

        current_status = st.selectbox(
            "Current status",
            [
                "Student",
                "Looking for a job",
                "Currently employed",
                "Freelancer",
                "Career switcher"
            ]
        )

    interests = st.text_input(
        "Career interests (optional)",
        placeholder="e.g. building apps, analysing data, research"
    )


# ============================================================
# INPUT 02 — RESUME AND SKILLS
# No manual resume-text entry field.
# ============================================================

section_heading(
    "02",
    "Your resume and skills",
    "Upload your resume and select the skills you can demonstrate."
)

with st.container(border=True):
    uploaded_resume = st.file_uploader(
        "Upload your resume",
        type=["pdf", "docx", "txt"],
        help="Supported formats: PDF, DOCX and TXT."
    )

    known_skills = st.multiselect(
        "Skills you already know",
        options=SKILL_LIBRARY,
        placeholder="Search and select your skills..."
    )

    custom_skills = st.text_input(
        "Additional skills (optional)",
        placeholder="Enter comma-separated skills not listed above"
    )

    skill_confidence = st.slider(
        "Self-rated confidence in your selected skills",
        min_value=1,
        max_value=5,
        value=3
    )

    st.caption(
        "Your confidence rating is contextual information, "
        "not a verified proficiency test."
    )


# ============================================================
# INPUT 03 — TARGET OPPORTUNITY
# ============================================================

section_heading(
    "03",
    "Choose your target opportunity",
    "Compare your profile with a role or a specific job posting."
)

with st.container(border=True):
    target_mode = st.radio(
        "Comparison mode",
        [
            "Career role",
            "Specific job description",
            "Compare multiple roles"
        ],
        horizontal=True
    )

    if target_mode == "Career role":
        target_role = st.selectbox(
            "Target role",
            list(ROLE_PROFILES.keys())
        )
        job_description = ""

    elif target_mode == "Specific job description":
        target_role = "Custom job description"

        job_description = st.text_area(
            "Paste the job description",
            placeholder=(
                "Paste responsibilities, required skills "
                "and qualifications from the job posting..."
            ),
            height=160
        )

    else:
        target_role = st.selectbox(
            "Primary role for the roadmap",
            list(ROLE_PROFILES.keys())
        )

        job_description = st.text_area(
            "Optional job description",
            placeholder="Paste a job description to include in the comparison...",
            height=120
        )


with st.expander("Advanced matching settings"):
    use_sbert = st.checkbox(
        "Use SBERT semantic matching",
        value=False,
        help=(
            "Downloads a sentence-transformer model on first use. "
            "If unavailable, the app falls back to TF-IDF."
        )
    )

    st.session_state["use_sbert"] = use_sbert

    st.caption(
        "Explainable score: 65% recognised skill coverage + "
        "35% text similarity. This is a heuristic, not a hiring probability."
    )


generate = st.button(
    "✦  Generate career intelligence report",
    use_container_width=True
)


# ============================================================
# RUN ANALYSIS
# ============================================================

if generate:

    if not full_name.strip():
        st.error("Please enter your name.")

    elif (
        target_mode == "Specific job description"
        and not job_description.strip()
    ):
        st.error("Please paste a job description to analyse.")

    else:
        with st.spinner("Analysing your profile and target opportunity..."):

            resume_text = extract_resume_text(uploaded_resume)

            custom_skill_list = [
                skill.strip()
                for skill in custom_skills.split(",")
                if skill.strip()
            ]

            resume_skills = detect_skills(resume_text)

            user_skills = sorted(
                set(known_skills + custom_skill_list + resume_skills),
                key=str.lower
            )

            profile_text = " ".join([
                resume_text,
                " ".join(user_skills),
                interests,
                education,
                experience_level,
                current_status
            ])

            if target_mode == "Specific job description":
                primary_role = "Custom job description"
                primary_job_text = job_description
                primary_required = detect_skills(job_description)

            else:
                primary_role = target_role
                role_data = ROLE_PROFILES[target_role]

                primary_job_text = " ".join(role_data["skills"])
                primary_required = list(role_data["skills"])

                if job_description.strip():
                    primary_job_text += " " + job_description
                    primary_required = sorted(set(
                        primary_required + detect_skills(job_description)
                    ))

            primary = analyse_match(
                profile_text,
                user_skills,
                primary_job_text,
                primary_role,
                primary_required
            )

            if target_mode == "Compare multiple roles":
                comparisons = []

                for role, data in ROLE_PROFILES.items():
                    comparisons.append(
                        analyse_match(
                            profile_text,
                            user_skills,
                            " ".join(data["skills"]),
                            role,
                            data["skills"]
                        )
                    )

                if job_description.strip():
                    comparisons.append(
                        analyse_match(
                            profile_text,
                            user_skills,
                            job_description,
                            "Pasted job description",
                            detect_skills(job_description)
                        )
                    )

                comparisons.sort(
                    key=lambda item: item["score"],
                    reverse=True
                )

            else:
                comparisons = [primary]

            if primary["score"] >= 75:
                fit_label = "Strong profile overlap"
                fit_copy = (
                    "Your profile overlaps with many recognised requirements. "
                    "Strengthen your evidence with projects and measurable outcomes."
                )

            elif primary["score"] >= 45:
                fit_label = "Developing profile overlap"
                fit_copy = (
                    "You have some relevant overlap. Focus on priority gaps "
                    "and demonstrate your skills through practical work."
                )

            else:
                fit_label = "Early-stage profile overlap"
                fit_copy = (
                    "Several requirements may need development. Start with "
                    "the foundations and build evidence through projects."
                )

            st.session_state["career_results"] = {
                "name": full_name.strip(),
                "education": education,
                "experience_level": experience_level,
                "status": current_status,
                "interests": interests,
                "confidence": skill_confidence,
                "resume_text": resume_text,
                "user_skills": user_skills,
                "primary": primary,
                "comparisons": comparisons,
                "fit_label": fit_label,
                "fit_copy": fit_copy,
                "target_mode": target_mode,
                "target_role": target_role,
                "date": date.today().isoformat()
            }

        st.success("Your career analysis is ready.")


# ============================================================
# RESULTS
# ============================================================

results = st.session_state.get("career_results")

if results:

    primary = results["primary"]

    section_heading(
        "04",
        "Your career snapshot",
        f"Analysis for {results['name']} · {primary['role']}"
    )

    st.markdown(f"""
    <div class="result-banner">
      <strong>{html.escape(results['fit_label'])}</strong>
      <p>{html.escape(results['fit_copy'])}
      The score is an explainable heuristic, not a validated hiring prediction.</p>
    </div>
    """, unsafe_allow_html=True)

    m1, m2, m3, m4 = st.columns(4)

    with m1:
        render_metric(
            "MATCH INDICATOR",
            f"{primary['score']}%",
            "Weighted profile overlap"
        )

    with m2:
        render_metric(
            "SKILL COVERAGE",
            f"{primary['coverage']}%",
            f"{len(primary['matched'])} of {len(primary['required'])} requirements"
        )

    with m3:
        render_metric(
            "MATCHED SKILLS",
            len(primary["matched"]),
            "Recognised target skills in your profile"
        )

    with m4:
        render_metric(
            "SKILL GAPS",
            len(primary["missing"]),
            "Requirements to prioritise"
        )

    st.markdown("")

    with st.container(border=True):
        st.markdown(
            '<div class="panel-title">Explainable match score</div>',
            unsafe_allow_html=True
        )

        st.markdown(
            f"- **Recognised skill coverage (65%):** {primary['coverage']}%.\n"
            f"- **Text similarity (35%):** {primary['similarity']}% using {primary['method']}.\n"
            f"- **Final indicator:** 65% × skill coverage + 35% × text similarity."
        )

        st.caption(
            "Keyword extraction may miss synonyms. Self-reported skills are "
            "not independently verified, and a high score does not guarantee employment."
        )


    # --------------------------------------------------------
    # SKILL STRENGTH AND MISSING SKILLS
    # --------------------------------------------------------

    section_heading(
        "05",
        "Skill strength and missing skills",
        "Review your existing overlap and prioritise the gaps."
    )

    left, right = st.columns(2)

    with left:
        with st.container(border=True):
            st.markdown(
                '<div class="panel-title">Skills already aligned</div>',
                unsafe_allow_html=True
            )
            render_chips(primary["matched"])

    with right:
        with st.container(border=True):
            st.markdown(
                '<div class="panel-title">Skills to prioritise</div>',
                unsafe_allow_html=True
            )
            render_chips(primary["missing"], missing=True)

    if primary["required"]:
        gap_df = pd.DataFrame([
            {
                "Requirement": skill,
                "Assessment": (
                    "Matched"
                    if skill in primary["matched"]
                    else "Gap"
                ),
                "Priority": (
                    "Maintain / deepen"
                    if skill in primary["matched"]
                    else "High"
                )
            }
            for skill in primary["required"]
        ])

        st.dataframe(
            gap_df,
            hide_index=True,
            use_container_width=True
        )

    else:
        st.info(
            "No recognised skills were extracted from this job description. "
            "Include explicit skills to improve the gap analysis."
        )


    # --------------------------------------------------------
    # MULTI-JOB COMPARISON
    # --------------------------------------------------------

    if results["target_mode"] == "Compare multiple roles":

        section_heading(
            "06",
            "Multi-role comparison",
            "Compare your profile with different career paths."
        )

        comparison_df = pd.DataFrame([
            {
                "Role": item["role"],
                "Match indicator": f"{item['score']}%",
                "Skill coverage": f"{item['coverage']}%",
                "Matched skills": len(item["matched"]),
                "Skill gaps": len(item["missing"])
            }
            for item in results["comparisons"]
        ])

        st.dataframe(
            comparison_df,
            hide_index=True,
            use_container_width=True
        )

        st.caption(
            "Scores are comparative heuristics, not employment probabilities."
        )


    # --------------------------------------------------------
    # CAREER SWITCH SIMULATOR
    # --------------------------------------------------------

    section_heading(
        "07",
        "Career switch simulator",
        "Explore how your profile aligns with another career direction."
    )

    switch_role = st.selectbox(
        "Simulate a move into",
        list(ROLE_PROFILES.keys()),
        key="switch_role"
    )

    switch_data = ROLE_PROFILES[switch_role]

    switch_match = analyse_match(
        " ".join([
            results["resume_text"],
            " ".join(results["user_skills"]),
            results["interests"],
            results["education"],
            results["experience_level"]
        ]),
        results["user_skills"],
        " ".join(switch_data["skills"]),
        switch_role,
        switch_data["skills"]
    )

    s1, s2, s3 = st.columns(3)

    with s1:
        render_metric(
            "SWITCH ALIGNMENT",
            f"{switch_match['score']}%",
            switch_role
        )

    with s2:
        render_metric(
            "TRANSFERABLE SKILLS",
            len(switch_match["matched"]),
            "Skills already aligned"
        )

    with s3:
        render_metric(
            "NEW SKILLS",
            len(switch_match["missing"]),
            "Skills to develop"
        )

    with st.container(border=True):
        st.markdown(
            '<div class="panel-title">Transferable strengths</div>',
            unsafe_allow_html=True
        )
        render_chips(switch_match["matched"])

        st.markdown(
            '<div class="panel-title">Suggested transition skills</div>',
            unsafe_allow_html=True
        )
        render_chips(switch_match["missing"], missing=True)


    # --------------------------------------------------------
    # CAREER READINESS ASSESSMENT
    # --------------------------------------------------------

    section_heading(
        "08",
        "Career readiness assessment",
        "Review the essentials before applying for a role."
    )

    readiness_items = [
        (
            "Resume uploaded",
            bool(results["resume_text"].strip()),
            "Upload a resume with relevant experience and projects."
        ),
        (
            "Core skills identified",
            len(results["user_skills"]) >= 3,
            "Add at least three skills you can demonstrate."
        ),
        (
            "Role requirements reviewed",
            bool(primary["required"]),
            "Use a role profile or a job description with clear requirements."
        ),
        (
            "Project evidence prepared",
            False,
            "Add portfolio links and measurable project outcomes."
        )
    ]

    readiness_score = round(
        sum(1 for _, done, _ in readiness_items if done)
        / len(readiness_items) * 100
    )

    render_metric(
        "READINESS CHECKLIST",
        f"{readiness_score}%",
        "Checklist completion, not a validated readiness test"
    )

    for label, done, advice in readiness_items:
        st.checkbox(
            f"{'✓' if done else '○'} {label} — {advice}",
            value=done,
            key="ready_" + re.sub(
                r"[^a-z0-9]+", "_", label.lower()
            )
        )


    # --------------------------------------------------------
    # LEARNING LINKS, CERTIFICATIONS AND PROJECTS
    # --------------------------------------------------------

    section_heading(
        "09",
        "Your learning roadmap",
        "Turn skill gaps into a practical learning plan."
    )

    role_data = ROLE_PROFILES.get(
        primary["role"],
        ROLE_PROFILES.get(
            results["target_role"],
            ROLE_PROFILES["Software Engineer"]
        )
    )

    if primary["missing"]:
        roadmap = [
            (
                "01",
                "Build the foundations",
                "Prioritise " + ", ".join(primary["missing"][:3]) +
                " and practise the basic concepts."
            ),
            (
                "02",
                "Apply the skills",
                "Complete a small project that demonstrates the new skills."
            ),
            (
                "03",
                "Show your evidence",
                "Document your approach, publish your project and update your resume."
            )
        ]
    else:
        roadmap = [
            (
                "01",
                "Deepen your skills",
                "Practise advanced tasks relevant to the target role."
            ),
            (
                "02",
                "Build portfolio evidence",
                "Create a realistic project and document your decisions."
            ),
            (
                "03",
                "Prepare for applications",
                "Tailor your resume and prepare examples of your impact."
            )
        ]

    for number, title, copy in roadmap:
        st.markdown(f"""
        <div class="roadmap-card">
          <div class="roadmap-number">{number}</div>
          <div>
            <div class="roadmap-title">{html.escape(title)}</div>
            <div class="roadmap-copy">{html.escape(copy)}</div>
          </div>
        </div>
        """, unsafe_allow_html=True)

    course_col, project_col = st.columns(2)

    with course_col:
        with st.container(border=True):
            st.markdown(
                '<div class="panel-title">Course recommendations</div>',
                unsafe_allow_html=True
            )

            for label, url in role_data["courses"]:
                st.markdown(f"- [{label}]({url})")

            st.markdown("")
            st.markdown(
                '<div class="panel-title">Certification recommendations</div>',
                unsafe_allow_html=True
            )

            for label, url in role_data["certifications"]:
                st.markdown(f"- [{label}]({url})")

    with project_col:
        with st.container(border=True):
            st.markdown(
                '<div class="panel-title">Project recommendations</div>',
                unsafe_allow_html=True
            )

            for index, project in enumerate(role_data["projects"], 1):
                st.markdown(f"**{index:02d}.** {project}")

            st.markdown("")
            st.markdown(
                '<div class="panel-title">Personalised career guidance</div>',
                unsafe_allow_html=True
            )

            if primary["missing"]:
                st.write(
                    "Start with **"
                    + ", ".join(primary["missing"][:3])
                    + "**. Practise through exercises, then demonstrate "
                    "your skills in a portfolio project."
                )
            else:
                st.write(
                    "Your provided skills cover the recognised requirements "
                    "in this role template. Deepen your skills, tailor your "
                    "resume and prepare examples of your impact."
                )


    # --------------------------------------------------------
    # JOB-MARKET CONTEXT
    # --------------------------------------------------------

    section_heading(
        "10",
        "Job-market context",
        "General role context and links to external job-market resources."
    )

    if primary["role"] == "Custom job description":
        market_text = (
            "For this custom posting, review the employer's requirements, "
            "location, seniority and posting date. This app does not retrieve "
            "live vacancy counts or salary statistics."
        )
    else:
        market_text = role_data["market"]

    with st.container(border=True):
        st.write(market_text)

        st.markdown(
            "[Search jobs on LinkedIn](https://www.linkedin.com/jobs/) · "
            "[Search jobs on Indeed India](https://in.indeed.com/) · "
            "[Explore occupations on O*NET](https://www.onetonline.org/)"
        )

        st.caption(
            "Market context is general guidance, not live market analytics. "
            "Vacancy counts and salaries vary by location and time."
        )


    # --------------------------------------------------------
    # EXPORT REPORT
    # --------------------------------------------------------

    section_heading(
        "11",
        "Your career report",
        "Download your analysis for future reference."
    )

    report = {
        "Candidate": results["name"],
        "Education": results["education"],
        "Experience level": results["experience_level"],
        "Current status": results["status"],
        "Target opportunity": primary["role"],
        "Report date": results["date"],
        "Match indicator": (
            f"{primary['score']}% — heuristic, not a hiring probability"
        ),
        "Skill coverage": f"{primary['coverage']}%",
        "Text similarity": (
            f"{primary['similarity']}% using {primary['method']}"
        ),
        "Matched skills": (
            ", ".join(primary["matched"]) or "None recognised"
        ),
        "Missing skills": (
            ", ".join(primary["missing"]) or "None recognised"
        ),
        "Identified skills": (
            ", ".join(results["user_skills"]) or "None identified"
        ),
        "Recommended projects": "\n".join(
            "- " + item for item in role_data["projects"]
        ),
        "Recommended courses": "\n".join(
            label + ": " + url
            for label, url in role_data["courses"]
        ),
        "Recommended certifications": "\n".join(
            label + ": " + url
            for label, url in role_data["certifications"]
        ),
        "Scoring methodology": (
            "65% recognised skill coverage + 35% text similarity. "
            "Not a validated recruitment prediction."
        )
    }

    report_df = pd.DataFrame([
        {"Section": key, "Details": value}
        for key, value in report.items()
    ])

    export_left, export_right = st.columns(2)

    with export_left:
        st.download_button(
            "Download CSV report",
            data=report_df.to_csv(index=False).encode("utf-8-sig"),
            file_name="skillbridge_career_report.csv",
            mime="text/csv",
            use_container_width=True
        )

    with export_right:
        try:
            pdf_bytes = build_pdf(report)

            st.download_button(
                "Download PDF report",
                data=pdf_bytes,
                file_name="skillbridge_career_report.pdf",
                mime="application/pdf",
                use_container_width=True
            )

        except Exception as error:
            st.warning(f"PDF export failed: {error}")

    st.caption(
        "SkillBridge AI provides educational guidance based on supplied "
        "information. It does not guarantee interviews, employment or salary outcomes."
    )


# ============================================================
# FOOTER
# ============================================================

st.markdown("""
<div class="footer">
  <span>✦ SkillBridge AI · Career Intelligence</span>
  <span>Understand your skills. Plan your next move.</span>
</div>
""", unsafe_allow_html=True)
