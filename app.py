
import re
from datetime import date

import pandas as pd
import streamlit as st
from sklearn.feature_extraction.text import TfidfVectorizer
from sklearn.metrics.pairwise import cosine_similarity
from fpdf import FPDF

try:
    from pypdf import PdfReader
except ImportError:
    PdfReader = None

try:
    from docx import Document
except ImportError:
    Document = None


# ============================================================
# PAGE CONFIGURATION
# ============================================================

st.set_page_config(
    page_title="SkillBridge AI | Career Intelligence",
    page_icon="✦",
    layout="wide",
    initial_sidebar_state="collapsed",
)


# ============================================================
# CUSTOM UI DESIGN
# ============================================================

st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=DM+Sans:wght@400;500;600;700&family=Manrope:wght@400;500;600;700;800&display=swap');

html, body, [class*="css"] {
    font-family: 'DM Sans', sans-serif;
}

.stApp {
    background:
        radial-gradient(ellipse at 10% 0%, rgba(91, 78, 230, .20), transparent 32%),
        radial-gradient(ellipse at 95% 12%, rgba(30, 170, 190, .10), transparent 28%),
        #080C18;
    color: #F0F3FF;
}

.block-container {
    max-width: 1220px;
    padding-top: 2rem;
    padding-bottom: 4rem;
}

h1, h2, h3, h4 {
    font-family: 'Manrope', sans-serif;
    letter-spacing: -0.035em;
}

#MainMenu, footer, header {
    visibility: hidden;
}

[data-testid="stSidebar"] {
    display: none;
}

.hero {
    padding: 2.4rem;
    border: 1px solid rgba(164, 174, 255, .20);
    border-radius: 26px;
    background: linear-gradient(
        115deg,
        rgba(36, 39, 88, .96),
        rgba(15, 29, 49, .94) 65%,
        rgba(13, 52, 62, .86)
    );
    box-shadow: 0 22px 70px rgba(0, 0, 0, .18);
    margin-bottom: 2rem;
}

.brand {
    color: #B8C0FF;
    text-transform: uppercase;
    letter-spacing: .18em;
    font-size: .75rem;
    font-weight: 800;
}

.hero h1 {
    font-size: clamp(2.3rem, 5vw, 3.7rem);
    line-height: 1.08;
    margin: .8rem 0;
}

.hero p {
    color: #C3CBE0;
    font-size: 1.02rem;
    line-height: 1.7;
    max-width: 740px;
}

.hero-note {
    margin-top: 1.3rem;
    color: #DCE5FF;
    font-size: .84rem;
}

.section-heading {
    font-size: 1.3rem;
    font-weight: 800;
    color: #F5F7FF;
    margin-bottom: .3rem;
}

.section-description {
    font-size: .88rem;
    color: #929DB8;
    margin-bottom: 1.2rem;
}

.panel {
    padding: 1.3rem;
    border: 1px solid rgba(148, 163, 184, .17);
    border-radius: 20px;
    background: linear-gradient(
        145deg,
        rgba(20, 28, 49, .90),
        rgba(14, 21, 37, .86)
    );
    margin-bottom: 1.2rem;
}

.step-number {
    display: inline-block;
    padding: .45rem .7rem;
    border-radius: 10px;
    background: linear-gradient(135deg, #8175FF, #4B9EEA);
    color: white;
    font-weight: 800;
    margin-right: .55rem;
}

.chip {
    display: inline-block;
    border-radius: 999px;
    border: 1px solid rgba(135, 149, 255, .25);
    background: rgba(101, 105, 236, .13);
    color: #DCE3FF;
    padding: 5px 10px;
    margin: 3px 4px 3px 0;
    font-size: .8rem;
}

.result-panel {
    padding: 1.2rem;
    border-radius: 18px;
    border: 1px solid rgba(148, 163, 184, .16);
    background: rgba(18, 27, 47, .92);
}

.stButton > button,
.stDownloadButton > button {
    border-radius: 12px;
    min-height: 2.9rem;
    font-weight: 800;
}

.stButton > button[kind="primary"] {
    background: linear-gradient(100deg, #7667FF, #398BE9);
    color: white;
    border: 0;
}

.stTextInput input,
.stTextArea textarea,
.stSelectbox div[data-baseweb="select"] > div,
.stMultiSelect div[data-baseweb="select"] > div {
    background-color: rgba(7, 13, 26, .70) !important;
    border-color: rgba(148, 163, 184, .24) !important;
    border-radius: 12px !important;
}

[data-testid="stFileUploader"] {
    border: 1px dashed rgba(139, 151, 255, .38);
    padding: .6rem;
    border-radius: 14px;
    background: rgba(8, 14, 28, .42);
}

div[data-testid="stMetric"] {
    background: rgba(18, 27, 47, .92);
    border: 1px solid rgba(148, 163, 184, .16);
    padding: 15px;
    border-radius: 16px;
}

div[data-testid="stMetricLabel"] {
    color: #AAB5CF;
}

hr {
    border-color: rgba(148, 163, 184, .15);
}

.small-muted {
    color: #8E9AB5;
    font-size: .82rem;
}
</style>
""", unsafe_allow_html=True)


# ============================================================
# SKILL LIBRARY
# ============================================================

SKILL_GROUPS = {
    "Programming": [
        "python", "java", "javascript", "typescript", "c++",
        "c#", "go", "rust", "php", "kotlin", "swift"
    ],
    "Data & AI": [
        "machine learning", "deep learning", "artificial intelligence",
        "data science", "data analysis", "pandas", "numpy",
        "scikit-learn", "tensorflow", "pytorch", "nlp",
        "computer vision", "generative ai", "llm", "statistics",
        "power bi", "tableau", "excel", "sql"
    ],
    "Web Development": [
        "html", "css", "react", "angular", "vue", "node.js",
        "express", "django", "flask", "fastapi", "streamlit",
        "rest api", "graphql"
    ],
    "Cloud & DevOps": [
        "aws", "azure", "google cloud", "docker", "kubernetes",
        "terraform", "jenkins", "ci/cd", "linux", "git",
        "github", "serverless"
    ],
    "Databases": [
        "mysql", "postgresql", "mongodb", "redis", "sqlite",
        "database design"
    ],
    "Cybersecurity": [
        "cybersecurity", "network security", "penetration testing",
        "ethical hacking", "cryptography", "owasp", "siem"
    ],
    "Core & Professional": [
        "data structures", "algorithms",
        "object-oriented programming", "problem solving",
        "communication", "teamwork", "leadership",
        "agile", "scrum", "testing", "unit testing", "debugging"
    ]
}

ALL_SKILLS = sorted(
    {skill for group in SKILL_GROUPS.values() for skill in group},
    key=len,
    reverse=True
)


# ============================================================
# CAREER ROLE PROFILES
# ============================================================

ROLES = {
    "Data Analyst": {
        "skills": [
            "sql", "excel", "python", "pandas", "statistics",
            "power bi", "data analysis", "communication"
        ],
        "projects": [
            "Build an interactive sales dashboard",
            "Analyze a public dataset and report business insights",
            "Create a SQL business case study"
        ],
        "courses": [
            "Google Data Analytics Certificate",
            "Microsoft Power BI Data Analyst learning path"
        ]
    },

    "Machine Learning Engineer": {
        "skills": [
            "python", "machine learning", "statistics",
            "scikit-learn", "pandas", "numpy",
            "deep learning", "git", "rest api"
        ],
        "projects": [
            "Train and evaluate a classification model",
            "Deploy a machine learning model as an API",
            "Build a reproducible machine learning pipeline"
        ],
        "courses": [
            "Machine Learning Specialization",
            "TensorFlow learning path"
        ]
    },

    "Python Developer": {
        "skills": [
            "python", "data structures", "algorithms",
            "object-oriented programming", "git", "sql",
            "rest api", "testing", "django"
        ],
        "projects": [
            "Build a CRUD REST API",
            "Create a test-covered automation tool",
            "Build a Django application with authentication"
        ],
        "courses": [
            "Python programming specialization",
            "Python certification learning path"
        ]
    },

    "Full Stack Developer": {
        "skills": [
            "javascript", "html", "css", "react",
            "node.js", "rest api", "sql", "git", "testing"
        ],
        "projects": [
            "Build a full-stack task manager",
            "Create an authenticated e-commerce prototype",
            "Deploy a responsive portfolio and API"
        ],
        "courses": [
            "Meta Front-End Developer",
            "Full-stack web development course"
        ]
    },

    "Cloud Engineer": {
        "skills": [
            "linux", "aws", "azure", "docker",
            "kubernetes", "git", "terraform", "ci/cd"
        ],
        "projects": [
            "Deploy a containerized web application",
            "Build a CI/CD pipeline",
            "Provision infrastructure using Terraform"
        ],
        "courses": [
            "AWS Cloud Practitioner",
            "Microsoft Azure Fundamentals"
        ]
    },

    "Cybersecurity Analyst": {
        "skills": [
            "network security", "linux", "cybersecurity",
            "siem", "cryptography", "python", "owasp"
        ],
        "projects": [
            "Build a home-lab security monitoring dashboard",
            "Document a web application threat model",
            "Analyze sample logs for suspicious activity"
        ],
        "courses": [
            "CompTIA Security+",
            "Google Cybersecurity Certificate"
        ]
    },

    "Software Engineer": {
        "skills": [
            "python", "java", "data structures", "algorithms",
            "object-oriented programming", "git", "sql",
            "testing", "problem solving"
        ],
        "projects": [
            "Build a well-tested application",
            "Document algorithmic problem solutions",
            "Create a REST API with persistent storage"
        ],
        "courses": [
            "Software design specialization",
            "Git and GitHub learning path"
        ]
    }
}


COURSE_SUGGESTIONS = {
    "python": "Python fundamentals and intermediate programming",
    "sql": "SQL queries, joins, aggregation, and window functions",
    "machine learning": "Supervised learning and model evaluation",
    "pandas": "Data manipulation with pandas",
    "power bi": "Data modeling and dashboards in Power BI",
    "javascript": "Modern JavaScript and asynchronous programming",
    "react": "React components, hooks, and state management",
    "docker": "Docker images, containers, and Compose",
    "aws": "AWS core services and cloud architecture",
    "git": "Git branching, pull requests, and collaboration",
    "statistics": "Statistics, probability, and hypothesis testing",
    "scikit-learn": "Machine learning pipelines with scikit-learn",
    "rest api": "REST API design, authentication, and testing",
    "linux": "Linux command line fundamentals",
    "html": "Semantic HTML and accessible page structure",
    "css": "Responsive layouts and modern CSS",
    "data structures": "Arrays, trees, graphs, and complexity analysis"
}


# ============================================================
# HELPER FUNCTIONS
# ============================================================

def extract_resume_text(uploaded_file):
    if uploaded_file is None:
        return ""

    try:
        filename = uploaded_file.name.lower()

        if filename.endswith(".pdf") and PdfReader:
            reader = PdfReader(uploaded_file)
            return "\n".join(
                page.extract_text() or ""
                for page in reader.pages
            )

        if filename.endswith(".docx") and Document:
            document = Document(uploaded_file)
            return "\n".join(
                paragraph.text
                for paragraph in document.paragraphs
            )

        if filename.endswith(".txt"):
            return uploaded_file.getvalue().decode(
                "utf-8", errors="ignore"
            )

    except Exception as error:
        st.error(f"Unable to read the uploaded resume: {error}")

    return ""


def extract_skills(text):
    text = (text or "").lower()

    aliases = {
        "sklearn": "scikit-learn",
        "nodejs": "node.js",
        "postgres": "postgresql",
        "powerbi": "power bi",
        "gcp": "google cloud",
        "oop": "object-oriented programming"
    }

    for alias, canonical in aliases.items():
        text = re.sub(
            r"\b" + re.escape(alias) + r"\b",
            canonical,
            text
        )

    detected = []

    for skill in ALL_SKILLS:
        pattern = (
            r"(?<![a-z0-9+#])"
            + re.escape(skill)
            + r"(?![a-z0-9+#])"
        )

        if re.search(pattern, text):
            detected.append(skill)

    return sorted(set(detected))


def calculate_similarity(resume_text, job_text):
    if not resume_text.strip() or not job_text.strip():
        return 0.0

    try:
        vectorizer = TfidfVectorizer(
            stop_words="english",
            ngram_range=(1, 2),
            max_features=5000
        )

        matrix = vectorizer.fit_transform(
            [resume_text, job_text]
        )

        score = cosine_similarity(
            matrix[0:1],
            matrix[1:2]
        )[0][0] * 100

        return round(float(score), 1)

    except ValueError:
        return 0.0


def generate_pdf_report(report):
    pdf = FPDF()
    pdf.set_auto_page_break(auto=True, margin=15)
    pdf.add_page()

    pdf.set_font("Helvetica", "B", 20)
    pdf.cell(0, 12, "SkillBridge AI - Career Report", ln=True)

    pdf.set_font("Helvetica", size=10)
    pdf.cell(
        0, 7,
        f"Generated: {date.today().isoformat()}",
        ln=True
    )

    pdf.ln(4)
    pdf.set_font("Helvetica", "B", 13)
    pdf.cell(0, 9, "Candidate: " + report["name"][:65], ln=True)
    pdf.cell(0, 9, "Target role: " + report["role"][:65], ln=True)
    pdf.cell(
        0, 9,
        f'Career fit indicator: {report["score"]:.1f}%',
        ln=True
    )

    sections = [
        ("Profile skills", report["resume_skills"]),
        ("Skills selected by candidate", report["known_skills"]),
        ("Target role skills", report["required_skills"]),
        ("Matched skills", report["matched_skills"]),
        ("Potential skill gaps", report["missing_skills"])
    ]

    for heading, items in sections:
        pdf.ln(3)
        pdf.set_font("Helvetica", "B", 12)
        pdf.multi_cell(0, 8, heading)

        pdf.set_font("Helvetica", size=10)

        for item in items or ["None detected"]:
            safe_item = (
                str(item)
                .encode("latin-1", "replace")
                .decode("latin-1")
            )
            pdf.multi_cell(0, 6, "- " + safe_item)

    return bytes(pdf.output())


# ============================================================
# HEADER
# ============================================================

st.markdown("""
<div class="hero">
    <div class="brand">✦ SkillBridge AI / Career intelligence</div>

    <h1>Build your next career move<br>with clarity.</h1>

    <p>
        Your skills, experience, and ambitions — brought together
        in one personal career workspace. Discover your role alignment,
        identify skill gaps, and find your next steps.
    </p>

    <div class="hero-note">
        ✦ Resume insights &nbsp; · &nbsp;
        Skill gap mapping &nbsp; · &nbsp;
        Personal learning roadmap
    </div>
</div>
""", unsafe_allow_html=True)


# ============================================================
# SECTION 1 — PERSONAL PROFILE
# ============================================================

st.markdown(
    '<div class="section-heading">'
    '<span class="step-number">01</span> Tell us about yourself'
    '</div>'
    '<div class="section-description">'
    'Build your profile so your career recommendations have context.'
    '</div>',
    unsafe_allow_html=True
)

with st.container(border=True):
    col1, col2, col3 = st.columns(3, gap="large")

    with col1:
        candidate_name = st.text_input(
            "Full name",
            placeholder="e.g. Alex Morgan"
        )

        education = st.selectbox(
            "Current education",
            [
                "Choose one...",
                "Undergraduate student",
                "Postgraduate student",
                "Graduate",
                "Diploma",
                "Working professional",
                "Other"
            ]
        )

    with col2:
        experience_level = st.selectbox(
            "Experience level",
            [
                "Fresher / no professional experience",
                "Internship experience",
                "0–2 years",
                "2–5 years",
                "5+ years"
            ]
        )

        years_experience = st.number_input(
            "Years of experience",
            min_value=0.0,
            max_value=50.0,
            value=0.0,
            step=0.5
        )

    with col3:
        current_status = st.selectbox(
            "Current status",
            [
                "Student",
                "Looking for first job",
                "Employed",
                "Career switcher",
                "Returning to work"
            ]
        )

        target_role = st.selectbox(
            "Career goal",
            list(ROLES.keys()) + ["Custom job description"]
        )


# ============================================================
# SECTION 2 — RESUME AND KNOWN SKILLS
# ============================================================

st.markdown(
    '<div class="section-heading">'
    '<span class="step-number">02</span> Your resume and skills'
    '</div>'
    '<div class="section-description">'
    'Upload your resume and select the skills you already know.'
    '</div>',
    unsafe_allow_html=True
)

with st.container(border=True):
    left, right = st.columns([1, 1], gap="large")

    with left:
        uploaded_resume = st.file_uploader(
            "Upload your resume",
            type=["pdf", "docx", "txt"],
            help="Supported formats: PDF, DOCX, and TXT."
        )

        extracted_text = (
            extract_resume_text(uploaded_resume)
            if uploaded_resume else ""
        )

        resume_text = st.text_area(
            "Resume content",
            value=extracted_text,
            height=220,
            placeholder=(
                "Paste your resume here if you do not have a file. "
                "Include your education, projects, internships, "
                "experience, and technical skills."
            )
        )

        if uploaded_resume and extracted_text:
            st.caption(
                f"Text extracted from {uploaded_resume.name}. "
                "You can edit the extracted content above."
            )

    with right:
        known_skills = st.multiselect(
            "Skills you already know",
            options=ALL_SKILLS,
            help=(
                "Select the skills you can genuinely use. "
                "Add skills outside the list below."
            )
        )

        custom_skills = st.text_input(
            "Other skills",
            placeholder="e.g. Spring Boot, Figma, Excel macros"
        )

        confidence = st.select_slider(
            "How confident are you in your skills?",
            options=[
                "Beginner",
                "Learning",
                "Comfortable",
                "Confident"
            ],
            value="Learning"
        )

        interests = st.text_input(
            "Interests or preferred domain (optional)",
            placeholder="e.g. healthcare, fintech, web applications"
        )

        if known_skills:
            st.markdown("**Your selected skills**")

            st.markdown(
                "".join(
                    f'<span class="chip">{skill.title()}</span>'
                    for skill in known_skills
                ),
                unsafe_allow_html=True
            )


# ============================================================
# SECTION 3 — TARGET ROLE
# ============================================================

st.markdown(
    '<div class="section-heading">'
    '<span class="step-number">03</span> Choose your target opportunity'
    '</div>'
    '<div class="section-description">'
    'Compare your profile with a career role or a specific job posting.'
    '</div>',
    unsafe_allow_html=True
)

with st.container(border=True):

    if target_role == "Custom job description":

        job_description = st.text_area(
            "Job description",
            height=190,
            placeholder=(
                "Paste the job description here, including "
                "responsibilities and required skills."
            )
        )

    else:

        role_info = ROLES[target_role]

        st.markdown(f"**{target_role} · common skills**")

        st.markdown(
            "".join(
                f'<span class="chip">{skill.title()}</span>'
                for skill in role_info["skills"]
            ),
            unsafe_allow_html=True
        )

        job_description = " ".join(
            role_info["skills"] + role_info["projects"]
        )

        st.caption(
            "These are illustrative role requirements. "
            "Paste a specific job description for a more targeted comparison."
        )

    additional_requirements = st.text_input(
        "Additional requirements (optional)",
        placeholder="e.g. FastAPI, PostgreSQL, certification"
    )

    analyze_button = st.button(
        "Generate my career insights  →",
        type="primary",
        use_container_width=True
    )


# ============================================================
# SECTION 4 — ANALYSIS ENGINE
# ============================================================

if analyze_button:

    if not candidate_name.strip():
        st.error("Please enter your name to create your career profile.")

    elif (
        not resume_text.strip()
        and not known_skills
        and not custom_skills.strip()
    ):
        st.error(
            "Please upload/paste your resume or select the skills you know."
        )

    elif (
        target_role == "Custom job description"
        and not job_description.strip()
    ):
        st.error("Please paste a job description for the custom role.")

    else:

        custom_skill_list = [
            skill.strip().lower()
            for skill in custom_skills.split(",")
            if skill.strip()
        ]

        profile_text = (
            resume_text
            + " "
            + " ".join(known_skills)
            + " "
            + " ".join(custom_skill_list)
        )

        detected_skills = extract_skills(profile_text)

        resume_skills = sorted(
            set(detected_skills + known_skills + custom_skill_list)
        )

        required_skills = extract_skills(
            job_description + " " + additional_requirements
        )

        matched_skills = [
            skill for skill in required_skills
            if skill in resume_skills
        ]

        missing_skills = [
            skill for skill in required_skills
            if skill not in resume_skills
        ]

        skill_coverage = (
            round(
                len(matched_skills) / len(required_skills) * 100,
                1
            )
            if required_skills else 0
        )

        text_score = calculate_similarity(
            profile_text,
            job_description + " " + additional_requirements
        )

        # Transparent heuristic: not a hiring probability.
        career_score = round(
            skill_coverage * 0.65 + text_score * 0.35,
            1
        )

        st.session_state["career_report"] = {
            "name": candidate_name.strip(),
            "education": education,
            "experience_level": experience_level,
            "years_experience": years_experience,
            "current_status": current_status,
            "confidence": confidence,
            "interests": interests,
            "role": target_role,
            "known_skills": known_skills + custom_skill_list,
            "resume_skills": resume_skills,
            "required_skills": required_skills,
            "matched_skills": matched_skills,
            "missing_skills": missing_skills,
            "skill_coverage": skill_coverage,
            "text_score": text_score,
            "score": career_score
        }


# ============================================================
# SECTION 5 — RESULTS
# ============================================================

report = st.session_state.get("career_report")

if report:

    st.markdown("---")

    st.markdown(
        '<div class="section-heading">'
        '<span class="step-number">✦</span> Your career intelligence'
        '</div>'
        '<div class="section-description">'
        'Your profile, target role alignment, and recommended next steps.'
        '</div>',
        unsafe_allow_html=True
    )

    st.caption(
        "The scores are heuristic guidance indicators, not hiring "
        "probabilities. They depend on the text and skills detected."
    )

    metric1, metric2, metric3, metric4 = st.columns(4)

    metric1.metric(
        "Career fit indicator",
        f'{report["score"]}%'
    )

    metric2.metric(
        "Skill coverage",
        f'{report["skill_coverage"]}%'
    )

    metric3.metric(
        "Matched skills",
        len(report["matched_skills"])
    )

    metric4.metric(
        "Potential gaps",
        len(report["missing_skills"])
    )

    if report["score"] >= 80:
        readiness = "Strong alignment"
    elif report["score"] >= 60:
        readiness = "Promising foundation"
    elif report["score"] >= 40:
        readiness = "Developing alignment"
    else:
        readiness = "Early-stage alignment"

    st.progress(
        int(report["score"]),
        text=readiness
    )

    st.markdown(
        f'**Profile:** {report["name"]} &nbsp; · &nbsp; '
        f'{report["experience_level"]} &nbsp; · &nbsp; '
        f'{report["current_status"]} &nbsp; · &nbsp; '
        f'{report["confidence"]} skill confidence',
        unsafe_allow_html=True
    )

    overview_tab, gaps_tab, roadmap_tab, export_tab = st.tabs(
        [
            "Overview",
            "Skill gap analysis",
            "Learning roadmap",
            "Export report"
        ]
    )

    # ---------------------- OVERVIEW ----------------------

    with overview_tab:

        left, right = st.columns(2)

        with left:
            st.markdown("#### Matching strengths")

            if report["matched_skills"]:
                st.markdown(
                    "".join(
                        f'<span class="chip">✓ {skill.title()}</span>'
                        for skill in report["matched_skills"]
                    ),
                    unsafe_allow_html=True
                )
            else:
                st.info(
                    "No direct matches detected. "
                    "Add more detail to your profile and resume."
                )

        with right:
            st.markdown("#### Skills in your profile")

            if report["resume_skills"]:
                st.markdown(
                    "".join(
                        f'<span class="chip">{skill.title()}</span>'
                        for skill in report["resume_skills"]
                    ),
                    unsafe_allow_html=True
                )
            else:
                st.info("No skills detected from the current profile.")

        st.markdown("#### How your score is calculated")

        st.write(
            "Career fit indicator = 65% detected skill coverage + "
            "35% TF-IDF text similarity. Experience level and confidence "
            "are collected for context; they do not artificially change "
            "the score."
        )

        st.caption(
            f'Text similarity: {report["text_score"]}% · '
            f'Education: {report["education"]} · '
            f'Years of experience: {report["years_experience"]}'
        )

        if report["interests"]:
            st.markdown(
                f'**Career interests:** {report["interests"]}'
            )

    # ---------------------- SKILL GAPS ----------------------

    with gaps_tab:

        if report["required_skills"]:

            gap_table = pd.DataFrame({
                "Target skill": [
                    skill.title()
                    for skill in report["required_skills"]
                ],
                "Status": [
                    "Matched"
                    if skill in report["matched_skills"]
                    else "Potential gap"
                    for skill in report["required_skills"]
                ],
                "Priority": [
                    "Covered"
                    if skill in report["matched_skills"]
                    else (
                        "High"
                        if skill in report["missing_skills"][:3]
                        else "Medium"
                    )
                    for skill in report["required_skills"]
                ]
            })

            st.dataframe(
                gap_table,
                use_container_width=True,
                hide_index=True
            )

            if report["missing_skills"]:

                st.warning(
                    "Potential areas to strengthen: "
                    + ", ".join(
                        skill.title()
                        for skill in report["missing_skills"]
                    )
                )

                st.write(
                    "Start with the top three gaps. Practise them through "
                    "small projects and add evidence to your portfolio "
                    "before listing them as skills."
                )

            else:

                st.success(
                    "All detected target skills appear in your profile. "
                    "Make sure your resume includes evidence for each skill."
                )

        else:

            st.info(
                "No target skills were detected. Add more details to the "
                "job description or additional requirements."
            )

    # ---------------------- ROADMAP ----------------------

    with roadmap_tab:

        missing = report["missing_skills"]

        if missing:

            roadmap_rows = []

            for index, skill in enumerate(missing[:9]):

                if index < 3:
                    phase = "Foundations"
                    timeline = "Weeks 1–2"

                elif index < 6:
                    phase = "Practice"
                    timeline = "Weeks 3–4"

                else:
                    phase = "Portfolio"
                    timeline = "Weeks 5–6"

                suggested_learning = COURSE_SUGGESTIONS.get(
                    skill,
                    f"Learn {skill} fundamentals and practise with a project"
                )

                roadmap_rows.append({
                    "Phase": phase,
                    "Timeline": timeline,
                    "Focus skill": skill.title(),
                    "Suggested learning": suggested_learning
                })

            st.write(
                "A starter six-week plan. Adjust the pace to your "
                "schedule and experience."
            )

            st.dataframe(
                pd.DataFrame(roadmap_rows),
                use_container_width=True,
                hide_index=True
            )

        else:

            st.success(
                "No skill gaps were detected in this comparison. "
                "Focus on portfolio depth, interview preparation, "
                "and evidence of your impact."
            )

        if report["role"] in ROLES:

            role_info = ROLES[report["role"]]

            st.markdown("#### Portfolio project ideas")

            for project in role_info["projects"]:
                st.markdown(f"- {project}")

            st.markdown("#### Learning paths to explore")

            for course in role_info["courses"]:
                st.markdown(f"- {course}")

        elif missing:

            st.markdown("#### Project ideas")

            for skill in missing[:3]:
                st.markdown(
                    f"- Build a project demonstrating **{skill.title()}**."
                )

    # ---------------------- EXPORT ----------------------

    with export_tab:

        st.write(
            "Download your analysis for future reference."
        )

        pdf_bytes = generate_pdf_report(report)

        st.download_button(
            "Download career report · PDF",
            data=pdf_bytes,
            file_name="skillbridge_career_report.pdf",
            mime="application/pdf",
            type="primary"
        )

        csv_data = pd.DataFrame({
            "Target skill": [
                skill.title()
                for skill in report["required_skills"]
            ],
            "Status": [
                "Matched"
                if skill in report["matched_skills"]
                else "Potential gap"
                for skill in report["required_skills"]
            ]
        })

        st.download_button(
            "Download skill checklist · CSV",
            data=csv_data.to_csv(index=False).encode("utf-8"),
            file_name="skillbridge_skill_analysis.csv",
            mime="text/csv"
        )

    st.markdown(
        '<div class="small-muted">'
        'SkillBridge AI is a career-planning prototype. Skill extraction '
        'is keyword-based, and role profiles are illustrative rather than '
        'live job-market data.'
        '</div>',
        unsafe_allow_html=True
    )


# ============================================================
# FOOTER
# ============================================================

st.markdown("---")

st.markdown(
    '<div class="small-muted" style="text-align:center;">'
    'SKILLBRIDGE AI · Your career journey, made clearer.'
    '</div>',
    unsafe_allow_html=True
)
