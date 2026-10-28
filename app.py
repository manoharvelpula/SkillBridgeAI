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

st.set_page_config(page_title="SkillBridge AI", page_icon="🚀", layout="wide")

st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=DM+Sans:wght@400;500;600;700&family=Space+Grotesk:wght@500;600;700&display=swap');
html,body,[class*="css"]{font-family:'DM Sans',sans-serif}
.stApp{background:linear-gradient(135deg,#0b1020,#10172b 60%,#101323);color:#f3f4f6}
h1,h2,h3{font-family:'Space Grotesk',sans-serif}
.hero{padding:1.5rem 1.7rem;border-radius:22px;background:linear-gradient(120deg,rgba(99,102,241,.25),rgba(14,165,233,.12));border:1px solid rgba(148,163,184,.22);margin-bottom:1.2rem}
.eyebrow{color:#a5b4fc;font-weight:700;letter-spacing:.13em;text-transform:uppercase;font-size:.76rem}
.muted{color:#aab4c8}
.pill{display:inline-block;padding:.25rem .6rem;border-radius:999px;margin:.15rem .18rem .15rem 0;background:rgba(99,102,241,.17);border:1px solid rgba(129,140,248,.25);color:#dbeafe;font-size:.82rem}
</style>
""", unsafe_allow_html=True)

SKILLS = {
"Programming":["python","java","javascript","typescript","c++","c#","go","rust","php","kotlin"],
"Data & AI":["machine learning","deep learning","artificial intelligence","data science","data analysis","pandas","numpy","scikit-learn","tensorflow","pytorch","nlp","computer vision","generative ai","llm","statistics","power bi","tableau","excel","sql"],
"Web Development":["html","css","react","angular","vue","node.js","express","django","flask","fastapi","streamlit","rest api","graphql"],
"Cloud & DevOps":["aws","azure","google cloud","docker","kubernetes","terraform","jenkins","ci/cd","linux","git","github","serverless"],
"Databases":["mysql","postgresql","mongodb","redis","sqlite","database design"],
"Cybersecurity":["cybersecurity","network security","penetration testing","ethical hacking","cryptography","owasp","siem"],
"Core & Professional":["data structures","algorithms","object-oriented programming","problem solving","communication","teamwork","leadership","agile","scrum","testing","unit testing","debugging"]
}
ALL_SKILLS = sorted({s for group in SKILLS.values() for s in group}, key=len, reverse=True)

ROLES = {
"Data Analyst":{"skills":["sql","excel","python","pandas","statistics","power bi","data analysis","communication"],"projects":["Build an interactive sales dashboard","Analyze a public dataset and report insights","Create a SQL business case study"],"courses":["Google Data Analytics Certificate","Microsoft Power BI Data Analyst learning path"]},
"Machine Learning Engineer":{"skills":["python","machine learning","statistics","scikit-learn","pandas","numpy","deep learning","git","rest api"],"projects":["Train and evaluate a classification model","Deploy an ML model as an API","Build a reproducible ML pipeline"],"courses":["Machine Learning Specialization","TensorFlow learning path"]},
"Python Developer":{"skills":["python","data structures","algorithms","object-oriented programming","git","sql","rest api","testing","django"],"projects":["Build a CRUD REST API","Create a test-covered automation tool","Build a Django app with authentication"],"courses":["Python programming specialization","Python certification learning path"]},
"Full Stack Developer":{"skills":["javascript","html","css","react","node.js","rest api","sql","git","testing"],"projects":["Build a full-stack task manager","Create an authenticated e-commerce prototype","Deploy a responsive portfolio and API"],"courses":["Meta Front-End Developer","Full-stack web development course"]},
"Cloud Engineer":{"skills":["linux","aws","azure","docker","kubernetes","git","terraform","ci/cd"],"projects":["Deploy a containerized web app","Build a CI/CD pipeline","Provision infrastructure with Terraform"],"courses":["AWS Cloud Practitioner","Microsoft Azure Fundamentals"]},
"Cybersecurity Analyst":{"skills":["network security","linux","cybersecurity","siem","cryptography","python","owasp"],"projects":["Build a home-lab security monitoring dashboard","Document a web app threat model","Analyze sample logs for suspicious activity"],"courses":["CompTIA Security+","Google Cybersecurity Certificate"]},
"Software Engineer":{"skills":["python","java","data structures","algorithms","object-oriented programming","git","sql","testing","problem solving"],"projects":["Build a well-tested application","Document algorithmic problem solutions","Create a REST API with persistent storage"],"courses":["Software design specialization","Git and GitHub learning path"]}
}

COURSES = {
"python":"Python fundamentals and intermediate programming","sql":"SQL queries, joins, aggregation, and window functions",
"machine learning":"Supervised learning, evaluation, and model validation","pandas":"Data manipulation with pandas",
"power bi":"Data modeling and dashboards in Power BI","javascript":"Modern JavaScript and asynchronous programming",
"react":"React components, hooks, and state management","docker":"Docker images, containers, and Compose",
"aws":"AWS core services and cloud architecture","git":"Git branching, pull requests, and collaboration",
"statistics":"Statistics, probability, and hypothesis testing","scikit-learn":"ML pipelines with scikit-learn",
"rest api":"REST API design, authentication, and testing","linux":"Linux command line fundamentals",
"html":"Semantic HTML and accessible structure","css":"Responsive layouts and modern CSS",
"data structures":"Arrays, trees, graphs, and complexity analysis"
}

def extract_text(upload):
    if not upload: return ""
    try:
        if upload.name.lower().endswith(".pdf") and PdfReader:
            return "\n".join(p.extract_text() or "" for p in PdfReader(upload).pages)
        if upload.name.lower().endswith(".docx") and Document:
            return "\n".join(p.text for p in Document(upload).paragraphs)
        if upload.name.lower().endswith(".txt"):
            return upload.getvalue().decode("utf-8", errors="ignore")
    except Exception as e:
        st.error(f"Unable to read file: {e}")
    return ""

def find_skills(text):
    text = (text or "").lower()
    aliases = {"sklearn":"scikit-learn","nodejs":"node.js","postgres":"postgresql","powerbi":"power bi","gcp":"google cloud","oop":"object-oriented programming"}
    for a,b in aliases.items():
        text = re.sub(r"\b"+re.escape(a)+r"\b", b, text)
    found=[]
    for s in ALL_SKILLS:
        if re.search(r"(?<![a-z0-9+#])"+re.escape(s)+r"(?![a-z0-9+#])",text):
            found.append(s)
    return sorted(set(found))

def text_similarity(a,b):
    if not a.strip() or not b.strip(): return 0.0
    try:
        m=TfidfVectorizer(stop_words="english",ngram_range=(1,2),max_features=5000).fit_transform([a,b])
        return round(float(cosine_similarity(m[0:1],m[1:2])[0][0])*100,1)
    except ValueError:
        return 0.0

def create_pdf(d):
    pdf=FPDF()
    pdf.set_auto_page_break(auto=True,margin=15)
    pdf.add_page()
    pdf.set_font("Helvetica","B",20); pdf.cell(0,12,"SkillBridge AI - Career Report",ln=True)
    pdf.set_font("Helvetica",size=10); pdf.cell(0,7,f"Generated: {date.today().isoformat()}",ln=True)
    pdf.ln(4); pdf.set_font("Helvetica","B",13)
    pdf.cell(0,9,"Candidate: "+d["name"][:65],ln=True)
    pdf.cell(0,9,"Target role: "+d["role"][:65],ln=True)
    pdf.cell(0,9,f'Career fit indicator: {d["score"]:.1f}%',ln=True); pdf.ln(3)
    sections=[("Resume skills",d["resume_skills"]),("Target skills",d["required"]),("Potential skill gaps",d["missing"]),("Matched skills",d["matched"]),("Suggested actions",["Practise a high-priority missing skill in a project.","Add measurable results and project links to your resume.","Validate each skill against real job postings."])]
    for title,items in sections:
        pdf.set_font("Helvetica","B",12); pdf.multi_cell(0,8,title)
        pdf.set_font("Helvetica",size=10)
        for item in items or ["None detected"]:
            safe=str(item).encode("latin-1","replace").decode("latin-1")
            pdf.multi_cell(0,6,"- "+safe)
        pdf.ln(2)
    return bytes(pdf.output())

st.sidebar.title("🚀 SkillBridge AI")
st.sidebar.caption("Career intelligence workspace")
page=st.sidebar.radio("Navigate",["Career Analyzer","Role Explorer","Skill Library","About"])
st.sidebar.divider()
st.sidebar.caption("Prototype · Keyword skill detection · TF-IDF text similarity")

st.markdown('<div class="hero"><div class="eyebrow">AI-powered career intelligence</div><h1>Turn your skills into your next opportunity.</h1><div class="muted">Understand role alignment, identify skill gaps, and build a practical learning plan.</div></div>',unsafe_allow_html=True)

if page=="Career Analyzer":
    st.subheader("01 · Build your career profile")
    c1,c2=st.columns(2,gap="large")
    with c1:
        name=st.text_input("Your name (optional)",placeholder="e.g. Alex Morgan")
        upload=st.file_uploader("Upload resume",type=["pdf","docx","txt"])
        file_text=extract_text(upload) if upload else ""
        resume=st.text_area("Resume text",value=file_text,height=250,placeholder="Paste resume content including skills, projects, education, and experience.")
        if upload and file_text: st.caption(f"Text extracted from {upload.name}; you can edit it above.")
    with c2:
        role=st.selectbox("Target career role",list(ROLES.keys())+["Custom job description"])
        level=st.selectbox("Experience level",["Student / Fresher","Entry level (0–2 years)","Intermediate (2–5 years)","Experienced (5+ years)"])
        if role=="Custom job description":
            jd=st.text_area("Paste job description",height=180,placeholder="Paste the role's job description...")
            extra=st.text_input("Additional required skills (optional)")
        else:
            info=ROLES[role]
            jd=" ".join(info["skills"]+info["projects"])
            extra=st.text_input("Add skills to assess (optional)",placeholder="e.g. FastAPI, PostgreSQL, Docker")
            st.markdown("**Common role skills**")
            st.markdown("".join(f'<span class="pill">{s.title()}</span>' for s in info["skills"]),unsafe_allow_html=True)
        analyze=st.button("Analyze my career fit →",type="primary",use_container_width=True)
    if analyze:
        if not resume.strip(): st.error("Upload a readable resume or paste your resume text first.")
        elif role=="Custom job description" and not jd.strip(): st.error("Please paste a job description.")
        else:
            rs=find_skills(resume); required=find_skills(jd+" "+extra)
            matched=[s for s in required if s in rs]
            missing=[s for s in required if s not in rs]
            coverage=round(len(matched)/len(required)*100,1) if required else 0
            similarity=text_similarity(resume,jd+" "+extra)
            score=round(coverage*.65+similarity*.35,1)
            st.session_state["report"]={"name":name.strip() or "Candidate","role":role,"level":level,"resume_skills":rs,"required":required,"matched":matched,"missing":missing,"coverage":coverage,"similarity":similarity,"score":score}
    d=st.session_state.get("report")
    if d:
        st.divider(); st.subheader("02 · Your career intelligence report")
        st.caption("Heuristic guidance only—not a hiring probability or a guarantee of job readiness.")
        a,b,c,e=st.columns(4)
        a.metric("Career fit indicator",f'{d["score"]}%'); b.metric("Skill coverage",f'{d["coverage"]}%')
        c.metric("Matched skills",len(d["matched"])); e.metric("Potential gaps",len(d["missing"]))
        st.progress(int(d["score"]),text="Strong alignment" if d["score"]>=80 else "Promising foundation" if d["score"]>=60 else "Developing alignment" if d["score"]>=40 else "Early-stage alignment")
        t1,t2,t3,t4=st.tabs(["Overview","Skill gaps","Learning roadmap","Export"])
        with t1:
            x,y=st.columns(2)
            with x:
                st.markdown("#### Matched strengths")
                st.markdown("".join(f'<span class="pill">✓ {s.title()}</span>' for s in d["matched"]) or "No direct target skill matches detected.",unsafe_allow_html=True)
            with y:
                st.markdown("#### Resume skills detected")
                st.markdown("".join(f'<span class="pill">{s.title()}</span>' for s in d["resume_skills"]) or "No skills from the built-in keyword library were detected.",unsafe_allow_html=True)
            st.markdown("#### Scoring method")
            st.write("Career fit indicator = 65% detected skill coverage + 35% TF-IDF text similarity. The selected experience level is shown for context and does not change the score.")
        with t2:
            if d["required"]:
                table=pd.DataFrame({"Skill":[s.title() for s in d["required"]],"Status":["Matched" if s in d["matched"] else "Potential gap" for s in d["required"]],"Priority":["Covered" if s in d["matched"] else ("High" if s in d["missing"][:3] else "Medium") for s in d["required"]]})
                st.dataframe(table,use_container_width=True,hide_index=True)
                if d["missing"]: st.warning("Potential gaps: "+", ".join(s.title() for s in d["missing"]))
                else: st.success("All detected target skills also appear in your resume text. Ensure your resume contains evidence for each skill.")
            else: st.info("No skills detected in target requirements. Add more detail to the job description.")
        with t3:
            missing=d["missing"]
            if missing:
                rows=[]
                for i,s in enumerate(missing[:9]):
                    rows.append({"Phase":"Foundations" if i<3 else "Practice" if i<6 else "Portfolio","Timeline":"Weeks 1–2" if i<3 else "Weeks 3–4" if i<6 else "Weeks 5–6","Priority skill":s.title(),"Suggested learning":COURSES.get(s,f"Learn {s} fundamentals and practise with a small project")})
                st.write("Suggested six-week sequence; adapt it to your schedule and background.")
                st.dataframe(pd.DataFrame(rows),use_container_width=True,hide_index=True)
            else: st.success("No target skill gaps were detected from the current text. Focus on portfolio depth and interview practice.")
            if d["role"] in ROLES:
                st.markdown("#### Portfolio project ideas")
                for project in ROLES[d["role"]]["projects"]: st.markdown("- "+project)
                st.markdown("#### Learning paths to explore")
                for course in ROLES[d["role"]]["courses"]: st.markdown("- "+course)
            elif missing:
                st.markdown("#### Project suggestion")
                for s in missing[:3]: st.markdown(f"- Build an end-to-end project that demonstrates **{s.title()}**.")
        with t4:
            st.download_button("Download career report (PDF)",create_pdf(d),"skillbridge_career_report.pdf","application/pdf",type="primary")
            csv=pd.DataFrame({"Target skill":[s.title() for s in d["required"]],"Status":["Matched" if s in d["matched"] else "Potential gap" for s in d["required"]]})
            st.download_button("Download skill checklist (CSV)",csv.to_csv(index=False).encode("utf-8"),"skillbridge_skill_analysis.csv","text/csv")

elif page=="Role Explorer":
    st.subheader("Explore career roles")
    role=st.selectbox("Choose a role",list(ROLES.keys()))
    info=ROLES[role]
    st.markdown(f"### {role}")
    st.markdown("**Core skills**")
    st.markdown("".join(f'<span class="pill">{s.title()}</span>' for s in info["skills"]),unsafe_allow_html=True)
    c1,c2=st.columns(2)
    with c1:
        st.markdown("#### Project ideas")
        for p in info["projects"]: st.markdown("- "+p)
    with c2:
        st.markdown("#### Courses / certifications to explore")
        for c in info["courses"]: st.markdown("- "+c)
    st.caption("Role profiles are illustrative examples, not live labor-market data. Check current job postings.")

elif page=="Skill Library":
    st.subheader("Skill library")
    st.write("This keyword library is used to detect skills in resume text and job descriptions.")
    group=st.selectbox("Category",list(SKILLS.keys()))
    st.markdown("".join(f'<span class="pill">{s.title()}</span>' for s in SKILLS[group]),unsafe_allow_html=True)
    st.caption("Keyword matching may miss synonyms, context, and skills not included in this list.")

else:
    st.subheader("About SkillBridge AI")
    st.markdown("""
**SkillBridge AI** is a career-planning prototype that helps users:
- Extract skills from uploaded PDF, DOCX, and TXT resumes or pasted text.
- Compare resume content with a role profile or custom job description.
- Review a transparent skill-coverage and text-similarity indicator.
- Identify potential gaps and create a starter learning roadmap.
- Export a PDF report and CSV skill checklist.

**Limitations**
- This version uses keyword-based skill extraction and TF-IDF similarity; it does not use a generative AI model.
- Built-in role profiles and learning suggestions are examples, not live job-market data.
- Scores are guidance indicators only and do not predict hiring outcomes.
- Text-based PDFs work best; scanned/image-only PDFs may need manual text entry.
""")
    st.info("Keep any future API keys in Streamlit Secrets, never in a public GitHub repository.")
