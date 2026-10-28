# SkillBridge AI — AI-Powered Career Intelligence System

A Streamlit career-planning prototype for resume/role comparison, keyword skill extraction, skill-gap analysis, learning roadmap suggestions, and PDF/CSV reports.

## Deploy on Streamlit Community Cloud

1. Upload `app.py`, `requirements.txt`, and `README.md` to the root of your GitHub repository.
2. Visit https://share.streamlit.io/ and sign in with GitHub.
3. Create an app, select your repository and branch, and set the main file path to `app.py`.
4. Click Deploy.

## Features
- Upload PDF, DOCX, or TXT resume; or paste resume text.
- Built-in technology role profiles and custom job description analysis.
- Skill coverage and TF-IDF text similarity indicators.
- Skill-gap table, learning roadmap, and project/course suggestions.
- Downloadable PDF report and CSV skill checklist.

## Limitations
This is a prototype. Skill detection is keyword-based and text similarity uses TF-IDF cosine similarity. The score is a heuristic indicator, not a hiring probability. Role profiles are illustrative, not live job-market data.
