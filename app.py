import streamlit as st
import re
import io
import os          # ← PHASE 3 NEW
import json        # ← PHASE 3 NEW
import time        # ← PHASE 3 NEW
import openai      # ← PHASE 3 NEW
from docx import Document
import pdfplumber
from dotenv import load_dotenv  # ← PHASE 3 NEW

# ─────────────────────────────────────────────
# PAGE CONFIG - must be the first Streamlit call
# ─────────────────────────────────────────────
st.set_page_config(
    page_title="Resume Tailoring Engine",
    layout="wide"
)

st.title("Resume Tailoring Engine")
st.markdown("Paste a job description and upload your resume. The engine will tailor your bullets automatically.")

st.divider()

# ─────────────────────────────────────────────
# PHASE 3 NEW - OpenAI client initialization
# ─────────────────────────────────────────────
load_dotenv()
client = openai.OpenAI(api_key=os.getenv("OPENAI_API_KEY"))

# ─────────────────────────────────────────────
# PHASE 3 NEW - Session state initialization
# ─────────────────────────────────────────────
if "keyword_json" not in st.session_state:
    st.session_state.keyword_json = None

# ─────────────────────────────────────────────
# SECTION 1 - JD CLEANING FUNCTION
# (unchanged from Phase 2)
# ─────────────────────────────────────────────

def clean_jd(raw_text: str) -> str:
    BOILERPLATE_PATTERNS = [
        "equal opportunity",
        "eoe",
        "we are an equal",
        "disability",
        "veteran status",
        "background check",
        "affirmative action",
        "all qualified applicants",
        "race, color",
        "national origin",
        "accommodation",
        "we celebrate diversity",
        "primary work location",
        "additional locations",
        "required skills",
        "optional skills",
        "usa ga",
        "covington",
    ]

    lines = raw_text.strip().split("\n")
    cleaned_lines = []
    seen = set()

    for line in lines:
        line = line.strip()
        if not line:
            continue
        line_lower = line.lower()
        if line_lower in seen:
            continue
        seen.add(line_lower)
        is_boilerplate = any(pattern in line_lower for pattern in BOILERPLATE_PATTERNS)
        if is_boilerplate:
            continue
        cleaned_lines.append(line)

    cleaned = "\n".join(cleaned_lines)
    cleaned = re.sub(r'\n{3,}', '\n\n', cleaned)
    return cleaned.strip()


# ─────────────────────────────────────────────
# SECTION 2 - SECTION MAP AND SHARED UTILITIES
# (unchanged from Phase 2)
# ─────────────────────────────────────────────

SECTION_MAP = {
    "education": [
        "education", "academic background", "academics",
        "educational background", "academic qualifications"
    ],
    "experience": [
        "experience", "work experience", "professional experience",
        "employment history", "work history", "relevant experience",
        "internship experience", "internships"
    ],
    "projects": [
        "projects", "personal projects", "academic projects",
        "portfolio", "project work", "key projects", "selected projects",
        "projects & portfolio", "projects and portfolio",
        "relevant project", "relevant projects",
        "featured project", "featured projects"
    ],
    "skills": [
        "skills", "technical skills", "core competencies",
        "technologies", "tools & technologies", "tools and technologies",
        "technical expertise", "competencies", "skills & tools",
        "technical skills & certifications",
        "skills & certifications",
        "technical skills and certifications",
        "core technical skills",
        "skills and certifications"
    ]
}

BULLET_CHARS = {"•", "-", "*", "–", "▪", "◦", "‣", "·"}

YEAR_PATTERN = re.compile(r'\b(19|20)\d{2}\b')

DATE_RANGE_PATTERN = re.compile(
    r'('
    r'(?:Jan(?:uary)?|Feb(?:ruary)?|Mar(?:ch)?|Apr(?:il)?|May|Jun(?:e)?|'
    r'Jul(?:y)?|Aug(?:ust)?|Sep(?:tember)?|Oct(?:ober)?|Nov(?:ember)?|Dec(?:ember)?)'
    r'\s+\d{4}'
    r'\s*[-–]\s*'
    r'(?:Present|'
    r'(?:Jan(?:uary)?|Feb(?:ruary)?|Mar(?:ch)?|Apr(?:il)?|May|Jun(?:e)?|'
    r'Jul(?:y)?|Aug(?:ust)?|Sep(?:tember)?|Oct(?:ober)?|Nov(?:ember)?|Dec(?:ember)?)'
    r'\s+\d{4})'
    r')',
    re.IGNORECASE
)


def detect_section(line: str) -> str | None:
    line_lower = line.lower().strip()
    line_clean = re.sub(r'[:\-_/|]+$', '', line_lower).strip()
    for section_name, variants in SECTION_MAP.items():
        if line_clean in variants:
            return section_name
    return None


def is_bullet_line(line: str) -> bool:
    stripped = line.strip()
    if not stripped:
        return False
    return stripped[0] in BULLET_CHARS


def extract_bullet_text(line: str) -> str:
    stripped = line.strip()
    if stripped and stripped[0] in BULLET_CHARS:
        return stripped[1:].strip()
    return stripped


# ─────────────────────────────────────────────
# SECTION 3 - EXPERIENCE BLOCK PARSER
# (unchanged from Phase 2)
# ─────────────────────────────────────────────

def parse_experience_block(lines: list) -> list:
    roles = []
    current_role = None
    header_buffer = []

    def split_title_company(text: str) -> tuple[str, str]:
        text = text.strip()
        if " | " in text:
            parts = text.split(" | ", 1)
            return parts[0].strip(), parts[1].strip()
        elif re.search(r' at ', text, re.IGNORECASE):
            parts = re.split(r' at ', text, maxsplit=1, flags=re.IGNORECASE)
            return parts[0].strip(), parts[1].strip()
        elif "," in text:
            parts = text.split(",", 1)
            return parts[0].strip(), parts[1].strip()
        else:
            return text, ""

    i = 0
    while i < len(lines):
        line = lines[i].strip()
        if not line:
            i += 1
            continue

        if is_bullet_line(line):
            if current_role is not None:
                current_role["bullets"].append(extract_bullet_text(line))
            i += 1
            continue

        if YEAR_PATTERN.search(line):
            date_match = DATE_RANGE_PATTERN.search(line)

            if date_match:
                if current_role is not None:
                    roles.append(current_role)

                extracted_dates = date_match.group(0).strip()
                title_company_raw = line[:date_match.start()].strip().rstrip('\t ')

                if not title_company_raw and header_buffer:
                    title_company_raw = " | ".join(header_buffer).strip()
                    header_buffer = []
                elif header_buffer:
                    title_company_raw = " | ".join(header_buffer) + " " + title_company_raw
                    title_company_raw = title_company_raw.strip()
                    header_buffer = []

                title, company = split_title_company(title_company_raw)
                current_role = {
                    "title": title,
                    "company": company,
                    "dates": extracted_dates,
                    "bullets": []
                }

            else:
                if current_role is not None:
                    roles.append(current_role)

                title_company_raw = " | ".join(header_buffer).strip() if header_buffer else ""
                header_buffer = []
                title, company = split_title_company(title_company_raw)
                current_role = {
                    "title": title,
                    "company": company,
                    "dates": line,
                    "bullets": []
                }

        else:
            if current_role is not None:
                if not current_role["company"]:
                    current_role["company"] = line
            else:
                header_buffer.append(line)

        i += 1

    if current_role is not None:
        roles.append(current_role)

    return roles


# ─────────────────────────────────────────────
# SECTION 4 - PROJECTS BLOCK PARSER
# (unchanged from Phase 2)
# ─────────────────────────────────────────────

def parse_projects_block(lines: list) -> list:
    projects = []
    current_project = None

    TECH_PREFIXES = ["technologies:", "tech:", "tech stack:", "stack:", "tools:", "built with:"]

    def make_project_from_text(raw: str) -> dict:
        clean = re.sub(r'\[.*?\]', '', raw).strip()
        clean = re.sub(r'https?://\S+', '', clean).strip()
        clean = re.sub(r'\s{2,}', ' ', clean).strip()

        if " | " in clean:
            parts = clean.split(" | ", 1)
            return {"name": parts[0].strip(), "tech": parts[1].strip(), "bullets": []}
        else:
            return {"name": clean, "tech": "", "bullets": []}

    for line in lines:
        line = line.strip()
        if not line:
            continue

        if is_bullet_line(line):
            bullet_text = extract_bullet_text(line)

            if " | " in bullet_text or "[github" in bullet_text.lower():
                if current_project is not None:
                    projects.append(current_project)
                current_project = make_project_from_text(bullet_text)
            else:
                if current_project is not None:
                    current_project["bullets"].append(bullet_text)
            continue

        line_lower = line.lower()

        is_tech_line = any(line_lower.startswith(prefix) for prefix in TECH_PREFIXES)
        if is_tech_line:
            if current_project is not None:
                for prefix in TECH_PREFIXES:
                    if line_lower.startswith(prefix):
                        current_project["tech"] = line[len(prefix):].strip()
                        break
            continue

        if current_project is not None:
            projects.append(current_project)
        current_project = make_project_from_text(line)

    if current_project is not None:
        projects.append(current_project)

    return projects


# ─────────────────────────────────────────────
# SECTION 5 - MASTER LINE PARSER
# (unchanged from Phase 2)
# ─────────────────────────────────────────────

def parse_text_lines(lines: list) -> dict:
    resume_object = {
        "header": "",
        "education": [],
        "experience": [],
        "projects": [],
        "skills": ""
    }

    current_section = "header"
    section_lines = {
        "header": [],
        "education": [],
        "experience": [],
        "projects": [],
        "skills": []
    }

    for line in lines:
        if not line.strip():
            continue
        detected = detect_section(line)
        if detected:
            current_section = detected
            continue
        section_lines[current_section].append(line)

    resume_object["header"] = " | ".join(
        l.strip() for l in section_lines["header"] if l.strip()
    )
    resume_object["education"] = [l.strip() for l in section_lines["education"] if l.strip()]
    resume_object["experience"] = parse_experience_block(section_lines["experience"])
    resume_object["projects"] = parse_projects_block(section_lines["projects"])
    resume_object["skills"] = " ".join(
        l.strip() for l in section_lines["skills"] if l.strip()
    )

    return resume_object


def build_plain_text(resume_object: dict) -> str:
    parts = []

    if resume_object["header"]:
        parts.append(resume_object["header"])

    if resume_object["education"]:
        parts.append("Education")
        parts.extend(resume_object["education"])

    if resume_object["experience"]:
        parts.append("Experience")
        for role in resume_object["experience"]:
            parts.append(f"{role['title']} {role['company']} {role['dates']}")
            parts.extend(role["bullets"])

    if resume_object["projects"]:
        parts.append("Projects")
        for proj in resume_object["projects"]:
            parts.append(f"{proj['name']} {proj['tech']}")
            parts.extend(proj["bullets"])

    if resume_object["skills"]:
        parts.append("Skills")
        parts.append(resume_object["skills"])

    return "\n".join(parts)


# ─────────────────────────────────────────────
# SECTION 6 - DOCX PARSER
# (unchanged from Phase 2)
# ─────────────────────────────────────────────

def parse_docx(file) -> tuple[dict, str]:
    try:
        doc = Document(io.BytesIO(file.read()))
    except Exception as e:
        raise ValueError(f"Could not open DOCX file. It may be corrupted or not a valid Word document. Detail: {e}")

    lines = []
    for para in doc.paragraphs:
        text = para.text.strip()
        if not text:
            continue

        style_name = para.style.name.lower() if para.style and para.style.name else ""
        is_list_style = (
            "list" in style_name or
            "bullet" in style_name or
            "item" in style_name
        )

        if is_list_style:
            text_check = re.sub(r'[:\-_/|]+$', '', text.lower().strip()).strip()
            is_section_heading = any(
                text_check in variants
                for variants in SECTION_MAP.values()
            )

            if is_section_heading:
                lines.append(text)
            else:
                lines.append("• " + text)
        else:
            lines.append(text)

    if not lines:
        raise ValueError("The DOCX file appears to be empty - no text was extracted.")

    resume_object = parse_text_lines(lines)
    plain_text = build_plain_text(resume_object)

    return resume_object, plain_text


# ─────────────────────────────────────────────
# SECTION 7 - PDF PARSER
# (unchanged from Phase 2)
# ─────────────────────────────────────────────

def parse_pdf(file) -> tuple[dict, str]:
    try:
        pdf_bytes = io.BytesIO(file.read())
        with pdfplumber.open(pdf_bytes) as pdf:
            all_text_parts = []
            for page in pdf.pages:
                page_text = page.extract_text()
                if page_text:
                    all_text_parts.append(page_text)
    except Exception as e:
        raise ValueError(f"Could not open PDF file. It may be corrupted or password-protected. Detail: {e}")

    full_text = "\n".join(all_text_parts)

    if len(full_text.strip()) < 200:
        raise ValueError(
            "Your PDF appears to be a scanned image - no readable text was found. "
            "Please upload a text-based PDF or convert your resume to DOCX."
        )

    lines = full_text.split("\n")
    resume_object = parse_text_lines(lines)
    plain_text = build_plain_text(resume_object)

    return resume_object, plain_text


# ─────────────────────────────────────────────
# PHASE 3 NEW - EXTRACTION SYSTEM PROMPT
# ─────────────────────────────────────────────

EXTRACTION_SYSTEM_PROMPT = """
You are a professional technical recruiter and senior resume analyst. Your job is to read a job description and extract exhaustive, precise keyword intelligence that a candidate will use to tailor their resume for maximum ATS and recruiter match.

You must return ONLY a valid JSON object. No preamble. No explanation. No markdown fencing. No backticks. Just the raw JSON object starting with { and ending with }.

The JSON object must have exactly this structure:

{
  "role_metadata": {
    "job_title": "string",
    "required_years_experience": "string",
    "industry": "string"
  },
  "must_have": [
    {
      "keyword": "string",
      "why_it_matters": "string",
      "placement": "string - one of: Summary, Skills, Experience bullet, Project, Tools",
      "wording_suggestion": "string"
    }
  ],
  "good_to_have": [
    {
      "keyword": "string",
      "why_it_matters": "string",
      "placement": "string - one of: Summary, Skills, Experience bullet, Project, Tools",
      "wording_suggestion": "string"
    }
  ],
  "rare_but_gold": [
    {
      "keyword": "string",
      "why_it_matters": "string",
      "placement": "string - one of: Summary, Skills, Experience bullet, Project, Tools",
      "wording_suggestion": "string"
    }
  ],
  "business_language": [
    "string - a complete phrase that mirrors how the hiring team describes the work"
  ]
}

RULES:

must_have: Every explicit tool, platform, method, skill, and responsibility named directly in the JD. Non-negotiable ATS screening terms. Be exhaustive - extract 25 to 40 keywords. Do not leave out any named tool, system, skill, or responsibility.

good_to_have: Secondary skills, adjacent tools, soft skills with technical context, and responsibilities mentioned briefly or implied. Extract 15 to 25 keywords.

rare_but_gold: High-signal differentiator phrases that mirror the hiring team's business language, implied priorities, or strategic framing. These are terms stronger candidates use that weaker candidates miss. Extract 10 to 15 phrases.

business_language: Extract 12 to 20 complete phrases that reflect the company's exact vocabulary for describing the work. Examples: 'translate ambiguous business questions into structured analytical problems', 'identify risks, opportunities, and performance gaps', 'build and maintain actionable datasets', 'validate and reconcile data against source systems'. These phrases go directly into bullet rewrites in Phase 5.

Do not duplicate keywords across categories. Do not pad with generic filler. Every keyword must be directly traceable to the JD text or to a strong inference from the role's described responsibilities.
"""


# ─────────────────────────────────────────────
# PHASE 3 NEW - KEYWORD EXTRACTION FUNCTION
# ─────────────────────────────────────────────

def extract_keywords(cleaned_jd: str) -> dict:
    messages = [
        {"role": "system", "content": EXTRACTION_SYSTEM_PROMPT},
        {"role": "user", "content": f"Here is the job description to analyze:\n\n{cleaned_jd}"}
    ]

    last_error = None

    for attempt in range(2):
        try:
            response = client.chat.completions.create(
                model="gpt-4o",
                messages=messages,
                temperature=0.2,
                max_tokens=4000
            )

            raw_text = response.choices[0].message.content.strip()

            # Strip markdown fencing if model returns it despite instructions
            if raw_text.startswith("```"):
                lines = raw_text.split("\n")
                raw_text = "\n".join(lines[1:-1]).strip()

            keyword_dict = json.loads(raw_text)

            # Validate required top-level keys
            required_keys = ["role_metadata", "must_have", "good_to_have", "rare_but_gold"]
            for key in required_keys:
                if key not in keyword_dict:
                    raise ValueError(f"Missing required key in response: '{key}'")

            # Validate categories are lists
            for category in ["must_have", "good_to_have", "rare_but_gold"]:
                if not isinstance(keyword_dict[category], list):
                    raise ValueError(f"'{category}' must be a list, got {type(keyword_dict[category])}")

            return keyword_dict  # ← success, exit immediately

        except json.JSONDecodeError as e:
            last_error = f"JSON parsing failed. Raw response:\n\n{raw_text}\n\nDetail: {str(e)}"
            if attempt == 0:
                time.sleep(1)
                continue
            raise ValueError(last_error)

        except ValueError as e:
            last_error = str(e)
            if attempt == 0:
                time.sleep(1)
                continue
            raise ValueError(last_error)

        except openai.APIConnectionError:
            raise ValueError("Could not connect to OpenAI API. Check your internet connection.")

        except openai.AuthenticationError:
            raise ValueError("OpenAI API key is invalid or missing. Check your .env file and confirm OPENAI_API_KEY is set correctly.")

        except openai.RateLimitError:
            raise ValueError("OpenAI rate limit hit. Wait 30 seconds and try again.")

        except openai.APIStatusError as e:
            raise ValueError(f"OpenAI API error {e.status_code}: {e.message}")

    raise ValueError(last_error or "Extraction failed after two attempts.")


# ─────────────────────────────────────────────
# STREAMLIT UI - MAIN LAYOUT
# (unchanged from Phase 2)
# ─────────────────────────────────────────────

col_left, col_right = st.columns([1, 1], gap="large")

# ── LEFT COLUMN: JD Input ──────────────────────
with col_left:
    st.subheader("Step 1 - Job Description")

    raw_jd = st.text_area(
        "Paste the full job description here",
        height=300,
        placeholder="Copy and paste the complete job posting text here...",
        key="jd_input"
    )

    jd_ready = False
    cleaned_jd = ""

    if raw_jd:
        cleaned_jd = clean_jd(raw_jd)
        char_count = len(cleaned_jd)
        st.caption(f"Cleaned JD: {char_count:,} characters")

        if char_count < 100:
            st.warning("Job description seems too short. Please paste the full JD text.")
        else:
            st.success("JD looks good. Ready for keyword extraction.")
            jd_ready = True

# ── RIGHT COLUMN: Resume Upload ────────────────
with col_right:
    st.subheader("Step 2 - Upload Resume")

    uploaded_file = st.file_uploader(
        "Upload your resume (DOCX or PDF)",
        type=["docx", "pdf"],
        key="resume_upload"
    )

    resume_object = None
    resume_text = ""
    resume_ready = False

    if uploaded_file is not None:
        file_ext = uploaded_file.name.split(".")[-1].lower()

        with st.spinner("Parsing your resume..."):
            try:
                if file_ext == "docx":
                    resume_object, resume_text = parse_docx(uploaded_file)
                elif file_ext == "pdf":
                    resume_object, resume_text = parse_pdf(uploaded_file)
                else:
                    st.error("Unsupported file type. Please upload a DOCX or PDF file.")

                if resume_object is not None:
                    resume_ready = True

            except ValueError as ve:
                st.error(str(ve))
            except Exception as e:
                st.error(f"Unexpected error while parsing your resume: {e}")

        if resume_ready:
            num_roles = len(resume_object.get("experience", []))
            num_projects = len(resume_object.get("projects", []))
            skills_found = bool(resume_object.get("skills", "").strip())

            st.success(
                f"Resume parsed - {num_roles} role(s), {num_projects} project(s), "
                f"Skills section {'found' if skills_found else 'NOT found'}."
            )

            if num_roles == 0:
                st.warning(
                    "No work experience roles were detected. "
                    "Check that your Experience section heading uses a standard label "
                    "like 'Work Experience', 'Professional Experience', or 'Experience'."
                )

            if num_projects == 0:
                st.warning(
                    "No projects detected. If your resume has a Projects section, "
                    "check that its heading uses a label like 'Projects', 'Key Projects', or 'Portfolio'."
                )

            with st.expander("View Parsed Resume Structure (for verification)"):
                st.json(resume_object)

# ─────────────────────────────────────────────
# PHASE 3 - KEYWORD EXTRACTION UI
# Replaces the old pipeline gate placeholder
# ─────────────────────────────────────────────

st.divider()

if not jd_ready and not resume_ready:
    st.info("Paste a job description and upload your resume to begin.")
elif not jd_ready:
    st.info("Waiting for a valid job description.")
elif not resume_ready:
    st.info("Waiting for a valid resume upload.")

if jd_ready and resume_ready:

    st.subheader("Step 3 - Keyword Extraction")

    run_extraction = st.button(
        "Run Keyword Extraction",
        type="primary",
        help="Sends the cleaned job description to gpt-4o and extracts structured keywords. Costs ~$0.01–$0.03."
    )

    if run_extraction:
        # Clear previous result so stale data doesn't linger if the JD changed
        st.session_state.keyword_json = None
        with st.spinner("Extracting keywords from job description... (5–15 seconds)"):
            try:
                keyword_dict = extract_keywords(cleaned_jd)
                st.session_state.keyword_json = keyword_dict
                st.success("Keyword extraction complete.")
            except ValueError as e:
                st.error(f"Extraction failed: {str(e)}")

    # Display results if extraction has run this session
    if st.session_state.keyword_json is not None:
        kw = st.session_state.keyword_json

        # Role metadata strip
        meta = kw.get("role_metadata", {})
        st.markdown(
            f"**Role:** {meta.get('job_title', 'Not extracted')}  |  "
            f"**Experience Required:** {meta.get('required_years_experience', 'Not specified')}  |  "
            f"**Industry:** {meta.get('industry', 'Not specified')}"
        )

        st.markdown("---")

        # Metric cards
        col1, col2, col3 = st.columns(3)
        with col1:
            st.metric(
                label="Must-Have Keywords",
                value=len(kw.get("must_have", [])),
                help="Non-negotiable ATS screening terms explicitly named in the JD"
            )
        with col2:
            st.metric(
                label="Good-to-Have Keywords",
                value=len(kw.get("good_to_have", [])),
                help="Secondary skills that strengthen match quality"
            )
        with col3:
            st.metric(
                label="Rare but Gold Keywords",
                value=len(kw.get("rare_but_gold", [])),
                help="High-signal differentiator phrases implied by the JD"
            )

        st.markdown("---")

        # Keyword cards - one expander per keyword, grouped by category
        st.markdown("#### Must-Have Keywords")
        for item in kw.get("must_have", []):
            with st.expander(f"**{item.get('keyword', 'Unknown')}** - {item.get('placement', '')}"):
                st.write(f"**Why it matters:** {item.get('why_it_matters', '')}")
                st.write(f"**Wording suggestion:** {item.get('wording_suggestion', '')}")

        st.markdown("#### Good-to-Have Keywords")
        for item in kw.get("good_to_have", []):
            with st.expander(f"**{item.get('keyword', 'Unknown')}** - {item.get('placement', '')}"):
                st.write(f"**Why it matters:** {item.get('why_it_matters', '')}")
                st.write(f"**Wording suggestion:** {item.get('wording_suggestion', '')}")

        st.markdown("#### Rare but Gold Keywords")
        for item in kw.get("rare_but_gold", []):
            with st.expander(f"**{item.get('keyword', 'Unknown')}** — {item.get('placement', '')}"):
                st.write(f"**Why it matters:** {item.get('why_it_matters', '')}")
                st.write(f"**Wording suggestion:** {item.get('wording_suggestion', '')}")

        if kw.get("business_language"):
            st.markdown("#### Role-Specific Business Language")
            st.caption("Use these phrases naturally in rewritten bullets where your experience supports them.")
            for phrase in kw.get("business_language", []):
                st.markdown(f"- {phrase}")

        with st.expander("Debug — Raw JSON Response", expanded=False):
            st.json(kw)