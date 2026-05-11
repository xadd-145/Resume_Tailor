import streamlit as st
import re
import io
from docx import Document
import pdfplumber

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
# SECTION 1 - JD CLEANING FUNCTION
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
        "relevant project", "relevant projects",   # ← your heading
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
# ─────────────────────────────────────────────

def parse_projects_block(lines: list) -> list:
    """
    Handles three project entry formats:
      Format A - "Project Name" on its own line (Normal style)
      Format B - "Project Name | Tech Stack\t[GitHub Link]" on its own line
      Format C - "• Project Name | Tech Stack\t[GitHub Link]" - a List Paragraph
                  that Word styled as a bullet, which parse_docx injected '• ' into.
                  Detected by presence of ' | ' after stripping the bullet char.
    """
    projects = []
    current_project = None

    TECH_PREFIXES = ["technologies:", "tech:", "tech stack:", "stack:", "tools:", "built with:"]

    def make_project_from_text(raw: str) -> dict:
        """Parses a project name line (with or without embedded tech stack) into a dict."""
        clean = re.sub(r'\[.*?\]', '', raw).strip()        # remove [GitHub Link]
        clean = re.sub(r'https?://\S+', '', clean).strip() # remove raw URLs
        clean = re.sub(r'\s{2,}', ' ', clean).strip()      # collapse whitespace/tabs

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

            # Format C detection: a bullet that contains ' | ' is almost certainly
            # a project name line that Word styled as a List Paragraph.
            # Real bullets never have ' | ' separating name from tech stack.
            if " | " in bullet_text or "[github" in bullet_text.lower():
                if current_project is not None:
                    projects.append(current_project)
                current_project = make_project_from_text(bullet_text)
            else:
                if current_project is not None:
                    current_project["bullets"].append(bullet_text)
            continue

        line_lower = line.lower()

        # Standalone tech stack line (Format A second line)
        is_tech_line = any(line_lower.startswith(prefix) for prefix in TECH_PREFIXES)
        if is_tech_line:
            if current_project is not None:
                for prefix in TECH_PREFIXES:
                    if line_lower.startswith(prefix):
                        current_project["tech"] = line[len(prefix):].strip()
                        break
            continue

        # Non-bullet, non-tech line = new project name (Format A or B)
        if current_project is not None:
            projects.append(current_project)
        current_project = make_project_from_text(line)

    if current_project is not None:
        projects.append(current_project)

    return projects


# ─────────────────────────────────────────────
# SECTION 5 - MASTER LINE PARSER
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
            # CRITICAL FIX: Before injecting a bullet marker, check if this
            # list-style paragraph is actually a section heading.
            # Some resumes apply list formatting to section headers by mistake.
            text_check = re.sub(r'[:\-_/|]+$', '', text.lower().strip()).strip()
            is_section_heading = any(
                text_check in variants
                for variants in SECTION_MAP.values()
            )

            if is_section_heading:
                lines.append(text)  # Treat as heading, not bullet
            else:
                lines.append("• " + text)  # Inject bullet marker
        else:
            lines.append(text)

    if not lines:
        raise ValueError("The DOCX file appears to be empty - no text was extracted.")

    resume_object = parse_text_lines(lines)
    plain_text = build_plain_text(resume_object)

    return resume_object, plain_text


# ─────────────────────────────────────────────
# SECTION 7 - PDF PARSER
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
# STREAMLIT UI - MAIN LAYOUT
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
# PIPELINE GATE - Both inputs must be ready
# ─────────────────────────────────────────────

st.divider()

if jd_ready and resume_ready:
    st.success("Both inputs are ready. Keyword Extraction (Phase 3) will go here next.")
elif not jd_ready and not resume_ready:
    st.info("Paste a job description and upload your resume to begin.")
elif not jd_ready:
    st.info("Waiting for a valid job description.")
elif not resume_ready:
    st.info("Waiting for a valid resume upload.")