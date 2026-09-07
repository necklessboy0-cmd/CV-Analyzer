import io
import json
import hashlib
import os
import re
from pathlib import Path
from typing import Any

import streamlit as st
from google import genai
from google.genai import types
from pypdf import PdfReader
from docx import Document
from docx.shared import Inches, Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.section import WD_SECTION

APP_TITLE = "CV Analyzer"
GEMINI_MODEL = "gemini-3.5-flash-lite"
MAX_CV_CHARS = 100_000
MAX_FILE_BYTES = 12 * 1024 * 1024
SUPPORTED_TYPES = ["pdf", "docx", "txt"]

WORKFLOW_STEPS = [
    "CV Received",
    "Document Structure",
    "Content Extraction",
    "Content Validation",
    "Initial CV Analysis",
    "Mistakes Identified",
    "Analysis Report",
    "CV Enhancement",
    "Enhanced CV Validation",
    "Re-Analysis",
    "Final Quality Check",
    "Final Result",
    "Generate Files",
    "Download",
]

st.set_page_config(page_title=APP_TITLE, page_icon="📄", layout="wide", initial_sidebar_state="collapsed")


def inject_css() -> None:
    st.markdown(
        """
        <style>
        @keyframes cvFloat { 0%{transform:translate3d(0,0,0) scale(1)} 50%{transform:translate3d(1.5%,-1%,0) scale(1.03)} 100%{transform:translate3d(0,0,0) scale(1)} }
        .stApp { background: radial-gradient(circle at 12% 8%, rgba(191,156,255,.24), transparent 28%), radial-gradient(circle at 88% 18%, rgba(255,211,92,.12), transparent 25%), linear-gradient(155deg,#f0e7ff 0%,#e9dcff 38%,#24163e 100%); color:#120d1c; }
        .stApp::before { content:""; position:fixed; inset:-12%; pointer-events:none; z-index:0; background:radial-gradient(circle at 20% 30%,rgba(255,255,255,.18) 0 2px,transparent 3px),radial-gradient(circle at 70% 60%,rgba(255,255,255,.13) 0 2px,transparent 3px); background-size:150px 150px,210px 210px; animation:cvFloat 24s ease-in-out infinite; }
        .block-container { position:relative; z-index:1; max-width:1180px; padding-top:2.2rem; padding-bottom:3rem; }
        .hero { padding:2rem 2.2rem; border:1px solid rgba(255,255,255,.65); border-radius:28px; background:rgba(255,255,255,.54); box-shadow:0 22px 70px rgba(32,15,63,.20); backdrop-filter:blur(14px); margin-bottom:1.2rem; }
        .hero h1 { margin:0; color:#b38300; font-size:clamp(2.3rem,5vw,4.2rem); letter-spacing:-.045em; font-weight:900; }
        .hero p { margin:.55rem 0 0; color:#241936; font-size:1.05rem; }
        .panel { border-radius:22px; padding:1.2rem 1.35rem; background:rgba(24,13,43,.76); border:1px solid rgba(255,220,122,.18); box-shadow:0 18px 55px rgba(10,5,20,.25); color:#f7f1ff; }
        .panel h3,.panel h4 { color:#f3c84f; }
        .metric-card { border-radius:20px; padding:1.25rem; background:rgba(255,255,255,.92); border:1px solid rgba(62,34,95,.12); box-shadow:0 14px 40px rgba(26,12,48,.13); min-height:145px; }
        .score-number { font-size:3.7rem; line-height:1; font-weight:900; color:#5d3297; }
        .score-label { color:#4c405b; font-weight:700; }
        .score-track { width:100%; height:12px; border-radius:99px; background:#e6dcf3; overflow:hidden; margin-top:.85rem; }
        .score-fill { height:100%; border-radius:99px; background:linear-gradient(90deg,#6d3bb5,#b38300); }
        .section-card { border-radius:18px; padding:1.05rem 1.15rem; margin:.65rem 0; background:rgba(255,255,255,.93); border-left:5px solid #6d3bb5; box-shadow:0 9px 28px rgba(26,12,48,.10); }
        .issue-card { border-radius:16px; padding:.9rem 1rem; margin:.55rem 0; background:rgba(255,255,255,.93); border-left:4px solid #b38300; color:#1c1328; }
        .small-note { color:#665a73; font-size:.88rem; }
        .workflow { margin:1rem 0 1.4rem; padding:1rem .8rem 1.05rem; border-radius:20px; background:rgba(255,255,255,.72); border:1px solid rgba(255,255,255,.65); box-shadow:0 12px 35px rgba(26,12,48,.10); overflow-x:auto; }
        .workflow-line { display:flex; align-items:flex-start; min-width:980px; }
        .wf-step { flex:1; min-width:68px; text-align:center; position:relative; }
        .wf-step:not(:last-child)::after { content:""; position:absolute; top:14px; left:50%; width:100%; height:2px; background:#d9cfe8; z-index:0; }
        .wf-dot { position:relative; z-index:1; width:30px; height:30px; margin:0 auto .35rem; border-radius:50%; display:flex; align-items:center; justify-content:center; font-weight:900; background:#fff; border:2px solid #cbbbdc; color:#8c7b9e; }
        .wf-step.done .wf-dot { background:#6d3bb5; border-color:#6d3bb5; color:#fff; }
        .wf-step.active .wf-dot { background:#fff7d7; border-color:#b38300; color:#8a6800; animation:wfSpin 1.3s linear infinite; }
        .wf-step.error .wf-dot { background:#fff0f0; border-color:#b42318; color:#b42318; }
        .wf-step.done:not(:last-child)::after { background:#6d3bb5; }
        .wf-label { font-size:.68rem; line-height:1.15; font-weight:800; color:#65576f; }
        .wf-step.done .wf-label,.wf-step.active .wf-label { color:#2c1942; }
        @keyframes wfSpin { from{transform:rotate(0)} to{transform:rotate(360deg)} }
        .status-banner { border-radius:15px; padding:.75rem 1rem; margin-bottom:1rem; background:rgba(24,13,43,.84); color:#f7f1ff; }
        .compare-good { color:#16794a; font-weight:900; }
        .compare-warn { color:#9a6700; font-weight:900; }
        div[data-testid="stFileUploader"] { border-radius:18px; }
        .stButton > button { border-radius:13px; font-weight:800; }
        </style>
        """,
        unsafe_allow_html=True,
    )


def get_api_key() -> str | None:
    try:
        if "GEMINI_API_KEY" in st.secrets:
            value = st.secrets["GEMINI_API_KEY"]
            if value:
                return str(value).strip()
    except Exception:
        pass
    value = os.getenv("GEMINI_API_KEY")
    return value.strip() if value else None


def workflow_reset() -> None:
    st.session_state.workflow = {step: "pending" for step in WORKFLOW_STEPS}
    st.session_state.workflow[WORKFLOW_STEPS[0]] = "active"


def set_workflow(step: str, status: str = "active", message: str = "") -> None:
    workflow = st.session_state.setdefault("workflow", {s: "pending" for s in WORKFLOW_STEPS})
    if step not in WORKFLOW_STEPS:
        return
    idx = WORKFLOW_STEPS.index(step)
    if status == "active":
        for i, name in enumerate(WORKFLOW_STEPS):
            if i < idx:
                workflow[name] = "done"
            elif i == idx:
                workflow[name] = "active"
            elif workflow[name] != "done":
                workflow[name] = "pending"
    elif status in {"done", "error"}:
        workflow[step] = status
        if status == "done" and idx + 1 < len(WORKFLOW_STEPS):
            workflow[WORKFLOW_STEPS[idx + 1]] = "pending"
    st.session_state.workflow_message = message
    render_workflow()


def render_workflow() -> None:
    workflow = st.session_state.get("workflow", {s: "pending" for s in WORKFLOW_STEPS})
    symbols = {"done": "✓", "active": "⟳", "pending": "○", "error": "✕"}
    html = '<div class="workflow"><div class="workflow-line">'
    for step in WORKFLOW_STEPS:
        status = workflow.get(step, "pending")
        html += f'<div class="wf-step {status}"><div class="wf-dot">{symbols[status]}</div><div class="wf-label">{step}</div></div>'
    html += "</div>"
    msg = st.session_state.get("workflow_message", "")
    if msg:
        html += f'<div class="status-banner"><b>{symbols.get(next((s for s,v in workflow.items() if v=="active"), "pending"), "○")}</b> {msg}</div>'
    html += "</div>"
    st.markdown(html, unsafe_allow_html=True)


def read_txt(data: bytes) -> str:
    for encoding in ("utf-8-sig", "utf-8", "utf-16", "latin-1"):
        try:
            return data.decode(encoding)
        except UnicodeDecodeError:
            continue
    raise ValueError("The TXT file could not be decoded as readable text.")


def extract_docx(data: bytes) -> dict[str, Any]:
    document = Document(io.BytesIO(data))
    chunks: list[str] = []
    for paragraph in document.paragraphs:
        text = paragraph.text.strip()
        if text:
            chunks.append(text)
    for table in document.tables:
        for row in table.rows:
            cells = [cell.text.strip() for cell in row.cells if cell.text.strip()]
            if cells:
                chunks.append(" | ".join(cells))
    text = "\n".join(chunks)
    # python-docx does not expose rendered page boundaries reliably. Explicit form feeds are honored.
    raw_pages = text.split("\f")
    pages = [{"page_number": i + 1, "text": normalize_text(p)} for i, p in enumerate(raw_pages) if normalize_text(p)]
    if not pages:
        pages = [{"page_number": 1, "text": ""}]
    return {"file_type": "docx", "page_count": len(pages), "pages": pages, "text": "\n\n".join(p["text"] for p in pages)}


def extract_pdf(data: bytes) -> dict[str, Any]:
    reader = PdfReader(io.BytesIO(data))
    pages = []
    for index, page in enumerate(reader.pages, start=1):
        text = normalize_text(page.extract_text() or "")
        pages.append({"page_number": index, "text": text})
    return {"file_type": "pdf", "page_count": len(pages), "pages": pages, "text": "\n\n".join(f"=== PAGE {p['page_number']} ===\n{p['text']}" for p in pages)}


def extract_text(filename: str, data: bytes) -> dict[str, Any]:
    suffix = Path(filename).suffix.lower()
    if suffix == ".pdf":
        return extract_pdf(data)
    if suffix == ".docx":
        return extract_docx(data)
    if suffix == ".txt":
        text = normalize_text(read_txt(data))
        raw_pages = text.split("\f")
        pages = [{"page_number": i + 1, "text": normalize_text(p)} for i, p in enumerate(raw_pages) if normalize_text(p)] or [{"page_number": 1, "text": ""}]
        return {"file_type": "txt", "page_count": len(pages), "pages": pages, "text": "\n\n".join(f"=== PAGE {p['page_number']} ===\n{p['text']}" for p in pages)}
    raise ValueError("Unsupported file type. Please upload PDF, DOCX, or TXT.")


def normalize_text(text: str) -> str:
    text = text.replace("\x00", "").replace("\r\n", "\n").replace("\r", "\n")
    return text.strip()


def detect_sections(text: str) -> list[str]:
    patterns = {
        "Contact Information": r"\b(contact|phone|mobile|email|linkedin|github|portfolio)\b",
        "Professional Summary": r"\b(summary|profile|objective|about me)\b",
        "Experience": r"\b(experience|employment|work history|professional experience)\b",
        "Education": r"\b(education|academic background|qualifications)\b",
        "Skills": r"\b(skills|technical skills|core competencies|competencies)\b",
        "Projects": r"\b(projects|selected projects|academic projects)\b",
        "Certifications": r"\b(certifications|certificates|licenses)\b",
        "Achievements": r"\b(achievements|awards|honors|accomplishments)\b",
        "Languages": r"\b(languages|language proficiency)\b",
    }
    lower = text.lower()
    return [name for name, pattern in patterns.items() if re.search(pattern, lower)]


def detect_sections_by_page(pages: list[dict[str, Any]]) -> dict[str, list[str]]:
    return {f"Page {p['page_number']}": detect_sections(p.get("text", "")) for p in pages}


def deterministic_checks(text: str, pages: list[dict[str, Any]]) -> dict[str, Any]:
    urls = re.findall(r"(?:https?://|www\.)[^\s<>()]+", text, flags=re.I)
    emails = re.findall(r"\b[A-Z0-9._%+-]+@[A-Z0-9.-]+\.[A-Z]{2,}\b", text, flags=re.I)
    lines = [line.strip() for line in text.splitlines() if line.strip()]
    return {
        "word_count": len(re.findall(r"\b[\w'-]+\b", text)),
        "character_count": len(text),
        "line_count": len(lines),
        "page_count": len(pages),
        "page_word_counts": {str(p["page_number"]): len(re.findall(r"\b[\w'-]+\b", p.get("text", ""))) for p in pages},
        "empty_pages": [p["page_number"] for p in pages if not p.get("text", "").strip()],
        "url_count": len(urls),
        "email_count": len(emails),
        "urls": urls[:50],
        "double_space_count": len(re.findall(r"[ \t]{2,}", text)),
        "space_before_punctuation_count": len(re.findall(r"[ \t]+[,.;:!?]", text)),
        "repeated_punctuation_count": len(re.findall(r"([!?.,;:])\1+", text)),
        "trailing_space_line_count": sum(1 for line in text.splitlines() if line.endswith((" ", "\t"))),
        "very_long_line_count": sum(1 for line in lines if len(line) > 140),
        "detected_sections": detect_sections(text),
        "sections_by_page": detect_sections_by_page(pages),
    }


def build_prompt(cv_text: str, checks: dict[str, Any]) -> str:
    return f"""
You are a senior CV/resume auditor and professional career-document editor.

The supplied CV is the ONLY source of truth. Never invent, infer, upgrade, or fabricate education,
employment, dates, skills, achievements, metrics, employers, job titles, links, credentials,
locations, personal details, or results. If information is missing or unclear, report it as missing
or unclear rather than assuming it.

IMPORTANT PAGE RULE:
- The document is supplied with explicit PAGE markers in original order.
- Analyze Page 1, then Page 2, then Page 3, etc. exactly in that order.
- Never move a heading or content from one page to another in your interpretation.
- Never treat a heading found on a later page as if it were the first heading of the document.
- Use page references in detailed errors whenever the evidence is page-specific.
- Page count is structural evidence, not a reason to penalize a candidate by itself.

The application already performed deterministic checks. Use them as evidence, but also perform your
own deep linguistic and professional review.

Audit all relevant areas:
1. Overall effectiveness and score from 0 to 100.
2. Executive summary.
3. Strengths.
4. Critical issues.
5. Section-by-section analysis of sections actually present.
6. Micro-level grammar, spelling, punctuation, capitalization, spacing, wording, consistency,
   awkward phrasing, repetition, and unclear statements. Identify real mistakes, not vague advice.
7. Missing information that would materially improve the CV.
8. ATS/readability/recruiter effectiveness without claiming to have visually rendered the file.
9. Link/contact analysis. Never create or fabricate URLs.
10. Actionable recommendations prioritized by impact.
11. Final professional assessment.

For every detailed error, provide the exact problem, short evidence, why it matters, a concrete fix,
and page number when known.

Deterministic checks:
{json.dumps(checks, ensure_ascii=False, indent=2)}

Return ONLY valid JSON with this structure:
{{
  "overall_score": 0,
  "score_explanation": "string",
  "executive_summary": "string",
  "strengths": ["string"],
  "critical_issues": ["string"],
  "section_analysis": [{{"section":"string","status":"strong|needs_improvement|missing|not_applicable","score":0,"findings":["string"],"recommendations":["string"]}}],
  "detailed_errors": [{{"category":"grammar|spelling|punctuation|spacing|capitalization|wording|consistency|formatting_text|content|other","severity":"high|medium|low","issue":"string","evidence":"string","why_it_matters":"string","fix":"string","page":0}}],
  "missing_information": ["string"],
  "recommendations": [{{"priority":"high|medium|low","recommendation":"string","reason":"string"}}],
  "link_analysis": {{"present":true,"findings":["string"]}},
  "final_assessment":"string"
}}

CV TEXT START
{cv_text}
CV TEXT END
""".strip()


def extract_json(text: str) -> dict[str, Any]:
    cleaned = text.strip()
    if cleaned.startswith("```"):
        cleaned = re.sub(r"^```(?:json)?\s*", "", cleaned, flags=re.I)
        cleaned = re.sub(r"\s*```$", "", cleaned)
    try:
        parsed = json.loads(cleaned)
        if isinstance(parsed, dict):
            return parsed
    except json.JSONDecodeError:
        pass
    start, end = cleaned.find("{"), cleaned.rfind("}")
    if start >= 0 and end > start:
        parsed = json.loads(cleaned[start:end + 1])
        if isinstance(parsed, dict):
            return parsed
    raise ValueError("Gemini returned an invalid JSON analysis.")


def as_list(value: Any) -> list[Any]:
    return value if isinstance(value, list) else []


def normalize_result(result: dict[str, Any]) -> dict[str, Any]:
    try:
        score = int(float(result.get("overall_score", 0)))
    except (TypeError, ValueError):
        score = 0
    result["overall_score"] = max(0, min(100, score))
    for key in ("score_explanation", "executive_summary", "final_assessment"):
        result[key] = str(result.get(key, "")).strip()
    for key in ("strengths", "critical_issues", "missing_information"):
        result[key] = [str(x) for x in as_list(result.get(key))]
    sections = []
    for item in as_list(result.get("section_analysis")):
        if not isinstance(item, dict):
            continue
        try: item_score = int(float(item.get("score", 0)))
        except (TypeError, ValueError): item_score = 0
        sections.append({"section":str(item.get("section","Section")),"status":str(item.get("status","needs_improvement")),"score":max(0,min(100,item_score)),"findings":[str(x) for x in as_list(item.get("findings"))],"recommendations":[str(x) for x in as_list(item.get("recommendations"))]})
    result["section_analysis"] = sections
    errors = []
    for item in as_list(result.get("detailed_errors")):
        if not isinstance(item, dict): continue
        try: page = int(item.get("page", 0) or 0)
        except (TypeError, ValueError): page = 0
        errors.append({"category":str(item.get("category","other")),"severity":str(item.get("severity","low")),"issue":str(item.get("issue","")),"evidence":str(item.get("evidence","")),"why_it_matters":str(item.get("why_it_matters","")),"fix":str(item.get("fix","")),"page":page})
    result["detailed_errors"] = errors
    recs = []
    for item in as_list(result.get("recommendations")):
        if isinstance(item, dict): recs.append({"priority":str(item.get("priority","medium")),"recommendation":str(item.get("recommendation","")),"reason":str(item.get("reason",""))})
        else: recs.append({"priority":"medium","recommendation":str(item),"reason":""})
    result["recommendations"] = recs
    link = result.get("link_analysis") if isinstance(result.get("link_analysis"), dict) else {}
    result["link_analysis"] = {"present":bool(link.get("present",False)),"findings":[str(x) for x in as_list(link.get("findings"))]}
    return result


def analyze_with_gemini(cv_text: str, checks: dict[str, Any], api_key: str) -> dict[str, Any]:
    client = genai.Client(api_key=api_key)
    response = client.models.generate_content(model=GEMINI_MODEL, contents=build_prompt(cv_text, checks), config=types.GenerateContentConfig(response_mime_type="application/json"))
    response_text = getattr(response, "text", None)
    if not response_text:
        raise ValueError("Gemini returned an empty analysis.")
    return normalize_result(extract_json(response_text))


def build_enhancement_prompt(cv_text: str, initial_result: dict[str, Any]) -> str:
    return f"""
You are a senior professional CV editor. Create an enhanced version of the CV below.

HARD FACTUAL RULES:
- The original CV is the sole source of truth.
- Do not invent or infer facts.
- Do not add employers, job titles, dates, degrees, institutions, skills, certifications,
  achievements, metrics, locations, links, contact details, or responsibilities that are not present.
- Do not silently delete meaningful factual information.
- Preserve all factual claims unless correcting an obvious language error without changing meaning.
- Do not create new URLs.

Your job is to fix genuine mistakes identified in the audit, improve grammar and wording, strengthen
clarity and professional presentation, improve section hierarchy and ATS readability, remove unnecessary
repetition, and make the document concise without losing factual content.

The original analysis is provided only to identify problems:
{json.dumps(initial_result, ensure_ascii=False, indent=2)}

Return ONLY valid JSON:
{{
  "enhanced_cv_text":"full enhanced CV text",
  "changes_made":["specific factual-safe change"],
  "preserved_information":["important preserved item"]
}}

ORIGINAL CV:
{cv_text}
""".strip()


def enhance_cv_with_gemini(cv_text: str, initial_result: dict[str, Any], api_key: str) -> dict[str, Any]:
    client = genai.Client(api_key=api_key)
    response = client.models.generate_content(model=GEMINI_MODEL, contents=build_enhancement_prompt(cv_text, initial_result), config=types.GenerateContentConfig(response_mime_type="application/json"))
    text = getattr(response, "text", None)
    if not text: raise ValueError("Gemini returned an empty enhanced CV.")
    data = extract_json(text)
    enhanced = str(data.get("enhanced_cv_text", "")).strip()
    if not enhanced: raise ValueError("Gemini returned an empty enhanced CV text.")
    if len(enhanced) > MAX_CV_CHARS: enhanced = enhanced[:MAX_CV_CHARS]
    return {"enhanced_cv_text": enhanced, "changes_made":[str(x) for x in as_list(data.get("changes_made"))], "preserved_information":[str(x) for x in as_list(data.get("preserved_information"))]}


def build_verification_prompt(original_text: str, enhanced_text: str, original_result: dict[str, Any], enhanced_result: dict[str, Any]) -> str:
    return f"""
You are the final quality-control auditor for an enhanced CV.
Compare the ORIGINAL and ENHANCED CV plus both analyses.

Verify:
1. Which original mistakes were fixed.
2. Which original mistakes remain.
3. Whether any new mistakes were introduced.
4. Whether factual information was lost, changed, or invented.
5. Whether grammar, wording, structure, consistency, ATS readability and professionalism improved.
6. Whether the enhanced CV is safe to deliver as a professional CV.

Never assume an improvement is correct if it changes a factual claim.

Return ONLY valid JSON:
{{
  "original_issue_count":0,
  "resolved_issue_count":0,
  "remaining_issue_count":0,
  "new_issue_count":0,
  "resolved_issues":["string"],
  "remaining_issues":["string"],
  "new_issues":["string"],
  "factual_changes":["string"],
  "information_loss":["string"],
  "verification_status":"PASS|REVIEW_RECOMMENDED",
  "verification_summary":"string",
  "improvements":["string"]
}}

ORIGINAL ANALYSIS:
{json.dumps(original_result, ensure_ascii=False, indent=2)}

ENHANCED ANALYSIS:
{json.dumps(enhanced_result, ensure_ascii=False, indent=2)}

ORIGINAL CV:
{original_text}

ENHANCED CV:
{enhanced_text}
""".strip()


def verify_enhancement(original_text: str, enhanced_text: str, original_result: dict[str, Any], enhanced_result: dict[str, Any], api_key: str) -> dict[str, Any]:
    client = genai.Client(api_key=api_key)
    response = client.models.generate_content(model=GEMINI_MODEL, contents=build_verification_prompt(original_text, enhanced_text, original_result, enhanced_result), config=types.GenerateContentConfig(response_mime_type="application/json"))
    text = getattr(response, "text", None)
    if not text: raise ValueError("Gemini returned an empty verification.")
    data = extract_json(text)
    def integer(key):
        try: return max(0, int(data.get(key, 0)))
        except (TypeError, ValueError): return 0
    return {
        "original_issue_count": integer("original_issue_count"),
        "resolved_issue_count": integer("resolved_issue_count"),
        "remaining_issue_count": integer("remaining_issue_count"),
        "new_issue_count": integer("new_issue_count"),
        "resolved_issues":[str(x) for x in as_list(data.get("resolved_issues"))],
        "remaining_issues":[str(x) for x in as_list(data.get("remaining_issues"))],
        "new_issues":[str(x) for x in as_list(data.get("new_issues"))],
        "factual_changes":[str(x) for x in as_list(data.get("factual_changes"))],
        "information_loss":[str(x) for x in as_list(data.get("information_loss"))],
        "verification_status":"PASS" if str(data.get("verification_status","")).upper()=="PASS" else "REVIEW_RECOMMENDED",
        "verification_summary":str(data.get("verification_summary","")),
        "improvements":[str(x) for x in as_list(data.get("improvements"))],
    }


def render_page_structure(structure: dict[str, Any], checks: dict[str, Any]) -> None:
    st.markdown("### Document Structure")
    st.info(f"{structure['page_count']} page(s) detected. Original page order is preserved for analysis.")
    cols = st.columns(min(4, max(1, structure["page_count"])))
    for idx, page in enumerate(structure["pages"]):
        with cols[idx % len(cols)]:
            word_count = checks["page_word_counts"].get(str(page["page_number"]), 0)
            st.markdown(f'<div class="metric-card"><b>Page {page["page_number"]}</b><br><span class="small-note">{word_count:,} words</span></div>', unsafe_allow_html=True)
    with st.expander("Show page-by-page extracted text"):
        for page in structure["pages"]:
            st.markdown(f"**Page {page['page_number']}**")
            st.text(page["text"][:12000] or "[No readable text extracted]")


def show_list(title: str, items: list[str], empty_text: str = "None identified.") -> None:
    st.markdown(f"### {title}")
    if not items:
        st.info(empty_text); return
    for item in items:
        st.markdown(f'<div class="issue-card">{item}</div>', unsafe_allow_html=True)


def render_analysis(result: dict[str, Any], checks: dict[str, Any], title: str = "Analysis Result") -> None:
    score = result["overall_score"]
    st.markdown(f"## {title}")
    col1, col2, col3 = st.columns([1.05,1.35,1.35])
    with col1:
        st.markdown(f'<div class="metric-card"><div class="score-number">{score}<span style="font-size:1.5rem;">/100</span></div><div class="score-label">Overall CV Score</div><div class="score-track"><div class="score-fill" style="width:{score}%"></div></div></div>', unsafe_allow_html=True)
    with col2:
        st.markdown(f'<div class="metric-card"><div style="font-size:.85rem;font-weight:800;color:#6d3bb5;">DOCUMENT</div><div style="font-size:1.7rem;font-weight:900;color:#20142d;">{checks["word_count"]:,} words</div><div class="small-note">{checks["character_count"]:,} characters · {checks["line_count"]:,} text lines · {checks["page_count"]} page(s)</div></div>', unsafe_allow_html=True)
    with col3:
        st.markdown(f'<div class="metric-card"><div style="font-size:.85rem;font-weight:800;color:#b38300;">CONTACT / LINKS</div><div style="font-size:1.7rem;font-weight:900;color:#20142d;">{checks["email_count"]} email · {checks["url_count"]} link(s)</div><div class="small-note">Detected sections: {len(checks["detected_sections"])}</div></div>', unsafe_allow_html=True)
    st.markdown("---")
    st.markdown("### Score Explanation")
    st.markdown(f'<div class="section-card">{result["score_explanation"]}</div>', unsafe_allow_html=True)
    st.markdown("### Executive Summary")
    st.markdown(f'<div class="section-card">{result["executive_summary"]}</div>', unsafe_allow_html=True)
    show_list("Strengths", result["strengths"], "No major strengths were returned.")
    show_list("Critical Issues", result["critical_issues"], "No critical issues were identified.")
    st.markdown("### Section-by-Section Analysis")
    for section in result["section_analysis"]:
        findings = "".join(f"<li>{x}</li>" for x in section["findings"])
        recs = "".join(f"<li>{x}</li>" for x in section["recommendations"])
        st.markdown(f'<div class="section-card"><div style="font-size:1.15rem;font-weight:900;">{section["section"]}</div><div style="margin:.3rem 0 .7rem;"><strong>Status:</strong> {section["status"].replace("_"," ").title()} &nbsp;|&nbsp; <strong>Score:</strong> {section["score"]}/100</div>{"<strong>Findings</strong><ul>"+findings+"</ul>" if findings else ""}{"<strong>Recommendations</strong><ul>"+recs+"</ul>" if recs else ""}</div>', unsafe_allow_html=True)
    st.markdown("### Mistakes & Issues Found")
    if result["detailed_errors"]:
        for error in result["detailed_errors"]:
            page = f' · Page {error["page"]}' if error["page"] else ""
            st.markdown(f'<div class="issue-card"><strong>{error["category"].title()} · {error["severity"].title()}{page}</strong><br><b>Problem:</b> {error["issue"]}<br><b>Evidence:</b> {error["evidence"]}<br><b>Why it matters:</b> {error["why_it_matters"]}<br><b>Recommended fix:</b> {error["fix"]}</div>', unsafe_allow_html=True)
    else: st.success("No specific micro-level mistakes were returned by the AI audit.")
    show_list("Missing Information", result["missing_information"], "No material missing information was identified.")
    st.markdown("### Recommendations")
    for item in result["recommendations"]:
        st.markdown(f'<div class="section-card"><strong>{item["priority"].title()} Priority</strong><br>{item["recommendation"]}{"<br><span class=\'small-note\">"+item["reason"]+"</span>" if item["reason"] else ""}</div>', unsafe_allow_html=True)
    st.markdown("### Link Analysis")
    show_list("", result["link_analysis"]["findings"], "No link-specific findings were returned.")
    st.markdown("### Automated Text Diagnostics")
    diag = [("Double spaces",checks["double_space_count"]),("Space before punctuation",checks["space_before_punctuation_count"]),("Repeated punctuation",checks["repeated_punctuation_count"]),("Trailing-space lines",checks["trailing_space_line_count"]),("Very long lines",checks["very_long_line_count"])]
    dcols = st.columns(5)
    for col,(label,value) in zip(dcols,diag):
        with col: st.metric(label,value)
    st.markdown("### Final Assessment")
    st.markdown(f'<div class="panel"><h3>Professional Conclusion</h3><p>{result["final_assessment"]}</p></div>', unsafe_allow_html=True)


def render_verification(verification: dict[str, Any], original_score: int, enhanced_score: int) -> None:
    st.markdown("## Final Quality Check")
    cols = st.columns(5)
    metrics = [("Original issues",verification["original_issue_count"]),("Resolved",verification["resolved_issue_count"]),("Remaining",verification["remaining_issue_count"]),("New issues",verification["new_issue_count"]),("Score change",f"{original_score} → {enhanced_score}")]
    for col,(label,value) in zip(cols,metrics):
        with col: st.metric(label,value)
    status = verification["verification_status"]
    if status == "PASS": st.success("✓ PASS — the enhanced CV passed the final quality-control gate.")
    else: st.warning("⚠ REVIEW RECOMMENDED — the enhanced CV should be reviewed before delivery.")
    if verification["verification_summary"]:
        st.markdown(f'<div class="section-card">{verification["verification_summary"]}</div>', unsafe_allow_html=True)
    show_list("Improvements Verified", verification["improvements"], "No improvement list was returned.")
    show_list("Resolved Issues", verification["resolved_issues"], "No resolved issues were returned.")
    show_list("Remaining Issues", verification["remaining_issues"], "None reported.")
    show_list("New Issues", verification["new_issues"], "No new issues identified.")
    show_list("Factual Changes Flagged", verification["factual_changes"], "No factual changes were flagged.")
    show_list("Information Loss", verification["information_loss"], "No information loss was flagged.")


def safe_filename(name: str) -> str:
    stem = Path(name).stem
    return re.sub(r"[^A-Za-z0-9._-]+", "_", stem).strip("._") or "enhanced_cv"


def build_docx(cv_text: str) -> bytes:
    doc = Document()
    section = doc.sections[0]
    section.top_margin = Inches(.65); section.bottom_margin = Inches(.65); section.left_margin = Inches(.7); section.right_margin = Inches(.7)
    styles = doc.styles
    styles["Normal"].font.name = "Arial"; styles["Normal"].font.size = Pt(10)
    lines = [line.strip() for line in cv_text.splitlines()]
    first_nonempty = next((x for x in lines if x), "Enhanced CV")
    p = doc.add_paragraph(); p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p.add_run(first_nonempty); r.bold = True; r.font.size = Pt(20)
    for line in lines[1:]:
        if not line: continue
        clean = re.sub(r"^[•●▪◦\-*]+\s*", "", line).strip()
        is_heading = bool(re.match(r"^(professional summary|summary|profile|experience|professional experience|education|skills|technical skills|projects|certifications|achievements|awards|languages|contact information)$", clean, re.I))
        if is_heading:
            p = doc.add_paragraph(); p.paragraph_format.space_before = Pt(8); p.paragraph_format.space_after = Pt(3)
            r = p.add_run(clean.upper()); r.bold = True; r.font.size = Pt(12)
        elif re.match(r"^[•●▪◦\-*]\s*", line):
            p = doc.add_paragraph(style="List Bullet"); p.paragraph_format.space_after = Pt(2); p.add_run(clean)
        else:
            p = doc.add_paragraph(clean); p.paragraph_format.space_after = Pt(3)
    out = io.BytesIO(); doc.save(out); return out.getvalue()


def build_pdf(cv_text: str) -> bytes:
    try:
        from reportlab.lib.enums import TA_CENTER
        from reportlab.lib.pagesizes import A4
        from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
        from reportlab.lib.units import inch
        from reportlab.platypus import SimpleDocTemplate, Paragraph
    except ModuleNotFoundError as exc:
        raise RuntimeError("PDF generation requires the 'reportlab' package. Add reportlab to requirements.txt and redeploy.") from exc
    from xml.sax.saxutils import escape
    out = io.BytesIO()
    doc = SimpleDocTemplate(out, pagesize=A4, rightMargin=.65*inch, leftMargin=.65*inch, topMargin=.55*inch, bottomMargin=.55*inch)
    styles = getSampleStyleSheet()
    title = ParagraphStyle("CVTitle", parent=styles["Title"], fontName="Helvetica-Bold", fontSize=20, leading=23, alignment=TA_CENTER, spaceAfter=8)
    heading = ParagraphStyle("CVHeading", parent=styles["Heading2"], fontName="Helvetica-Bold", fontSize=11.5, leading=14, spaceBefore=8, spaceAfter=4)
    body = ParagraphStyle("CVBody", parent=styles["BodyText"], fontName="Helvetica", fontSize=9.5, leading=12.5, spaceAfter=3)
    bullet = ParagraphStyle("CVBullet", parent=body, leftIndent=12, firstLineIndent=-7)
    lines = [line.strip() for line in cv_text.splitlines()]
    story = []
    first = next((x for x in lines if x), "Enhanced CV")
    story.append(Paragraph(escape(first), title))
    heading_names = {"professional summary","summary","profile","experience","professional experience","education","skills","technical skills","projects","certifications","achievements","awards","languages","contact information"}
    for line in lines[1:]:
        if not line: continue
        clean = re.sub(r"^[•●▪◦\-*]+\s*", "", line).strip()
        if clean.lower() in heading_names: story.append(Paragraph(escape(clean.upper()), heading))
        elif re.match(r"^[•●▪◦\-*]\s*", line): story.append(Paragraph(escape("• " + clean), bullet))
        else: story.append(Paragraph(escape(clean), body))
    doc.build(story); return out.getvalue()


def build_txt(cv_text: str) -> bytes:
    return cv_text.encode("utf-8")


def clear_for_new_upload(filename: str, data: bytes) -> None:
    upload_hash = hashlib.sha256(data).hexdigest()
    previous_hash = st.session_state.get("cv_upload_hash")
    if previous_hash != upload_hash:
        for key in (
            "cv_result", "cv_checks", "cv_filename", "cv_structure",
            "cv_text", "original_cv_text", "original_cv_bytes",
            "original_cv_structure", "enhanced_cv", "enhanced_checks",
            "enhanced_result", "verification", "output_files",
            "enhancement_complete", "cv_upload_hash"
        ):
            st.session_state.pop(key, None)
        workflow_reset()
        st.session_state.cv_filename = filename
        st.session_state.cv_upload_hash = upload_hash


def handle_error(exc: Exception) -> None:
    message = str(exc).strip()
    lower = message.lower()
    if "429" in lower or "rate" in lower or "quota" in lower:
        st.error("Gemini rate limit or quota reached. Please try again later.")
    elif "401" in lower or "403" in lower or "api key" in lower or "authentication" in lower:
        st.error("Gemini authentication failed. Check GEMINI_API_KEY in Streamlit Secrets.")
    elif "json" in lower or "response_mime_type" in lower:
        st.error("Gemini returned a response that the analyzer could not parse as JSON. Please try again.")
    elif "reportlab" in lower or "module named" in lower:
        st.error("A required Python package is missing. Make sure requirements.txt includes reportlab, python-docx, pypdf, google-genai, and streamlit, then redeploy.")
    elif "pdf" in lower or "document" in lower or "decode" in lower:
        st.error("The document could not be read. Try exporting it again as PDF, DOCX, or TXT.")
    elif message:
        st.error(f"Processing error: {message}")
    else:
        st.error("The CV could not be processed. Please try again.")
    with st.expander("Technical error details", expanded=False):
        st.code(message or repr(exc))


def main() -> None:
    inject_css()
    st.markdown('<div class="hero"><h1>CV Analyzer</h1><p>Deep CV auditing, page-aware mistake detection, AI enhancement, re-analysis, quality verification, and professional downloads.</p></div>', unsafe_allow_html=True)
    render_workflow()
    st.markdown('<div class="panel"><h3>Upload your CV</h3><p>Supported formats: PDF, DOCX, TXT. Content is extracted in the app and analyzed with Gemini. Your CV remains the source of truth for every AI operation.</p></div>', unsafe_allow_html=True)
    uploaded = st.file_uploader("Choose a CV file", type=SUPPORTED_TYPES, help="Upload a readable PDF, DOCX, or TXT CV.")
    if not uploaded:
        st.info("Upload a CV to begin the workflow."); return
    data = uploaded.getvalue()
    clear_for_new_upload(uploaded.name, data)
    if not data: st.error("The uploaded file is empty."); return
    if len(data) > MAX_FILE_BYTES: st.error("The file is too large. Please upload a CV smaller than 12 MB."); return

    api_key = get_api_key()
    if not api_key:
        st.error("Gemini API key not found. Add GEMINI_API_KEY to Streamlit Secrets or the GEMINI_API_KEY environment variable.")
        return

    if st.button("Analyze CV", type="primary", use_container_width=True):
        try:
            workflow_reset(); set_workflow("CV Received","done","CV received successfully.")
            set_workflow("Document Structure","active","Detecting page count and preserving original page order.")
            structure = extract_text(uploaded.name, data)
            st.session_state.cv_structure = structure
            set_workflow("Document Structure","done",f"{structure['page_count']} page(s) detected; original order preserved.")
            set_workflow("Content Extraction","active","Extracting readable text page by page.")
            text = normalize_text(structure["text"])
            if not text or not any(p.get("text") for p in structure["pages"]): raise ValueError("No readable text was extracted. If this is a scanned/image-only PDF, use a text-based PDF or DOCX/TXT version.")
            if len(text) > MAX_CV_CHARS: text = text[:MAX_CV_CHARS]; st.warning(f"The extracted CV text exceeded {MAX_CV_CHARS:,} characters, so the analysis used the first portion of the document.")
            st.session_state.cv_text = text
            # Immutable source-of-truth snapshot used by every later AI stage.
            st.session_state.original_cv_text = text
            st.session_state.original_cv_bytes = data
            st.session_state.original_cv_structure = structure
            set_workflow("Content Extraction","done","Content extracted with page boundaries retained and original source locked.")
            set_workflow("Content Validation","active","Validating extracted content and page structure.")
            checks = deterministic_checks(text, structure["pages"])
            if checks["page_count"] < 1: raise ValueError("No document pages were detected.")
            st.session_state.cv_checks = checks
            set_workflow("Content Validation","done","Extraction and structure validation passed.")
            set_workflow("Initial CV Analysis","active","Gemini is performing the deep initial CV audit.")
            result = analyze_with_gemini(text, checks, api_key)
            st.session_state.cv_result = result
            set_workflow("Initial CV Analysis","done","Initial CV audit completed.")
            set_workflow("Mistakes Identified","active","Separating concrete mistakes and issues from general recommendations.")
            set_workflow("Mistakes Identified","done",f"{len(result['detailed_errors'])} detailed issue(s) identified.")
            set_workflow("Analysis Report","active","Preparing the complete original-CV analysis report.")
            set_workflow("Analysis Report","done","Initial analysis report is ready.")
        except Exception as exc:
            active = next((s for s,v in st.session_state.get("workflow",{}).items() if v=="active"), None)
            if active: set_workflow(active,"error","Processing stopped because an error occurred.")
            handle_error(exc)

    if "cv_result" in st.session_state and "cv_checks" in st.session_state:
        filename = st.session_state.get("cv_filename", uploaded.name)
        st.caption(f"Showing analysis for: {filename}")
        render_page_structure(st.session_state["cv_structure"], st.session_state["cv_checks"])
        render_analysis(st.session_state["cv_result"], st.session_state["cv_checks"], "Initial Analysis Result")

        if not st.session_state.get("enhancement_complete"):
            st.markdown("---")
            st.markdown("## ✨ AI CV Enhancement")
            st.markdown(
                '<div class="panel"><h3>Enhance My CV</h3>'
                '<p><strong>Source:</strong> The exact CV already uploaded and analyzed above. '
                'No second upload is required. AI enhancement uses the locked original CV as the '
                'single factual source, then continues through validation, re-analysis, final quality check, '
                'final result, and file generation.</p></div>',
                unsafe_allow_html=True
            )
            if st.button("✨ Enhance My CV", type="primary", use_container_width=True):
                try:
                    set_workflow("CV Enhancement","active","Gemini is creating a fact-preserving professional version of the SAME CV already analyzed.")
                    original_cv_text = st.session_state.get("original_cv_text") or st.session_state.get("cv_text")
                    original_cv_result = st.session_state.get("cv_result")
                    if not original_cv_text:
                        raise RuntimeError("Original CV text is missing from the current session. Please click Analyze CV once more for this uploaded CV.")
                    if not original_cv_result:
                        raise RuntimeError("Initial CV analysis is missing. Please complete Analyze CV before enhancement.")
                    enhanced = enhance_cv_with_gemini(original_cv_text, original_cv_result, api_key)
                    st.session_state.enhanced_cv = enhanced
                    set_workflow("CV Enhancement","done","Enhanced CV generated without intentionally adding new facts.")
                    set_workflow("Enhanced CV Validation","active","Checking the enhanced CV for content loss, factual changes, and readability.")
                    enhanced_text = enhanced["enhanced_cv_text"]
                    enhanced_pages = [{"page_number":1,"text":enhanced_text}]
                    enhanced_checks = deterministic_checks(enhanced_text, enhanced_pages)
                    if enhanced_checks["word_count"] < 10: raise ValueError("The enhanced CV is unexpectedly short and failed validation.")
                    st.session_state.enhanced_checks = enhanced_checks
                    set_workflow("Enhanced CV Validation","done","Enhanced CV validation passed.")
                    set_workflow("Re-Analysis","active","Re-analyzing the enhanced CV against the original audit criteria.")
                    enhanced_result = analyze_with_gemini(enhanced_text, enhanced_checks, api_key)
                    st.session_state.enhanced_result = enhanced_result
                    set_workflow("Re-Analysis","done","Enhanced CV re-analysis completed.")
                    set_workflow("Final Quality Check","active","Comparing original and enhanced versions for resolved, remaining, and new issues.")
                    verification = verify_enhancement(original_cv_text, enhanced_text, original_cv_result, enhanced_result, api_key)
                    st.session_state.verification = verification
                    set_workflow("Final Quality Check","done",f"Final verification: {verification['verification_status']}.")
                    set_workflow("Final Result","active","Preparing the final comparison and delivery result.")
                    set_workflow("Final Result","done","Final result prepared.")
                    set_workflow("Generate Files","active","Generating professional PDF, DOCX, and TXT files.")
                    base = safe_filename(filename)
                    st.session_state.output_files = {"pdf":build_pdf(enhanced_text),"docx":build_docx(enhanced_text),"txt":build_txt(enhanced_text),"base":base}
                    set_workflow("Generate Files","done","Output files generated.")
                    set_workflow("Download","active","Files are ready for download.")
                    set_workflow("Download","done","Enhanced CV files are ready.")
                    st.session_state.enhancement_complete = True
                except Exception as exc:
                    active = next((s for s,v in st.session_state.get("workflow",{}).items() if v=="active"), None)
                    if active: set_workflow(active,"error","Enhancement workflow stopped because an error occurred.")
                    handle_error(exc)

        if st.session_state.get("enhancement_complete"):
            st.markdown("---")
            verification = st.session_state["verification"]
            render_verification(verification, st.session_state["cv_result"]["overall_score"], st.session_state["enhanced_result"]["overall_score"])
            st.markdown("## 🎯 Final Result")
            original_score = st.session_state["cv_result"]["overall_score"]
            enhanced_score = st.session_state["enhanced_result"]["overall_score"]
            delta = enhanced_score - original_score
            label = "Improved" if delta > 0 else "Unchanged" if delta == 0 else "Lower"
            st.markdown(f'<div class="section-card"><h3>Original: {original_score}/100 → Enhanced: {enhanced_score}/100</h3><p><strong>{label}:</strong> {delta:+d} points · {verification["resolved_issue_count"]} resolved · {verification["remaining_issue_count"]} remaining · {verification["new_issue_count"]} new</p></div>', unsafe_allow_html=True)
            st.markdown("### Enhanced CV Preview")
            st.text_area("Enhanced CV", st.session_state.enhanced_cv["enhanced_cv_text"], height=520, label_visibility="collapsed")
            show_list("Changes Made", st.session_state.enhanced_cv["changes_made"], "No change list was returned.")
            st.markdown("### Enhanced CV Re-Analysis")
            render_analysis(st.session_state.enhanced_result, st.session_state.enhanced_checks, "Enhanced CV Analysis")
            st.markdown("---")
            st.markdown("## 📥 Download Enhanced CV")
            files = st.session_state.output_files
            c1,c2,c3 = st.columns(3)
            with c1: st.download_button("Download Enhanced PDF", files["pdf"], file_name=f"{files['base']}_enhanced.pdf", mime="application/pdf", use_container_width=True)
            with c2: st.download_button("Download Enhanced DOCX", files["docx"], file_name=f"{files['base']}_enhanced.docx", mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document", use_container_width=True)
            with c3: st.download_button("Download Enhanced TXT", files["txt"], file_name=f"{files['base']}_enhanced.txt", mime="text/plain", use_container_width=True)


if __name__ == "__main__":
    main()
