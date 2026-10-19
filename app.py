"""CVtro — CV Review & Analysis Platform (Streamlit + Gemini)

Pipeline: extract -> validate -> chunk -> embed (vectors) -> retrieve (basic / advanced RAG)
-> analyze (+ field-fit) -> recheck -> ask user about missing parts -> generate (no invention)
-> fact-check (local + AI) -> re-score -> download.

requirements.txt: streamlit, google-genai, pypdf, python-docx, numpy, reportlab
secrets: GEMINI_API_KEY (st.secrets or environment variable)
"""
from __future__ import annotations

import hashlib
import html
import io
import json
import math
import os
import re
import zlib
from typing import Any

import numpy as np
import streamlit as st
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Inches, Pt
from google import genai
from google.genai import types
from pypdf import PdfReader

APP_TITLE = "CVtro"
GEN_MODEL = os.getenv("GEMINI_MODEL", "gemini-3.5-flash-lite")
EMBED_MODEL = os.getenv("GEMINI_EMBED_MODEL", "gemini-embedding-001")
MAX_CV_CHARS = 100_000
MAX_FILE_BYTES = 12 * 1024 * 1024
HASH_DIM = 512
SECTION_PATTERNS = {
    "summary": r"(summary|profile|objective|about me)",
    "experience": r"(experience|employment|work history|internships?)",
    "education": r"(education|academic|qualifications?)",
    "skills": r"(skills|competenc\w+|technical skills)",
    "projects": r"(projects?)",
    "certifications": r"(certifications?|courses|training)",
}
CORE_SECTIONS = ["summary", "experience", "education", "skills"]

st.set_page_config(page_title=APP_TITLE, page_icon="✦", layout="wide", initial_sidebar_state="collapsed")


# --------------------------------------------------------------------------- UI
def inject_css() -> None:
    st.markdown(
        """<style>
        .stApp{background:radial-gradient(ellipse at 50% 0%,#2a1d12 0%,#0e0b09 62%);color:#f3ead7}
        .block-container{max-width:1100px;padding-top:2rem;padding-bottom:4rem}
        h1,h2,h3{font-family:Georgia,'Playfair Display',serif!important;color:#d4a84b!important}
        .logo{font-family:Georgia,serif;font-weight:900;font-size:3rem;color:#d4a84b;line-height:1}
        .logo i{color:#f3ead7;font-style:normal}
        .tag{font-family:Georgia,serif;font-size:clamp(1.6rem,4vw,2.6rem);font-weight:900;letter-spacing:.03em;
             background:linear-gradient(180deg,#f3ead7,#d4a84b);-webkit-background-clip:text;background-clip:text;color:transparent}
        .card{background:rgba(26,20,16,.85);border:1px solid rgba(212,168,75,.25);border-radius:14px;padding:1rem 1.2rem;margin:.6rem 0}
        .note{background:#f1e3b0;color:#2a2114;border-radius:2px 2px 26px 2px;padding:1rem;min-height:150px;
              box-shadow:0 10px 24px rgba(0,0,0,.5);margin-bottom:1rem}
        .note h4{font-family:Georgia,serif;margin:0 0 .3rem;color:#2a2114}
        .pill{display:inline-block;padding:2px 11px;border-radius:99px;border:1px solid #d4a84b;font-size:.8rem;margin:0 6px 6px 0}
        .pill.g{border-color:#8fd19e;color:#8fd19e}.pill.b{border-color:#e58a7a;color:#e58a7a}
        .dia{list-style:none;padding:0}.dia li{position:relative;padding:.45rem 0 .45rem 1.6rem;border-bottom:1px solid rgba(212,168,75,.15)}
        .dia li::before{content:"";position:absolute;left:.25rem;top:.85rem;width:9px;height:9px;background:#d4a84b;transform:rotate(45deg)}
        .step{padding:.3rem 0;border-bottom:1px solid rgba(212,168,75,.12);font-size:.92rem}
        .cvp{background:#f7f1e3;color:#2a2114;padding:1.6rem;border-radius:6px;font-family:Georgia,serif;line-height:1.5}
        .cvp h4{margin:0;font-size:1.5rem;color:#2a2114}
        .cvp h5{margin:1rem 0 .2rem;border-bottom:1px solid #2a2114;text-transform:uppercase;letter-spacing:.08em;color:#2a2114}
        .stButton>button{border-radius:99px;font-weight:700;border:1px solid #d4a84b}
        </style>""",
        unsafe_allow_html=True,
    )


class Logger:
    """Pipeline log: every step is recorded with a recheck status and redrawn live."""

    ICON = {"run": "…", "ok": "✓", "warn": "⚠", "err": "✕"}

    def __init__(self, box) -> None:
        self.box = box
        st.session_state.setdefault("log", [])

    def add(self, step: str, status: str = "ok", msg: str = "") -> int:
        st.session_state.log.append([step, status, msg])
        self.draw()
        return len(st.session_state.log) - 1

    def update(self, idx: int, status: str, msg: str = "") -> None:
        st.session_state.log[idx][1:] = [status, msg]
        self.draw()

    def draw(self) -> None:
        rows = "".join(
            f'<div class="step">{self.ICON[s]} <b>{html.escape(a)}</b> '
            f'<span style="color:#a89a82">{html.escape(m)}</span></div>'
            for a, s, m in st.session_state.log
        )
        self.box.markdown(f'<div class="card">{rows or "No steps yet."}</div>', unsafe_allow_html=True)


# ------------------------------------------------------------------ API / Gemini
def get_api_key() -> str | None:
    try:
        if "GEMINI_API_KEY" in st.secrets and st.secrets["GEMINI_API_KEY"]:
            return str(st.secrets["GEMINI_API_KEY"]).strip()
    except Exception:
        pass
    v = os.getenv("GEMINI_API_KEY")
    return v.strip() if v else None


def extract_json(text: str) -> dict[str, Any]:
    t = re.sub(r"^```(?:json)?|```$", "", text.strip(), flags=re.M).strip()
    a, b = t.find("{"), t.rfind("}")
    if a < 0 or b < 0:
        raise ValueError("Model did not return JSON.")
    return json.loads(t[a : b + 1])


def gjson(client: genai.Client, prompt: str) -> dict[str, Any]:
    last: Exception | None = None
    for _ in range(2):
        r = client.models.generate_content(
            model=GEN_MODEL,
            contents=prompt,
            config=types.GenerateContentConfig(response_mime_type="application/json", temperature=0.2),
        )
        try:
            return extract_json(r.text or "")
        except Exception as exc:  # retry once on malformed JSON
            last = exc
    raise ValueError(f"Could not parse model JSON: {last}")


def count_tokens(client: genai.Client, text: str) -> tuple[int, str]:
    try:
        return client.models.count_tokens(model=GEN_MODEL, contents=text).total_tokens, "Gemini tokenizer"
    except Exception:
        return math.ceil(len(text) / 4), "estimate (chars/4)"


# ------------------------------------------------------------ extraction / checks
def extract_text(name: str, data: bytes) -> str:
    ext = name.lower().rsplit(".", 1)[-1]
    if ext == "pdf":
        return "\n".join((p.extract_text() or "") for p in PdfReader(io.BytesIO(data)).pages)
    if ext == "docx":
        d = Document(io.BytesIO(data))
        parts = [p.text for p in d.paragraphs]
        for t in d.tables:
            for row in t.rows:
                parts.append(" | ".join(c.text.strip() for c in row.cells))
        return "\n".join(parts)
    return data.decode("utf-8", errors="ignore")


def normalize_text(t: str) -> str:
    t = t.replace("\r", "")
    t = re.sub(r"[ \t]+", " ", t)
    return re.sub(r"\n{3,}", "\n\n", t).strip()


def detect_sections(text: str) -> list[str]:
    found = []
    for key, pat in SECTION_PATTERNS.items():
        for line in text.splitlines():
            s = line.strip().strip(":").lower()
            if 0 < len(s) <= 45 and re.fullmatch(rf".*\b{pat}\b.*", s):
                found.append(key)
                break
    return found


def deterministic_checks(text: str) -> dict[str, Any]:
    return {
        "words": len(text.split()),
        "sections": detect_sections(text),
        "email": bool(re.search(r"[\w.+-]+@[\w-]+\.[\w.]+", text)),
        "phone": bool(re.search(r"(\+?\d[\d\s().-]{8,}\d)", text)),
    }


# ------------------------------------------------------------- vectors / RAG
def chunk_text(text: str, max_words: int = 70) -> list[str]:
    lines = [l.strip() for l in text.splitlines() if l.strip()]
    chunks, cur, w = [], [], 0
    for l in lines:
        n = len(l.split())
        if w + n > max_words and cur:
            chunks.append("\n".join(cur))
            cur = cur[-1:]
            w = len(cur[0].split())
        cur.append(l)
        w += n
    if cur:
        chunks.append("\n".join(cur))
    return chunks


def tok(s: str) -> list[str]:
    return re.findall(r"[a-z0-9+#.]{2,}", s.lower())


def hashed_embed(texts: list[str]) -> np.ndarray:
    """Deterministic offline fallback: unigram+bigram feature hashing, sublinear TF, L2-normalised."""
    m = np.zeros((len(texts), HASH_DIM), dtype=np.float32)
    for r, t in enumerate(texts):
        w = tok(t)
        counts: dict[str, int] = {}
        for f in w + [a + "_" + b for a, b in zip(w, w[1:])]:
            counts[f] = counts.get(f, 0) + 1
        for f, c in counts.items():
            h = zlib.crc32(f.encode())
            m[r, h % HASH_DIM] += (1 if (h >> 16) & 1 else -1) * (1 + math.log(c))
    return m / (np.linalg.norm(m, axis=1, keepdims=True) + 1e-9)


def embed_texts(client: genai.Client, texts: list[str], task: str) -> np.ndarray:
    out: list[list[float]] = []
    for i in range(0, len(texts), 50):
        r = client.models.embed_content(model=EMBED_MODEL, contents=texts[i : i + 50], config=types.EmbedContentConfig(task_type=task))
        out += [e.values for e in r.embeddings]
    m = np.array(out, dtype=np.float32)
    return m / (np.linalg.norm(m, axis=1, keepdims=True) + 1e-9)


def build_index(client: genai.Client, chunks: list[str]) -> dict[str, Any]:
    try:
        mat, kind = embed_texts(client, chunks, "RETRIEVAL_DOCUMENT"), "gemini"
    except Exception:
        mat, kind = hashed_embed(chunks), "local-hash"
    return {"chunks": chunks, "mat": mat, "kind": kind, "tok": [tok(c) for c in chunks]}


def embed_queries(client: genai.Client, index: dict[str, Any], qs: list[str]) -> np.ndarray:
    if index["kind"] == "gemini":
        try:
            return embed_texts(client, qs, "RETRIEVAL_QUERY")
        except Exception:
            pass
    return hashed_embed(qs) if index["kind"] == "local-hash" else np.zeros((len(qs), index["mat"].shape[1]), dtype=np.float32)


def bm25(q: str, index: dict[str, Any]) -> np.ndarray:
    docs, n = index["tok"], len(index["tok"])
    avg = sum(len(d) for d in docs) / max(n, 1)
    scores = np.zeros(n)
    for w in set(tok(q)):
        df = sum(1 for d in docs if w in d)
        if not df:
            continue
        idf = math.log(1 + (n - df + .5) / (df + .5))
        for i, d in enumerate(docs):
            tf = d.count(w)
            if tf:
                scores[i] += idf * tf * 2.2 / (tf + 1.2 * (.25 + .75 * len(d) / max(avg, 1)))
    return scores


def retrieve(client: genai.Client, index: dict[str, Any], queries: list[str], k: int, advanced: bool) -> list[tuple[int, float]]:
    n = len(index["chunks"])
    Q = embed_queries(client, index, queries)
    sims = Q @ index["mat"].T
    best = sims.max(axis=0)
    if not advanced:  # BASIC RAG: pure vector similarity
        return [(int(i), float(best[i])) for i in np.argsort(-best)[:k]]
    fused = np.zeros(n)  # ADVANCED RAG: hybrid (vector + BM25) -> RRF -> MMR
    for qi, q in enumerate(queries):
        for r, i in enumerate(np.argsort(-sims[qi])):
            fused[i] += 1 / (60 + r)
        for r, i in enumerate(np.argsort(-bm25(q, index))):
            fused[i] += 1 / (60 + r)
    mx = fused.max() or 1.0
    pool, sel = list(np.argsort(-fused)), []
    while pool and len(sel) < k:
        def mmr(i: int) -> float:
            div = max((float(index["mat"][i] @ index["mat"][j]) for j in sel), default=0.0)
            return .7 * fused[i] / mx - .3 * div
        pick = max(pool, key=mmr)
        pool.remove(pick)
        sel.append(int(pick))
    return [(i, float(best[i])) for i in sel]


# ----------------------------------------------------------------- prompts
ANALYSIS_SCHEMA = (
    '{"score":0-100,"summary":"1-2 sentences",'
    '"field_fit":{"score":0-100,"verdict":"Fits|Partial|Does not fit","matched":["..."],"gaps":["..."]},'
    '"notes":[{"title":"Skills|Experience|Format|Grammar","score":0-10,"text":"short","fix":"short"}],'
    '"issues":[{"section":"","problem":"","evidence":"EXACT quote copied from the CV text","fix":"rewrite using only facts already in the CV"}],'
    '"missing":[{"section":"","why":"","question":"question to ask the user"}],"recommendations":["..."]}'
)


def analysis_prompt(field: str, checks: dict[str, Any], ctx: str) -> str:
    return (
        f'You are a strict, honest CV reviewer. Target field: "{field}".\n'
        f"Deterministic facts: {checks['words']} words; section headings found: {', '.join(checks['sections']) or 'none'}; "
        f"email {'present' if checks['email'] else 'absent'}; phone {'present' if checks['phone'] else 'absent'}.\n"
        f"RETRIEVED CV CHUNKS:\n{ctx}\n\n"
        "Rules: judge ONLY what is in the chunks; never invent facts; 'evidence' must be an exact quote; 'missing' lists only "
        "standard sections absent from the CV; judge honestly whether the CV fits the field ('Does not fit' is allowed). "
        "Always give 4 notes (Skills, Experience, Format, Grammar).\nReturn JSON only, schema: " + ANALYSIS_SCHEMA
    )


def norm(s: Any) -> str:
    return re.sub(r"\s+", " ", str(s)).strip().lower()


def verify_analysis(a: dict[str, Any], text: str, found: list[str]) -> tuple[dict[str, Any], int]:
    src = norm(text)
    issues = [x for x in a.get("issues", []) if x.get("evidence") and norm(x["evidence"]) in src]
    dropped = len(a.get("issues", [])) - len(issues)
    a["issues"] = issues
    missing = a.get("missing") or []
    for k in CORE_SECTIONS:
        if k not in found and not any(k in norm(m.get("section", "")) for m in missing):
            missing.append({"section": k.title(), "why": "No heading detected in the CV.",
                            "question": f"Do you want to add a {k} section? If yes, share the real details."})
    a["missing"] = missing
    a["score"] = max(0, min(100, int(a.get("score") or 0)))
    a.setdefault("field_fit", {})
    return a, dropped


def cv_to_text(cv: dict[str, Any]) -> str:
    head = "\n".join(x for x in [cv.get("name", ""), " | ".join(cv.get("contact") or [])] if x)
    body = "\n\n".join(s["title"].upper() + "\n" + "\n".join("- " + b for b in s["bullets"]) for s in cv.get("sections", []))
    return head + "\n\n" + body


def generation_prompt(field: str, text: str, adds: list[str], fixes: list[str], remaining: str) -> str:
    return (
        f'Rewrite this CV for the field "{field}".\nSOURCE CV:\n{text}\n\n'
        f"USER-PROVIDED ADDITIONS (verbatim facts from the user):\n{chr(10).join(adds) or 'none'}\n\n"
        f"APPROVED FIXES:\n{chr(10).join(fixes) or 'none'}\n"
        + (f"\nREMAINING ISSUES TO FIX WITHOUT ADDING NEW FACTS:\n{remaining}\n" if remaining else "")
        + "\nHARD RULES: use ONLY facts present in SOURCE CV or USER-PROVIDED ADDITIONS. Never invent employers, dates, degrees, "
        "skills, numbers, metrics, tools or sections. Omit any section that has no source content. A summary is allowed only if "
        "built purely from stated facts. Improve grammar, tense, bullet structure and strong verbs; order sections for the field; "
        "plain ATS-safe headings; keep contact details exactly as in the source.\n"
        'Return JSON only: {"name":"","contact":["..."],"sections":[{"title":"","bullets":["..."]}]}'
    )


def numbers_in(s: str) -> list[str]:
    return [n.rstrip(".,") for n in re.findall(r"\d[\d,.%+]*", s)]


# ---------------------------------------------------------------- file output
def safe_filename(name: str) -> str:
    return re.sub(r"[^\w\-]+", "_", os.path.splitext(name)[0]).strip("_") or "CV"


def build_txt(cv: dict[str, Any]) -> bytes:
    return cv_to_text(cv).encode("utf-8")


def build_md(cv: dict[str, Any]) -> bytes:
    s = f"# {cv.get('name', '')}\n{' | '.join(cv.get('contact') or [])}\n\n"
    s += "\n\n".join(f"## {x['title']}\n" + "\n".join("- " + b for b in x["bullets"]) for x in cv["sections"])
    return s.encode("utf-8")


def build_docx(cv: dict[str, Any]) -> bytes:
    d = Document()
    for sec in d.sections:
        sec.left_margin = sec.right_margin = Inches(.8)
        sec.top_margin = sec.bottom_margin = Inches(.7)
    d.styles["Normal"].font.name = "Calibri"
    d.styles["Normal"].font.size = Pt(10.5)
    p = d.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p.add_run(cv.get("name", ""))
    r.bold, r.font.size = True, Pt(20)
    if cv.get("contact"):
        c = d.add_paragraph(" | ".join(cv["contact"]))
        c.alignment = WD_ALIGN_PARAGRAPH.CENTER
    for s in cv["sections"]:
        h = d.add_paragraph()
        hr = h.add_run(s["title"].upper())
        hr.bold, hr.font.size = True, Pt(11.5)
        for b in s["bullets"]:
            d.add_paragraph(b, style="List Bullet")
    out = io.BytesIO()
    d.save(out)
    return out.getvalue()


def build_pdf(cv: dict[str, Any]) -> bytes:
    from reportlab.lib.enums import TA_CENTER
    from reportlab.lib.pagesizes import A4
    from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
    from reportlab.lib.units import inch
    from reportlab.platypus import Paragraph, SimpleDocTemplate
    from xml.sax.saxutils import escape

    out = io.BytesIO()
    doc = SimpleDocTemplate(out, pagesize=A4, rightMargin=.65 * inch, leftMargin=.65 * inch, topMargin=.55 * inch, bottomMargin=.55 * inch)
    ss = getSampleStyleSheet()
    title = ParagraphStyle("T", parent=ss["Title"], fontName="Helvetica-Bold", fontSize=20, leading=23, alignment=TA_CENTER, spaceAfter=4)
    sub = ParagraphStyle("S", parent=ss["BodyText"], fontSize=9.5, alignment=TA_CENTER, spaceAfter=6)
    head = ParagraphStyle("H", parent=ss["Heading2"], fontName="Helvetica-Bold", fontSize=11.5, spaceBefore=8, spaceAfter=3)
    body = ParagraphStyle("B", parent=ss["BodyText"], fontSize=9.5, leading=12.5, leftIndent=12, firstLineIndent=-7, spaceAfter=2)
    story = [Paragraph(escape(cv.get("name", "")), title), Paragraph(escape(" | ".join(cv.get("contact") or [])), sub)]
    for s in cv["sections"]:
        story.append(Paragraph(escape(s["title"].upper()), head))
        story += [Paragraph(escape("• " + b), body) for b in s["bullets"]]
    doc.build(story)
    return out.getvalue()


def cv_html(cv: dict[str, Any]) -> str:
    e = html.escape
    return (f"<h4>{e(cv.get('name', ''))}</h4><div>{e(' · '.join(cv.get('contact') or []))}</div>"
            + "".join(f"<h5>{e(s['title'])}</h5><ul>" + "".join(f"<li>{e(b)}</li>" for b in s["bullets"]) + "</ul>" for s in cv["sections"]))


# ------------------------------------------------------------------- renderers
def render_lab(ing: dict[str, Any], hits: list[tuple[int, float]] | None) -> None:
    idx = ing["index"]
    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Tokens", f"{ing['tokens']:,}", help=ing["token_src"])
    c2.metric("Chunks", len(idx["chunks"]))
    c3.metric("Vector dims", idx["mat"].shape[1])
    c4.metric("Embeddings", "Gemini" if idx["kind"] == "gemini" else "Local hash")
    st.caption("Embedding of chunk C1 (first 64 dimensions)")
    st.bar_chart(idx["mat"][0][:64], height=110)
    if idx["kind"] != "gemini":
        st.caption("⚠ Gemini embeddings were unavailable, so deterministic local hashed vectors are used.")
    if hits:
        st.markdown("**Retrieved context (cosine similarity)**")
        for i, sim in hits:
            st.progress(max(0.0, min(1.0, sim)), text=f"C{i + 1} · similarity {sim:.3f} · ~{math.ceil(len(idx['chunks'][i]) / 4)} tokens")
            st.caption(idx["chunks"][i][:220].replace("\n", " "))


def render_analysis(a: dict[str, Any], field: str) -> None:
    ff = a.get("field_fit", {})
    st.markdown(f"## Review: {html.escape(field)}")
    c1, c2, c3 = st.columns([1, 1, 2])
    c1.metric("ATS score", f"{a['score']}/100")
    c2.metric("Field fit", f"{int(ff.get('score') or 0)}/100", ff.get("verdict", ""))
    c3.write(a.get("summary", ""))
    pills = "".join(f'<span class="pill g">✓ {html.escape(str(x))}</span>' for x in ff.get("matched", []))
    pills += "".join(f'<span class="pill b">✕ {html.escape(str(x))}</span>' for x in ff.get("gaps", []))
    st.markdown(pills, unsafe_allow_html=True)
    notes = a.get("notes", [])
    if notes:
        cols = st.columns(min(4, len(notes)))
        for col, n in zip(cols, notes):
            col.markdown(f'<div class="note"><h4>{html.escape(str(n.get("title")))} · {int(n.get("score") or 0)}/10</h4>'
                         f'<div>{html.escape(str(n.get("text", "")))}</div><hr style="border-top:1px dashed #2a2114">'
                         f'<small>{html.escape(str(n.get("fix", "")))}</small></div>', unsafe_allow_html=True)
    st.markdown("### Key recommendations")
    st.markdown('<ul class="dia">' + "".join(f"<li>{html.escape(str(x))}</li>" for x in a.get("recommendations", [])) + "</ul>", unsafe_allow_html=True)


# ------------------------------------------------------------------- pipelines
def run_analysis(client: genai.Client, ing: dict[str, Any], field: str, advanced: bool, L: Logger) -> dict[str, Any]:
    idx = ing["index"]
    queries = [f"{field} skills experience qualifications"]
    if advanced:
        i = L.add("3 · Query expansion", "run")
        try:
            r = gjson(client, f'Return JSON {{"queries":["..."]}} with 5 short search queries (max 8 words each) to find CV evidence relevant to the field "{field}".')
            queries += [str(q) for q in r.get("queries", [])[:5]]
            L.update(i, "ok", f"{len(queries)} queries")
        except Exception as exc:
            L.update(i, "warn", f"skipped: {exc}")
    i = L.add("4 · Retrieval", "run", "hybrid vector+BM25 → RRF → MMR" if advanced else "vector cosine top-k")
    hits = retrieve(client, idx, queries, 8 if advanced else 4, advanced)
    if not hits:
        L.update(i, "err", "nothing retrieved")
        raise RuntimeError("Retrieval returned no chunks.")
    L.update(i, "ok", f"{len(hits)} chunks · top similarity {hits[0][1]:.3f}")
    st.session_state.hits = hits
    ctx = "\n\n".join(f"[C{k + 1}]\n{idx['chunks'][k]}" for k, _ in hits)
    i = L.add("5 · AI analysis", "run", "grounded in retrieved chunks")
    a = gjson(client, analysis_prompt(field, ing["checks"], ctx))
    L.update(i, "ok", f"score {a.get('score')} · field fit {a.get('field_fit', {}).get('verdict')}")
    i = L.add("6 · Recheck analysis", "run")
    a, dropped = verify_analysis(a, ing["text"], ing["checks"]["sections"])
    L.update(i, "ok", f"{dropped} ungrounded issue(s) removed · {len(a['issues'])} verified against CV text")
    return a


def run_generation(client: genai.Client, ing: dict[str, Any], field: str, adds: list[str], fixes: list[str], L: Logger) -> tuple[dict[str, Any], dict[str, Any]]:
    remaining, cv, sc = "", {}, {}
    for p in (1, 2):
        i = L.add(f"7.{p} · Generate CV", "run", "facts from your CV + your answers only")
        src = ing["text"] + "\n" + "\n".join(adds)
        cv = gjson(client, generation_prompt(field, ing["text"], adds, fixes, remaining))
        cv.setdefault("sections", [])
        for s in cv["sections"]:
            s["bullets"] = [str(b) for b in s.get("bullets", [])]
        L.update(i, "ok", f"{len(cv_to_text(cv).split())} words drafted")

        i = L.add(f"8.{p} · Fact check (local numbers/contacts)", "run")
        sn, removed = norm(src), 0
        for s in cv["sections"]:
            keep = [b for b in s["bullets"] if all(n.lower() in sn for n in numbers_in(b))]
            removed += len(s["bullets"]) - len(keep)
            s["bullets"] = keep
        cv["contact"] = [c for c in (cv.get("contact") or []) if norm(c) in sn]
        L.update(i, "warn" if removed else "ok", f"{removed} line(s) with unsupported numbers removed" if removed else "every number traced to your sources")

        i = L.add(f"9.{p} · Fact check (AI grounding)", "run")
        try:
            v = gjson(client, "Compare the GENERATED CV against SOURCE (CV + user additions). List every generated line containing a claim "
                      "(employer, role, date, degree, skill, tool, metric) NOT supported by SOURCE.\n"
                      f"SOURCE:\n{src}\n\nGENERATED:\n{cv_to_text(cv)}\n\n"
                      'Return JSON only: {"unsupported":[{"line":"exact generated line","reason":""}]}')
            bad = [norm(x.get("line", "")) for x in v.get("unsupported", []) if x.get("line")]
            n = 0
            for s in cv["sections"]:
                keep = [b for b in s["bullets"] if not any(u and (u in norm(b) or norm(b) in u) for u in bad)]
                n += len(s["bullets"]) - len(keep)
                s["bullets"] = keep
            L.update(i, "warn" if n else "ok", f"{n} unsupported line(s) removed" if n else "no unsupported claims found")
        except Exception as exc:
            L.update(i, "warn", f"skipped: {exc}")
        cv["sections"] = [s for s in cv["sections"] if s["bullets"]]

        i = L.add(f"10.{p} · Re-score generated CV", "run")
        sc = gjson(client, f'Score this CV for the field "{field}" (ATS readiness, clarity, grammar, structure, field relevance). '
                   "Be honest: 100 only if nothing can be improved from the facts present.\nCV:\n" + cv_to_text(cv) +
                   '\nReturn JSON only: {"score":0-100,"field_fit":0-100,"issues":[{"problem":"","fix":"fix possible WITHOUT new facts, or \'needs user info\'"}],"summary":""}')
        L.update(i, "ok", f"score {sc.get('score')} · field fit {sc.get('field_fit')}")
        fixable = [x for x in sc.get("issues", []) if "needs user info" not in str(x.get("fix", "")).lower()]
        if int(sc.get("score") or 0) >= 100 or not fixable:
            break
        remaining = "\n".join(f"{x.get('problem')} → {x.get('fix')}" for x in fixable)
    return cv, sc


def reset_for(key: str) -> None:
    for k in list(st.session_state.keys()):
        if k != "_key":
            del st.session_state[k]
    st.session_state["_key"] = key


# ------------------------------------------------------------------------- main
def main() -> None:
    inject_css()
    st.markdown('<div class="logo">CV<i>tro</i></div><div class="tag">BETWEEN YOU AND YOUR SUCCESS</div>', unsafe_allow_html=True)
    st.caption("Upload your CV, name your target field, and get a grounded RAG review and a no-invention rewrite.")
    api_key = get_api_key()
    if not api_key:
        st.error("GEMINI_API_KEY is missing. Add it to Streamlit secrets or set it as an environment variable.")
        st.stop()
    client = genai.Client(api_key=api_key)

    tab_up, tab_paste = st.tabs(["Upload CV", "Paste text"])
    with tab_up:
        up = st.file_uploader("Drag & drop your CV (PDF, DOCX or TXT)", type=["pdf", "docx", "txt"])
    with tab_paste:
        pasted = st.text_area("Paste CV text", height=160)
    raw_name, raw_bytes = (up.name, up.getvalue()) if up else (("pasted.txt", pasted.encode()) if pasted.strip() else (None, None))
    if raw_name is None:
        st.info("Upload or paste a CV to begin.")
        return
    if len(raw_bytes) > MAX_FILE_BYTES:
        st.error("File is larger than 12 MB.")
        return

    key = hashlib.sha256(raw_bytes).hexdigest()
    if st.session_state.get("_key") != key:
        reset_for(key)

    st.markdown("## Pipeline")
    L = Logger(st.empty())

    if "ing" not in st.session_state:
        try:
            i = L.add("1 · Extract text", "run", raw_name)
            text = normalize_text(extract_text(raw_name, raw_bytes))[:MAX_CV_CHARS]
            checks = deterministic_checks(text)
            if checks["words"] < 40:
                L.update(i, "err", f"{checks['words']} words — too little text (scanned PDF?). Paste the text instead.")
                return
            L.update(i, "ok", f"{checks['words']} words · sections: {', '.join(checks['sections']) or 'none'} · email {'found' if checks['email'] else 'not found'}")
            i = L.add("2 · Chunk + embed + tokens", "run")
            chunks = chunk_text(text)
            index = build_index(client, chunks)
            norms = np.linalg.norm(index["mat"], axis=1)
            tokens, tsrc = count_tokens(client, text)
            ok = bool(np.all(np.abs(norms - 1) < 0.01))
            L.update(i, "ok" if ok else "warn", f"{len(chunks)} chunks · {index['mat'].shape[1]}-dim {index['kind']} vectors · unit-norm check {'passed' if ok else 'failed'} · {tokens} tokens")
            st.session_state.ing = {"text": text, "checks": checks, "index": index, "tokens": tokens, "token_src": tsrc, "name": raw_name}
        except Exception as exc:
            st.error(f"Could not process the CV: {exc}")
            return
    else:
        L.draw()
    ing = st.session_state.ing

    st.markdown("## Vector Lab")
    render_lab(ing, st.session_state.get("hits"))

    st.markdown("## Review setup")
    c1, c2 = st.columns(2)
    field = c1.text_input("Target field (what is this CV for?)", placeholder="e.g. Audit & Assurance, Data Analyst, Teaching")
    mode = c2.radio("Retrieval mode", ["Advanced RAG", "Basic RAG"], horizontal=True,
                    help="Advanced: query expansion, hybrid vector+BM25, RRF fusion, MMR. Basic: vector similarity only.")
    if st.button("Analyze CV", type="primary"):
        if not field.strip():
            st.warning("Please enter the target field first.")
        else:
            try:
                st.session_state.pop("cv", None)
                st.session_state.analysis = run_analysis(client, ing, field.strip(), mode == "Advanced RAG", L)
                st.session_state.field = field.strip()
                st.rerun()
            except Exception as exc:
                L.add("Error", "err", str(exc))

    a = st.session_state.get("analysis")
    if not a:
        return
    render_analysis(a, st.session_state.field)

    st.markdown("## Your decisions")
    st.caption("Nothing is added unless you choose to and provide the facts yourself.")
    adds: list[str] = []
    for i, m in enumerate(a["missing"]):
        st.markdown(f"**{m.get('section')}** — {m.get('why')}  \n*{m.get('question')}*")
        if st.checkbox("Yes, I want to add this", key=f"add_{i}"):
            t = st.text_area("Write the real details (only what you type is used)", key=f"txt_{i}", height=90)
            if t.strip():
                adds.append(f"{m.get('section')}: {t.strip()}")
    fixes: list[str] = []
    if a["issues"]:
        st.markdown("**Proposed changes — tick what you approve**")
        for i, x in enumerate(a["issues"]):
            if st.checkbox(f"{x.get('section')}: {x.get('problem')}  →  {x.get('fix')}", value=True, key=f"fix_{i}", help=f"Evidence: “{x.get('evidence')}”"):
                fixes.append(f"{x.get('section')}: {x.get('fix')}")

    if st.button("✨ Generate my best CV (target 100)", type="primary"):
        try:
            cv, sc = run_generation(client, ing, st.session_state.field, adds, fixes, L)
            if not cv["sections"]:
                st.error("Nothing verifiable remained to build a CV from. Add real details above and try again.")
            else:
                st.session_state.cv, st.session_state.sc = cv, sc
        except Exception as exc:
            L.add("Error", "err", str(exc))

    cv, sc = st.session_state.get("cv"), st.session_state.get("sc")
    if not cv:
        return
    st.markdown("## Your generated CV")
    m1, m2 = st.columns(2)
    m1.metric("Score", f"{int(sc.get('score') or 0)}/100", f"{int(sc.get('score') or 0) - a['score']:+d} vs original")
    m2.metric("Field fit", f"{int(sc.get('field_fit') or 0)}/100")
    st.write(sc.get("summary", ""))
    gaps = [x for x in sc.get("issues", []) if "needs user info" in str(x.get("fix", "")).lower()]
    if int(sc.get("score") or 0) < 100:
        st.info("The score is below 100 because the remaining gaps need real information from you — nothing was invented to inflate it."
                + ("\n\n" + "\n".join(f"- {x.get('problem')}" for x in gaps) if gaps else ""))
    st.markdown(f'<div class="cvp">{cv_html(cv)}</div>', unsafe_allow_html=True)

    base = safe_filename(ing["name"])
    d1, d2, d3, d4 = st.columns(4)
    try:
        d1.download_button("Download PDF", build_pdf(cv), f"{base}_CVtro.pdf", "application/pdf", use_container_width=True)
    except Exception as exc:
        d1.caption(f"PDF unavailable: {exc}")
    d2.download_button("Download DOCX", build_docx(cv), f"{base}_CVtro.docx",
                       "application/vnd.openxmlformats-officedocument.wordprocessingml.document", use_container_width=True)
    d3.download_button("Download TXT", build_txt(cv), f"{base}_CVtro.txt", "text/plain", use_container_width=True)
    d4.download_button("Download MD", build_md(cv), f"{base}_CVtro.md", "text/markdown", use_container_width=True)


if __name__ == "__main__":
    main()
