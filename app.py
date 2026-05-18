"""Skripsi Beres — Daftar Pustaka Generator + Format Skripsi Auto-Fix.

Two-feature toolkit untuk mahasiswa tingkat akhir:
1. Daftar Pustaka Generator: input DOI/judul/URL → format APA/MLA/IEEE/Harvard
2. Format Skripsi Auto-Fix: upload .docx → auto-fix margin/font/spasi/heading
"""

from __future__ import annotations

import io
import re
from typing import Any

import requests
from flask import Flask, jsonify, render_template, request, send_file
from docx import Document
from docx.shared import Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING

app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = 16 * 1024 * 1024  # 16 MB

# -----------------------------------------------------------------------------
# Crossref API helper
# -----------------------------------------------------------------------------

CROSSREF_API = "https://api.crossref.org/works"
DOI_REGEX = re.compile(r"\b(10\.\d{4,9}/[-._;()/:A-Z0-9]+)\b", re.IGNORECASE)


def fetch_crossref_by_doi(doi: str) -> dict[str, Any] | None:
    """Resolve a DOI to a Crossref work record."""
    try:
        r = requests.get(
            f"{CROSSREF_API}/{doi}",
            headers={"User-Agent": "SkripsiBeres/1.0 (mailto:contact@skripsi-beres.app)"},
            timeout=10,
        )
        if r.status_code == 200:
            return r.json().get("message")
    except requests.RequestException:
        return None
    return None


def search_crossref_by_title(title: str, limit: int = 5) -> list[dict[str, Any]]:
    """Search Crossref by free-text query (judul/keyword)."""
    try:
        r = requests.get(
            CROSSREF_API,
            params={"query.bibliographic": title, "rows": limit},
            headers={"User-Agent": "SkripsiBeres/1.0 (mailto:contact@skripsi-beres.app)"},
            timeout=10,
        )
        if r.status_code == 200:
            return r.json().get("message", {}).get("items", [])
    except requests.RequestException:
        return []
    return []


# -----------------------------------------------------------------------------
# Citation formatters
# -----------------------------------------------------------------------------

def _authors(work: dict[str, Any]) -> list[tuple[str, str]]:
    """Return [(family, given), ...]."""
    return [
        (a.get("family", "").strip(), a.get("given", "").strip())
        for a in work.get("author", [])
        if a.get("family")
    ]


def _year(work: dict[str, Any]) -> str:
    issued = work.get("issued", {}).get("date-parts", [[None]])
    return str(issued[0][0]) if issued and issued[0][0] else "n.d."


def _title(work: dict[str, Any]) -> str:
    titles = work.get("title", [])
    return titles[0] if titles else ""


def _container(work: dict[str, Any]) -> str:
    c = work.get("container-title", [])
    return c[0] if c else ""


def _volume_issue_pages(work: dict[str, Any]) -> str:
    parts = []
    if v := work.get("volume"):
        parts.append(str(v))
    if i := work.get("issue"):
        parts[-1] = f"{parts[-1]}({i})" if parts else f"({i})"
    page = work.get("page", "")
    return ", ".join(parts) + (f", {page}" if page else "")


def _doi(work: dict[str, Any]) -> str:
    return work.get("DOI", "")


def format_apa(work: dict[str, Any]) -> str:
    """APA 7th edition."""
    auths = _authors(work)
    if not auths:
        author_str = ""
    elif len(auths) == 1:
        f, g = auths[0]
        author_str = f"{f}, {''.join(p[0] + '.' for p in g.split() if p)}"
    elif len(auths) <= 20:
        parts = [f"{f}, {''.join(p[0] + '.' for p in g.split() if p)}" for f, g in auths]
        author_str = ", ".join(parts[:-1]) + ", & " + parts[-1]
    else:
        parts = [f"{f}, {''.join(p[0] + '.' for p in g.split() if p)}" for f, g in auths[:19]]
        last_f, last_g = auths[-1]
        author_str = ", ".join(parts) + f", ... {last_f}, {''.join(p[0] + '.' for p in last_g.split() if p)}"

    year = _year(work)
    title = _title(work)
    container = _container(work)
    vip = _volume_issue_pages(work)
    doi = _doi(work)

    out = f"{author_str} ({year}). {title}."
    if container:
        out += f" *{container}*"
        if vip:
            out += f", {vip}"
        out += "."
    if doi:
        out += f" https://doi.org/{doi}"
    return out


def format_mla(work: dict[str, Any]) -> str:
    """MLA 9th edition."""
    auths = _authors(work)
    if not auths:
        author_str = ""
    elif len(auths) == 1:
        f, g = auths[0]
        author_str = f"{f}, {g}"
    elif len(auths) == 2:
        f1, g1 = auths[0]
        f2, g2 = auths[1]
        author_str = f"{f1}, {g1}, and {g2} {f2}"
    else:
        f, g = auths[0]
        author_str = f"{f}, {g}, et al"

    year = _year(work)
    title = _title(work)
    container = _container(work)
    vip = _volume_issue_pages(work)

    out = f'{author_str}. "{title}."'
    if container:
        out += f" *{container}*"
        if vip:
            out += f", {vip}"
    out += f", {year}."
    return out


def format_ieee(work: dict[str, Any], ref_num: int = 1) -> str:
    """IEEE."""
    auths = _authors(work)
    parts = []
    for f, g in auths:
        initials = " ".join(p[0] + "." for p in g.split() if p)
        parts.append(f"{initials} {f}")
    if len(parts) > 6:
        author_str = ", ".join(parts[:1]) + " et al."
    else:
        author_str = ", ".join(parts)

    year = _year(work)
    title = _title(work)
    container = _container(work)
    vip = _volume_issue_pages(work)

    out = f'[{ref_num}] {author_str}, "{title},"'
    if container:
        out += f" *{container}*"
        if vip:
            out += f", {vip}"
    out += f", {year}."
    return out


def format_harvard(work: dict[str, Any]) -> str:
    """Harvard."""
    auths = _authors(work)
    if not auths:
        author_str = ""
    elif len(auths) == 1:
        f, g = auths[0]
        initials = "".join(p[0] + "." for p in g.split() if p)
        author_str = f"{f}, {initials}"
    elif len(auths) <= 3:
        parts = [f"{f}, {''.join(p[0] + '.' for p in g.split() if p)}" for f, g in auths]
        author_str = ", ".join(parts[:-1]) + " and " + parts[-1]
    else:
        f, g = auths[0]
        author_str = f"{f}, {''.join(p[0] + '.' for p in g.split() if p)} et al."

    year = _year(work)
    title = _title(work)
    container = _container(work)
    vip = _volume_issue_pages(work)
    doi = _doi(work)

    out = f"{author_str} ({year}) '{title}'"
    if container:
        out += f", *{container}*"
        if vip:
            out += f", {vip}"
    out += "."
    if doi:
        out += f" doi: {doi}."
    return out


def format_vancouver(work: dict[str, Any], ref_num: int = 1) -> str:
    """Vancouver."""
    auths = _authors(work)
    parts = []
    for f, g in auths:
        initials = "".join(p[0] for p in g.split() if p)
        parts.append(f"{f} {initials}")
    if len(parts) > 6:
        author_str = ", ".join(parts[:6]) + ", et al"
    else:
        author_str = ", ".join(parts)

    year = _year(work)
    title = _title(work)
    container = _container(work)
    vip = _volume_issue_pages(work)

    out = f"{ref_num}. {author_str}. {title}."
    if container:
        out += f" {container}."
        out += f" {year}"
        if vip:
            out += f";{vip}"
        out += "."
    else:
        out += f" {year}."
    return out


FORMATTERS = {
    "apa": format_apa,
    "mla": format_mla,
    "ieee": format_ieee,
    "harvard": format_harvard,
    "vancouver": format_vancouver,
}


# -----------------------------------------------------------------------------
# DOCX auto-fixer
# -----------------------------------------------------------------------------

KAMPUS_TEMPLATES = {
    "umum": {
        "name": "Umum (Standar Indonesia)",
        "font": "Times New Roman",
        "size": 12,
        "line_spacing": 1.5,
        "margin_top": 3.0,
        "margin_right": 3.0,
        "margin_bottom": 3.0,
        "margin_left": 4.0,
        "first_line_indent": 1.27,
    },
    "ui": {
        "name": "Universitas Indonesia",
        "font": "Times New Roman",
        "size": 12,
        "line_spacing": 1.5,
        "margin_top": 3.0,
        "margin_right": 3.0,
        "margin_bottom": 3.0,
        "margin_left": 4.0,
        "first_line_indent": 1.27,
    },
    "ugm": {
        "name": "Universitas Gadjah Mada",
        "font": "Times New Roman",
        "size": 12,
        "line_spacing": 2.0,
        "margin_top": 4.0,
        "margin_right": 3.0,
        "margin_bottom": 3.0,
        "margin_left": 4.0,
        "first_line_indent": 1.27,
    },
    "itb": {
        "name": "Institut Teknologi Bandung",
        "font": "Times New Roman",
        "size": 12,
        "line_spacing": 1.5,
        "margin_top": 3.0,
        "margin_right": 2.5,
        "margin_bottom": 2.5,
        "margin_left": 4.0,
        "first_line_indent": 1.27,
    },
    "ipb": {
        "name": "Institut Pertanian Bogor",
        "font": "Times New Roman",
        "size": 12,
        "line_spacing": 1.5,
        "margin_top": 3.0,
        "margin_right": 3.0,
        "margin_bottom": 3.0,
        "margin_left": 4.0,
        "first_line_indent": 1.27,
    },
    "telkom": {
        "name": "Telkom University",
        "font": "Times New Roman",
        "size": 11,
        "line_spacing": 1.5,
        "margin_top": 3.0,
        "margin_right": 2.5,
        "margin_bottom": 2.5,
        "margin_left": 3.5,
        "first_line_indent": 1.0,
    },
    "binus": {
        "name": "Bina Nusantara University",
        "font": "Times New Roman",
        "size": 12,
        "line_spacing": 2.0,
        "margin_top": 3.0,
        "margin_right": 3.0,
        "margin_bottom": 3.0,
        "margin_left": 4.0,
        "first_line_indent": 1.27,
    },
}


def autofix_docx(file_stream: io.BytesIO, template_key: str = "umum") -> tuple[io.BytesIO, dict[str, Any]]:
    """Apply formatting fixes to an uploaded .docx according to a kampus template."""
    tpl = KAMPUS_TEMPLATES.get(template_key, KAMPUS_TEMPLATES["umum"])

    doc = Document(file_stream)

    stats = {
        "paragraphs_fixed": 0,
        "sections_fixed": 0,
        "headings_fixed": 0,
        "template": tpl["name"],
    }

    # Fix margins on every section
    for section in doc.sections:
        section.top_margin = Cm(tpl["margin_top"])
        section.right_margin = Cm(tpl["margin_right"])
        section.bottom_margin = Cm(tpl["margin_bottom"])
        section.left_margin = Cm(tpl["margin_left"])
        stats["sections_fixed"] += 1

    # Fix paragraphs (font, size, line spacing, alignment)
    for p in doc.paragraphs:
        style_name = (p.style.name or "").lower() if p.style else ""
        is_heading = style_name.startswith("heading")

        # Line spacing — body text only; headings keep their own spacing
        if not is_heading:
            p.paragraph_format.line_spacing_rule = WD_LINE_SPACING.MULTIPLE
            p.paragraph_format.line_spacing = tpl["line_spacing"]

        # Justify body paragraphs by default; headings stay as-is
        if not is_heading and p.text.strip():
            p.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY

        # Font + size on every run
        for run in p.runs:
            run.font.name = tpl["font"]
            if is_heading:
                # Bigger sizes for h1/h2/h3 progressively
                level = 1
                m = re.search(r"heading\s*(\d+)", style_name)
                if m:
                    level = int(m.group(1))
                heading_size = {1: 16, 2: 14, 3: 13}.get(level, 12)
                run.font.size = Pt(heading_size)
                run.font.bold = True
                run.font.color.rgb = RGBColor(0x09, 0x09, 0x0B)
                stats["headings_fixed"] += 1
            else:
                run.font.size = Pt(tpl["size"])

        stats["paragraphs_fixed"] += 1

    # Save
    out = io.BytesIO()
    doc.save(out)
    out.seek(0)
    return out, stats


# -----------------------------------------------------------------------------
# Routes
# -----------------------------------------------------------------------------

@app.route("/")
def index():
    return render_template("landing.html")


@app.route("/app")
def workflow():
    return render_template("app.html", templates=KAMPUS_TEMPLATES)


@app.route("/api/cite", methods=["POST"])
def api_cite():
    """Generate a single citation from DOI/URL/title input."""
    data = request.get_json(silent=True) or {}
    query = (data.get("query") or "").strip()
    style = (data.get("style") or "apa").lower()

    if not query:
        return jsonify({"error": "Input tidak boleh kosong"}), 400
    if style not in FORMATTERS:
        return jsonify({"error": f"Style tidak dikenal: {style}"}), 400

    # Try DOI first (covers raw DOI + DOI URL)
    doi_match = DOI_REGEX.search(query)
    work = None
    if doi_match:
        work = fetch_crossref_by_doi(doi_match.group(1))

    # Fallback: title search
    candidates = []
    if not work:
        candidates = search_crossref_by_title(query, limit=5)
        if candidates:
            work = candidates[0]

    if not work:
        return jsonify({
            "error": "Referensi tidak ditemukan. Coba input DOI lengkap atau judul yang lebih spesifik.",
            "source": "crossref",
        }), 404

    citation = FORMATTERS[style](work)

    # Lighter alternatives for "did you mean"
    alternatives = []
    if candidates and len(candidates) > 1:
        for c in candidates[1:5]:
            alternatives.append({
                "title": _title(c),
                "year": _year(c),
                "container": _container(c),
                "doi": _doi(c),
            })

    return jsonify({
        "citation": citation,
        "style": style,
        "source": "crossref",
        "metadata": {
            "title": _title(work),
            "year": _year(work),
            "container": _container(work),
            "doi": _doi(work),
            "authors": [f"{f}, {g}" for f, g in _authors(work)],
        },
        "alternatives": alternatives,
    })


@app.route("/api/cite-bulk", methods=["POST"])
def api_cite_bulk():
    """Bulk citation: each line is one query."""
    data = request.get_json(silent=True) or {}
    text = (data.get("text") or "").strip()
    style = (data.get("style") or "apa").lower()

    if not text:
        return jsonify({"error": "Input tidak boleh kosong"}), 400
    if style not in FORMATTERS:
        return jsonify({"error": f"Style tidak dikenal: {style}"}), 400

    lines = [l.strip() for l in text.splitlines() if l.strip()]
    results = []
    for idx, line in enumerate(lines, start=1):
        doi_match = DOI_REGEX.search(line)
        work = None
        if doi_match:
            work = fetch_crossref_by_doi(doi_match.group(1))
        if not work:
            cands = search_crossref_by_title(line, limit=1)
            if cands:
                work = cands[0]

        if not work:
            results.append({"index": idx, "input": line, "ok": False, "error": "Tidak ditemukan"})
            continue

        formatter = FORMATTERS[style]
        if style in ("ieee", "vancouver"):
            citation = formatter(work, idx)
        else:
            citation = formatter(work)

        results.append({"index": idx, "input": line, "ok": True, "citation": citation})

    return jsonify({"results": results, "style": style, "count": len(results)})


@app.route("/api/format-docx", methods=["POST"])
def api_format_docx():
    """Auto-fix a .docx file according to a kampus template."""
    if "file" not in request.files:
        return jsonify({"error": "File tidak ditemukan"}), 400
    f = request.files["file"]
    if not f.filename.lower().endswith(".docx"):
        return jsonify({"error": "Hanya file .docx yang didukung"}), 400

    template_key = (request.form.get("template") or "umum").lower()
    if template_key not in KAMPUS_TEMPLATES:
        return jsonify({"error": f"Template tidak dikenal: {template_key}"}), 400

    try:
        out, stats = autofix_docx(io.BytesIO(f.read()), template_key)
    except Exception as e:  # noqa: BLE001
        return jsonify({"error": f"Gagal memproses file: {e}"}), 500

    base = (f.filename or "skripsi").rsplit(".", 1)[0]
    download_name = f"{base}_formatted.docx"

    return send_file(
        out,
        as_attachment=True,
        download_name=download_name,
        mimetype="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    )


@app.route("/api/templates")
def api_templates():
    return jsonify(KAMPUS_TEMPLATES)


@app.route("/health")
def health():
    return jsonify({"status": "ok", "service": "skripsi-beres"})


if __name__ == "__main__":
    app.run(host="0.0.0.0", port=5099, debug=True)
