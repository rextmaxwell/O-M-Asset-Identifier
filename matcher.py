
port os
import re
import io
import hashlib
import time
from typing import List, Dict, Tuple, Optional

import pandas as pd
from rapidfuzz import fuzz
from docx import Document

# PDF imports
import fitz  # PyMuPDF
import pdfplumber

from PIL import Image
import pytesseract


# ---------- Text Extraction ----------

def extract_text_pdf(path: str, use_ocr_if_empty: bool = True, max_pages_ocr: int = 5) -> str:
    """
    Extract text from a PDF. Try normal text first (PyMuPDF). If little or no text,
    optionally OCR up to max_pages_ocr pages.
    """
    text_chunks = []

    try:
        with fitz.open(path) as doc:
            for page in doc:
                text = page.get_text("text")
                if text and text.strip():
                    text_chunks.append(text)
    except Exception as e:
        print(f"[WARN] PyMuPDF failed on {path}: {e}")

    text = "\n".join(text_chunks).strip()

    # If text is too short, try pdfplumber as fallback
    if len(text) < 100:
        try:
            with pdfplumber.open(path) as pdf:
                for page in pdf.pages:
                    t = page.extract_text() or ""
                    if t.strip():
                        text_chunks.append(t)
            text = "\n".join(text_chunks).strip()
        except Exception as e:
            print(f"[WARN] pdfplumber failed on {path}: {e}")

    # OCR if still minimal
    if use_ocr_if_empty and len(text) < 100:
        try:
            ocr_texts = []
            with fitz.open(path) as doc:
                pages_to_ocr = min(len(doc), max_pages_ocr)
                for i in range(pages_to_ocr):
                    page = doc.load_page(i)
                    pix = page.get_pixmap(matrix=fitz.Matrix(2, 2))  # upscale for OCR quality
                    img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
                    ocr_texts.append(pytesseract.image_to_string(img))
            text = "\n".join(ocr_texts).strip()
        except Exception as e:
            print(f"[WARN] OCR failed on {path}: {e}")

    return text


def extract_text_docx(path: str) -> str:
    try:
        doc = Document(path)
        return "\n".join(p.text for p in doc.paragraphs)
    except Exception as e:
        print(f"[WARN] DOCX extraction failed on {path}: {e}")
        return ""


def extract_text_any(path: str) -> str:
    ext = os.path.splitext(path)[1].lower()
    if ext == ".pdf":
        return extract_text_pdf(path)
    elif ext in (".docx",):
        return extract_text_docx(path)
    else:
        return ""


# ---------- Identifier Parsing ----------

ID_PATTERNS = [
    # Asset tag formats (customize to your org)
    r'\b[A-Z]{2,5}-\d{3,8}\b',        # ABC-123456
    r'\b[A-Z]{2,5}\d{3,8}\b',         # ABC123456
    r'\bTAG[-\s]?\d{3,8}\b',          # TAG-123456
]

SERIAL_PATTERNS = [
    r'\b(?:S\/?N|Serial(?:\sNo\.|\sNumber)?|SERIAL\sNO\.?)[:\s#\-]*([A-Z0-9\-]{5,20})\b',
    r'\bSN[:\s#\-]*([A-Z0-9\-]{5,20})\b',
]

MODEL_PATTERNS = [
    r'\bModel(?:\sNo\.|\sNumber)?[:\s#\-]*([A-Z0-9\-\._]{2,30})\b',
    r'\b(?:Type|Model)\s([A-Z0-9\-\._]{2,30})\b',
]

MANUFACTURER_PATTERNS = [
    r'\bManufacturer[:\s]*([A-Za-z0-9&\-\., ]{2,50})\b',
    r'\bMade by[:\s]*([A-Za-z0-9&\-\., ]{2,50})\b',
]

def find_with_patterns(text: str, patterns: List[str]) -> List[str]:
    found = []
    for pat in patterns:
        for m in re.finditer(pat, text, flags=re.IGNORECASE):
            if m.groups():
                val = m.group(1)
            else:
                val = m.group(0)
            val = val.strip().strip(":#- ").upper()
            if val and val not in found:
                found.append(val)
    return found


def parse_identifiers(text: str) -> Dict[str, List[str]]:
    return {
        "asset_ids": find_with_patterns(text, ID_PATTERNS),
        "serials": find_with_patterns(text, SERIAL_PATTERNS),
        "models": find_with_patterns(text, MODEL_PATTERNS),
        "manufacturers": find_with_patterns(text, MANUFACTURER_PATTERNS),
    }


# ---------- Normalization ----------

def normalize(s: Optional[str]) -> str:
    if s is None:
        return ""
    s = str(s)
    s = re.sub(r'[_\-\.\(\)\[\],]', ' ', s)
    s = re.sub(r'\s+', ' ', s).strip().lower()
    return s


def file_hash(path: str, algo: str = "sha256", chunk: int = 1024 * 1024) -> str:
    h = hashlib.new(algo)
    with open(path, "rb") as f:
        for block in iter(lambda: f.read(chunk), b''):
            h.update(block)
    return h.hexdigest()


# ---------- Scoring ----------

def score_candidate(file_meta: Dict, asset_row: Dict, signals: Dict) -> Tuple[int, List[str]]:
    score = 0
    reasons = []

    # Exact ID match
    candidate_ids = set(signals["asset_ids"])
    asset_tokens = set()
    for col in ("asset_id", "external_id", "tag", "name"):
        if col in asset_row and asset_row[col]:
            asset_tokens |= set(re.findall(r'[A-Z]{2,5}-\d{3,8}|[A-Z]{2,5}\d{3,8}', str(asset_row[col]).upper()))
    if candidate_ids & asset_tokens:
        score += 50
        reasons.append("id_match")

    # Serial match
    asset_serial = str(asset_row.get("serial", "")).upper()
    if asset_serial:
        if asset_serial in signals["serials"]:
            score += 25; reasons.append("serial_match")

    # Model match
    asset_model = str(asset_row.get("model", "")).upper()
    if asset_model:
        if asset_model in signals["models"]:
            score += 20; reasons.append("model_match")
        else:
            # Some models include hyphens/spaces; try loose contain
            for m in signals["models"]:
                if len(m) >= 4 and (m in asset_model or asset_model in m):
                    score += 15; reasons.append("model_partial"); break

    # Manufacturer match
    asset_mfr = normalize(asset_row.get("manufacturer", ""))
    for m in signals["manufacturers"]:
        if fuzz.token_set_ratio(asset_mfr, normalize(m)) >= 90:
            score += 10; reasons.append("manufacturer")

    # Fuzzy name match (title-like)
    if signals.get("title_terms"):
        sim = fuzz.token_set_ratio(normalize(asset_row.get("name", "")), normalize(signals["title_terms"]))
        if sim >= 90:
            score += 20; reasons.append(f"fuzzy_name_{sim}")
        elif sim >= 80:
            score += 10; reasons.append(f"fuzzy_name_{sim}")

    # Folder/project hint
    if asset_row.get("project") and asset_row["project"] and asset_row["project"].lower() in file_meta["dir"].lower():
        score += 10; reasons.append("folder_hint")

    # Hash bonus (if provided)
    if file_meta.get("hash") and asset_row.get("file_hash") and file_meta["hash"] == asset_row["file_hash"]:
        score = max(score, 100); reasons.append("hash_match")

    return score, reasons


# ---------- Matching ----------

def guess_title_terms(text: str) -> str:
    """
    Heuristic: take the top-of-doc lines (first 20 lines),
    remove boilerplate words, keep meaningful tokens.
    """
    lines = [ln.strip() for ln in text.splitlines()[:20] if ln.strip()]
    header = " ".join(lines)
    header = re.sub(r'\b(operations|maintenance|manual|instructions|guide|table of contents)\b', ' ', header, flags=re.I)
    header = re.sub(r'\s+', ' ', header).strip()
    return header


def match_files_to_assets(file_paths: List[str], assets_df: pd.DataFrame, compute_hash: bool = False) -> List[Dict]:
    # Pre-normalize assets
    assets = assets_df.to_dict(orient="records")

    results = []
    for path in file_paths:
        try:
            text = extract_text_any(path)
            signals = parse_identifiers(text)
            signals["title_terms"] = guess_title_terms(text)

            meta = {
                "file_path": path,
                "dir": os.path.dirname(path),
                "name": os.path.basename(path),
                "size": os.path.getsize(path),
                "mtime": os.path.getmtime(path),
                "hash": file_hash(path) if compute_hash else None,
            }

            scored = []
            for a in assets:
                sc, reasons = score_candidate(meta, a, signals)
                scored.append({
                    "asset_id": a.get("asset_id", None),
                    "name": a.get("name", ""),
                    "score": sc,
                    "reasons": ", ".join(reasons) if reasons else "",
                    "manufacturer": a.get("manufacturer", ""),
                    "model": a.get("model", ""),
                    "serial": a.get("serial", ""),
                    "external_id": a.get("external_id", ""),
                    "project": a.get("project", ""),
                })

            scored.sort(key=lambda x: x["score"], reverse=True)
            top = scored[:5]

            results.append({
                "file_path": path,
                "signals": signals,
                "top_candidates": top,
                "auto_choice": top[0] if top and top[0]["score"] >= 80 else None
            })
        except Exception as e:
            results.append({
                "file_path": path,
                "error": str(e),
                "top_candidates": []
            })

    return results


# ---------- Utility ----------

def save_results_csv(matches: List[Dict], save_path: str):
    rows = []
    for r in matches:
        for c in r.get("top_candidates", []):
            rows.append({
                "file_path": r["file_path"],
                "asset_id": c["asset_id"],
                "asset_name": c["name"],
                "score": c["score"],
                "reasons": c["reasons"],
                "manufacturer": c["manufacturer"],
                "model": c["model"],
                "serial": c["serial"],
                "external_id": c["external_id"],
                "project": c["project"],
            })
