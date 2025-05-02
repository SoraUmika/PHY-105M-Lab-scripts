
from __future__ import annotations
import argparse
import io
import json
import os
import re
import sys
import time
import textwrap
from pathlib import Path
from typing import Any, Dict, List

import fitz  # PyMuPDF
import pytesseract
from PIL import Image
from docx import Document
from docx2pdf import convert
from openai import OpenAI, RateLimitError

# ----------------------------
# Configuration
# ----------------------------
TOTAL_POINTS = 30                    
MODEL_NAME = "gpt-4.1"         
DEEP_SEEK_MODEL = "Deepseek-R1"     
TEMP = 0.0                            
MAX_RETRIES = 3                       
REQUEST_TIMEOUT = 120                 
API_KEY = ""       
MIN_RUBRIC_LEN = 100  # characters
client = OpenAI(api_key=API_KEY)

# ----------------------------
# Helpers – (PDF/DOCX → text)
# ----------------------------
DEBUG_SAVE_PROMPTS = False 

DEFAULT_RUBRIC = textwrap.dedent("""

Scoring key per category: 0 Missing, 1‑2 Poor, 3‑4 Good, 5 Excellent.
Total available points: 30.
""")
def extract_text_from_pdf(pdf: Path) -> str:
    try:
        doc = fitz.open(pdf)
    except Exception:
        return ""
    text: List[str] = [page.get_text() for page in doc]
    return "\n".join(text)


def extract_text_from_docx(docx: Path) -> str:
    try:
        doc = Document(docx)
    except Exception:
        return ""
    parts = [p.text for p in doc.paragraphs]
    for t in doc.tables:
        for r in t.rows:
            for c in r.cells:
                parts.append(c.text)
    return "\n".join(parts)


def parse_lab_report(fp: Path, tmp: Path) -> str:
    if fp.suffix.lower() == ".pdf":
        return extract_text_from_pdf(fp)
    try:
        pdf = tmp / f"{fp.stem}.pdf"
        convert(str(fp), str(pdf))
        return extract_text_from_pdf(pdf)
    except Exception:
        return extract_text_from_docx(fp)

# ----------------------------
# Grading via OpenAI (core)
# ----------------------------

def _json(raw: str) -> Dict[str, Any]:
    """Return dict from the first JSON‑like object found in *raw*.
    Accept either strict JSON (double‑quoted) **or** Python‑style dict with
    single quotes – falling back to `ast.literal_eval` when needed.
    """
    import ast, json
    # strip code fences if the model wrapped the reply
    cleaned = re.sub(r"```[a-zA-Z]*", "", raw).strip("`")
    # grab the first brace‑to‑brace chunk
    m = re.search(r"\{.*\}", cleaned, flags=re.S)
    if not m:
        raise ValueError("no object in response")
    blob = m.group(0)
    try:
        return json.loads(blob)
    except json.JSONDecodeError:
        # last resort: single quotes → double quotes for keys/strings
        try:
            return ast.literal_eval(blob)
        except Exception as exc:
            raise ValueError(f"could not parse JSON/py‑dict: {exc}")

def _sum(cs):
    return sum(int(c.get("score", 0)) for c in cs)

def grade(rubric: str, manual: str, student: str, tag: str) -> Dict[str, Any]:
    sys_msg = "You are a strict grading assistant."  # keep short
    usr = textwrap.dedent(f"""
    Rubric:
    {rubric}
    ---
    Manual:
    {manual}
    ---
    Report:
    {student}
    Return only JSON {{'categories':[{{'name':'','score':0,'max':0,'feedback':''}}],'total':0,'overall_feedback':''}} with total = {TOTAL_POINTS}.
    """)
    for k in range(MAX_RETRIES):
        try:
            raw = client.chat.completions.create(
                timeout=REQUEST_TIMEOUT,
                model=MODEL_NAME,
                temperature=TEMP,
                messages=[{"role": "system", "content": sys_msg}, {"role": "user", "content": usr}],
            ).choices[0].message.content
            data = _json(raw)
            if _sum(data["categories"]) == data["total"] == TOTAL_POINTS:
                return data
        except Exception as e:
            print("retry", k+1, e)
            time.sleep(1)
    raise RuntimeError("grading failed")

# ----------------------------
# Rubric/manual loaders
# ----------------------------

def _read(p: Path) -> str:
    if p.suffix.lower() == ".pdf":
        return extract_text_from_pdf(p)
    if p.suffix.lower() == ".docx":
        return extract_text_from_docx(p)
    return p.read_text(encoding="utf-8")


def load_rubric(folder: Path, arg: Path | None) -> str:
    if arg:
        text = _read(arg)
        print("Rubric:", arg.name)
    else:
        cand = next((p for p in folder.iterdir() if "rubric" in p.name.lower() and p.suffix.lower() in {".docx", ".pdf", ".txt"}), None)
        text = _read(cand) if cand else ""
        print("Rubric:", cand.name if cand else "<built‑in>")
    return text if len(text) > MIN_RUBRIC_LEN else DEFAULT_RUBRIC

# ----------------------------
# Orchestrator
# ----------------------------

def grade_folder(folder: Path, rubric_arg: Path | None, manual_arg: Path | None):
    folder = folder.resolve()
    tmp = folder / "converted_pdf"; tmp.mkdir(exist_ok=True)
    out = folder / "graded_reports"; out.mkdir(exist_ok=True)

    rubric_text = load_rubric(folder, rubric_arg)

    manual_path = manual_arg or next((p for p in folder.iterdir() if "lab" in p.name.lower() and p.suffix.lower()==".pdf"), None)
    if not manual_path:
        sys.exit("Lab manual PDF not found (use --manual)")
    manual_text = _read(manual_path)
    print("Manual:", manual_path.name)

    for fp in folder.iterdir():
        if fp.suffix.lower() not in {".pdf", ".doc", ".docx"} or "rubric" in fp.name.lower() or fp == manual_path:
            continue
        print("Grading", fp.name)
        student = parse_lab_report(fp, tmp)
        if not student.strip():
            print("  skip: no text")
            continue
        try:
            res = grade(rubric_text, manual_text, student, fp.stem)
        except Exception as e:
            print("  fail:", e)
            continue
        (out/f"{fp.stem}_grade.json").write_text(json.dumps(res, indent=2, ensure_ascii=False))
        summary = [f"Score {res['total']}/{TOTAL_POINTS}", "", "Breakdown:"]
        summary += [f"- {c['name']}: {c['score']}/{c['max']} — {c['feedback']}" for c in res['categories']]
        summary += ["", "Overall:", res['overall_feedback']]
        (out/f"{fp.stem}_grade.txt").write_text("\n".join(summary))
        print("  saved")

# ----------------------------
# CLI
# ----------------------------

def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("folder", type=Path)
    ap.add_argument("--rubric", type=Path)
    ap.add_argument("--manual", type=Path)
    args = ap.parse_args()
    grade_folder(args.folder, args.rubric, args.manual)

if __name__ == "__main__":
    main()
