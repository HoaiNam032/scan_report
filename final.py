import re
import os
import unicodedata
from dataclasses import dataclass
from typing import List, Dict, Tuple
from datetime import date

import numpy as np
import pandas as pd
import cv2
from pdf2image import convert_from_path
import pytesseract
from PIL import Image, ImageOps, ImageEnhance, ImageFilter
from openpyxl import load_workbook

# ================= CONFIG =================
POPPLER_PATH = r"D:\poppler-25.11.0\Library\bin"
#PDF_PATH = r"QĐ 884.pdf"   # đổi file đầu vào tại đây
OUTPUT_EXCEL = "ocr_output.xlsx"

pytesseract.pytesseract.tesseract_cmd = r"C:\Program Files\Tesseract-OCR\tesseract.exe"
os.environ["TESSDATA_PREFIX"] = r"C:\Program Files\Tesseract-OCR"

Image.MAX_IMAGE_PIXELS = None

DEBUG = True
DEBUG_SAVE = False
DEBUG_DIR = r"D:\scan\ocr_debug"

DPI_TABLE = 320
DPI_TEXT = 300

# table config
CMT_COL_INDEX = 4
STT_OFFSET = 4
DOB_MIN_YEAR = 1900
DOB_MAX_YEAR = date.today().year
# =========================================


# =========================================================
# COMMON UTILS
# =========================================================
def norm_spaces(s: str) -> str:
    s = s.replace("\r", "\n")
    s = s.replace("\n", " ")
    s = re.sub(r"\s+", " ", s)
    return s.strip()

def strip_accents(s: str) -> str:
    s = s.replace("đ", "d").replace("Đ", "D")
    s = unicodedata.normalize("NFD", s)
    s = "".join(ch for ch in s if unicodedata.category(ch) != "Mn")
    return s

def fix_ocr_digits(s: str) -> str:
    return s.translate(str.maketrans({
        "O": "0", "o": "0",
        "I": "1", "l": "1", "|": "1"
    }))

def fix_ocr_vietnamese_common(s: str) -> str:
    if not s:
        return s
    s = re.sub(r"\bThành\s*phô\b", "Thành phố", s, flags=re.I)
    s = re.sub(r"\bThị\s*trần\b", "Thị trấn", s, flags=re.I)
    return s

def get_all_pdfs_in_input_folder(folder="input") -> list[str]:
    os.makedirs(folder, exist_ok=True)
    pdfs = []
    for f in os.listdir(folder):
        if f.lower().endswith(".pdf"):
            pdfs.append(os.path.join(folder, f))
    return sorted(pdfs)

def save_excel_force_text(path: str, id_col_name: str):
    wb = load_workbook(path)
    ws = wb.active
    header = [ws.cell(row=1, column=c).value for c in range(1, ws.max_column + 1)]
    if id_col_name not in header:
        wb.save(path)
        return
    col_idx = header.index(id_col_name) + 1
    for r in range(2, ws.max_row + 1):
        cell = ws.cell(row=r, column=col_idx)
        if cell.value is None:
            continue
        cell.value = str(cell.value)
        cell.number_format = "@"
    wb.save(path)


# =========================================================
# LAYOUT DETECTION (TABLE vs NON-TABLE)
# =========================================================
def find_lines_1d(proj: np.ndarray, min_gap: int, thresh_ratio: float) -> List[int]:
    mx = float(proj.max()) if proj.max() != 0 else 1.0
    idx = np.where(proj >= mx * thresh_ratio)[0]
    if len(idx) == 0:
        return []

    groups = []
    s = int(idx[0]); p = int(idx[0])
    for i in idx[1:]:
        i = int(i)
        if i - p > 1:
            groups.append((s, p))
            s = i
        p = i
    groups.append((s, p))

    centers = [int((a + b) // 2) for a, b in groups]
    centers.sort()

    merged = []
    for c in centers:
        if not merged or (c - merged[-1] > min_gap):
            merged.append(c)
    return merged

def detect_table_grid(gray: np.ndarray) -> Tuple[List[int], List[int]]:
    _, thr = cv2.threshold(gray, 0, 255, cv2.THRESH_BINARY_INV + cv2.THRESH_OTSU)
    H, W = thr.shape[:2]

    vert_kernel = cv2.getStructuringElement(cv2.MORPH_RECT, (1, max(20, H // 35)))
    hor_kernel = cv2.getStructuringElement(cv2.MORPH_RECT, (max(20, W // 35), 1))

    vertical = cv2.erode(thr, vert_kernel, iterations=1)
    vertical = cv2.dilate(vertical, vert_kernel, iterations=2)

    horizontal = cv2.erode(thr, hor_kernel, iterations=1)
    horizontal = cv2.dilate(horizontal, hor_kernel, iterations=2)

    x_proj = vertical.sum(axis=0)
    y_proj = horizontal.sum(axis=1)

    x_lines = find_lines_1d(x_proj, min_gap=max(8, W // 220), thresh_ratio=0.30)
    y_lines = find_lines_1d(y_proj, min_gap=max(8, H // 220), thresh_ratio=0.30)

    if len(x_lines) < 5 or len(y_lines) < 5:
        x_lines = find_lines_1d(x_proj, min_gap=max(6, W // 260), thresh_ratio=0.22)
        y_lines = find_lines_1d(y_proj, min_gap=max(6, H // 260), thresh_ratio=0.22)

    return sorted(x_lines), sorted(y_lines)

def detect_layout_type(pdf_path: str) -> str:
    """
    Auto detect:
    - nếu trang đầu có đủ line ngang dọc kiểu bảng -> "table"
    - ngược lại -> "text"
    """
    pages = convert_from_path(pdf_path, dpi=220, poppler_path=POPPLER_PATH, first_page=1, last_page=1)
    if not pages:
        return "text"

    img = cv2.cvtColor(np.array(pages[0].convert("RGB")), cv2.COLOR_RGB2BGR)
    gray = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)

    x_lines, y_lines = detect_table_grid(gray)

    if DEBUG:
        print(f"[LAYOUT DETECT] x_lines={len(x_lines)}, y_lines={len(y_lines)}")

    # Ngưỡng gợi ý:
    # - bảng thường có >=3 line dọc và khá nhiều line ngang
    if len(x_lines) >= 3 and len(y_lines) >= 8:
        return "table"
    return "text"


# =========================================================
# PIPELINE A: TABLE OCR (DOB + CMT/CCCD)
# =========================================================
def prep_for_ocr(gray_crop: np.ndarray, scale=2.2, pad_ratio=0.08) -> np.ndarray:
    h, w = gray_crop.shape[:2]
    pad = int(min(h, w) * pad_ratio)
    if h > 2 * pad and w > 2 * pad:
        gray_crop = gray_crop[pad:h - pad, pad:w - pad]

    gray_crop = cv2.resize(gray_crop, None, fx=scale, fy=scale, interpolation=cv2.INTER_CUBIC)
    gray_crop = cv2.GaussianBlur(gray_crop, (3, 3), 0)
    _, bw = cv2.threshold(gray_crop, 0, 255, cv2.THRESH_BINARY + cv2.THRESH_OTSU)
    return bw

def prep_for_ocr_variant(gray_crop: np.ndarray, scale: float, pad_ratio: float,
                         mode: str = "otsu", blur: str = "gauss") -> np.ndarray:
    h, w = gray_crop.shape[:2]
    pad = int(min(h, w) * pad_ratio)
    if h > 2 * pad and w > 2 * pad:
        gray_crop = gray_crop[pad:h - pad, pad:w - pad]

    g = cv2.resize(gray_crop, None, fx=scale, fy=scale, interpolation=cv2.INTER_CUBIC)

    if blur == "gauss":
        g = cv2.GaussianBlur(g, (3, 3), 0)
    elif blur == "median":
        g = cv2.medianBlur(g, 3)

    if mode == "adaptive":
        bw = cv2.adaptiveThreshold(g, 255, cv2.ADAPTIVE_THRESH_GAUSSIAN_C,
                                   cv2.THRESH_BINARY, 31, 8)
    else:
        _, bw = cv2.threshold(g, 0, 255, cv2.THRESH_BINARY + cv2.THRESH_OTSU)
    return bw

def thin_strokes(bw: np.ndarray, erode_iter: int = 1) -> np.ndarray:
    if erode_iter == 0:
        return bw
    inv = 255 - bw
    k = cv2.getStructuringElement(cv2.MORPH_RECT, (2, 2))
    inv = cv2.erode(inv, k, iterations=erode_iter)
    return 255 - inv

def tesseract_digits_with_conf(img: np.ndarray, psm: int = 7, numeric_mode: bool = True) -> tuple[str, float]:
    cfg = f"--oem 1 --psm {psm} -c tessedit_char_whitelist=0123456789"
    if numeric_mode:
        cfg += " -c classify_bln_numeric_mode=1"

    data = pytesseract.image_to_data(img, lang="eng", config=cfg, output_type=pytesseract.Output.DICT)

    items = []
    n = len(data.get("text", []))
    best_conf = 0.0

    for i in range(n):
        w = (data["text"][i] or "").strip()
        if not w or not re.fullmatch(r"\d+", w):
            continue
        conf = float(data["conf"][i]) if data["conf"][i] != "-1" else 0.0
        best_conf = max(best_conf, conf)
        top = int(data["top"][i]); left = int(data["left"][i])
        items.append((top, left, w, conf))

    if items:
        items.sort(key=lambda x: (x[0], x[1]))
        return " ".join([x[2] for x in items]), best_conf

    txt = pytesseract.image_to_string(img, lang="eng", config=cfg)
    txt = re.sub(r"\s+", " ", txt).strip()
    return txt, 0.0

def tesseract_date_with_conf(img: np.ndarray, psm: int = 7) -> tuple[str, float]:
    cfg = f"--oem 1 --psm {psm} -c tessedit_char_whitelist=0123456789/.-"
    data = pytesseract.image_to_data(img, lang="eng", config=cfg, output_type=pytesseract.Output.DICT)

    parts, confs = [], []
    n = len(data.get("text", []))
    for i in range(n):
        t = (data["text"][i] or "").strip()
        if not t or not re.search(r"[0-9]", t):
            continue
        parts.append(t)
        try:
            c = float(data["conf"][i]) if data["conf"][i] != "-1" else 0.0
        except Exception:
            c = 0.0
        confs.append(c)

    raw = " ".join(parts).strip()
    best_conf = max(confs) if confs else 0.0

    if not raw:
        raw = pytesseract.image_to_string(img, lang="eng", config=cfg)
        raw = re.sub(r"\s+", " ", raw).strip()

    return raw, best_conf

def _is_valid_date(d: int, m: int, y: int) -> bool:
    try:
        if not (DOB_MIN_YEAR <= y <= DOB_MAX_YEAR):
            return False
        _ = date(y, m, d)
        return True
    except Exception:
        return False

def _year_candidates_from_raw(yraw: str) -> List[Dict]:
    out = []

    # yy
    if len(yraw) == 2 and yraw.isdigit():
        yy = int(yraw)
        y1, y2 = 1900 + yy, 2000 + yy
        if DOB_MIN_YEAR <= y1 <= DOB_MAX_YEAR:
            out.append({"year": y1, "penalty": 8, "note": "yy->19xx"})
        if DOB_MIN_YEAR <= y2 <= DOB_MAX_YEAR:
            out.append({"year": y2, "penalty": 20, "note": "yy->20xx"})
        return out

    # yyyy
    if len(yraw) == 4 and yraw.isdigit():
        y = int(yraw)
        if DOB_MIN_YEAR <= y <= DOB_MAX_YEAR:
            out.append({"year": y, "penalty": 0, "note": "direct_yyyy"})

        for lead in ("1", "2"):
            yr2 = int(lead + yraw[1:])
            if DOB_MIN_YEAR <= yr2 <= DOB_MAX_YEAR:
                p = 12 if lead == "1" else 16
                out.append({"year": yr2, "penalty": p, "note": f"repair_lead->{lead}"})

        if y > DOB_MAX_YEAR:
            yy = int(yraw[-2:])
            y19, y20 = 1900 + yy, 2000 + yy
            if DOB_MIN_YEAR <= y19 <= DOB_MAX_YEAR:
                out.append({"year": y19, "penalty": 25, "note": "repair_from_last2_19xx"})
            if DOB_MIN_YEAR <= y20 <= DOB_MAX_YEAR:
                out.append({"year": y20, "penalty": 35, "note": "repair_from_last2_20xx"})
        return out

    return out

# ✅ Chỉ giữ 1 phiên bản parse_dob_candidates (không trùng)
def parse_dob_candidates(raw: str) -> List[Dict]:
    if not raw:
        return []

    s = raw.strip()
    s = s.replace("\\", "/").replace("-", "/").replace(".", "/")
    s = re.sub(r"\s+", "", s)
    s = re.sub(r"[^0-9/]", "", s)

    cands = []

    def _push(d_str: str, m_str: str, y_str: str, base_penalty: int, src_note: str):
        if not d_str or not m_str or not y_str:
            return
        if not d_str.isdigit() or not m_str.isdigit() or not y_str.isdigit():
            return

        d = int(d_str)
        m = int(m_str)
        y_opts = _year_candidates_from_raw(y_str)

        for yo in y_opts:
            y = yo["year"]
            if _is_valid_date(d, m, y):
                cands.append({
                    "dob": f"{d:02d}/{m:02d}/{y:04d}",
                    "penalty": base_penalty + yo["penalty"],
                    "note": f"{src_note}|{yo['note']}"
                })

    # có dấu / - .
    for mm in re.finditer(r"(\d{1,2})/(\d{1,2})/(\d{2,4})", s):
        dd, mm_, yy = mm.groups()
        _push(dd, mm_, yy, base_penalty=0, src_note="sep")

    # số dính
    digit_runs = re.findall(r"\d{6,10}", s)
    for dr in digit_runs:
        L = len(dr)
        if L == 8:
            _push(dr[:2], dr[2:4], dr[4:8], base_penalty=6, src_note="compact_8_ddmmyyyy")
        elif L == 7:
            _push(dr[:1], dr[1:3], dr[3:7], base_penalty=14, src_note="compact_7_1-2-4")
            _push(dr[:2], dr[2:3], dr[3:7], base_penalty=16, src_note="compact_7_2-1-4")
        elif L == 6:
            _push(dr[:2], dr[2:4], dr[4:6], base_penalty=10, src_note="compact_6_ddmmyy")

    # loại trùng, giữ penalty nhỏ nhất
    best_by_dob = {}
    for c in cands:
        k = c["dob"]
        if k not in best_by_dob or c["penalty"] < best_by_dob[k]["penalty"]:
            best_by_dob[k] = c

    return sorted(best_by_dob.values(), key=lambda x: x["penalty"])

def ocr_dob_smart(gray_cell: np.ndarray) -> Tuple[str, List[Dict]]:
    attempts = [
        {"tag": "dob_o1_s2.2_p0.08_psm7",  "scale": 2.2, "pad": 0.08, "psm": 7,  "mode": "otsu",     "blur": "gauss"},
        {"tag": "dob_o2_s2.8_p0.04_psm7",  "scale": 2.8, "pad": 0.04, "psm": 7,  "mode": "otsu",     "blur": "gauss"},
        {"tag": "dob_o3_s3.0_p0.02_psm7",  "scale": 3.0, "pad": 0.02, "psm": 7,  "mode": "otsu",     "blur": "gauss"},
        {"tag": "dob_o4_s3.0_p0.02_psm6",  "scale": 3.0, "pad": 0.02, "psm": 6,  "mode": "otsu",     "blur": "gauss"},
        {"tag": "dob_a2_s3.0_p0.02_psm7",  "scale": 3.0, "pad": 0.02, "psm": 7,  "mode": "adaptive", "blur": "none"},
        {"tag": "dob_a3_s3.0_p0.02_psm6",  "scale": 3.0, "pad": 0.02, "psm": 6,  "mode": "adaptive", "blur": "none"},
        {"tag": "dob_m1_s2.8_p0.02_psm7",  "scale": 2.8, "pad": 0.02, "psm": 7,  "mode": "otsu",     "blur": "median"},
        {"tag": "dob_p13_s3.0_p0.02_psm13","scale": 3.0, "pad": 0.02, "psm": 13, "mode": "otsu",     "blur": "none"},
    ]

    runs = []
    for a in attempts:
        bw = prep_for_ocr_variant(gray_cell, a["scale"], a["pad"], a["mode"], a["blur"])
        raw, conf = tesseract_date_with_conf(bw, psm=a["psm"])
        cand_objs = parse_dob_candidates(raw)

        picked = cand_objs[0]["dob"] if cand_objs else ""
        cand_penalty = cand_objs[0]["penalty"] if cand_objs else 999

        score = 0.0
        if picked:
            score += 120
            score += max(0, 40 - cand_penalty)

        score += conf
        score += (0.10 - a["pad"]) * 20

        if re.search(r"\d{1,2}[\/\.-]\d{1,2}[\/\.-]\d{4}", raw):
            score += 18
        if re.search(r"\d{1,2}[\/\.-]\d{1,2}[\/\.-]\d{2}$", raw):
            score -= 10

        m4 = re.search(r"(\d{4})", raw)
        if m4 and int(m4.group(1)) > DOB_MAX_YEAR:
            score -= 8

        runs.append({
            "attempt": a["tag"],
            "psm": a["psm"],
            "scale": a["scale"],
            "pad_ratio": a["pad"],
            "bw": bw,
            "raw": raw,
            "conf": conf,
            "cands": [c["dob"] for c in cand_objs],
            "picked": picked,
            "score": score
        })

    valid = [r for r in runs if r["picked"]]
    if not valid:
        return "", runs

    stats, counts = {}, {}
    for r in valid:
        p = r["picked"]
        stats[p] = stats.get(p, 0.0) + float(r["score"])
        counts[p] = counts.get(p, 0) + 1

    # vote theo số lần xuất hiện + tổng điểm
    best_dob = sorted(stats.keys(), key=lambda x: (counts[x], stats[x]), reverse=True)[0]
    return best_dob, runs

def parse_cmt_candidates(raw: str) -> List[str]:
    runs = re.findall(r"\d{6,20}", raw or "")
    cands = []
    for r in runs:
        if len(r) == 7:
            r = "0" + r
        if 8 <= len(r) <= 12:
            cands.append(r)
    return cands

def choose_best_cmt(cands: List[str]) -> str:
    if not cands:
        return ""
    for L in (12, 9, 8, 10, 11):
        arr = [x for x in cands if len(x) == L]
        if arr:
            return arr[-1]
    return cands[-1]

@dataclass
class Attempt:
    tag: str
    scale: float
    pad_ratio: float
    psm: int
    erode_iter: int = 0
    prior: float = 1.0

def ocr_cell_with_attempt(gray_cell: np.ndarray, att: Attempt) -> Dict:
    bw = prep_for_ocr(gray_cell, scale=att.scale, pad_ratio=att.pad_ratio)
    if att.erode_iter > 0:
        bw = thin_strokes(bw, erode_iter=att.erode_iter)

    raw, conf = tesseract_digits_with_conf(bw, psm=att.psm, numeric_mode=True)
    cands = parse_cmt_candidates(raw)
    picked = choose_best_cmt(cands)

    score = 0.0
    if len(picked) == 12:
        score += 200
    elif len(picked) == 9:
        score += 150
    elif len(picked) == 8:
        score += 140
    elif len(picked) in (10, 11):
        score += 120

    score += conf
    score += (0.10 - att.pad_ratio) * 50
    score += min(len(raw), 30) * 0.3
    score *= att.prior

    return {
        "bw": bw,
        "raw": raw,
        "conf": conf,
        "cands": cands,
        "picked": picked,
        "score": score,
        "attempt": att.tag,
        "psm": att.psm,
        "scale": att.scale,
        "pad_ratio": att.pad_ratio,
        "erode_iter": att.erode_iter,
    }

def compute_row_confidence(final_pick: str, runs: List[Dict]) -> Dict:
    valid = [r for r in runs if 8 <= len(r.get("picked", "")) <= 12]
    if not final_pick or not valid:
        return {
            "ocr_conf_pct": 0.0,
            "agreement_pct": 0.0,
            "best_tess_conf": 0.0,
            "score_margin_pct": 0.0,
            "band": "LOW"
        }

    same = [r for r in valid if r["picked"] == final_pick]
    best_tess_conf = max([float(r.get("conf", 0.0)) for r in same], default=0.0)

    agreement = len(same) / max(len(valid), 1)

    sorted_scores = sorted([float(r["score"]) for r in valid], reverse=True)
    best_score = sorted_scores[0] if sorted_scores else 0.0
    second_score = sorted_scores[1] if len(sorted_scores) > 1 else 0.0
    score_margin = max(0.0, (best_score - second_score) / best_score) if best_score > 0 else 0.0

    if len(final_pick) == 12:
        len_bonus = 1.0
    elif len(final_pick) == 9:
        len_bonus = 0.9
    elif len(final_pick) == 8:
        len_bonus = 0.85
    else:
        len_bonus = 0.75

    est = (
        0.45 * (best_tess_conf / 100.0) +
        0.30 * agreement +
        0.15 * score_margin +
        0.10 * len_bonus
    )

    est_pct = round(est * 100, 2)
    band = "HIGH" if est_pct >= 90 else ("MEDIUM" if est_pct >= 75 else "LOW")

    return {
        "ocr_conf_pct": est_pct,
        "agreement_pct": round(agreement * 100, 2),
        "best_tess_conf": round(best_tess_conf, 2),
        "score_margin_pct": round(score_margin * 100, 2),
        "band": band
    }

def ocr_cmt_smart(gray_cell: np.ndarray) -> Tuple[str, List[Dict], Dict]:
    attempts = [
        Attempt("a1_s2.2_p0.08_psm7", 2.2, 0.08, 7, 0, prior=1.05),
        Attempt("a1_s2.2_p0.08_psm6", 2.2, 0.08, 6, 0),
        Attempt("a2_s2.3_p0.08_psm7", 2.3, 0.08, 7, 0),
        Attempt("a3_s2.2_p0.04_psm7", 2.2, 0.04, 7, 0),
        Attempt("a3_s2.2_p0.04_psm6", 2.2, 0.04, 6, 0),
        Attempt("a4_s3.0_p0.02_psm6", 3.0, 0.02, 6, 0),
        Attempt("a5_s3.0_p0.02_psm11", 3.0, 0.02, 11, 0),
        Attempt("custom_s3.0_p0.02_psm7_er2", 3.0, 0.02, 7, 2),
    ]

    all_runs = [ocr_cell_with_attempt(gray_cell, att) for att in attempts]
    valid_runs = [r for r in all_runs if 8 <= len(r.get("picked", "")) <= 12]

    best = max(valid_runs, key=lambda x: x["score"]) if valid_runs else max(all_runs, key=lambda x: x["score"])
    final_pick = best.get("picked", "")

    conf_info = compute_row_confidence(final_pick, all_runs)
    conf_info["selected_attempt"] = best.get("attempt", "")
    conf_info["selected_score"] = round(float(best.get("score", 0.0)), 2)

    return final_pick, all_runs, conf_info

def run_table_pipeline(pdf_path: str) -> pd.DataFrame:
    pages = convert_from_path(pdf_path, dpi=DPI_TABLE, poppler_path=POPPLER_PATH)
    results = []

    for page_idx, pil in enumerate(pages, start=1):
        img = cv2.cvtColor(np.array(pil.convert("RGB")), cv2.COLOR_RGB2BGR)
        gray = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)

        x_lines, y_lines = detect_table_grid(gray)
        print(f"[TABLE] Page {page_idx}: x_lines={len(x_lines)}, y_lines={len(y_lines)}")

        if len(y_lines) < 3:
            continue

        # Chọn cột DOB + CMT
        if len(x_lines) >= (CMT_COL_INDEX + 2):
            dob_x1, dob_x2 = x_lines[CMT_COL_INDEX - 1], x_lines[CMT_COL_INDEX]
            cmt_x1, cmt_x2 = x_lines[CMT_COL_INDEX], x_lines[CMT_COL_INDEX + 1]
        elif len(x_lines) >= 3:
            dob_x1, dob_x2 = x_lines[0], x_lines[1]
            cmt_x1, cmt_x2 = x_lines[1], x_lines[2]
        elif len(x_lines) >= 2:
            dob_x1, dob_x2 = None, None
            cmt_x1, cmt_x2 = x_lines[0], x_lines[-1]
        else:
            continue

        pad = 3
        if dob_x1 is not None:
            dob_x1 += pad; dob_x2 -= pad
        cmt_x1 += pad; cmt_x2 -= pad

        for r in range(1, len(y_lines) - 1):
            stt = r + STT_OFFSET
            y1, y2 = y_lines[r], y_lines[r + 1]
            if y2 - y1 < 12:
                continue

            cmt_cell = gray[y1 + pad:y2 - pad, cmt_x1:cmt_x2]
            if cmt_cell.size == 0:
                continue

            cmt, cmt_runs, cmt_conf = ocr_cmt_smart(cmt_cell)

            dob = ""
            if dob_x1 is not None:
                dob_cell = gray[y1 + pad:y2 - pad, dob_x1:dob_x2]
                if dob_cell.size > 0:
                    dob, _ = ocr_dob_smart(dob_cell)

            if 8 <= len(cmt) <= 12:
                results.append({
                    "Ngày tháng năm sinh": dob,
                    "Số CMTND": str(cmt),
                    "Độ tin cậy OCR (%)": float(cmt_conf["ocr_conf_pct"])
                })

    df = pd.DataFrame(results)
    if df.empty:
        return pd.DataFrame(columns=["Ngày tháng năm sinh", "Số CMTND", "Độ tin cậy OCR (%)"])

    df["Số CMTND"] = df["Số CMTND"].astype(str)
    df["Ngày tháng năm sinh"] = df["Ngày tháng năm sinh"].astype(str)
    df = df.drop_duplicates(subset=["Số CMTND"], keep="first").reset_index(drop=True)
    return df


# =========================================================
# PIPELINE B: NON-TABLE OCR (QĐ style text blocks)
# =========================================================
def ocr_text_vie(img: Image.Image) -> str:
    return pytesseract.image_to_string(img, lang="vie", config="--oem 1 --psm 6")

def preprocess_digits_for_id(img: Image.Image, scale: int = 2) -> Image.Image:
    g = img.convert("L")
    w, h = g.size
    g = g.resize((w * scale, h * scale), Image.LANCZOS)
    g = ImageOps.autocontrast(g)
    g = ImageEnhance.Sharpness(g).enhance(2.2)
    g = ImageEnhance.Contrast(g).enhance(1.6)
    g = g.filter(ImageFilter.MaxFilter(3))
    return g

def ocr_ids_on_page(img: Image.Image) -> list[str]:
    img2 = preprocess_digits_for_id(img, scale=2)

    cfgs = [
        "--oem 1 --psm 6 -c tessedit_char_whitelist=0123456789;/:-",
        "--oem 1 --psm 11 -c tessedit_char_whitelist=0123456789;/:-",
        "--oem 1 --psm 6 -c tessedit_char_whitelist=0123456789 -c classify_bln_numeric_mode=1",
    ]

    texts = []
    for cfg in cfgs:
        try:
            texts.append(pytesseract.image_to_string(img2, lang="eng", config=cfg))
        except Exception:
            pass

    big = fix_ocr_digits(" ".join(texts))
    big = re.sub(r"\s+", " ", big).strip()

    if DEBUG:
        print("PAGE_DIGITS_OCR(sample):", big[:500])

    anchor_slash6  = re.compile(r"(?<!\d)/\s*6\b", flags=re.I)
    anchor_4slash6 = re.compile(r"(?<!\d)4\s*/\s*6\s*:?", flags=re.I)

    found = []

    def collect(anchor_regex):
        for m in anchor_regex.finditer(big):
            tail = big[m.end(): m.end() + 140]
            digits = re.sub(r"\D", "", tail)
            mm = re.search(r"0\d{11}", digits)
            if mm:
                found.append((m.start(), mm.group(0)))

    collect(anchor_slash6)
    collect(anchor_4slash6)

    found.sort(key=lambda x: x[0])

    out, seen = [], set()
    for _, id12 in found:
        if id12 not in seen:
            seen.add(id12)
            out.append(id12)

    if DEBUG:
        print("IDS_FROM_ANCHORS:", out)

    return out

def split_records(full_text: str) -> list[str]:
    text = full_text.replace("\r", "\n")

    # nhận "Họ và tên" có thể có số thứ tự phía trước và dấu phân cách linh hoạt
    pat = re.compile(
        r"(?:\b\d+\.\d+\.\s*)?H[oọ]\s*v[aà]\s*t[êe]n\s*(?:[:\-–—;]|\s)\s*",
        flags=re.I
    )

    starts = [m.start() for m in pat.finditer(text)]
    if not starts:
        return []

    blocks = []
    for i, s in enumerate(starts):
        e = starts[i + 1] if i + 1 < len(starts) else len(text)
        blocks.append(norm_spaces(text[s:e]))
    return blocks


def build_plain_and_map(original: str):
    plain_chars, mp = [], []
    for i, ch in enumerate(original):
        base = strip_accents(ch)
        for bc in base:
            plain_chars.append(bc)
            mp.append(i)
    return "".join(plain_chars), mp

def extract_name(block: str) -> str:
    b = norm_spaces(block)
    m = re.search(
        r"H[oọ]\s*v[aà]\s*t[êe]n\s*[:\-–—;]?\s*(.+?)"
        r"(?=Gi[ơo]i\s*t[ií]nh|Ng[aà]y|S[oô]\s*|Qu[oô]c|$)",
        b, flags=re.I
    )
    if not m:
        return ""
    name = m.group(1).strip()

    pn = strip_accents(name).lower()
    cut = re.search(r"\b(?:g|c|q)ioi\s*tinh\b", pn)
    if cut:
        plain, mp = build_plain_and_map(name)
        name = name[: mp[cut.start()]].strip()
    return name

def extract_gender(block: str) -> str:
    plain = strip_accents(block).lower()
    m = re.search(r"\b(?:g|c|q)ioi\s*tinh\s*[:\- ]*\s*(nam|nu)\b", plain)
    if not m:
        return ""
    return "Nam" if m.group(1) == "nam" else "Nữ"

def extract_dob_text(block: str) -> str:
    plain = strip_accents(block).lower()
    m = re.search(r"ngay\s*,?\s*thang\s*,?\s*nam\s*sinh\s*[:\- ]*(\d{1,2}/\d{1,2}/\d{4})", plain)
    return m.group(1) if m else ""

# def extract_id_from_record_text(block: str) -> str:
#     block2 = fix_ocr_digits(norm_spaces(block))
#     plain = strip_accents(block2).lower()

#     m = re.search(
#         r"(so\s*dinh\s*danh\s*ca\s*nhan\s*/\s*ho\s*chieu|"
#         r"so\s*dinh\s*danh|cccd|cmnd|ho\s*chieu)\s*[:\-]?\s*(.+?)"
#         r"(?=ngay\s*cap|noi\s*cap|gioi\s*tinh|quoc\s*tich|;|$)",
#         plain, flags=re.I
#     )
#     if not m:
#         return ""

#     tail = m.group(2)
#     digits = re.sub(r"\D", "", tail)
#     mm = re.search(r"0\d{11}", digits)
#     return mm.group(0) if mm else ""

def extract_id_from_record_text(block: str) -> str:
    block2 = fix_ocr_digits(norm_spaces(block))
    plain = strip_accents(block2).lower()

    # 1) Ưu tiên CCCD trước
    m_cccd = re.search(
        r"(can\s*cuoc\s*cong\s*dan|cccd)\s*(?:so)?\s*[:\-]?\s*([0-9\W]{6,60})",
        plain, flags=re.I
    )
    if m_cccd:
        tail = m_cccd.group(2)
        digits = re.sub(r"\D", "", tail)
        m12 = re.search(r"\d{12}", digits)
        if m12:
            return m12.group(0)

    # 2) Nếu không thấy CCCD thì mới lấy số định danh / hộ chiếu
    m_other = re.search(
        r"(so\s*dinh\s*danh\s*ca\s*nhan\s*/\s*ho\s*chieu|"
        r"so\s*dinh\s*danh|cmnd|ho\s*chieu)\s*[:\-]?\s*([0-9\W]{6,60})",
        plain, flags=re.I
    )
    if not m_other:
        return ""

    tail = m_other.group(2)
    digits = re.sub(r"\D", "", tail)

    # ưu tiên 12 số, nếu không có thì lấy 9 số
    m12 = re.search(r"\d{12}", digits)
    if m12:
        return m12.group(0)
    m9 = re.search(r"\d{9}", digits)
    return m9.group(0) if m9 else ""


def clean_address(addr: str) -> str:
    if not addr:
        return ""
    addr = addr.replace("|", " ").replace("_", " ")
    addr = re.sub(r"\s+", " ", addr).strip(" -:;,.")
    addr = fix_ocr_vietnamese_common(addr)

    addr = re.split(r"\b\d{1,2}/\d{1,2}/\d{4}\b", addr, maxsplit=1)[0]
    addr = re.split(r"\d{8,}", addr, maxsplit=1)[0]

    ap = strip_accents(addr).lower()
    cut = re.search(r"\b(so\s*dinh\s*danh|so\s*chung\s*minh\s*nhan\s*dan|cccd|cmnd|ho\s*chieu)\b", ap)
    if cut:
        plain, mp = build_plain_and_map(addr)
        addr = addr[: mp[cut.start()]].strip(" -:;,. ")

    return addr.strip(" -:;,. ")

def extract_current_address(block: str) -> str:
    original = norm_spaces(block)
    plain, mp = build_plain_and_map(original)
    p = plain.lower()

    label_pat = (
        r"(?:noi\s*o\s*hien\s*tai|cho\s*o\s*hien\s*tai|"
        r"dia\s*chi\s*hien\s*tai|noi\s*cu\s*tru|noi\s*thuong\s*tru)"
    )
    end_pat = (
        r"(?:so\s*dinh\s*danh|so\s*chung\s*minh\s*nhan\s*dan|"
        r"cccd|cmnd|ho\s*chieu|gioi\s*tinh|ngay\s*cap|noi\s*cap|quoc\s*tich|$)"
    )

    m = re.search(rf"{label_pat}\s*[:\-]?\s*(.+?)(?={end_pat})", p, flags=re.I)
    if not m:
        return ""

    s, e = m.start(1), m.end(1)
    addr = original[mp[s]: mp[e - 1] + 1]
    return clean_address(addr)

def run_text_pipeline(pdf_path: str) -> pd.DataFrame:
    pages = convert_from_path(pdf_path, dpi=DPI_TEXT, poppler_path=POPPLER_PATH)

    page_texts = []
    all_anchor_ids = []

    for idx, img in enumerate(pages):
        txt = ocr_text_vie(img)
        page_texts.append(txt)

        ids_page = ocr_ids_on_page(img)
        all_anchor_ids.extend(ids_page)

        if DEBUG:
            print(f"\n=== PAGE {idx+1} ===")
            print("ANCHOR_IDS_PAGE:", ids_page)

    full_text = "\n".join(page_texts)
    recs = split_records(full_text)

    if DEBUG:
        print("\n==============================")
        print("TOTAL_RECORDS:", len(recs))
        print("ALL_ANCHOR_IDS (doc order):", all_anchor_ids)

    ids_queue = list(all_anchor_ids)
    rows = []

    for i, rec in enumerate(recs):
        name = extract_name(rec)
        if not name:
            continue

        gender = extract_gender(rec)
        dob = extract_dob_text(rec)
        address = extract_current_address(rec)

        id_text = extract_id_from_record_text(rec)
        from_where = "record_text"

        if len(id_text) == 12:
            if id_text in ids_queue:
                ids_queue.remove(id_text)
        else:
            id_text = ids_queue.pop(0) if ids_queue else ""
            from_where = "anchor_queue"

        if DEBUG:
            print(f"rec#{i} name={name} dob={dob} -> id={id_text} ({from_where}) | queue_left={len(ids_queue)}")

        rows.append({
            "Họ và tên": name,
            "Giới tính": gender,
            "Ngày sinh": dob,
            "Nơi ở hiện tại": address,
            "Số định danh/Hộ chiếu": id_text
        })

    df = pd.DataFrame(rows)
    if df.empty:
        return pd.DataFrame(columns=["Họ và tên", "Giới tính", "Ngày sinh", "Nơi ở hiện tại", "Số định danh/Hộ chiếu"])

    df = df.drop_duplicates(subset=["Họ và tên"]).reset_index(drop=True)
    df["Số định danh/Hộ chiếu"] = df["Số định danh/Hộ chiếu"].astype(str)
    return df


# =========================================================
# MAIN ROUTER
# =========================================================
def main():
    pdf_files = get_all_pdfs_in_input_folder("input")
    if not pdf_files:
        print("❌ Không có file PDF nào trong thư mục input/")
        return

    for pdf_path in pdf_files:
        try:
            base_name = os.path.splitext(os.path.basename(pdf_path))[0]
            output_excel = os.path.join("output", f"{base_name}_OCR.xlsx")
            os.makedirs("output", exist_ok=True)

            layout = detect_layout_type(pdf_path)
            print(f"\n>>> FILE: {pdf_path}")
            print(f">>> AUTO DETECT LAYOUT = {layout.upper()}")

            if layout == "table":
                df = run_table_pipeline(pdf_path)
                df.to_excel(output_excel, index=False)
                if "Số CMTND" in df.columns:
                    save_excel_force_text(output_excel, id_col_name="Số CMTND")
            else:
                df = run_text_pipeline(pdf_path)
                df.to_excel(output_excel, index=False)
                if "Số định danh/Hộ chiếu" in df.columns:
                    save_excel_force_text(output_excel, id_col_name="Số định danh/Hộ chiếu")

            print(f"✅ Xuất {len(df)} dòng → {output_excel}")

        except Exception as e:
            print(f"❌ Lỗi file {pdf_path}: {e}")

if __name__ == "__main__":
    main()