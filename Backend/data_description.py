import os
import sys
import json
from pathlib import Path
from datetime import datetime
import pandas as pd
import numpy as np
from PIL import Image
import io
import shutil
import openpyxl
import pytesseract
from pathlib import Path

# Optional embeddings
try:
    from sentence_transformers import SentenceTransformer
    SENTENCE_MODEL = SentenceTransformer("all-MiniLM-L6-v2")
except Exception:
    SENTENCE_MODEL = None

IMAGE_EXTS = {".png", ".jpg", ".jpeg", ".bmp", ".gif", ".tiff", ".webp"}
CHUNK_SIZE = 10000  # default chunk size for large CSVs


def mkdirp(p: Path):
    p.mkdir(parents=True, exist_ok=True)
    return p

def is_image_path_like(s: str):
    if not isinstance(s, str):
        return False
    s_lower = s.lower().strip()
    if s_lower.startswith("http://") or s_lower.startswith("https://"):
        return any(s_lower.endswith(ext) for ext in IMAGE_EXTS)
    return any(s_lower.endswith(ext) for ext in IMAGE_EXTS)


def ocr_image(image_path: Path):
    try:
        img = Image.open(image_path)
        text = pytesseract.image_to_string(img)
        return text.strip()
    except Exception:
        return ""


def maybe_compute_embedding(text: str):
    if not text or SENTENCE_MODEL is None:
        return None
    try:
        emb = SENTENCE_MODEL.encode([text], show_progress_bar=False)
        return emb[0].tolist()
    except Exception:
        return None


def analyze_dataframe(df: pd.DataFrame):
    total_rows = int(len(df))
    num_cols = list(df.select_dtypes(include=[np.number]).columns)
    #cat_cols = list(df.select_dtypes(include=["object", "category"]).columns)
    dt_cols = list(df.select_dtypes(include=["datetime64[ns]"]).columns)

    missing_per_col = df.isnull().sum().to_dict()
    missing_percent = {k: (v / total_rows * 100) if total_rows else 0 for k, v in missing_per_col.items()}

    numeric_summary = {}
    if num_cols:
        desc = df[num_cols].describe().to_dict()
        for col, stats in desc.items():
            numeric_summary[col] = {k: float(v) for k, v in stats.items()}

    dup_rows = int(df.duplicated().sum())

    return {
        "total_records": total_rows,
        "total_variables": len(df.columns),
        "missing_counts": {k: int(v) for k, v in missing_per_col.items() if v > 0},
        "missing_percent": {k: float(v) for k, v in missing_percent.items() if v > 0},
        "numerical_summary": numeric_summary,
        #"categorical_insights": categorical,
        "duplicate_rows": dup_rows,
        "numerical_columns": num_cols,
        #"categorical_columns": cat_cols,
        "datetime_columns": dt_cols,
        "completeness_percentage": 100 - sum(missing_percent.values()) / len(df.columns) if df.columns.any() else 100
    }


def save_dataframe(df: pd.DataFrame, out_dir: Path, name_prefix: str):
    ts = datetime.now().strftime("%Y%m%d%H%M%S")
    csv_name = f"{name_prefix}_{ts}.csv"
    path = out_dir / csv_name
    df.to_csv(path, index=False)
    return str(path)


def extract_images_from_openpyxl(workbook_path: Path, images_out: Path):
    wb = openpyxl.load_workbook(workbook_path, data_only=True)
    extracted = []
    for sheet in wb.worksheets:
        imgs = getattr(sheet, "_images", []) or []
        for idx, img in enumerate(imgs):
            try:
                img_obj = img._data() if hasattr(img, "_data") else None
            except Exception:
                img_obj = None

            saved = None
            try:
                if img_obj:
                    b = img_obj.getvalue() if hasattr(img_obj, "getvalue") else img_obj
                elif hasattr(img, "path") and Path(img.path).exists():
                    b = Path(img.path).read_bytes()
                else:
                    b = None

                if b:
                    bio = io.BytesIO(b)
                    pil = Image.open(bio)
                    ext = pil.format.lower()
                    saved_name = f"{sheet.title}_img_{idx}.{ext}"
                    outp = images_out / saved_name
                    pil.save(outp)
                    saved = str(outp)
                else:
                    saved = None
            except Exception:
                saved = None

            anchor = None
            try:
                anchor = getattr(img, "anchor", None)
                if anchor is not None:
                    frm = getattr(anchor, "_from", None)
                    if frm is not None:
                        anchor = f"r{frm.row+1}c{frm.col+1}"
            except Exception:
                anchor = None

            if saved:
                extracted.append({
                    "sheet": sheet.title,
                    "saved_path": saved,
                    "anchor": anchor
                })
    return extracted


def process_excel(path: Path, upload_dir: Path, output_dir: Path):
    base_dir = Path(__file__).parent
    tables_out = mkdirp(base_dir /"output"/ "tables")
    images_out = mkdirp(base_dir / "output"/"images")

    extracted_images = extract_images_from_openpyxl(path, images_out)
    wb = pd.ExcelFile(path, engine="openpyxl")
    sheets_info = []

    for sheet_name in wb.sheet_names:
        df = wb.parse(sheet_name, dtype=object)
        has_images = any(e.get("sheet") == sheet_name for e in extracted_images)
        sheet_json = {"sheet_name": sheet_name}

        if has_images or df.shape[0] > 0:
            saved_csv = save_dataframe(df, tables_out, sheet_name)
            sheet_json["saved_table_csv"] = saved_csv
            sheet_json["table_description"] = analyze_dataframe(df)

        sheet_images = [i for i in extracted_images if i.get("sheet") == sheet_name]
        images_descr = []
        for im in sheet_images:
            p = Path(im["saved_path"])
            ocr_text = ocr_image(p)
            emb = maybe_compute_embedding(ocr_text)
            images_descr.append({
                "saved_path": im["saved_path"],
                "anchor": im.get("anchor"),
                "ocr_text": ocr_text,
                "ocr_embedding": emb
            })
        if images_descr:
            sheet_json["images_info"] = images_descr

        if not has_images and df.shape[0] > 0 and df.select_dtypes(include=[np.number, 'object', 'category', 'datetime64[ns]']).shape[1] > 0:
            sheet_json["dataset_description"] = analyze_dataframe(df)

        sheets_info.append(sheet_json)

    return sheets_info


def process_csv(path: Path, upload_dir: Path, output_dir: Path):
    tables_out = mkdirp(output_dir / "tables")
    images_out = mkdirp(output_dir / "images")
    total_rows = sum(1 for _ in open(path, encoding="utf-8", errors="ignore")) - 1

    it = pd.read_csv(path, dtype=object, chunksize=CHUNK_SIZE)
    sheets_info = []
    df_all = pd.DataFrame()
    for chunk in it:
        df_all = pd.concat([df_all, chunk], ignore_index=True)

    image_cols = []
    for col in df_all.columns:
        if any(is_image_path_like(str(v)) for v in df_all[col].dropna().head(50)):
            image_cols.append(col)

    has_images = bool(image_cols)
    sheet_json = {"sheet_name": path.stem}

    if has_images:
        saved_csv = save_dataframe(df_all, tables_out, path.stem)
        sheet_json["saved_table_csv"] = saved_csv
        sheet_json["table_description"] = analyze_dataframe(df_all)

        images_descr = []
        for col in image_cols:
            for val in df_all[col].dropna().unique():
                p = Path(val)
                if p.exists():
                    ocr_text = ocr_image(p)
                    emb = maybe_compute_embedding(ocr_text)
                    images_descr.append({
                        "saved_path": str(p),
                        "ocr_text": ocr_text,
                        "ocr_embedding": emb
                    })
        if images_descr:
            sheet_json["images_info"] = images_descr

    else:
        sheet_json["dataset_description"] = analyze_dataframe(df_all)

    sheets_info.append(sheet_json)
    return sheets_info


def process_uploaded_file(file_path: str):
    p = Path(file_path)
    if not p.exists():
        raise FileNotFoundError(f"File not found: {p}")

    # Use predefined input/output folders under backend/
    base_dir = Path(__file__).parent
    uploads_dir = mkdirp(base_dir/ "input")
    outputs_dir = mkdirp(base_dir/ "output")

    suffix = p.suffix.lower()
    result = {
        "file_name": p.name
    }

    if suffix in {".xlsx", ".xlsm", ".xltx", ".xltm"}:
        result["sheets"] = process_excel(p, uploads_dir, outputs_dir)
    elif suffix == ".csv":
        result["sheets"] = process_csv(p, uploads_dir, outputs_dir)
    else:
        raise ValueError("Unsupported file type. Provide CSV or Excel file.")

    json_path = outputs_dir / "data_description.json"
    with open(json_path, "w", encoding="utf-8") as f:
        json.dump(result, f, indent=2, ensure_ascii=False)

    return {"uploads_dir": str(uploads_dir), "outputs_dir": str(outputs_dir), "description_json": str(json_path)}

def get_latest_file(directory, extensions=None):
    if extensions is None:
        extensions = ['.csv', '.xlsx', '.xlsm', '.xltx', '.xltm']
    dir_path = Path(directory)
    files = [f for f in dir_path.iterdir() if f.suffix.lower() in extensions and f.is_file()]

    if not files:
        return None

    latest_file = max(files, key=lambda f: f.stat().st_mtime)
    return str(latest_file)

# def main():
#     default_dir = r"backend\input\csv"
#     latest_file = get_latest_file(default_dir)

#     if latest_file is None or not os.path.exists(latest_file):
#         print(f"No input files found in {default_dir}")
#         return
    
#     print("Processing latest file:", latest_file)
#     try:
#         res = process_uploaded_file(latest_file)
#         print("Done. Results JSON:", res["description_json"])
#     except Exception as e:
#         print("Error during processing:", e)

# if __name__ == "__main__":
#     main()