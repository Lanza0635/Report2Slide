# -*- coding: utf-8 -*-
"""
Rapordan Fotoğraf — Excel'deki URL sütunlarından görselleri indirir.
Her sayfa için ayrı klasör (sayfa adı) oluşturulur.
"""

from __future__ import annotations

import io
import os
import shutil
import tempfile
import textwrap
import zipfile
from collections import defaultdict
from concurrent.futures import ThreadPoolExecutor, as_completed
from typing import Any, Iterator
from urllib.parse import urlparse

import pandas as pd
import requests
from PIL import Image, ImageDraw, ImageFont

# ---- yardımcılar ----

_BAD_FS = '<>:"/\\|?*\n\r\t'


def _safe_segment(name: str, max_len: int = 120) -> str:
    s = str(name).strip()
    for c in _BAD_FS:
        s = s.replace(c, "_")
    s = s.strip().strip(".")
    if not s:
        s = "sheet"
    return s[:max_len]


def _unique_sheet_folder_name(sheet_name: str, used: defaultdict[str, int]) -> str:
    """
    Aynı Excel'de sheet adları farklı olsa da güvenli ad çakışırsa alt klasörleri ayırır.
    İlk kullanım: 'Mağaza_Raporu', ikinci çakışma: 'Mağaza_Raporu_2'
    """
    base = _safe_segment(sheet_name)
    used[base] += 1
    n = used[base]
    if n == 1:
        return base
    return f"{base}_{n}"


def _is_url(val: Any) -> bool:
    if val is None or (isinstance(val, float) and pd.isna(val)):
        return False
    s = str(val).strip()
    if len(s) < 8:
        return False
    return s.startswith("http://") or s.startswith("https://")


def _guess_ext(url: str, content_type: str | None) -> str:
    ct = (content_type or "").lower()
    if "png" in ct:
        return ".png"
    if "jpeg" in ct or "jpg" in ct:
        return ".jpg"
    if "webp" in ct:
        return ".webp"
    if "gif" in ct:
        return ".gif"
    path = urlparse(url).path.lower()
    for ext in (".png", ".jpg", ".jpeg", ".webp", ".gif"):
        if path.endswith(ext):
            return ".jpg" if ext == ".jpeg" else ext
    return ".jpg"


_UA = "Mozilla/5.0 (compatible; RapordanFotograf/1.0)"


def _download(url: str, session: requests.Session) -> tuple[bytes | None, str | None]:
    try:
        r = session.get(url, timeout=25, stream=True)
        r.raise_for_status()
        data = r.content
        if not data:
            return None, "empty response"
        return data, r.headers.get("Content-Type")
    except Exception as e:
        return None, str(e)


def _download_parallel(url: str) -> tuple[bytes | None, str | None]:
    """Oturum paylaşımı yok — ThreadPool içinde güvenli."""
    try:
        r = requests.get(
            url,
            timeout=25,
            headers={"User-Agent": _UA},
        )
        r.raise_for_status()
        data = r.content
        if not data:
            return None, "empty response"
        return data, r.headers.get("Content-Type")
    except Exception as e:
        return None, str(e)


def _get_font(size: int = 17):
    candidates = []
    windir = os.environ.get("WINDIR", "")
    if windir:
        candidates.append(os.path.join(windir, "Fonts", "arial.ttf"))
        candidates.append(os.path.join(windir, "Fonts", "segoeui.ttf"))
    candidates.append("/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf")
    for path in candidates:
        if path and os.path.isfile(path):
            try:
                return ImageFont.truetype(path, size)
            except OSError:
                pass
    return ImageFont.load_default()


def _overlay_lines_for_cell(
    row: pd.Series,
    df: pd.DataFrame,
    filename_columns: list[str],
    col_safe: str,
    pos: int,
) -> list[str]:
    """Her seçili sütun için ayrı satır: 'Başlık: değer' (alt alta)."""
    lines: list[str] = []
    for fc in filename_columns:
        if fc not in df.columns:
            continue
        try:
            cell = row[fc]
        except Exception:
            continue
        if pd.isna(cell):
            continue
        val = str(cell).strip()
        if not val:
            continue
        show_val = _safe_segment(val, 100)
        lines.append(f"{fc}: {show_val}")
    if lines:
        return lines
    return [f"{col_safe} #{pos}"]


def _apply_overlay(img_bytes: bytes, lines: list[str], ext: str) -> bytes:
    """Sol alt köşede çok satırlı etiket (sütunlar alt alta)."""
    raw = [str(x).strip() for x in lines if x and str(x).strip()]
    if not raw:
        return img_bytes

    im = Image.open(io.BytesIO(img_bytes))
    im = im.convert("RGBA")
    w, h = im.size
    draw = ImageDraw.Draw(im)
    # Daha küçük, zarif tip (önceki: ~13–20 px)
    font = _get_font(max(10, min(14, h // 52)))
    spacing = 5
    pad_x, pad_y = 7, 6
    margin = 10
    max_text_w = max(80, w - 2 * margin - 2 * pad_x)
    wrap_w = max(22, min(44, max_text_w // 9))

    wrapped: list[str] = []
    for line in raw:
        if len(line) <= wrap_w:
            wrapped.append(line)
        else:
            wrapped.extend(textwrap.wrap(line, width=wrap_w) or [line])

    text = "\n".join(wrapped)
    if not text.strip():
        return img_bytes

    try:
        bbox = draw.multiline_textbbox((0, 0), text, font=font, spacing=spacing)
        tw, th = bbox[2] - bbox[0], bbox[3] - bbox[1]
    except Exception:
        bbox = draw.textbbox((0, 0), text, font=font)
        tw, th = bbox[2] - bbox[0], bbox[3] - bbox[1]

    x0 = margin
    y0 = h - th - 2 * pad_y - margin
    x1 = min(w - margin, x0 + tw + 2 * pad_x)
    y1 = h - margin
    overlay = Image.new("RGBA", (w, h), (0, 0, 0, 0))
    od = ImageDraw.Draw(overlay)
    od.rectangle([x0, y0, x1, y1], fill=(15, 23, 42, 142))
    im = Image.alpha_composite(im, overlay)
    draw = ImageDraw.Draw(im)
    text_fill = (248, 250, 252, 255)
    try:
        draw.multiline_text(
            (x0 + pad_x, y0 + pad_y),
            text,
            fill=text_fill,
            font=font,
            spacing=spacing,
        )
    except Exception:
        draw.text((x0 + pad_x, y0 + pad_y), text, fill=text_fill, font=font)

    buf = io.BytesIO()
    ext_l = ext.lower()
    if ext_l in (".jpg", ".jpeg"):
        im.convert("RGB").save(buf, format="JPEG", quality=92)
    elif ext_l == ".png":
        im.save(buf, format="PNG")
    elif ext_l == ".webp":
        im.convert("RGBA").save(buf, format="WEBP", quality=88)
    elif ext_l == ".gif":
        im.convert("RGB").save(buf, format="PNG")
    else:
        im.convert("RGB").save(buf, format="JPEG", quality=92)
    return buf.getvalue()


def _filename_parts_from_row(row: pd.Series, df: pd.DataFrame, filename_columns: list[str]) -> list[str]:
    parts: list[str] = []
    for fc in filename_columns:
        if fc not in df.columns:
            continue
        try:
            cell = row[fc]
        except Exception:
            continue
        if pd.isna(cell):
            continue
        parts.append(_safe_segment(str(cell), 50))
    return parts


def count_photo_urls_for_options(upload_folder: str, filename: str, options: dict) -> int:
    """Seçimlere göre Excel’deki toplam fotoğraf URL sayısı (işlemden önce hızlı sayım)."""
    filepath = os.path.join(upload_folder, filename)
    if not os.path.isfile(filepath):
        raise FileNotFoundError(f"File not found: {filepath}")

    columns_by_sheet: dict = options.get("columns_by_sheet") or {}
    if not isinstance(columns_by_sheet, dict) or not columns_by_sheet:
        raise ValueError("Select at least one photo column for one or more sheets.")

    columns_by_sheet = {
        k: [c for c in (v or []) if c] for k, v in columns_by_sheet.items() if v
    }
    columns_by_sheet = {k: v for k, v in columns_by_sheet.items() if v}
    if not columns_by_sheet:
        raise ValueError(
            "Select at least one column per included sheet, or disable the sheet."
        )

    xl = pd.ExcelFile(filepath)
    available = set(xl.sheet_names)
    return _count_photo_urls(filepath, columns_by_sheet, available)


def _count_photo_urls(filepath: str, columns_by_sheet: dict, available: set[str]) -> int:
    n = 0
    for sheet_name, col_list in columns_by_sheet.items():
        if sheet_name not in available:
            continue
        df = pd.read_excel(filepath, sheet_name=sheet_name)
        for col in col_list:
            if col not in df.columns:
                continue
            for _, row in df.iterrows():
                if _is_url(row[col]):
                    n += 1
    return n


def iter_photo_extraction(upload_folder: str, filename: str, options: dict) -> Iterator[dict[str, Any]]:
    """
    Olaylar: meta (total), progress (current, total, sheet), done (zip_path, saved, errors), error (message)
    """
    filepath = os.path.join(upload_folder, filename)
    if not os.path.isfile(filepath):
        yield {"type": "error", "message": f"File not found: {filepath}"}
        return

    columns_by_sheet: dict = options.get("columns_by_sheet") or {}
    if not isinstance(columns_by_sheet, dict) or not columns_by_sheet:
        yield {"type": "error", "message": "Select at least one photo column for one or more sheets."}
        return

    filename_columns = [c for c in (options.get("filename_columns") or []) if c]
    overlay_labels = bool(options.get("overlay_labels"))

    columns_by_sheet = {
        k: [c for c in (v or []) if c] for k, v in columns_by_sheet.items() if v
    }
    columns_by_sheet = {k: v for k, v in columns_by_sheet.items() if v}
    if not columns_by_sheet:
        yield {
            "type": "error",
            "message": "Select at least one column per included sheet, or disable the sheet.",
        }
        return

    xl = pd.ExcelFile(filepath)
    available = set(xl.sheet_names)
    total_urls = _count_photo_urls(filepath, columns_by_sheet, available)
    yield {"type": "meta", "total": total_urls}

    saved = 0
    errors: list[str] = []
    temp_root: str | None = None

    try:
        temp_root = tempfile.mkdtemp(prefix="rf_photos_")
        root_base = temp_root
        sheet_folder_used: defaultdict[str, int] = defaultdict(int)

        # 1) İş planı (sıra korunur: dosya adları ve seq_plain doğru kalır)
        jobs: list[dict[str, Any]] = []
        sheet_dfs: dict[str, pd.DataFrame] = {}

        for sheet_name, col_list in columns_by_sheet.items():
            if sheet_name not in available:
                errors.append(f"Sheet missing (skipped): {sheet_name}")
                continue

            folder_name = _unique_sheet_folder_name(sheet_name, sheet_folder_used)
            target_dir = os.path.join(root_base, folder_name)
            os.makedirs(target_dir, exist_ok=True)

            df = pd.read_excel(filepath, sheet_name=sheet_name)
            sheet_dfs[sheet_name] = df
            for col in col_list:
                if col not in df.columns:
                    errors.append(f"{sheet_name} / column missing: {col}")
                    continue

                col_safe = _safe_segment(col, 60)
                seq_plain = 0
                for pos, (_, row) in enumerate(df.iterrows(), start=1):
                    val = row[col]
                    if not _is_url(val):
                        continue
                    url = str(val).strip()
                    parts = _filename_parts_from_row(row, df, filename_columns)
                    prefix = "_".join(parts) if parts else ""
                    if not prefix:
                        seq_plain += 1
                    jobs.append(
                        {
                            "idx": len(jobs),
                            "url": url,
                            "sheet_name": sheet_name,
                            "folder_name": folder_name,
                            "target_dir": target_dir,
                            "col": col,
                            "col_safe": col_safe,
                            "pos": pos,
                            "row": row.copy(),
                            "fname_prefix": prefix,
                            "seq_plain": seq_plain if not prefix else 0,
                        }
                    )

        if not jobs:
            yield {
                "type": "error",
                "message": "No photos could be downloaded. Check URL cells and selected columns.",
            }
            return

        # 2) Paralel indirme (ağ gecikmesi üst üste binmez)
        n_jobs = len(jobs)
        workers = min(24, max(4, min(n_jobs, (os.cpu_count() or 4) * 5)))
        fetch_out: dict[int, tuple[bytes | None, str | None]] = {}
        with ThreadPoolExecutor(max_workers=workers) as pool:
            futs = {
                pool.submit(_download_parallel, j["url"]): j["idx"] for j in jobs
            }
            for fut in as_completed(futs):
                idx = futs[fut]
                try:
                    fetch_out[idx] = fut.result()
                except Exception as e:
                    fetch_out[idx] = (None, str(e))

        # 3) Sırayla kaydet, etiketle, ilerleme gönder
        for j in sorted(jobs, key=lambda x: x["idx"]):
            idx = j["idx"]
            sheet_name = j["sheet_name"]
            folder_name = j["folder_name"]
            target_dir = j["target_dir"]
            col = j["col"]
            col_safe = j["col_safe"]
            pos = j["pos"]
            url = j["url"]
            row = j["row"]
            df = sheet_dfs[sheet_name]

            data, ctype = fetch_out.get(idx, (None, "no result"))
            if data is None:
                errors.append(f"{sheet_name} [{col}] row {pos}: {ctype or 'download failed'}")
                continue

            ext = _guess_ext(url, ctype)
            fname_prefix = j.get("fname_prefix") or ""

            if overlay_labels:
                o_lines = _overlay_lines_for_cell(
                    row, df, filename_columns, col_safe, pos
                )
            else:
                o_lines = []

            if overlay_labels and o_lines:
                try:
                    data = _apply_overlay(data, o_lines, ext)
                except Exception as ex:
                    errors.append(f"{sheet_name} [{col}] row {pos} overlay: {ex}")

            if fname_prefix:
                base = f"{fname_prefix}_{col_safe}_{pos:04d}"
            else:
                sp = int(j.get("seq_plain") or 0)
                base = f"{col_safe}_{pos:04d}_{sp:03d}"

            fname = base + ext
            if len(fname) > 180:
                fname = base[: 160 - len(ext)] + ext

            fpath = os.path.join(target_dir, fname)
            trial = 0
            while os.path.exists(fpath):
                trial += 1
                fname = f"{base}_{trial}{ext}"
                fpath = os.path.join(target_dir, fname)

            with open(fpath, "wb") as f:
                f.write(data)
            saved += 1
            yield {
                "type": "progress",
                "current": saved,
                "total": total_urls,
                "sheet": folder_name,
            }

        if saved == 0:
            yield {
                "type": "error",
                "message": "No photos could be downloaded. Check URL cells and selected columns.",
            }
            return

        zip_fd, zip_path = tempfile.mkstemp(suffix=".zip", prefix="rapor_fotograflar_")
        os.close(zip_fd)
        # JPEG/PNG zaten sıkıştırılmış; DEFLATE hem yavaş hem kazanç az — STORE hızlı paketler
        with zipfile.ZipFile(zip_path, "w", zipfile.ZIP_STORED) as zf:
            for root, _, files in os.walk(temp_root):
                for fn in files:
                    full = os.path.join(root, fn)
                    arc = os.path.relpath(full, temp_root)
                    arc = arc.replace("\\", "/")
                    zf.write(full, arc)
        yield {
            "type": "done",
            "zip_path": zip_path,
            "saved": saved,
            "errors": errors,
        }

    finally:
        if temp_root and os.path.isdir(temp_root):
            shutil.rmtree(temp_root, ignore_errors=True)


def run_photo_extraction(upload_folder: str, filename: str, options: dict) -> Any:
    """
    options:
      columns_by_sheet: dict[str, list[str]]
      filename_columns: list[str] — dosya adına eklenecek sütun başlıkları
      overlay_labels: bool — seçilen bilgileri görsel üzerine yaz (satır satır)
    Her zaman ZIP dosyası üretir; istemci dosyayı indirir.
    """
    last_done: dict | None = None
    for ev in iter_photo_extraction(upload_folder, filename, options):
        if ev["type"] == "error":
            raise ValueError(ev.get("message", "Error"))
        if ev["type"] == "done":
            last_done = ev
    if not last_done:
        raise RuntimeError("Processing did not complete.")
    return {
        "mode": "zip",
        "zip_path": last_done["zip_path"],
        "saved": last_done["saved"],
        "errors": last_done.get("errors", []),
    }
