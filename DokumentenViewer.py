# ========================= NEW.PY — FINAL 2025 (TEST BUILD) =========================
#  Chr. Carstensen Logistik — Dokumenten Viewer & Merger
#
#  Features:
#   - config.ini load/save on close
#   - filiale selection (10/40/50) influences default target
#   - ONLY Ziel row hidden by default, toggle with Alt+N
#   - AufNr validation depends on filiale:
#       * Filiale 10 -> exactly 8 digits
#       * Filiale 40/50 -> exactly 9 digits
#   - WinSped check (auto on complete AUF / F5 / button). Save blocked if no result
#   - OCR: select area with mouse -> detect digits -> insert into Auftragsnummer
#     - auto highlight selection result
#     - if multiple candidates -> press 1..9 to insert
# ======================================================================

import os
import re
import sys
import uuid
import shutil
import configparser
from datetime import datetime
from pathlib import Path

import pymssql
import pymupdf as fitz
from PyPDF2 import PdfReader, PdfWriter
from PIL import Image, ImageTk

import tkinter as tk
from tkinter import ttk, filedialog, messagebox

# Deskew optional
try:
    import cv2
    import numpy as np
except:
    cv2 = None

# OCR optional
try:
    import pytesseract
except:
    pytesseract = None


WINSPED_SQL = r"""
USE winsped;

SELECT TOP 1
    -- ================= Auftrag / Referenzen =================
    SLAuf.AufNr                     AS FBNR,          -- Auftragsnummer
    SLAuf.RefNr                     AS Reference2,    -- Refnr
    TT.TrackandTraceEmail           AS TT,             -- Track&Trace Mail
    
    SLAuf.LiefNr                    AS Liefnr,         -- Lieferschein-Nr (falls vorhanden)
    SLAuf.AufDatum                  As Aufdate,
    -- ================= Absender =============================
    Abskun.NAME1                    AS SenderName,
    Abskun.Strasse                  AS SenderAddressLine1,
    Abskun.PLZ                      AS SenderZip,
    Abskun.Ort                      AS SenderCity,
    Abskun.LKZ                      AS SenderCountry,
    Abskun.Tel                      AS SenderPhoneNo,
    Abskun.Email                    AS SenderMail,

    -- ================= Empfänger ============================
    Empkun.NAME1                    AS RecipientName,
    Empkun.Strasse                  AS RecipientAddressLine1,
    Empkun.PLZ                      AS RecipientZip,
    Empkun.Ort                      AS RecipientCity,
    Empkun.LKZ                      AS RecipientCountry,
    Empkun.Tel                      AS RecipientPhoneNo,
    Empkun.Email                    AS RecipientMail,

    -- ================= Mengen / Termine =====================
    SLAuf.ColliAnzSu                AS TotalColli,
    SLAuf.EntVonDat                 AS DeliveryDate,
    SLAuf.TatsGew                   AS TotalWeight,
    SLAuf.QMAnz                     AS TotalVolume,
    SLAuf.LMAnz                     AS TotalLDM

FROM XXASLAuf SLAuf
LEFT JOIN xxakun Abskun ON Abskun.Kundennr = SLAuf.UrAbsNr
LEFT JOIN xxakun Empkun ON Empkun.Kundennr = SLAuf.EmpNr
LEFT JOIN XXAAufExt TT  ON TT.AufIntNr     = SLAuf.AufIntNr

WHERE
    SLAuf.AufArt <> 'D'
    AND SLAuf.AufNr = %s;

"""


HELP_TEXT = """\
DokumentenViewer – Hilfe (Kurzanleitung)

Zweck
- Dokumente anzeigen (PDF/JPG/PNG), Seiten zusammenführen
- PDF + TXT/LIS erzeugen und ablegen

Bedienung
1) Load -> Dateien aus Quelle laden
2) AUFNR:
   - manuell eintragen ODER
   - OCR: Bereich mit der Maus markieren (linke Taste ziehen)
3) Dokumenttyp wählen
4) Filiale wählen (10/40/50)
5) WinSped prüfen (automatisch / F5 / кнопка)
6) Save -> PDF/TXT erzeugen, Quelle wird gelöscht, nächste Datei wird geladen

WinSped-Regel
- Wenn AUF im WinSped NICHT найден -> Save заблокирован

Filiale-Regel (AUFNR)
- Filiale 10: ровно 8 цифр
- Filiale 40/50: ровно 9 цифр

Hotkeys
- F1          Hilfe öffnen
- F5          WinSped prüfen
- Ctrl+S      Save
- Alt+N       Ziel ein-/ausblenden
- Pfeile      Navigation (←/→ Datei, ↑/↓ PDF-Seiten)
- Delete      Datei löschen

Hinweis
- Ziel только если нужно (Alt+N).
- Настройки сохраняются при выходе (config.ini).
"""


# ========================= OCR CONFIG (Windows) =========================
TESSERACT_EXE = r"C:\Program Files\Tesseract-OCR\tesseract.exe"
if pytesseract is not None and TESSERACT_EXE.strip():
    pytesseract.pytesseract.tesseract_cmd = TESSERACT_EXE


# ========================= APPDATA CONFIG ==============================

APPDATA_DIR = os.path.join(os.getenv("APPDATA"), "Carstensen", "DokumentenViewer")
os.makedirs(APPDATA_DIR, exist_ok=True)
CONFIG_FILE = os.path.join(APPDATA_DIR, "config.ini")


# ============================ CONSTANTS ================================

DOC_TYPES = ["Eingangsbelege", "Abliefernachweis", "Lademittel"]
DEFAULT_SOURCE = r"C:\Users\Public\Documents\ScanDoc\test"


# ========================= GLOBAL VARIABLES ============================

files = []
current_index = 0
current_file_path = None

current_is_pdf = False
current_page_index = 0
current_page_count = 1
current_rotation = 0

tk_img_preview = None
current_full_img = None
current_preview_scale = 1.0

sel_rect_id = None
sel_start = None

ocr_candidates = []
ocr_overlay_ids = []
ocr_selected_popup = None

ziel_visible = False
target_is_custom = False

last_winsPed_ok = False


# =========================== UTIL FUNCS ===============================

def safe_print(*args):
    txt = " ".join(str(a) for a in args)
    print(txt.encode("ascii", "ignore").decode("ascii"))


def resource_path(rel_path: str) -> str:
    base = getattr(sys, "_MEIPASS", str(Path(__file__).resolve().parent))
    return str(Path(base) / rel_path)


def extract_aufnr_from_filename(fname: str):
    m = re.match(r"^(\d{8,9})", Path(fname).stem)
    return m.group(1) if m else None


def rotate_before_save(img: Image.Image, angle: int):
    if angle == 0:
        return img
    return img.rotate(-angle, expand=True)


def default_target_for_filiale(filiale: str) -> str:
    filiale = (filiale or "").strip()
    if filiale == "10":
        return r"N:\DMS_IMPORT\DMS_OUT"
    return rf"N:\DMS_IMPORT\{filiale}_DMS_OUT"


def required_auf_len(filiale: str) -> int:
    return 8 if (filiale or "").strip() == "10" else 9


def validate_aufnr(num: str, filiale: str):
    if not num:
        return False, "Auftragsnummer fehlt."
    if not num.isdigit():
        return False, "Nur Ziffern erlaubt."
    need = required_auf_len(filiale)
    if len(num) != need:
        return False, f"Filiale {filiale}: Auftragsnummer muss {need}-stellig sein."
    return True, ""


# ========================= CONFIG LOAD/SAVE ============================

def load_config():
    cfg = configparser.ConfigParser()
    if os.path.exists(CONFIG_FILE):
        try:
            cfg.read(CONFIG_FILE, encoding="utf-8")
        except Exception as e:
            safe_print("Config read failed:", e)

    source = cfg.get("paths", "source", fallback=DEFAULT_SOURCE)
    target = cfg.get("paths", "target", fallback="")
    filiale = cfg.get("ui", "filiale", fallback="10")
    doctype = cfg.get("ui", "doctype", fallback=DOC_TYPES[0])
    return source, target, filiale, doctype


def save_config():
    cfg = configparser.ConfigParser()
    cfg["paths"] = {
        "source": entry_source.get().strip(),
        "target": entry_target.get().strip(),
    }
    cfg["ui"] = {
        "filiale": combo_filiale.get().strip(),
        "doctype": combo_doctype.get().strip(),
    }
    with open(CONFIG_FILE, "w", encoding="utf-8") as f:
        cfg.write(f)


# =========================== SQL / DB ===============================

def kennwort(paramm):
    value = ""
    filename = r"\\srv-dc2\DATEN$\Wiki\DMS_NEW\key.zip"
    try:
        with open(filename, "r", encoding="utf-8") as file:
            for line in file:
                line = line.strip()
                if line.startswith(paramm):
                    value = line[len(paramm) + 3:]
                    break
    except Exception as e:
        safe_print("Key read error:", e)
    return value

def fix_encoding(s):
    if not isinstance(s, str):
        return s
    try:
        return s.encode("latin1").decode("utf-8")
    except Exception:
        return s

def get_db_connection():
    server = kennwort("Data_Source")
    user = kennwort("user_id")
    password = kennwort("password")
    db_name = kennwort("DefaultDatabase")
    if not server or not user or not password or not db_name:
        raise RuntimeError("DB credentials not found (key.zip).")
    return pymssql.connect(server=server, user=user, password=password, database=db_name, charset="utf8")


# --- grouped WinSped panel fields (human friendly) ---
PANEL_GROUPS = [
    ("Auftrag / Referenzen", [
        ("Aufnr", "FBNR"),                  # SLAuf.aufnr (у тебя в SQL как 'FBNR')
        ("Liefnr", "Liefnr"),               # пока нет в SQL -> будет пусто, но UI готов
        ("Refnr", "Reference2"),            # SLAuf.refnr (у тебя 'Reference2')
        ("TrackandTraceEmail", "TT"),       # TT.TrackandTraceEmail (у тебя 'TT')
        ("Aufdate", "Aufdate"),             # пока нет в SQL -> будет пусто, но UI готов
    ]),

    ("Absender (Abskun)", [
        ("SenderName", "SenderName"),
        ("SenderAddress", "SenderAddressLine1"),
        ("SenderZip", "SenderZip"),
        ("SenderCity", "SenderCity"),
        ("SenderCountry", "SenderCountry"),
        ("SenderPhone", "SenderPhoneNo"),
        ("SenderMail", "SenderMail"),
    ]),

    ("Empfänger (Empkun)", [
        ("RecipientName", "RecipientName"),
        ("RecipientAddress", "RecipientAddressLine1"),
        ("RecipientZip", "RecipientZip"),
        ("RecipientCity", "RecipientCity"),
        ("RecipientCountry", "RecipientCountry"),
        ("RecipientPhone", "RecipientPhoneNo"),
        ("RecipientMail", "RecipientMail"),
    ]),

    ("Mengen / Termine", [
        ("TotalColli", "TotalColli"),
        ("DeliveryDate", "DeliveryDate"),
        ("TotalWeight", "TotalWeight"),
        ("TotalVolume", "TotalVolume"),
        ("TotalLDM", "TotalLDM"),
    ]),
]



def safe_str(v):
    return "" if v is None else str(v)


def update_winsPed_panel(row_dict, msg=""):
    # status
    lbl_db_status.config(text=msg)

    # clear all fields
    for key, var in panel_vars.items():
        var.set("")

    if not row_dict:
        return

    # fill grouped fields
    for group_title, fields in PANEL_GROUPS:
        for label, col in fields:
            if col in row_dict:
                panel_vars[label].set(safe_str(row_dict[col]))
            else:
                # column not present in SQL -> leave empty
                panel_vars[label].set(fix_encoding(row_dict[col]))


def set_save_enabled(enabled: bool):
    btn_save.config(state=("normal" if enabled else "disabled"))


def winsPed_query(aufnr: str):
    global last_winsPed_ok

    last_winsPed_ok = False
    update_winsPed_panel(None, msg="")
    set_save_enabled(False)

    fil = combo_filiale.get().strip()
    ok, msg = validate_aufnr(aufnr, fil)
    if not ok:
        update_winsPed_panel(None, msg=msg)
        return

    try:
        conn = get_db_connection()
        cur = conn.cursor(as_dict=True)
        cur.execute(WINSPED_SQL, (aufnr,))
        rows = cur.fetchall()
        cur.close()
        conn.close()
    except Exception as e:
        update_winsPed_panel(None, msg=f"DB error: {e}")
        return

    if not rows:
        update_winsPed_panel(None, msg="No data found for AUF. Save is blocked.")
        return

    last_winsPed_ok = True
    update_winsPed_panel(rows[0], msg=f"OK: {len(rows)} row(s)")
    set_save_enabled(True)


def maybe_autofetch_winsPed(event=None):
    auf = entry_aufnr.get().strip()
    fil = combo_filiale.get().strip()
    need = required_auf_len(fil)

    if not auf.isdigit():
        set_save_enabled(False)
        update_winsPed_panel(None, msg="")
        return

    if len(auf) == need:
        winsPed_query(auf)
    else:
        set_save_enabled(False)
        update_winsPed_panel(None, msg="AUF not complete.")


# =========================== OCR HELPERS ===============================

def ocr_candidates_from_crop(pil_crop: Image.Image, min_len: int) -> list[str]:
    if pytesseract is None:
        raise RuntimeError("OCR not available (pytesseract/Tesseract not installed).")

    txt = pytesseract.image_to_string(
        pil_crop,
        config="--psm 6 -c tessedit_char_whitelist=0123456789"
    )
    nums = re.findall(r"\d+", txt)
    nums = [n for n in nums if len(n) >= min_len]

    out, seen = [], set()
    for n in nums:
        if n not in seen:
            seen.add(n)
            out.append(n)

    out.sort(key=len, reverse=True)
    return out


def close_candidate_popup():
    global ocr_selected_popup
    if ocr_selected_popup is not None:
        try:
            ocr_selected_popup.destroy()
        except:
            pass
        ocr_selected_popup = None


def clear_ocr_overlay():
    for oid in ocr_overlay_ids:
        try:
            canvas_preview.delete(oid)
        except:
            pass
    ocr_overlay_ids.clear()


def show_ocr_overlay(left, top, right, bottom, text=None, ok=True):
    clear_ocr_overlay()
    outline = "green" if ok else "red"
    rid = canvas_preview.create_rectangle(left, top, right, bottom, outline=outline, width=3)
    ocr_overlay_ids.append(rid)
    if text:
        tid = canvas_preview.create_text(left, max(0, top - 18), anchor="nw",
                                         text=text, fill=outline, font=("Arial", 12, "bold"))
        ocr_overlay_ids.append(tid)


def choose_candidate_by_index(idx: int):
    close_candidate_popup()
    if idx < 0 or idx >= len(ocr_candidates):
        return

    num = ocr_candidates[idx]
    fil = combo_filiale.get().strip()
    ok, msg = validate_aufnr(num, fil)
    if not ok:
        messagebox.showerror("Fehler", f"OCR erkannt: {num}\n{msg}")
        return

    entry_aufnr.delete(0, tk.END)
    entry_aufnr.insert(0, num)
    maybe_autofetch_winsPed()


def show_candidate_popup():
    global ocr_selected_popup

    close_candidate_popup()
    ocr_selected_popup = tk.Toplevel(root)
    ocr_selected_popup.title("OCR Auswahl")
    ocr_selected_popup.geometry("420x220")
    ocr_selected_popup.transient(root)
    ocr_selected_popup.grab_set()

    tk.Label(ocr_selected_popup, text="Mehrere Nummern erkannt. Drücke 1..9:",
             font=("Arial", 11, "bold")).pack(pady=10)

    frame = tk.Frame(ocr_selected_popup)
    frame.pack(fill="both", expand=True, padx=10)

    for i, n in enumerate(ocr_candidates[:9]):
        tk.Label(frame, text=f"{i+1}) {n}", anchor="w", font=("Consolas", 11)).pack(fill="x")

    def on_key(event):
        ch = event.char.strip()
        if ch.isdigit():
            k = int(ch)
            if 1 <= k <= min(9, len(ocr_candidates)):
                choose_candidate_by_index(k - 1)
        if event.keysym in ("Escape",):
            close_candidate_popup()

    ocr_selected_popup.bind("<Key>", on_key)
    ttk.Button(ocr_selected_popup, text="Cancel (Esc)", command=close_candidate_popup).pack(pady=10)


# =========================== RENDER PREVIEW ===========================

def render_current_page():
    global tk_img_preview, current_full_img, current_preview_scale

    if not current_file_path:
        return

    try:
        if current_is_pdf:
            doc = fitz.open(current_file_path)
            page = doc.load_page(current_page_index)
            pix = page.get_pixmap(matrix=fitz.Matrix(1.5, 1.5))
            img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
            doc.close()
        else:
            img = Image.open(current_file_path).convert("RGB")

        if current_rotation != 0:
            img = img.rotate(-current_rotation, expand=True)

        current_full_img = img

        preview = img.copy()
        preview.thumbnail((1100, 1400))

        current_preview_scale = preview.width / img.width if img.width else 1.0

        tk_img_preview = ImageTk.PhotoImage(preview)

        canvas_preview.delete("all")
        canvas_preview.config(width=preview.width, height=preview.height)
        canvas_preview.create_image(0, 0, anchor="nw", image=tk_img_preview)

        clear_ocr_overlay()

    except Exception as e:
        canvas_preview.delete("all")
        canvas_preview.create_text(20, 20, anchor="nw", text=f"Fehler: {e}")
        current_full_img = None
        clear_ocr_overlay()


# =========================== LOAD FILES ===============================

def load_current_file():
    global current_file_path, current_index
    global current_is_pdf, current_page_count, current_page_index
    global current_rotation

    if not files:
        label_filename.config(text="(keine Dateien)")
        canvas_preview.delete("all")
        return

    current_file_path = files[current_index]
    current_rotation = 0

    label_filename.config(text=os.path.basename(current_file_path))

    ext = current_file_path.lower()
    current_is_pdf = ext.endswith(".pdf")

    if current_is_pdf:
        doc = fitz.open(current_file_path)
        current_page_count = doc.page_count
        doc.close()
        current_page_index = 0
    else:
        current_page_count = 1
        current_page_index = 0

    auf = extract_aufnr_from_filename(current_file_path)
    if auf:
        entry_aufnr.delete(0, tk.END)
        entry_aufnr.insert(0, auf)
        maybe_autofetch_winsPed()

    render_current_page()


def load_files():
    global files, current_index
    src = entry_source.get().strip()
    if not os.path.isdir(src):
        messagebox.showerror("Fehler", "Quelle existiert nicht.")
        return

    files = sorted(
        os.path.join(src, f)
        for f in os.listdir(src)
        if f.lower().endswith((".pdf", ".jpg", ".jpeg", ".png"))
    )

    if not files:
        messagebox.showinfo("Info", "Keine Dateien gefunden.")
        return

    current_index = 0
    load_current_file()


# =========================== NAVIGATION ===============================

def next_file(event=None):
    global current_index
    if current_index < len(files) - 1:
        current_index += 1
        load_current_file()


def prev_file(event=None):
    global current_index
    if current_index > 0:
        current_index -= 1
        load_current_file()


def next_page(event=None):
    global current_page_index
    if current_is_pdf and current_page_index < current_page_count - 1:
        current_page_index += 1
        render_current_page()


def prev_page(event=None):
    global current_page_index
    if current_is_pdf and current_page_index > 0:
        current_page_index -= 1
        render_current_page()


def rotate_page():
    global current_rotation
    current_rotation = (current_rotation + 90) % 360
    render_current_page()


# =========================== MERGE PDF/JPG =============================

def merge_pdfs(paths, out_pdf):
    writer = PdfWriter()
    for p in paths:
        if os.path.exists(p):
            try:
                reader = PdfReader(p)
                for page in reader.pages:
                    writer.add_page(page)
            except Exception as e:
                safe_print("merge error:", e)

    tmp = out_pdf + ".tmp"
    with open(tmp, "wb") as f:
        writer.write(f)
    shutil.move(tmp, out_pdf)


def append_image_to_pdf(image_path, final_pdf_path, dest_folder):
    img = Image.open(image_path).convert("RGB")
    if current_rotation != 0:
        img = img.rotate(-current_rotation, expand=True)
    temp_pdf = os.path.join(dest_folder, "__tmp_img.pdf")
    img.save(temp_pdf, "PDF", resolution=300.0)

    if os.path.exists(final_pdf_path):
        merge_pdfs([final_pdf_path, temp_pdf], final_pdf_path)
        os.remove(temp_pdf)
    else:
        shutil.move(temp_pdf, final_pdf_path)


def append_pdf_to_pdf(src_pdf, final_pdf):
    if current_rotation == 0:
        rotated = src_pdf
    else:
        rotated = final_pdf + ".rot.pdf"
        doc = fitz.open(src_pdf)
        out = fitz.open()

        for i in range(doc.page_count):
            page = doc.load_page(i)
            pix = page.get_pixmap()
            pil = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
            pil = rotate_before_save(pil, current_rotation)

            tmp_page = rotated + f"_p{i}.pdf"
            pil.save(tmp_page, "PDF", resolution=300)
            out.insert_pdf(fitz.open(tmp_page))
            os.remove(tmp_page)

        out.save(rotated)
        out.close()
        doc.close()

    if os.path.exists(final_pdf):
        merge_pdfs([final_pdf, rotated], final_pdf)
    else:
        shutil.move(rotated, final_pdf)

    if rotated.endswith(".rot.pdf") and os.path.exists(rotated):
        os.remove(rotated)


# ============================= LIS CREATION ===========================

def create_lis(aufnr, doctype, pdf_filename_only, folder):
    os.makedirs(folder, exist_ok=True)

    ref = uuid.uuid4().hex[:12] + "-" + datetime.now().strftime("%Y%m%d%H%M%S")
    today = datetime.now().strftime("%Y%m%d")

    lis_name = pdf_filename_only.replace(".pdf", ".txt")
    lis_path = os.path.join(folder, lis_name)

    ENDE_LINE = (
        f"ENDE|{ref}|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|"
        "0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|1|0|1|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|"
        "0|0|0|0|0|0|0|0|0|"
    )

    with open(lis_path, "w", encoding="latin-1", errors="replace") as f:
        f.write(f"START|{ref}|{today}|||GetMyInvoices|||Carstensen||||||||||||\n")
        f.write(f"DMSDOK|{ref}|1|WINSPED|#JJJJ#|{doctype}|{pdf_filename_only}||\n")
        f.write(f"DMSSW|{ref}|1|1|AUFNR|{aufnr}|\n")
        f.write(ENDE_LINE + "\n")

    return lis_path


# ============================= SAVE FILE ===============================

def save_file(event=None):
    global current_file_path

    if not last_winsPed_ok:
        messagebox.showerror("Fehler", "WinSped: AUF not found. Save is blocked.")
        return

    if not current_file_path:
        return

    fil = combo_filiale.get().strip()
    aufnr = entry_aufnr.get().strip()
    doctype = combo_doctype.get().strip()
    ziel = entry_target.get().strip()

    ok, msg = validate_aufnr(aufnr, fil)
    if not ok:
        messagebox.showerror("Fehler", msg)
        return

    dest_folder = os.path.join(ziel, f"{aufnr}_{doctype}")
    os.makedirs(dest_folder, exist_ok=True)

    final_pdf = os.path.join(dest_folder, f"{aufnr}_{doctype}.pdf")
    pdf_name_only = os.path.basename(final_pdf)

    try:
        if current_file_path.lower().endswith(".pdf"):
            append_pdf_to_pdf(current_file_path, final_pdf)
        else:
            append_image_to_pdf(current_file_path, final_pdf, dest_folder)
    except Exception as e:
        messagebox.showerror("Fehler", f"SAVE ERROR:\n{e}")
        return

    lis_name_only = pdf_name_only.replace(".pdf", ".txt")
    lis_path = os.path.join(dest_folder, lis_name_only)
    if not os.path.exists(lis_path):
        create_lis(aufnr, doctype, pdf_name_only, dest_folder)

    if os.path.exists(current_file_path):
        try:
            os.remove(current_file_path)
        except Exception as e:
            safe_print("Delete failed:", current_file_path, e)

    next_file()


# ============================= DELETE FILE ============================

def delete_file(event=None):
    global current_index
    if not current_file_path:
        return

    if not messagebox.askyesno("Delete", f"Datei löschen?\n{current_file_path}"):
        return

    try:
        os.remove(current_file_path)
    except Exception as e:
        messagebox.showerror("Fehler", str(e))
        return

    del files[current_index]

    if not files:
        label_filename.config(text="")
        canvas_preview.delete("all")
        entry_aufnr.delete(0, tk.END)
        set_save_enabled(False)
        update_winsPed_panel(None, msg="")
        return

    if current_index >= len(files):
        current_index = len(files) - 1

    load_current_file()


# =========================== MOUSE SELECTION / OCR ===========================

def on_sel_start(event):
    global sel_start, sel_rect_id
    sel_start = (event.x, event.y)
    close_candidate_popup()
    clear_ocr_overlay()
    if sel_rect_id is not None:
        canvas_preview.delete(sel_rect_id)
        sel_rect_id = None


def on_sel_move(event):
    global sel_rect_id
    if sel_start is None:
        return
    x0, y0 = sel_start
    x1, y1 = event.x, event.y
    if sel_rect_id is not None:
        canvas_preview.coords(sel_rect_id, x0, y0, x1, y1)
    else:
        sel_rect_id = canvas_preview.create_rectangle(x0, y0, x1, y1, outline="red", width=2)


def on_sel_end(event):
    global sel_start, ocr_candidates

    if sel_start is None or current_full_img is None:
        sel_start = None
        return

    x0, y0 = sel_start
    x1, y1 = event.x, event.y
    sel_start = None

    left = min(x0, x1)
    top = min(y0, y1)
    right = max(x0, x1)
    bottom = max(y0, y1)

    if (right - left) < 10 or (bottom - top) < 10:
        return

    scale = current_preview_scale if current_preview_scale > 0 else 1.0
    L = int(left / scale)
    T = int(top / scale)
    R = int(right / scale)
    B = int(bottom / scale)

    L = max(0, min(L, current_full_img.width - 1))
    T = max(0, min(T, current_full_img.height - 1))
    R = max(1, min(R, current_full_img.width))
    B = max(1, min(B, current_full_img.height))

    crop = current_full_img.crop((L, T, R, B))

    fil = combo_filiale.get().strip()
    need = required_auf_len(fil)
    min_len = max(6, need)

    try:
        ocr_candidates = ocr_candidates_from_crop(crop, min_len=min_len)
    except Exception as e:
        show_ocr_overlay(left, top, right, bottom, text="OCR error", ok=False)
        messagebox.showerror("Fehler", f"OCR error:\n{e}")
        return

    if not ocr_candidates:
        show_ocr_overlay(left, top, right, bottom, text="No number", ok=False)
        messagebox.showerror("Fehler", "OCR: keine Nummer erkannt.")
        return

    show_ocr_overlay(left, top, right, bottom, text=f"Found: {len(ocr_candidates)}", ok=True)

    if len(ocr_candidates) == 1:
        choose_candidate_by_index(0)
        return

    show_candidate_popup()


# ============================== UI ==================================

root = tk.Tk()
root.title("Chr. Carstensen LOGISTIK – Dokumenten Viewer")
root.geometry("1500x950")
root.configure(bg="#f2f2f2")

# Icon
try:
    root.iconbitmap(resource_path("carstensen.ico"))
except Exception as e:
    safe_print("iconbitmap failed:", e)

# Help
help_window = None

def show_help(event=None):
    global help_window
    if help_window is not None and help_window.winfo_exists():
        help_window.deiconify()
        help_window.lift()
        help_window.focus_force()
        return

    help_window = tk.Toplevel(root)
    help_window.title("Hilfe – DokumentenViewer")
    help_window.geometry("820x520")
    help_window.transient(root)

    frame = tk.Frame(help_window)
    frame.pack(fill="both", expand=True, padx=10, pady=10)

    txt = tk.Text(frame, wrap="word", font=("Consolas", 11))
    txt.pack(side="left", fill="both", expand=True)

    scroll = ttk.Scrollbar(frame, orient="vertical", command=txt.yview)
    scroll.pack(side="right", fill="y")
    txt.configure(yscrollcommand=scroll.set)

    txt.insert("1.0", HELP_TEXT)
    txt.config(state="disabled")

    def on_close():
        global help_window
        help_window.destroy()
        help_window = None

    ttk.Button(help_window, text="Schließen", command=on_close).pack(pady=(0, 10))
    help_window.protocol("WM_DELETE_WINDOW", on_close)


# ---------- PATH INPUTS ----------
frame_paths = tk.Frame(root, bg="#f2f2f2")
frame_paths.pack(pady=10)

# Quelle row
frame_src = tk.Frame(frame_paths, bg="#f2f2f2")
frame_src.grid(row=0, column=0, sticky="w")

tk.Label(frame_src, text="Quelle:", bg="#f2f2f2").grid(row=0, column=0)
entry_source = ttk.Entry(frame_src, width=80)
entry_source.grid(row=0, column=1, padx=5)

def choose_source():
    p = filedialog.askdirectory()
    if p:
        entry_source.delete(0, tk.END)
        entry_source.insert(0, p)

ttk.Button(frame_src, text="Durchsuchen", command=choose_source).grid(row=0, column=2, padx=5)

# Ziel row (toggle)
frame_dst = tk.Frame(frame_paths, bg="#f2f2f2")
frame_dst.grid(row=1, column=0, sticky="w")

tk.Label(frame_dst, text="Ziel:", bg="#f2f2f2").grid(row=0, column=0, padx=(0, 5))
entry_target = ttk.Entry(frame_dst, width=80)
entry_target.grid(row=0, column=1, padx=5)

def choose_target():
    global target_is_custom
    p = filedialog.askdirectory()
    if p:
        entry_target.delete(0, tk.END)
        entry_target.insert(0, p)
        target_is_custom = True

ttk.Button(frame_dst, text="Durchsuchen", command=choose_target).grid(row=0, column=2, padx=5)

def mark_target_custom(event=None):
    global target_is_custom
    target_is_custom = True

entry_target.bind("<KeyRelease>", mark_target_custom)


# ---------- AUFNR / DOCTYPE / FILIALE ----------
frame_info = tk.Frame(root, bg="#f2f2f2")
frame_info.pack()

tk.Label(frame_info, text="Auftragsnummer:", bg="#f2f2f2").grid(row=0, column=0)
entry_aufnr = ttk.Entry(frame_info, width=20)
entry_aufnr.grid(row=0, column=1, padx=5)
entry_aufnr.bind("<KeyRelease>", maybe_autofetch_winsPed)

tk.Label(frame_info, text="Dokumenttyp:", bg="#f2f2f2").grid(row=0, column=2, padx=(20, 5))
combo_doctype = ttk.Combobox(frame_info, values=DOC_TYPES, width=30, state="readonly")
combo_doctype.grid(row=0, column=3)

tk.Label(frame_info, text="Filiale:", bg="#f2f2f2").grid(row=0, column=4, padx=(20, 5))
combo_filiale = ttk.Combobox(frame_info, values=["10", "40", "50"], width=6, state="readonly")
combo_filiale.grid(row=0, column=5)

def on_filiale_change(event=None):
    global target_is_custom
    fil = combo_filiale.get().strip()

    # reset AUF if too long
    auf = entry_aufnr.get().strip()
    need = required_auf_len(fil)
    if len(auf) > need:
        entry_aufnr.delete(0, tk.END)
        entry_aufnr.insert(0, auf[:need])

    # update target if not customized
    if not target_is_custom:
        entry_target.delete(0, tk.END)
        entry_target.insert(0, default_target_for_filiale(fil))

    maybe_autofetch_winsPed()

combo_filiale.bind("<<ComboboxSelected>>", on_filiale_change)


# ---------- BUTTONS ----------
frame_btn = tk.Frame(root, bg="#f2f2f2")
frame_btn.pack(pady=10)

ttk.Button(frame_btn, text="Load", width=12, command=load_files).grid(row=0, column=0, padx=4)
ttk.Button(frame_btn, text="↑ Page", width=12, command=prev_page).grid(row=0, column=1, padx=4)
ttk.Button(frame_btn, text="↓ Page", width=12, command=next_page).grid(row=0, column=2, padx=4)
ttk.Button(frame_btn, text="Rotate ↻", width=12, command=rotate_page).grid(row=0, column=3, padx=4)
ttk.Button(frame_btn, text="<< Prev", width=12, command=prev_file).grid(row=0, column=4, padx=4)
ttk.Button(frame_btn, text="Next >>", width=12, command=next_file).grid(row=0, column=5, padx=4)

btn_save = ttk.Button(frame_btn, text="Save", width=12, command=save_file)
btn_save.grid(row=0, column=6, padx=4)

ttk.Button(frame_btn, text="Delete", width=12, command=delete_file).grid(row=0, column=7, padx=4)
ttk.Button(frame_btn, text="Hilfe (F1)", width=12, command=show_help).grid(row=0, column=8, padx=4)
ttk.Button(frame_btn, text="WinSped (F5)", width=12,
           command=lambda: winsPed_query(entry_aufnr.get().strip())
).grid(row=0, column=9, padx=4)

set_save_enabled(False)

# ---------- FILE NAME ----------
label_filename = tk.Label(root, text="", bg="#f2f2f2", font=("Arial", 14, "bold"))
label_filename.pack(pady=5)

# ---------- PREVIEW AREA + RIGHT PANEL ----------
preview_frame = tk.Frame(root, bg="#f2f2f2")
preview_frame.pack(fill="both", expand=True)

# left with scrolling container
left_frame = tk.Frame(preview_frame, bg="#f2f2f2")
left_frame.pack(side="left", fill="both", expand=True)

scroll_canvas = tk.Canvas(left_frame, bg="white")
scroll_canvas.pack(side="left", fill="both", expand=True)

scroll_y = ttk.Scrollbar(left_frame, orient="vertical", command=scroll_canvas.yview)
scroll_y.pack(side="right", fill="y")
scroll_canvas.configure(yscrollcommand=scroll_y.set)

inner = tk.Frame(scroll_canvas, bg="white")
scroll_canvas.create_window((0, 0), window=inner, anchor="nw")

def update_scroll(event=None):
    scroll_canvas.configure(scrollregion=scroll_canvas.bbox("all"))

inner.bind("<Configure>", update_scroll)

# actual preview canvas (for image + OCR selection)
canvas_preview = tk.Canvas(inner, bg="white", highlightthickness=0)
canvas_preview.pack(padx=10, pady=10)

canvas_preview.bind("<ButtonPress-1>", on_sel_start)
canvas_preview.bind("<B1-Motion>", on_sel_move)
canvas_preview.bind("<ButtonRelease-1>", on_sel_end)

# right: DB info panel
right_frame = tk.Frame(preview_frame, bg="#f7f7f7", width=420)
right_frame.pack(side="right", fill="y")

tk.Label(right_frame, text=" WinSped ", bg="#f7f7f7",
         font=("Arial", 12, "bold")).pack(pady=(10, 5))

lbl_db_status = tk.Label(right_frame, text="", bg="#f7f7f7", fg="#444")
lbl_db_status.pack(pady=(0, 10))

panel_vars = {}
panel_form = tk.Frame(right_frame, bg="#f7f7f7")
panel_form.pack(fill="y", padx=10)

row_i = 0

for group_title, fields in PANEL_GROUPS:
    # group header
    hdr = tk.Label(
        panel_form,
        text=group_title,
        bg="#f7f7f7",
        anchor="w",
        font=("Arial", 10, "bold")
    )
    hdr.grid(row=row_i, column=0, columnspan=2, sticky="w", pady=(10, 4))
    row_i += 1

    # group fields
    for label, col in fields:
        tk.Label(panel_form, text=label + ":", bg="#f7f7f7", anchor="w") \
            .grid(row=row_i, column=0, sticky="w", pady=2)

        v = tk.StringVar()
        panel_vars[label] = v

        e = ttk.Entry(panel_form, textvariable=v, width=42, state="readonly")
        e.grid(row=row_i, column=1, sticky="w", pady=2)

        row_i += 1


# ---------- Toggle ONLY Ziel row (Alt+N) ----------
def toggle_ziel_visibility(event=None):
    global ziel_visible
    ziel_visible = not ziel_visible
    if ziel_visible:
        frame_dst.grid()
    else:
        frame_dst.grid_remove()

root.bind_all("<Alt-n>", toggle_ziel_visibility)
root.bind_all("<Alt-N>", toggle_ziel_visibility)


# ---------- HOTKEYS ----------
root.bind("<Control-s>", save_file)
root.bind("<Delete>", delete_file)
root.bind("<Left>", prev_file)
root.bind("<Right>", next_file)
root.bind("<Up>", prev_page)
root.bind("<Down>", next_page)
root.bind("<F1>", show_help)
root.bind("<F5>", lambda e: winsPed_query(entry_aufnr.get().strip()))


# ---------- LOAD CONFIG ----------
source, target, filiale, doctype = load_config()
if filiale not in ["10", "40", "50"]:
    filiale = "10"
combo_filiale.set(filiale)

entry_source.delete(0, tk.END)
entry_source.insert(0, source)

if not target.strip():
    target = default_target_for_filiale(filiale)

entry_target.delete(0, tk.END)
entry_target.insert(0, target)

target_is_custom = (target.strip() != default_target_for_filiale(filiale))

combo_doctype.set(doctype if doctype in DOC_TYPES else DOC_TYPES[0])

# Hide Ziel by default
frame_dst.grid_remove()
ziel_visible = False


# ---------- SAVE CONFIG ON CLOSE ----------
def on_close():
    try:
        save_config()
    except Exception as e:
        safe_print("Config save failed:", e)
    root.destroy()

root.protocol("WM_DELETE_WINDOW", on_close)

root.mainloop()
