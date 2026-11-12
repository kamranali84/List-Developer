"""
LIST DEVELOPER  -  valve_list1.py
=================================================================
Reads P&IDs, specifications, AutoCAD (DXF), PDF, Excel and Word files and
builds Line List, Valve List, Equipment List, Instrument List, Tag List and
any CUSTOM list, with a customer-specific format editor and PDF output.

INSTALL (once):
    pip install streamlit pandas openpyxl reportlab pymupdf pdfplumber python-docx ezdxf xlrd
    OCR (needed for P&IDs whose text is outlined / scanned):
        pip install pytesseract pillow pypdfium2 numpy   +  install the Tesseract program
        Windows: https://github.com/UB-Mannheim/tesseract/wiki   (then set its path in Settings if not auto-found)

RUN (light theme recommended for the bluish-green look):
    streamlit run valve_list1.py --theme.base light
"""
import datetime
import io
import collections
import difflib
import json
import math
import os
import re
import tempfile
from xml.sax.saxutils import escape

import pandas as pd
import streamlit as st

# ---------- optional libraries -------------------------------------------------
try:
    import fitz  # PyMuPDF
except Exception:
    fitz = None
try:
    import pypdf
except Exception:
    pypdf = None
try:
    import pdfplumber
except Exception:
    pdfplumber = None
try:
    import openpyxl
except Exception:
    openpyxl = None
try:
    import pypdfium2 as pdfium
except Exception:
    pdfium = None
from concurrent.futures import ThreadPoolExecutor
from copy import copy
try:
    import docx  # python-docx
except Exception:
    docx = None
try:
    import ezdxf
except Exception:
    ezdxf = None
try:
    import numpy as np
except Exception:
    np = None
try:
    import pytesseract
    from PIL import Image, ImageDraw
except Exception:
    pytesseract = None
    Image = None
    ImageDraw = None

st.set_page_config(page_title="List Developer", page_icon="📋", layout="wide")
try:  # force a light theme with black text (applies from the next start of the app)
    _cfg_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), ".streamlit")
    _cfg = os.path.join(_cfg_dir, "config.toml")
    if not os.path.exists(_cfg):
        os.makedirs(_cfg_dir, exist_ok=True)
        with open(_cfg, "w") as _f:
            _f.write('[theme]\nbase = "light"\ntextColor = "#000000"\nbackgroundColor = "#ffffff"\n'
                     'secondaryBackgroundColor = "#ecfdf5"\nprimaryColor = "#0f766e"\n')
except Exception:
    pass

# =============================================================================
#  STYLE  (bluish-green gradient background)
# =============================================================================
st.markdown(
    """
<style>
.stApp{
  background: linear-gradient(135deg,#6ee7b7 0%,#5eead4 30%,#67e8f9 65%,#7dd3fc 100%);
  background-attachment: fixed;
}
/* ALL TEXT BLACK */
.stApp, .stApp *, [data-baseweb="popover"] *, [data-testid="stSidebar"] *{color:#000000 !important;}
header[data-testid="stHeader"]{background:transparent;}
section[data-testid="stSidebar"]{background:rgba(236,253,245,0.96);}
.hero{padding:6px 4px 14px 4px;}
.hero h1{font-size:2.6rem;margin:0;letter-spacing:1px;}
.hero p{margin:2px 0 0 2px;font-size:1.02rem;}
div[data-baseweb="tab-list"]{gap:6px;flex-wrap:wrap;}
button[data-baseweb="tab"]{background:rgba(255,255,255,.55);border-radius:10px 10px 0 0;padding:8px 14px;}
button[data-baseweb="tab"][aria-selected="true"]{background:rgba(255,255,255,.98);font-weight:800;}
div[data-baseweb="tab-panel"]{background:rgba(255,255,255,.96);border-radius:0 14px 14px 14px;padding:18px 20px;}
input, textarea, [data-baseweb="select"] > div{background:#ffffff !important;}
.stButton>button, .stDownloadButton>button{background:#ccfbf1 !important;border:1.5px solid #0f766e !important;font-weight:600;}
.stButton>button:hover, .stDownloadButton>button:hover{background:#99f6e4 !important;}
.kpi{background:linear-gradient(135deg,#ecfeff,#d1fae5);border-left:6px solid #0f766e;border-radius:12px;
     padding:12px 14px;text-align:center;box-shadow:0 2px 6px rgba(0,0,0,.12);}
.kpi .n{font-size:1.9rem;font-weight:800;line-height:1.1;}
.kpi .t{font-size:.82rem;text-transform:uppercase;letter-spacing:.6px;}
</style>
""",
    unsafe_allow_html=True,
)

# =============================================================================
#  ENGINEERING KNOWLEDGE  (ISA tags, valve codes, equipment codes)
# =============================================================================
VAR_MAP = {
    "A": "Analysis", "B": "Burner/Combustion", "C": "Conductivity", "D": "Density",
    "E": "Voltage", "F": "Flow", "G": "Gauging", "H": "Hand", "I": "Current",
    "J": "Power", "K": "Time", "L": "Level", "M": "Moisture", "P": "Pressure",
    "Q": "Quantity", "R": "Radiation", "S": "Speed/Frequency", "T": "Temperature",
    "U": "Multivariable", "V": "Vibration", "W": "Weight/Force", "X": "Unclassified",
    "Y": "Event/State", "Z": "Position",
}
FUNC_MAP = {
    "I": "Indicator", "T": "Transmitter", "C": "Controller", "S": "Switch", "E": "Element",
    "G": "Gauge/Glass", "R": "Recorder", "A": "Alarm", "H": "High", "L": "Low",
    "V": "Valve", "Y": "Relay/Converter", "Z": "Final element", "Q": "Totalizer",
    "D": "Differential", "X": "Accessory", "W": "Well", "P": "Test point", "K": "Control station",
    "F": "Ratio", "J": "Scan",
}
INSTR_LAST = set("ITCSEGRAHLVYZQXWPK")
INSTR_OVERRIDE = {
    "ZSO": ("Position", "Limit switch - Open"), "ZSC": ("Position", "Limit switch - Closed"),
    "SV": ("Solenoid", "Solenoid valve"),
}
INSTR_EXCLUDE = {
    "REV", "DWG", "ISO", "SH", "NO", "DOC", "PID", "API", "ASME", "ASTM", "IS", "EN", "DIN",
    "BS", "TAG", "SPEC", "ITEM", "PO", "MR", "TR", "UNIT", "AREA", "NOTE", "DN", "PN", "CL",
    "SCH", "LINE", "TYPE", "SIZE", "SHEET", "PAGE", "DATE", "TEXT", "VIEW", "ISSUE", "SEE",
    "PSI", "PSIG", "BAR", "KPA", "MPA", "LPM", "KGF", "GPM", "SS", "TIE", "TEST", "STAGE",
    "PLAN", "HOLD", "FLOW", "STEP", "TANK", "SKID", "PACK",
}
VALVE_MAP = {
    "HV": "Hand Valve (Manual)", "XV": "On/Off Valve (Actuated)", "SDV": "Shutdown Valve",
    "ESDV": "Emergency Shutdown Valve", "BDV": "Blowdown Valve", "FV": "Flow Control Valve",
    "FCV": "Flow Control Valve", "PCV": "Pressure Control Valve", "LCV": "Level Control Valve",
    "TCV": "Temperature Control Valve", "PV": "Pressure Control Valve", "LV": "Level Control Valve",
    "TV": "Temperature Control Valve", "PSV": "Pressure Safety Valve", "PRV": "Pressure Relief Valve",
    "TSV": "Thermal Safety Valve", "PSE": "Rupture Disc", "MOV": "Motor Operated Valve",
    "SOV": "Solenoid Valve", "NRV": "Non-Return (Check) Valve", "CV": "Control/Check Valve",
    "GV": "Gate Valve", "BV": "Ball Valve", "GLV": "Globe Valve", "CHV": "Check Valve",
    "ZV": "Position Valve", "RV": "Relief Valve", "SV": "Solenoid Valve", "DBB": "Double Block & Bleed",
    "VL": "Manual Valve", "SRV": "Safety Relief Valve", "SX": "Valve Assembly / Manifold",
}
EQUIP_MAP = {
    "P": "Pump", "V": "Vessel", "E": "Heat Exchanger", "C": "Column/Tower", "K": "Compressor",
    "T": "Tank", "R": "Reactor", "H": "Heater", "F": "Furnace/Filter", "S": "Separator",
    "D": "Drum", "M": "Mixer/Motor", "B": "Boiler/Blower", "TK": "Storage Tank", "HX": "Heat Exchanger",
    "PK": "Package Unit", "AG": "Agitator", "SD": "Skid / Module Unit", "TB": "Terminal / Junction Box",
}
MANUAL_FN = {"VL", "HV", "BV", "GV", "GLV", "CHV"}
SELF_FN = {"SRV", "PSV", "PRV", "RV", "TSV", "NRV", "CV"}

DEFAULT_PATTERNS = {
    "line": r'(?<![A-Z0-9])(?P<size>\d{1,2}(?:-\d/\d)?|\d/\d)\s*(?:"|IN)?\s*-\s*(?P<service>[A-Z]{1,4})\s*-\s*(?P<seq>\d{2,6})\s*-\s*(?P<spec>[A-Z0-9]{2,6})(?:\s*-\s*(?P<insul>[A-Z]{1,3}|\d{2,3}))?(?![A-Z0-9])',
    "valve": r'(?<![A-Z0-9])(?P<type>ESDV|SDV|BDV|XV|VL|SRV|SX|HV|FCV|PCV|LCV|TCV|FV|PV|LV|TV|PSV|PRV|TSV|PSE|MOV|SOV|NRV|CV|GV|BV|GLV|CHV|ZV|RV|SV|DBB)(?:\s*-\s*|\s+)?(?P<num>\d{2,5})(?P<sfx>[A-Z]?)(?![A-Z0-9])',
    "equipment": r'(?<![A-Z0-9])(?P<type>TK|HX|PK|AG|[PVECKTRHFSDMB])\s*-\s*(?P<num>\d{2,5})(?P<sfx>[A-Z]{0,2})(?![A-Z0-9])',
    "instrument": r'(?<![A-Z0-9])(?P<func>[A-Z]{2,4})(?:\s*-\s*|\s+)(?P<loop>\d{2,5})(?P<sfx>[A-Z]?)(?![A-Z0-9])',
}

DWG_FULL = re.compile(r"^[A-Z]{1,3}-\d{3,6}(?:-[A-Z0-9]{1,6}){2,}$")
DWG_FIND = re.compile(r"(?<![A-Z0-9-])[A-Z]{1,3}-\d{3,6}(?:-[A-Z0-9]{1,6}){1,}(?![A-Z0-9-])")
PID_DEFAULT_SVC = ("BACKWASH, FILTRATE, FEED, PERMEATE, AIR, DRAIN, VENT, WASTE, INTEGRITY TEST, CLEAN-IN-PLACE, CIP, "
                   "CHEMICAL, COMPRESSED AIR, WATER, STEAM, RETURN, SUCTION, DISCHARGE, INLET, OUTLET, FLUSH, "
                   "CONCENTRATE, SAMPLE, GAS, OIL, ACID, CAUSTIC, SLUDGE, NITROGEN, CONVEYING, AUXILIARY, TRANSPORT, ASH")

LIST_NAMES = ["Valve List", "Equipment List", "Line List", "Instrument List", "Tag List"]
HIDE_BY_DEFAULT = {"Prefix", "Number", "Suffix", "Context", "Valve Code", "Plant Unit", "System Code", "Sheet No"}
BASE_KEYS = {"Category", "Tag", "Source", "Loc", "Context"}


# =============================================================================
#  SESSION DEFAULTS
# =============================================================================
def _init_state():
    ss = st.session_state
    for k, v in DEFAULT_PATTERNS.items():
        ss.setdefault("pat_" + k, v)
    ss.setdefault("meta_customer", "")
    ss.setdefault("meta_project", "")
    ss.setdefault("meta_docno", "")
    ss.setdefault("meta_rev", "A")
    ss.setdefault("meta_by", "")
    ss.setdefault("sty_size", "A4")
    ss.setdefault("sty_orient", "Landscape")
    ss.setdefault("sty_color", "#0f766e")
    ss.setdefault("sty_font", 7)
    ss.setdefault("sty_sno", True)
    ss.setdefault("lists", {})
    ss.setdefault("raw", pd.DataFrame())
    ss.setdefault("log", [])
    ss.setdefault("run_id", 0)
    ss.setdefault("extras", {})          # list name -> [custom column names]
    ss.setdefault("tpl_loaded", {})      # list name -> {field: {order, include, header}}
    ss.setdefault("tpl_ver", 0)
    ss.setdefault("tpl_out", {})         # list name -> edited template DataFrame
    ss.setdefault("final", {})           # list name -> formatted DataFrame
    ss.setdefault("custom_rules", [])    # [{"Name":..., "Pattern":...}]
    ss.setdefault("pdf_bytes", {})
    ss.setdefault("logo", None)
    ss.setdefault("pid_on", True)
    ss.setdefault("pid_ocr", True)
    ss.setdefault("pid_ocr_text", True)
    ss.setdefault("pid_tess", "")
    ss.setdefault("pid_prefix", "")
    ss.setdefault("pid_sv", False)
    ss.setdefault("pid_right", 10)
    ss.setdefault("pid_svc", PID_DEFAULT_SVC)
    ss.setdefault("pid_vmap", "")
    ss.setdefault("prefix_hints", [])
    ss.setdefault("fmtbook", None)       # active customer Excel format workbook
    ss.setdefault("final_fmt", {})
    ss.setdefault("final_grp", {})


_init_state()


# =============================================================================
#  FILE READERS  -> list of (source, location, text)
# =============================================================================
def normalise(t: str) -> str:
    if not t:
        return ""
    for a, b in (("“", '"'), ("”", '"'), ("″", '"'), ("–", "-"), ("—", "-"),
                 ("‐", "-"), ("‑", "-"), ("\u00a0", " ")):
        t = t.replace(a, b)
    return t.upper()


def read_pdf(name, data, use_ocr, log):
    blocks, pages, ocr_pages = [], 0, 0
    if fitz:
        doc = fitz.open(stream=data, filetype="pdf")
        for i, page in enumerate(doc, 1):
            pages += 1
            txt = page.get_text("text")
            if len(txt.strip()) < 20 and use_ocr and pytesseract and Image:
                pix = page.get_pixmap(dpi=250)
                img = Image.open(io.BytesIO(pix.tobytes("png")))
                txt = pytesseract.image_to_string(img, config="--psm 11")
                ocr_pages += 1
            blocks.append((name, f"p{i}", txt))
    elif pypdf:
        rd = pypdf.PdfReader(io.BytesIO(data))
        for i, page in enumerate(rd.pages, 1):
            pages += 1
            blocks.append((name, f"p{i}", page.extract_text() or ""))
    else:
        log.append((name, "PDF", 0, "Install pymupdf (pip install pymupdf)"))
        return blocks
    empty = sum(1 for b in blocks if len(b[2].strip()) < 20)
    note = f"{pages} pages"
    if ocr_pages:
        note += f", OCR on {ocr_pages}"
    if empty and not use_ocr:
        note += f" | {empty} pages have no text (scanned? enable OCR)"
    log.append((name, "PDF", sum(len(b[2]) for b in blocks), note))
    return blocks


def read_dxf(name, data, log):
    if data[:4] == b"AC10" or data[:2] == b"AC":
        log.append((name, "DWG", 0, "Binary DWG not readable - save/convert to DXF (AutoCAD SAVEAS or ODA File Converter)"))
        return []
    if not ezdxf:
        log.append((name, "DXF", 0, "Install ezdxf (pip install ezdxf)"))
        return []
    with tempfile.NamedTemporaryFile(delete=False, suffix=".dxf") as tmp:
        tmp.write(data)
        path = tmp.name
    blocks = []
    try:
        doc = ezdxf.readfile(path)

        def collect(layout):
            out = []
            for e in layout:
                t = e.dxftype()
                try:
                    if t == "TEXT":
                        out.append(e.dxf.text)
                    elif t == "MTEXT":
                        out.append(e.plain_text())
                    elif t == "INSERT":
                        for a in e.attribs:
                            out.append(a.dxf.text)
                    elif t in ("ATTRIB",):
                        out.append(e.dxf.text)
                except Exception:
                    pass
            return "\n".join(out)

        for lay in doc.layouts:
            blocks.append((name, f"Layout:{lay.name}", collect(lay)))
        for blk in doc.blocks:
            if not blk.name.startswith("*"):
                blocks.append((name, f"Block:{blk.name}", collect(blk)))
        log.append((name, "DXF", sum(len(b[2]) for b in blocks), f"{len(blocks)} layouts/blocks"))
    except Exception as ex:
        log.append((name, "DXF", 0, f"Error: {ex}"))
    finally:
        os.unlink(path)
    return blocks


def read_excel(name, data, log):
    blocks = []
    try:
        sheets = pd.read_excel(io.BytesIO(data), sheet_name=None, header=None, dtype=str)
        for sh, df in sheets.items():
            rows = [" ; ".join(str(c) for c in r if pd.notna(c)) for r in df.values.tolist()]
            blocks.append((name, f"Sheet:{sh}", "\n".join(rows)))
        log.append((name, "Excel", sum(len(b[2]) for b in blocks), f"{len(blocks)} sheets"))
    except Exception as ex:
        log.append((name, "Excel", 0, f"Error: {ex}"))
    return blocks


def read_docx(name, data, log):
    if not docx:
        log.append((name, "Word", 0, "Install python-docx (pip install python-docx)"))
        return []
    try:
        d = docx.Document(io.BytesIO(data))
        parts = [p.text for p in d.paragraphs]
        for tb in d.tables:
            for row in tb.rows:
                parts.append(" ; ".join(c.text for c in row.cells))
        for sec in d.sections:
            parts += [p.text for p in sec.header.paragraphs] + [p.text for p in sec.footer.paragraphs]
        txt = "\n".join(parts)
        log.append((name, "Word", len(txt), f"{len(d.paragraphs)} paragraphs, {len(d.tables)} tables"))
        return [(name, "Document", txt)]
    except Exception as ex:
        log.append((name, "Word", 0, f"Error: {ex}"))
        return []


def read_files(files, use_ocr):
    blocks, log = [], []
    for f in files:
        name, data = f.name, f.getvalue()
        ext = name.lower().rsplit(".", 1)[-1] if "." in name else ""
        if ext == "pdf":
            blocks += read_pdf(name, data, use_ocr, log)
        elif ext in ("dxf", "dwg"):
            blocks += read_dxf(name, data, log)
        elif ext in ("xlsx", "xlsm", "xls"):
            blocks += read_excel(name, data, log)
        elif ext == "docx":
            blocks += read_docx(name, data, log)
        elif ext == "doc":
            log.append((name, "DOC", 0, "Old .doc not supported - save as .docx"))
        elif ext in ("csv", "txt"):
            txt = data.decode("utf-8", errors="ignore")
            blocks.append((name, "Text", txt))
            log.append((name, ext.upper(), len(txt), "read"))
        else:
            log.append((name, ext, 0, "Unsupported file type"))
    return blocks, log


# =============================================================================
#  SPATIAL P&ID READER  (uses text + pipe geometry positions via pdfplumber)
# =============================================================================
SIZE_ONLY = re.compile(r'^(?P<size>\d{1,2}(?:-\d/\d)?|\d/\d)\s*(?:"|IN|INCH)$')
LINE_LABEL = re.compile(r'^(?P<size>\d{1,2}(?:-\d/\d)?|\d/\d)\s*"\s*[- ]\s*(?P<spec>[A-Z0-9][A-Z0-9 \-/\.]*)$')
TAG_RX = re.compile(r'(?<![A-Z0-9])(?:(?P<pre>[A-Z]{2,3})\s?-\s?)?(?P<fn>[A-Z]{2,4})(?:\s?-\s?|\s)(?P<num>\d{3,5}[A-Z]?)(?![A-Z0-9]|-\d)')
FN_ITEM = re.compile(r'^(?:(?P<pre>[A-Z]{2,3})\s?-\s?)?(?P<fn>[A-Z]{2,4})$')
NUM_ITEM = re.compile(r'^\d{3,5}[A-Z]?$')

def build_items(page):
    items = []
    fp = page.filter(lambda o: o.get("object_type") != "char" or o.get("upright", True))
    words = fp.extract_words(x_tolerance=2.5, y_tolerance=2, keep_blank_chars=False, extra_attrs=["size"])
    words.sort(key=lambda w: (round(w["top"] / 2), w["x0"]))
    cur = None
    for w in words:
        if cur and abs(w["top"] - cur["top"]) <= 2 and 0 <= w["x0"] - cur["x1"] <= 0.55 * cur["size"]:
            cur["text"] += " " + w["text"]; cur["x1"] = w["x1"]; cur["bottom"] = max(cur["bottom"], w["bottom"])
        else:
            if cur: items.append(cur)
            cur = dict(text=w["text"], x0=w["x0"], x1=w["x1"], top=w["top"], bottom=w["bottom"], size=w["size"], rot=False)
    if cur: items.append(cur)
    rot = [c for c in page.chars if not c.get("upright", True)]
    rot.sort(key=lambda c: (c["matrix"][1] > 0, c["x0"]))
    groups = []
    for c in rot:
        d = c["matrix"][1] > 0
        for g in groups:
            if g["d"] == d and abs(g["x0"] - c["x0"]) < 1.5:
                g["c"].append(c); break
        else:
            groups.append(dict(d=d, x0=c["x0"], c=[c]))
    for g in groups:
        cs = sorted(g["c"], key=lambda c: c["top"], reverse=g["d"])
        run = [cs[0]]
        def flush(run):
            t = "".join(c["text"] for c in run).strip()
            if t:
                items.append(dict(text=t, x0=min(c["x0"] for c in run), x1=max(c["x1"] for c in run),
                                  top=min(c["top"] for c in run), bottom=max(c["bottom"] for c in run),
                                  size=run[0]["size"], rot=True))
        for c in cs[1:]:
            prev = run[-1]
            gap = (prev["top"] - c["bottom"]) if g["d"] else (c["top"] - prev["bottom"])
            if gap > 0.6 * c["size"]:
                flush(run); run = [c]
            else:
                run.append(c)
        flush(run)
    for it in items:
        it["cx"] = (it["x0"] + it["x1"]) / 2; it["cy"] = (it["top"] + it["bottom"]) / 2
    return items


VALVE_FN = {"XV","VL","PV","SRV","SX","SV","HV","FV","FCV","PCV","LCV","TCV","LV","TV","PSV","PRV","TSV","SDV","ESDV","BDV","MOV","SOV","NRV","CV","GV","BV","ZV","RV"}
ACTUATED = {"XV","SDV","ESDV","BDV","FV","FCV","PCV","LCV","TCV","LV","TV","MOV","ZV"}
FAIL = {"FC":"Fail Closed","FO":"Fail Open","FL":"Fail Last"}
NOTE_REF = re.compile(r"\(NOTE\s+(\d+)\)")


def _dist(a, b):
    return math.hypot(a[0] - b[0], a[1] - b[1])


def merge_blocks(items):
    """Join consecutive text lines (left-aligned or centred, tightly spaced) into multi-line blocks."""
    its = sorted([i for i in items if not i["rot"]], key=lambda i: (round(i["x0"]), i["top"]))
    used, blocks = set(), []
    for k, a in enumerate(its):
        if k in used:
            continue
        blk = [a]; used.add(k)
        grew = True
        while grew:
            grew = False
            last = blk[-1]
            for j, b in enumerate(its):
                if j in used:
                    continue
                gap = b["top"] - last["bottom"]
                if -1 <= gap <= 0.6 * last["size"] and (abs(b["x0"] - last["x0"]) <= 4 or abs(b["cx"] - last["cx"]) <= 6):
                    blk.append(b); used.add(j); grew = True; break
        blocks.append(dict(text=" ".join(x["text"] for x in blk),
                           cx=sum(x["cx"] for x in blk) / len(blk), cy=sum(x["cy"] for x in blk) / len(blk)))
    return blocks


def clean_service(t):
    t = re.sub(r"\bNOT BY UCC\b", " ", t)
    t = re.sub(r"\bD-\d{4,6}-\d+-L\d-\d+\b", " ", t)
    t = re.sub(r"\s+", " ", t).strip()
    t = re.split(r"\s\d[\d.\-]*\s*(?:GPM|SCFM|PSIG|PSI|CFM|LPM|M3/H|NM3/H|KG/H|BAR|MM|KPA)\b", t)[0]
    return t.strip(" -.,;")


def build_pipes(page, s, right_pct):
    W = page.width
    segs = []
    def add(a, b):
        dx, dy = abs(a[0] - b[0]), abs(a[1] - b[1])
        if dy < 1 and dx >= 50 * s and max(a[0], b[0]) < W * (1 - right_pct / 100):
            segs.append(["H", (a[1] + b[1]) / 2, min(a[0], b[0]), max(a[0], b[0])])
        elif dx < 1 and dy >= 50 * s and a[0] < W * (1 - right_pct / 100):
            segs.append(["V", (a[0] + b[0]) / 2, min(a[1], b[1]), max(a[1], b[1])])
    for l in page.lines:
        if not l.get("dash"):
            add((l["x0"], l["top"]), (l["x1"], l["bottom"]))
    for c in page.curves:
        if c.get("linewidth", 0) >= 0.5 and not c.get("dash") and c.get("pts"):
            for a, b in zip(c["pts"], c["pts"][1:]):
                add(a, b)
    pipes = []
    for o in "HV":
        ss = sorted([x for x in segs if x[0] == o], key=lambda x: (x[1], x[2]))
        cur = None
        for x in ss:
            if cur and abs(x[1] - cur["c"]) <= 1.5 and x[2] - cur["hi"] <= 60 * s:
                cur["hi"] = max(cur["hi"], x[3])
            else:
                if cur: pipes.append(cur)
                cur = dict(o=o, c=x[1], lo=x[2], hi=x[3])
        if cur: pipes.append(cur)
    return pipes


def pipe_for(pipes, p, s):
    best = None
    for pp in pipes:
        perp, along = (abs(p[1] - pp["c"]), p[0]) if pp["o"] == "H" else (abs(p[0] - pp["c"]), p[1])
        if perp <= 25 * s and pp["lo"] - 60 * s <= along <= pp["hi"] + 60 * s:
            if best is None or perp < best[0]:
                best = (perp, pp)
    return best[1] if best else None


def on_pipe(pp, c, perp_max, ext=20):
    """(along-axis coordinate) if the label centre c lies beside pipe pp, else None."""
    if pp["o"] == "H":
        perp, along = abs(c[1] - pp["c"]), c[0]
    else:
        perp, along = abs(c[0] - pp["c"]), c[1]
    if perp <= perp_max and pp["lo"] - ext <= along <= pp["hi"] + ext:
        return along
    return None


def parse_page(page, svc_words, right_pct=10, bottom_pct=4):
    s = page.width / 3024.0
    all_items = build_items(page)
    W, H = page.width, page.height
    prefix_hints = sorted({i["text"] for i in all_items if re.fullmatch(r"[A-Z]{2,4}\d{2,4}-", i["text"])})
    dc = [(i["size"], i["cy"] + i["cx"], m.group(0)) for i in all_items for m in DWG_FIND.finditer(i["text"])]
    dwg = max(dc)[2] if dc else ""
    notes = {}
    nums_ = [i for i in all_items if not i["rot"] and re.fullmatch(r"\d{1,2}\.", i["text"]) and i["cx"] > W * 0.85]
    if nums_:
        top0 = min(n["top"] for n in nums_)
        stop = [i["top"] for i in all_items if i["text"].startswith("CONNECTION SCHEDULE") and i["top"] > top0]
        limit = min(stop) if stop else top0 + 250 * s
        xmin = min(n["x1"] for n in nums_) - 2
        nums_ = sorted([n for n in nums_ if n["top"] < limit], key=lambda n: n["top"])
        lines_ = sorted([i for i in all_items if not i["rot"] and i["cx"] > W * 0.85 and i["x0"] > xmin
                         and top0 - 3 <= i["top"] < limit and not re.fullmatch(r"\d{1,2}\.", i["text"])],
                        key=lambda i: i["top"])
        prev_bottom = None
        for ln in lines_:
            if prev_bottom is not None and ln["top"] - prev_bottom > 1.2 * ln["size"]:
                break
            prev_bottom = ln["bottom"]
            cur = [n for n in nums_ if n["top"] <= ln["top"] + 3]
            if cur:
                k = cur[-1]["text"].rstrip(".")
                notes[k] = (notes.get(k, "") + " " + ln["text"]).strip()
    items = [i for i in all_items if i["cx"] < W * (1 - right_pct / 100) and i["cy"] < H * (1 - bottom_pct / 100)]
    tags = []
    def add(pre, fn, num, x, y, ref):
        if not num.startswith("00"):
            tags.append(dict(pre=pre or "", fn=fn, num=num, x=x, y=y, ref=ref))
    used = set()
    for n in [i for i in items if not i["rot"] and NUM_ITEM.match(i["text"])]:
        best = None
        for k, f in enumerate(items):
            if f["rot"] or not FN_ITEM.match(f["text"]):
                continue
            gap = n["top"] - f["bottom"]
            if -3 <= gap <= 6 and abs(f["cx"] - n["cx"]) <= 16 * s + 8 and (best is None or gap < best[0]):
                best = (gap, k, f)
        if not best:
            continue
        _, k, f = best
        m = FN_ITEM.match(f["text"]); pre, fn = m.group("pre"), m.group("fn")
        ref = False
        above = [a for a in items if a is not f and not a["rot"] and -3 <= f["top"] - a["bottom"] <= 6
                 and abs(a["cx"] - f["cx"]) <= 16 * s + 8]
        if above:
            a = min(above, key=lambda a: f["top"] - a["bottom"])
            if a["text"] == "TO":
                ref = True
            elif not pre and re.fullmatch(r"[A-Z]{2,3}", a["text"]):
                pre = a["text"]
        used.add(k)
        add(pre, fn, n["text"], n["cx"], f["cy"], ref)
    for k, it in enumerate(items):
        if k in used:
            continue
        for m in TAG_RX.finditer(it["text"]):
            add(m.group("pre"), m.group("fn"), m.group("num"), it["cx"], it["cy"],
                it["text"][:m.start()].rstrip().endswith("TO"))
    refs = [t for t in tags if t["ref"]]
    keys = {(t["pre"], t["fn"], t["num"]) for t in tags}
    tags = [t for t in tags if not any(o[0] == t["pre"] and o[2] == t["num"] and o[1] != t["fn"]
                                        and o[1].startswith(t["fn"]) for o in keys)]
    uniq = {}
    for t in tags:
        u = uniq.setdefault((t["pre"], t["fn"], t["num"]), dict(t, n=0, real=[]))
        if not t["ref"]:
            u["n"] += 1; u["real"].append((t["x"], t["y"]))
    sizes, fcs, svcs, desc_items = [], [], [], []
    for it in items:
        tx = it["text"]
        m = SIZE_ONLY.match(tx) or LINE_LABEL.match(tx)
        if m:
            sizes.append((it["cx"], it["cy"], m.group("size").replace(" ", "") + '"', tx))
        if tx in FAIL:
            fcs.append((it["cx"], it["cy"], FAIL[tx]))
    blocks_ = merge_blocks(items)
    for b in blocks_:
        if NOTE_REF.search(b["text"]) or "MANIFOLD" in b["text"]:
            desc_items.append((b["cx"], b["cy"], TAG_RX.sub("", b["text"]).strip()))
    cands = [(i["cx"], i["cy"], i["text"]) for i in items if i["rot"]] + \
            [(b["cx"], b["cy"], b["text"]) for b in blocks_]
    kw = [re.compile(r"\b" + re.escape(w.strip().upper()) + r"\b") for w in svc_words if w.strip()]
    for x, y, tx in cands:
        c = clean_service(tx)
        if len(c) > 2 and any(k.search(c.upper()) for k in kw) and not TAG_RX.search(c):
            if True:
                svcs.append((x, y, c))
    pipes = build_pipes(page, s, right_pct)
    vals = [u for u in uniq.values() if u["fn"] in VALVE_FN and u["n"] > 0]
    for u in vals:
        u["pos"] = u["real"][0]; u["fail"] = ""
    elig = [u for u in vals if u["fn"] in ACTUATED]
    pairs = sorted((_dist(u["pos"], (f[0], f[1])), i, j) for i, u in enumerate(elig) for j, f in enumerate(fcs))
    ua, fa = set(), set()
    for d, i, j in pairs:
        if d <= 300 * s and i not in ua and j not in fa:
            ua.add(i); fa.add(j)
            elig[i]["pos"] = (fcs[j][0], fcs[j][1]); elig[i]["fail"] = fcs[j][2]
    rows = []
    for u in vals:
        p = u["pos"]
        pp = pipe_for(pipes, p, s)
        size = size_src = svc = ""
        if pp:
            c = [(abs(a - (p[0] if pp["o"] == "H" else p[1])), z) for z in sizes
                 for a in [on_pipe(pp, (z[0], z[1]), 30 * s)] if a is not None]
            if c:
                z = min(c, key=lambda t: t[0])[1]
                size, size_src = z[2], z[3] + " [same pipe]"
            c = [(abs(a - (p[0] if pp["o"] == "H" else p[1])), z) for z in svcs
                 for a in [on_pipe(pp, (z[0], z[1]), 45 * s, 260 * s)] if a is not None]
            if c:
                svc = min(c, key=lambda t: t[0])[1][2]
        if not svc and "TUBE" in size_src.upper():
            cand = [t for t in refs if abs(t["y"] - p[1]) <= 30 * s and abs(t["x"] - p[0]) <= 450 * s]
            if cand:
                t = min(cand, key=lambda t: abs(t["x"] - p[0]))
                svc = "Tubing to " + "-".join(x for x in (t["pre"], t["fn"], t["num"]) if x)
        if not size and sizes:
            z = min(sizes, key=lambda z: _dist(p, (z[0], z[1])))
            if _dist(p, (z[0], z[1])) <= 350 * s:
                size, size_src = z[2], z[3] + " [nearest - verify]"
        if not svc and svcs:
            z = min(svcs, key=lambda z: _dist(p, (z[0], z[1])))
            if _dist(p, (z[0], z[1])) <= 200 * s:
                svc = z[2] + " (nearest - verify)"
        note = ""
        if u["fn"] == "SX":
            near = [d for d in desc_items if _dist(u["pos"], (d[0], d[1])) <= 90 * s]
            for d in near:
                m = NOTE_REF.search(d[2])
                if m and m.group(1) in notes:
                    note = re.sub(r"[\s.]+[A-Z]?$", "", notes[m.group(1)]).capitalize()
                elif "MANIFOLD" in d[2] and not note:
                    note = d[2].strip(" -").capitalize()
        acc = sorted({o["fn"] for o in uniq.values() if o["pre"] == u["pre"] and o["num"] == u["num"]
                      and o["fn"] != u["fn"] and o["n"] > 0 and (o["fn"] not in VALVE_FN or o["fn"] == "SV")}
                     ) if u["fn"] in ACTUATED else []
        rows.append(dict(pre=u["pre"], fn=u["fn"], num=u["num"], pos=p, fail=u["fail"], acc=acc, size=size,
                         size_src=size_src, svc=svc, note=note, pipe=bool(pp), count=u["n"]))
    others = [u for u in uniq.values() if u["fn"] not in VALVE_FN and u["n"] > 0]
    return rows, others, prefix_hints, dwg



# =============================================================================
#  OCR READER  (for P&IDs whose text was converted to outlines / scanned drawings)
# =============================================================================
L2D = {"O": "0", "Q": "0", "D": "0", "I": "1", "L": "1", "S": "5", "B": "8", "Z": "2", "G": "6", "T": "7"}
D2L = {"0": "O", "1": "I", "5": "S", "8": "B", "2": "Z", "6": "G", "4": "A"}
KNOWN_FN = ["ZSO", "ZSC", "FCV", "XV", "PIT", "PT", "PI", "PSL", "PSH", "PSLL", "PSHH", "TI", "TT", "TIT", "FI", "FT",
            "FIT", "LI", "LT", "LIT", "LSL", "LSH", "LSLL", "LSHH", "SV", "PCV", "LCV", "TCV", "HV", "ZS", "PDT",
            "PDI", "PDIT", "AIT", "ESDV", "SDV", "BDV", "MOV", "FE", "TE", "PG", "TG", "PSV", "PRV"]
ACCESSORY_FN = {"ZSO", "ZSC", "ZS", "SV"}
OCR_WL = "-c tessedit_char_whitelist=ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789_-"
OCR_DIGITS = "-c tessedit_char_whitelist=0123456789"
BOM_TYPE = {"KGV": "Knife Gate", "BALL": "Ball", "BUTTERFLY": "Butterfly", "CHECK": "Check", "PLUG": "Plug",
            "FLOAT TYPE": "Float", "ROOT ISOLATION": "Root Isolation", "PRESSURE RELIEF": "Pressure Relief",
            "GATE": "Gate", "GLOBE": "Globe", "NEEDLE": "Needle", "DIAPHRAGM": "Diaphragm"}
BOM_OPR = {"MAN OPR": "Manual", "MANOPR": "Manual", "AC": "Air cylinder operated", "SAV": "Solenoid air valve",
           "LS": "Limit switches", "W/LSO": "with open limit switch", "LSO": "open limit switch",
           "MOV": "Motor operated", "PNEU": "Pneumatic"}


TESS_CANDIDATES = [
    r"C:\Program Files\Tesseract-OCR\tesseract.exe",
    r"C:\Program Files (x86)\Tesseract-OCR\tesseract.exe",
    os.path.expanduser(r"~\AppData\Local\Tesseract-OCR\tesseract.exe"),
    os.path.expanduser(r"~\AppData\Local\Programs\Tesseract-OCR\tesseract.exe"),
]


def tess_ready():
    """Point pytesseract at the path typed in Settings (or a usual install folder), then test it."""
    if not (pytesseract and Image):
        return False
    typed = str(st.session_state.get("pid_tess", "")).strip().strip('"')
    paths = ([typed] if typed else []) + TESS_CANDIDATES
    for p in [None] + paths:                      # None = whatever is on PATH
        try:
            if p is not None:
                if not os.path.exists(p):
                    continue
                pytesseract.pytesseract.tesseract_cmd = p
            pytesseract.get_tesseract_version.__wrapped__() if hasattr(pytesseract.get_tesseract_version, "__wrapped__") \
                else pytesseract.get_tesseract_version()
            return True
        except Exception:
            continue
    return False


def _fix(s, kind):
    m = L2D if kind == "d" else D2L
    return "".join(m.get(c, c) for c in s)


def norm_l1(s):
    s = re.sub(r"[^A-Z0-9]", "", s.upper())
    if len(s) < 6:
        return None
    s = _fix(s[:2], "d") + _fix(s[2:5], "l") + _fix(s[5:6], "d") + s[6:].replace("O", "0").replace("Q", "0")
    return s if re.fullmatch(r"\d{2}[A-Z]{3}\d[A-Z0-9]{0,2}", s) else None


def norm_l2(s):
    s = re.sub(r"[^A-Z0-9_]", "", s.upper().replace("-", "_"))
    parts = s.split("_")
    a = parts[0]
    if len(a) != 5:
        return None
    a = _fix(a[:2], "l") + _fix(a[2:], "d")
    if not re.fullmatch(r"[A-Z]{2}\d{3}", a):
        return None
    if len(parts) > 1 and len(parts[1]) == 4:
        b = _fix(parts[1][:2], "l") + _fix(parts[1][2:], "d")
        if re.fullmatch(r"[A-Z]{2}\d{2}", b):
            return a + "_" + b
    return a


def norm_fn(s):
    s = re.sub(r"[^A-Z0-9]", "", s.upper())
    if not s:
        return None
    s = s if s.isalpha() else _fix(s, "l")
    if s in KNOWN_FN:
        return s
    m = difflib.get_close_matches(s, KNOWN_FN, n=1, cutoff=0.6)
    return m[0] if m else None


def parse_bubble(text):
    lines = [l.strip() for l in text.splitlines() if l.strip()]
    l1 = l2 = None
    for l in lines:
        if l1 is None and norm_l1(l):
            l1 = norm_l1(l)
        elif l2 is None and norm_l2(l):
            l2 = norm_l2(l)
    fn = norm_fn(lines[-1]) if lines else None
    if not fn:
        fn = next((norm_fn(l) for l in reversed(lines) if norm_fn(l)), None)
    return [l1, l2, fn]


def make_renderer(pdoc, pno, W, H):
    """render(bbox_in_pts(top-left origin), scale) -> grayscale PIL image of just that area"""
    pg = pdoc[pno]
    def render(bb, scale):
        x0, t, x1, b = bb
        x0, t, x1, b = max(0, x0), max(0, t), min(W, x1), min(H, b)
        return pg.render(scale=scale, crop=(x0, H - b, W - x1, t)).to_pil().convert("L")
    return render


def _clamp(page, bb):
    x0, t, x1, b = bb
    return (max(page.bbox[0], x0), max(page.bbox[1], t), min(page.bbox[2], x1), min(page.bbox[3], b))


def _circle_img(render, c, scale):
    img = render((c["x0"], c["top"], c["x1"], c["bottom"]), scale)
    w, h = img.size
    mask = Image.new("L", (w, h), 0)
    ImageDraw.Draw(mask).ellipse((w * .035, h * .035, w * .965, h * .965), fill=255)
    return Image.composite(img, Image.new("L", (w, h), 255), mask)


def _tess(img, cfg):
    try:
        return pytesseract.image_to_string(img, config=cfg)
    except Exception:
        return ""


def read_bubbles(render, circs, workers=4):
    """OCR every tag bubble (3 stacked lines: system / loop-or-equipment / function); majority vote over variants."""
    variants = [(5, "--psm 6 " + OCR_WL), (5, "--psm 11 " + OCR_WL), (8, "--psm 6 " + OCR_WL), (3, "--psm 6 " + OCR_WL)]
    votes = [[[], [], []] for _ in circs]
    with ThreadPoolExecutor(workers) as ex:
        for sc, cfg in variants:
            pics = [_circle_img(render, c, sc) for c in circs]     # render sequentially (pdfium is not thread-safe)
            for i, t in enumerate(ex.map(lambda im: _tess(im, cfg), pics)):
                p = parse_bubble(t)
                for k in range(3):
                    if p[k]:
                        votes[i][k].append(p[k])
            del pics
    out = []
    for v in votes:
        ratios = []
        for k in range(3):
            c = collections.Counter(v[k])
            ratios.append(c.most_common(1)[0][1] / len(variants) if c else 0)
        out.append([_vote(v[0]), _vote(v[1]), _vote(v[2]), min(ratios)])
    return out


def read_balloons(render, circs, workers=4):
    """item-number balloons: several OCR variants, majority vote -> list of candidate-count dicts"""
    variants = [(6, "--psm 7"), (4, "--psm 7"), (6, "--psm 8"), (8, "--psm 10"), (5, "--psm 7")]
    votes = [collections.Counter() for _ in circs]
    with ThreadPoolExecutor(workers) as ex:
        for sc, cfg in variants:
            pics = [_circle_img(render, c, sc) for c in circs]
            for i, o in enumerate(ex.map(lambda im: _tess(im, cfg + " " + OCR_DIGITS).strip(), pics)):
                if o.isdigit() and 0 < int(o) < 100:
                    votes[i][int(o)] += 1
            del pics
    return votes


def find_bom_box(page):
    ims = [i for i in page.images if i["x1"] - i["x0"] > 300 and i["bottom"] - i["top"] > 150]
    if not ims:
        return None
    i = max(ims, key=lambda i: (i["x1"] - i["x0"]) * (i["bottom"] - i["top"]))
    return (i["x0"] - 5, i["top"] - 5, i["x1"] + 5, i["bottom"] + 5)


def _strip_grid(img):
    a = np.array(img.convert("L")) < 140
    rows, cols = a.mean(axis=1) > 0.45, a.mean(axis=0) > 0.45
    def grow(m, k=3):
        o = m.copy()
        for d in range(1, k + 1):
            o[d:] |= m[:-d]; o[:-d] |= m[d:]
        return o
    a[grow(rows), :] = False
    a[:, grow(cols)] = False
    return Image.fromarray(np.where(a, 0, 255).astype("uint8"))


def read_bom(page, render):
    box = find_bom_box(page)
    if not box:
        return {}
    txt = _tess(_strip_grid(render(box, 4)), "--psm 6")
    rows = []
    for ln in txt.splitlines():
        ln = ln.strip()
        m = re.match(r"^\W*(\d{1,2})\b(.*?)\bBO\W+(\d{1,3})\W+(.+)$", ln)
        if m:
            rows.append([int(m.group(1)), int(m.group(3)), m.group(4).strip()])
            continue
        m = re.match(r"^\W*(?:(\d{1,2})\W+)?(?:.*?\s)?BO\W+(\d{1,3})\W+(.+)$", ln)     # item number lost by OCR
        if m:
            rows.append([int(m.group(1)) if m.group(1) else None, int(m.group(2)), m.group(3).strip()])
            continue
        if re.search(r"NOT\s*US", ln.upper()):
            m = re.match(r"^\W*(\d{1,2})\b", ln)
            rows.append([int(m.group(1)) if m else None, 0, "NOT USED"])
    taken = {r[0] for r in rows if r[0]}
    free = [n for n in range(1, max(taken | {len(rows)}) + 1) if n not in taken]
    for r in rows:
        if r[0] is None and free:
            r[0] = free.pop(0)
    return {r[0]: dict(qty=r[1], desc=re.sub(r"[;|_]+\s*$", "", r[2]).strip()) for r in rows if r[0]}


def parse_valve_desc(desc):
    """'VALVE-BUTTERFLY, 10", AC, SAV, LS' -> (type, size, operator) or None if not a valve."""
    d = desc.upper().replace("”", '"')
    if not d.startswith("VALVE"):
        return None
    body = re.sub(r"^VALVE\W*", "", d)
    typ = body.split(",")[0].strip()
    typ = BOM_TYPE.get(typ, typ.title()) + " Valve"
    m = re.search(r"(\d+(?:-\d/\d|/\d+)?)\s*\"", d) or re.search(r",\s*(\d+(?:/\d+)?)\b", d)
    size = (m.group(1) + '"') if m else ""
    rest = d[m.end():] if m else ""
    ops = []
    if re.search(r"MAN\s*OPR", rest):
        ops.append("Manual")
    if re.search(r"\bAC\b", rest):
        ops.append("Air cylinder operated")
    if re.search(r"\bSAV\b", rest):
        ops.append("Solenoid air valve")
    if re.search(r"\bLS\b", rest):
        ops.append("Limit switches")
    if re.search(r"LSO", rest):
        ops.append("Open limit switch")
    return typ, size, ", ".join(ops)


def ocr_text_items(page, render, scale=1.6):
    """Whole-sheet OCR of the free text (line labels, services, sizes) -> items in PDF points."""
    img = render((0, 0, page.width, page.height), scale)
    d = pytesseract.image_to_data(img, config="--psm 11", output_type=pytesseract.Output.DICT)
    lines = collections.OrderedDict()
    for i, tx in enumerate(d["text"]):
        if tx.strip() and float(d["conf"][i]) > 40:
            lines.setdefault((d["block_num"][i], d["par_num"][i], d["line_num"][i]), []).append(i)
    items = []
    for idx in lines.values():
        x0 = min(d["left"][i] for i in idx) / scale
        x1 = max(d["left"][i] + d["width"][i] for i in idx) / scale
        t0 = min(d["top"][i] for i in idx) / scale
        t1 = max(d["top"][i] + d["height"][i] for i in idx) / scale
        items.append(dict(text=" ".join(d["text"][i] for i in idx).upper().replace("[", "(").replace("]", ")"),
                          x0=x0, x1=x1, top=t0, bottom=t1, size=t1 - t0, rot=False,
                          cx=(x0 + x1) / 2, cy=(t0 + t1) / 2))
    return items


def _cluster(members, dia):
    """chain neighbouring bubbles on the same row (ZSO / ZSC / valve) into clusters"""
    ms = sorted(members, key=lambda m: (round(m["cy"] / 8), m["cx"]))
    cl = []
    for m in ms:
        if cl and abs(cl[-1][-1]["cy"] - m["cy"]) < 8 and m["cx"] - cl[-1][-1]["cx"] < dia * 1.25:
            cl[-1].append(m)
        else:
            cl.append([m])
    return cl


def _vote(vals):
    vals = [v for v in vals if v]
    return collections.Counter(vals).most_common(1)[0][0] if vals else None


def parse_page_ocr(page, render, svc_words, right_pct=10, do_text=True, workers=4, status=None, bom_default=None):
    W, H = page.width, page.height
    s = W / 3024.0
    say = status or (lambda m: None)
    right = W * (1 - right_pct / 100)
    circ = [c for c in page.curves if abs(c["width"] - c["height"]) < 2.5 and len(c.get("pts", [])) >= 3
            and c["x1"] < right]
    tagc = [c for c in circ if 0.014 * W <= c["width"] <= 0.028 * W]
    ballc = [c for c in circ if 0.0065 * W <= c["width"] <= 0.0125 * W]
    say(f"reading {len(tagc)} tag bubbles")
    bub = read_bubbles(render, tagc, workers) if tagc else []
    members = [dict(l1=b[0], l2=b[1], fn=b[2], agree=b[3], cx=(c["x0"] + c["x1"]) / 2, cy=(c["top"] + c["bottom"]) / 2,
                    dia=c["width"]) for b, c in zip(bub, tagc)]
    members = [m for m in members if m["fn"] or m["l1"] or m["l2"]]
    # --- canonical system prefix (first 5 chars) by sheet-wide voting
    cnt = collections.Counter(m["l1"][:5] for m in members if m["l1"])
    canon = {k for k, v in cnt.items() if v >= 3}
    for m in members:
        if m["l1"] and canon:
            for cand in [m["l1"]] + [m["l1"][:i] + m["l1"][i + 1:] for i in range(len(m["l1"]))]:
                if cand[:5] in canon:
                    m["l1"] = cand
                    break
            else:
                best = difflib.get_close_matches(m["l1"][:5], list(canon), n=1, cutoff=0.6)
                if best:
                    m["l1"] = best[0] + m["l1"][5:]
    # --- cluster consensus (ZSO / ZSC / valve bubbles of one valve share system + equipment code)
    dia = max((m["dia"] for m in members), default=68)
    for cl in _cluster(members, dia):
        if len(cl) < 2:
            continue
        l1 = _vote([m["l1"] for m in cl]); l2 = _vote([(m["l2"] or "").split("_")[0] for m in cl])
        for m in cl:
            if l1:
                m["l1"] = l1
            if l2:
                sfx = "_" + m["l2"].split("_")[1] if m["l2"] and "_" in m["l2"] else ""
                m["l2"] = l2 + sfx
    for m in members:                                    # KKS measuring-circuit letters follow the variable
        if m["fn"] and m["l2"] and not m["l2"].startswith("AA") and m["fn"] not in ACCESSORY_FN:
            want = {"P": "CP", "T": "CT", "F": "CF", "L": "CL"}.get(m["fn"][0])
            if want:
                m["l2"] = want + m["l2"][2:]
    for fnx, dflt in (("ZSO", "XB01"), ("ZSC", "XB02")):
        sf = collections.Counter(m["l2"].split("_")[1] for m in members
                                 if m["fn"] == fnx and m["l2"] and "_" in m["l2"])
        dom = sf.most_common(1)[0][0] if sf and sf.most_common(1)[0][1] >= 3 else dflt
        for m in members:
            if m["fn"] == fnx and m["l2"]:
                m["l2"] = m["l2"].split("_")[0] + "_" + dom
    # --- group into valves (KKS class AA or a valve function) and instruments
    bases = collections.OrderedDict()
    for m in members:
        if m["l1"] and m["l2"]:
            key = (m["l1"], m["l2"].split("_")[0])
        else:
            key = ("?", f"{m['cx']:.0f},{m['cy']:.0f}")
        bases.setdefault(key, []).append(m)
    say("reading bill of materials")
    bom = read_bom(page, render) if find_bom_box(page) else {}
    bom = bom or dict(bom_default or {})              # BOM is often printed on sheet 1 only
    pipes = build_pipes(page, s, right_pct)
    def on_axis(c):
        mx, my = (c["x0"] + c["x1"]) / 2, (c["top"] + c["bottom"]) / 2
        return any((pp["o"] == "H" and abs(my - pp["c"]) < 5 and pp["lo"] - 40 <= mx <= pp["hi"] + 40) or
                   (pp["o"] == "V" and abs(mx - pp["c"]) < 5 and pp["lo"] - 40 <= my <= pp["hi"] + 40) for pp in pipes)
    ballc = [c for c in ballc if not on_axis(c)]
    say(f"reading {len(ballc)} item balloons")
    votes = read_balloons(render, ballc, workers) if ballc else []
    items = []
    if do_text:
        say("reading free text (services, line sizes)")
        items = [i for i in ocr_text_items(page, render) if i["cx"] < right]
    sizes, svcs = [], []
    for it in items:
        m = SIZE_ONLY.match(it["text"]) or LINE_LABEL.match(it["text"])
        if m:
            sizes.append((it["cx"], it["cy"], m.group("size").replace(" ", "") + '"', it["text"]))
        mm = re.search(r"(\d{2,3})\s*NB\s*\(?\s*(\d+(?:-\d/\d)?)\s*[\"”']?\s*\)?", it["text"])
        if mm:
            sizes.append((it["cx"], it["cy"], mm.group(2) + '"', it["text"]))
    blocks_ = merge_blocks(items)
    kw = [re.compile(r"\b" + re.escape(w.strip().upper()) + r"\b") for w in svc_words if w.strip()]
    for b in blocks_:
        c = clean_service(b["text"])
        if 2 < len(c) <= 45 and not re.search(r"\(|SEE |NOTE|DRAWING", c.upper()) and any(k.search(c.upper()) for k in kw):
            svcs.append((b["cx"], b["cy"], c))
    note_svcs = []                                   # e.g. "250 NB (10") CONVEYING AIR VALVE, AIR CYLINDER OPERATED ..."
    for b in blocks_:
        m = re.search(r"\d+\s*NB\s*\(?[\d\"\-/ ]*\)?\s*(.*?)\bVALVE\b(.*)", b["text"].upper())
        if m:
            sp = re.sub(r"^[^A-Z]*[A-Z0-9]{0,2}\)\s*", "", m.group(1)).strip(" ,()")
            manual = "MANUAL" in sp
            sp = re.sub(r"\b(ISOLATION|MANUAL)\b", "", sp)
            sp = re.sub(r"\s+", " ", sp).strip(" ,")
            if sp.count("(") > sp.count(")"):
                sp += ")"
            if 3 < len(sp) <= 60:
                note_svcs.append(dict(svc=sp, ac="AIR CYLINDER" in m.group(2), manual=manual, x=b["cx"], y=b["cy"]))
    # --- resolve balloon numbers: majority vote; near-ties (3/5, 1/7 ...) decided by size consistency with the sheet
    sheet_sizes = {z[2] for z in sizes}
    balloons = []
    for v, c in zip(votes, ballc):
        cands = [(n, k) for n, k in v.most_common() if n in bom]
        if not cands or cands[0][1] < 3:
            continue
        n = cands[0][0]
        if len(cands) > 1 and cands[1][1] >= cands[0][1] - 1 and sheet_sizes:
            def ok(x):
                vt_ = parse_valve_desc(bom[x]["desc"])
                return bool(vt_ and vt_[1] in sheet_sizes)
            good = [x for x, _ in cands[:2] if ok(x)]
            if len(good) == 1:
                n = good[0]
        balloons.append(dict(n=n, cx=(c["x0"] + c["x1"]) / 2, cy=(c["top"] + c["bottom"]) / 2))
    # --- valve rows
    vrows, others, used_balloons = [], [], set()
    valve_bases = {}
    for key, ms in bases.items():
        fns = {m["fn"] for m in ms}
        is_valve = key[1].startswith("AA") or bool(fns & VALVE_FN)
        if is_valve and key[0] != "?":
            valve_bases[key] = ms
        else:
            for m in ms:
                if m["fn"] in ACCESSORY_FN:
                    continue
                others.append(dict(pre="", fn=m["fn"] or "", num="", pos=(m["cx"], m["cy"]),
                                   tag=(" ".join(x for x in (m["l1"], m["l2"]) if x) or "?"),
                                   conf="" if (m["l1"] and m["l2"] and m["fn"] and m["agree"] >= 0.75)
                                   else "OCR low confidence - verify"))
    vb_list = list(valve_bases.items())
    cand_all = [b for b in balloons if parse_valve_desc(bom[b["n"]]["desc"])]
    pairs = []
    for vi, (key, ms) in enumerate(vb_list):
        fns = {m["fn"] for m in ms}
        cx = sum(m["cx"] for m in ms) / len(ms); cy = sum(m["cy"] for m in ms) / len(ms)
        need = None
        if fns & {"FCV", "XV", "PCV", "LCV", "TCV", "SV", "ZSC"}:
            need = ("AC", "SAV")
        elif fns & {"ZSO"}:
            need = ("LSO",)
        pool = [b for b in cand_all if need and any(n in bom[b["n"]]["desc"].upper() for n in need)] or cand_all
        for bi, b in enumerate(pool):
            pairs.append((math.hypot(b["cx"] - cx, b["cy"] - cy), vi, id(b), b))
    pairs.sort(key=lambda t: t[0])
    assign, ub = {}, set()
    for d, vi, bid, b in pairs:
        if vi not in assign and bid not in ub and d < 600 * s:
            assign[vi] = b; ub.add(bid)
    for vi, (key, ms) in enumerate(vb_list):
        fns = [m["fn"] for m in ms if m["fn"]]
        main = next((f for f in fns if f in VALVE_FN), None)
        cx = sum(m["cx"] for m in ms) / len(ms); cy = sum(m["cy"] for m in ms) / len(ms)
        b = assign.get(vi)
        vt = parse_valve_desc(bom[b["n"]]["desc"]) if b else None
        assumed = ""
        if not vt:                                   # balloon unreadable: pick the BOM item that fits the accessories
            need_ = ("AC", "SAV") if set(fns) & {"FCV", "XV", "PCV", "LCV", "TCV", "SV", "ZSC"} else (("LSO",) if "ZSO" in fns else None)
            fit = [n_ for n_, v_ in bom.items() if parse_valve_desc(v_["desc"]) and need_ and
                   any(x in v_["desc"].upper() for x in need_) and parse_valve_desc(v_["desc"])[1] in sheet_sizes]
            if len(fit) == 1:
                vt = parse_valve_desc(bom[fit[0]]["desc"])
                assumed = f"BOM item {fit[0]} assumed from accessories and line size"
                b = dict(n=fit[0], cx=cx, cy=cy)
        pos = (b["cx"], b["cy"]) if b else (cx, cy)
        size = vt[1] if vt else ""
        pp = pipe_for(pipes, pos, s)
        svc = ""
        if pp:
            along0 = pos[0] if pp["o"] == "H" else pos[1]
            c = [(abs(a - along0), z) for z in svcs for a in [on_pipe(pp, (z[0], z[1]), 45 * s, 260 * s)] if a is not None]
            if c:
                svc = min(c, key=lambda t: t[0])[1][2]
            if not size:
                c = [(abs(a - along0), z) for z in sizes for a in [on_pipe(pp, (z[0], z[1]), 30 * s)] if a is not None]
                if c:
                    size = min(c, key=lambda t: t[0])[1][2]
        if not svc and svcs:
            z = min(svcs, key=lambda z: math.hypot(pos[0] - z[0], pos[1] - z[1]))
            if math.hypot(pos[0] - z[0], pos[1] - z[1]) <= 250 * s:
                svc = z[2] + " (nearest - verify)"
        if not svc or "nearest" in svc:
            opr_ = vt[2] if vt else ""
            pick = [n_ for n_ in note_svcs if (n_["ac"] and "Air cylinder" in opr_) or (n_["manual"] and "Manual" in opr_)]
            if pick:
                svc = min(pick, key=lambda n_: math.hypot(n_["x"] - pos[0], n_["y"] - pos[1]))["svc"] + " (from note)"
        full = all(m["l1"] and m["l2"] and m["agree"] >= 0.75 for m in ms)
        vrows.append(dict(pre=key[0], fn=main or "AA", num=key[1], pos=pos, tag=f"{key[0]} {key[1]}", fail="",
                          acc=sorted({f for f in fns if f != main}),
                          size=size, size_src=("BOM item " + str(b["n"])) if b else "", svc=svc, note="",
                          vtype=vt[0] if vt else "", operator=vt[2] if vt else "",
                          item=b["n"] if b else "", bom=bom[b["n"]]["desc"] if b else "",
                          conf=assumed or ("" if (full and b) else "OCR/association - verify"), count=1))
        if b and not assumed:
            used_balloons.add(id(b))
        for m in ms:                                       # accessories as instrument rows
            if m["fn"] in ACCESSORY_FN and m["l1"] and m["l2"]:
                others.append(dict(pre="", fn=m["fn"], num="", pos=(m["cx"], m["cy"]),
                                   tag=f"{m['l1']} {m['l2']}", assoc=f"{key[0]} {key[1]}",
                                   conf="" if m["agree"] >= 0.75 else "OCR low confidence - verify"))
    # --- valves that carry no tag (only an item balloon)
    for b in cand_all:
        if id(b) in used_balloons:
            continue
        vt = parse_valve_desc(bom[b["n"]]["desc"])
        pos = (b["cx"], b["cy"])
        pp = pipe_for(pipes, pos, s)
        svc = ""
        if pp:
            along0 = pos[0] if pp["o"] == "H" else pos[1]
            c = [(abs(a - along0), z) for z in svcs for a in [on_pipe(pp, (z[0], z[1]), 45 * s, 260 * s)] if a is not None]
            if c:
                svc = min(c, key=lambda t: t[0])[1][2]
        vrows.append(dict(pre="", fn="", num="", pos=pos, tag=f"Item {b['n']} (no tag on drawing)", fail="", acc=[],
                          size=vt[1], size_src=f"BOM item {b['n']}", svc=svc, note="", vtype=vt[0], operator=vt[2],
                          item=b["n"], bom=bom[b["n"]]["desc"],
                          conf=f"Untagged - count = balloons on sheet (BOM qty for all sheets: {bom[b['n']]['qty']})", count=1))
    # --- drawing number (largest drawing-number text in the title block)
    dwg = ""
    cands = [(i["size"], i["cy"] + i["cx"], m.group(0)) for i in items for m in DWG_FIND.finditer(i["text"])]
    if cands:
        dwg = max(cands)[2]
    return vrows, others, dwg, bom


def pid_cfg():
    s = st.session_state
    vm = {}
    for ln in s["pid_vmap"].splitlines():
        if "=" in ln:
            k, v = ln.split("=", 1)
            vm[k.strip().upper()] = v.strip()
    return dict(svc_words=[w.strip().upper() for w in s["pid_svc"].split(",") if w.strip()],
                right_pct=s["pid_right"], include_sv=s["pid_sv"], valve_map=vm, prefix=s["pid_prefix"].strip(),
                ocr=s["pid_ocr"], ocr_text=s["pid_ocr_text"], tess=s["pid_tess"].strip(),
                workers=max(1, min(4, os.cpu_count() or 2)))


def operator_of(fn, acc):
    if fn in ACTUATED:
        t = "Pneumatic actuator"
        if "SV" in acc:
            t += ", solenoid-operated"
        if "ZSO" in acc or "ZSC" in acc:
            t += ", with limit switches"
        return t
    if fn in MANUAL_FN:
        return "Manual"
    if fn in SELF_FN:
        return "Self-operated"
    return ""


def _split_num(num):
    m = re.match(r"(\d+)([A-Z]?)$", num)
    return (m.group(1), m.group(2)) if m else (num, "")


def pid_extract(name, data, cfg, log, status=None):
    """Geometry-aware P&ID reading: valve tag, type, size, service per valve + instrument tags.
    Pages whose text is real are read directly; pages with outlined/scanned text are read by OCR."""
    if not pdfplumber:
        log.append((name, "P&ID", 0, "Install pdfplumber (pip install pdfplumber) for position-based reading"))
        return [], []
    say = status or (lambda m: None)
    vmap = dict(VALVE_MAP)
    vmap.update(cfg["valve_map"])
    rows, hints, notes = [], set(), []
    tess_ready()                                 # sets the Tesseract path (typed or auto-detected)

    def tag_of(pre, fn, num):
        return cfg["prefix"] + "-".join(x for x in (pre, fn, num) if x)

    def full(tag):
        return tag if tag.startswith("Item ") else cfg["prefix"] + tag

    def emit(pn, vrows, others, dwg):
        base = {"Source": name, "Loc": f"p{pn}",
                "P&ID Number": dwg or (DWG_FIND.search(name.upper()) or [os.path.splitext(name)[0]])[0]}
        valve_tags = {(r["pre"], r["num"]): (full(r["tag"]) if r.get("tag") else tag_of(r["pre"], r["fn"], r["num"]))
                      for r in vrows if not (r["fn"] == "SV" and not cfg["include_sv"])}
        extra_instr = []
        for r in vrows:
            n, sfx = _split_num(r["num"])
            if r["fn"] == "SV" and not cfg["include_sv"]:
                extra_instr.append(dict(pre=r["pre"], fn="SV", num=r["num"], pos=r["pos"]))
                continue
            vtype = r.get("vtype") or vmap.get(r["fn"], r["fn"] or "Valve")
            if r["note"]:
                vtype = r["note"]
            tag = full(r["tag"]) if r.get("tag") else tag_of(r["pre"], r["fn"], r["num"])
            opr = r["operator"] if "operator" in r else operator_of(r["fn"], r["acc"])
            extra = {}
            if "item" in r:
                extra = {"BOM Item": r["item"], "BOM Description": r["bom"], "Reading Check": r["conf"]}
            rows.append(dict(base, Category="Valve", Tag=tag, Type=vtype,
                             Description=(r["size"] + " " if r["size"] else "") + vtype +
                             (", " + r["fail"] if r["fail"] else ""),
                             Operator=opr,
                             **{"Valve Size": r["size"], "Service": r["svc"], "Fail Position": r["fail"],
                                "Accessories": ", ".join(r["acc"]), "Size Source": r["size_src"]}, **extra,
                             Prefix=r["fn"], Number=n, Suffix=sfx,
                             Context=f'x={int(r["pos"][0])}, y={int(r["pos"][1])} pt'))
        for o in others + extra_instr:
            fn = o["fn"]
            if fn in INSTR_OVERRIDE:
                mv, fu = INSTR_OVERRIDE[fn]
            elif len(fn) >= 2 and fn[0] in VAR_MAP and fn[-1] in INSTR_LAST and fn not in INSTR_EXCLUDE:
                mv, fu = VAR_MAP[fn[0]], " / ".join(FUNC_MAP.get(c, c) for c in fn[1:])
            else:
                n, sfx = _split_num(o["num"])
                pos = o.get("pos") or (o.get("real") or [(0, 0)])[0]
                if fn in EQUIP_MAP:
                    rows.append(dict(base, Category="Equipment", Tag=tag_of(o["pre"], fn, o["num"]),
                                     Type=EQUIP_MAP[fn], Prefix=fn, Number=n, Suffix=sfx,
                                     Context=f'x={int(pos[0])}, y={int(pos[1])} pt'))
                elif o.get("tag") and fn == "":
                    mv, fu = "", ""
                    rows.append(dict(base, Category="Instrument", Tag=full(o["tag"]), **{"Reading Check": o.get("conf", "")},
                                     Context=f'x={int(pos[0])}, y={int(pos[1])} pt'))
                continue
            n, sfx = _split_num(o["num"])
            pos = o.get("pos") or (o.get("real") or [(0, 0)])[0]
            tag = full(o["tag"]) if o.get("tag") else tag_of(o["pre"], fn, o["num"])
            assoc = full(o["assoc"]) if o.get("assoc") else (valve_tags.get((o["pre"], o["num"]), "") if o["pre"] else "")
            extra = {"Reading Check": o["conf"]} if o.get("conf") else {}
            rows.append(dict(base, Category="Instrument", Tag=tag,
                             **{"Measured Variable": mv, "Function": fu, "Loop No": n, "Associated Valve": assoc},
                             **extra, Prefix=fn, Number=n, Suffix=sfx, Context=f'x={int(pos[0])}, y={int(pos[1])} pt'))

    try:
        pdoc = pdfium.PdfDocument(data) if (pdfium and cfg.get("ocr")) else None
        with pdfplumber.open(io.BytesIO(data)) as pdf:
            bom_doc = {}
            if pdoc is not None and np is not None and tess_ready():          # find the bill of materials first
                for pn, page in enumerate(pdf.pages, 1):
                    if len(page.chars) < 400 and find_bom_box(page):
                        bom_doc = read_bom(page, make_renderer(pdoc, pn - 1, page.width, page.height))
                        if bom_doc:
                            break
            for pn, page in enumerate(pdf.pages, 1):
                outlined = len(page.chars) < 400
                if outlined:
                    if not cfg.get("ocr"):
                        notes.append(f"p{pn}: text is outlined (no selectable text) - enable OCR")
                        continue
                    missing = []
                    if not pytesseract or not Image:
                        missing.append("python packages pytesseract + pillow")
                    elif not tess_ready():
                        missing.append("the Tesseract program (not found - install it or set its path in Settings)")
                    if pdoc is None:
                        missing.append("python package pypdfium2")
                    if np is None:
                        missing.append("python package numpy (pip install numpy)")
                    if missing:
                        notes.append(f"p{pn}: text is outlined, OCR is required but missing: " + "; ".join(missing))
                        continue
                    rend = make_renderer(pdoc, pn - 1, page.width, page.height)
                    vrows, others, dwg, bom = parse_page_ocr(
                        page, rend, cfg["svc_words"], cfg["right_pct"], cfg.get("ocr_text", True),
                        cfg.get("workers", 4), lambda m, pn=pn: say(f"{name} p{pn}: {m}"), bom_doc)
                    notes.append(f"p{pn}: read by OCR (BOM items found: {len(bom)})")
                else:
                    vrows, others, ph, dwg = parse_page(page, cfg["svc_words"], cfg["right_pct"])
                    hints.update(ph)
                emit(pn, vrows, others, dwg)
        nv = sum(1 for r in rows if r["Category"] == "Valve")
        ne = sum(1 for r in rows if r["Category"] == "Equipment")
        log.append((name, "P&ID geometry", len(rows),
                    f"{nv} valves, {len(rows) - nv - ne} instrument tags, {ne} equipment. " + "; ".join(notes)))
    except Exception as ex:
        log.append((name, "P&ID geometry", 0, f"Error: {ex}"))
    return rows, sorted(hints)


# =============================================================================
#  EXTRACTION ENGINE
# =============================================================================
def natural_key(s):
    return [int(p) if p.isdigit() else p for p in re.split(r"(\d+)", str(s))]


def natural_str(s):
    return re.sub(r"\d+", lambda m: m.group().zfill(8), str(s))


def safe_compile(key):
    pat = st.session_state.get("pat_" + key, DEFAULT_PATTERNS[key])
    try:
        return re.compile(pat)
    except re.error:
        return re.compile(DEFAULT_PATTERNS[key])


def scan_blocks(blocks):
    rx = {k: safe_compile(k) for k in DEFAULT_PATTERNS}
    custom = []
    for r in st.session_state["custom_rules"]:
        if r.get("Name") and r.get("Pattern"):
            try:
                custom.append((r["Name"], re.compile(r["Pattern"])))
            except re.error:
                pass
    rows = []
    for src, loc, text in blocks:
        t = normalise(text)
        if not t.strip():
            continue
        claimed = []

        def free(m):
            return not any(m.start() < e and m.end() > s for s, e in claimed)

        def ctx(m):
            return t[max(0, m.start() - 35): m.end() + 35].replace("\n", " ").strip()

        base = dict(Source=src, Loc=loc)
        line_pos = []
        # ---- lines
        for m in rx["line"].finditer(t):
            g = m.groupdict()
            size = (g.get("size") or "").strip()
            ln = f'{size}"-{g.get("service")}-{g.get("seq")}-{g.get("spec")}'
            if g.get("insul"):
                ln += f'-{g["insul"]}'
            claimed.append(m.span())
            line_pos.append((m.start(), ln, size, g.get("spec")))
            rows.append(dict(base, Category="Line", Tag=ln, Size=size + '"', Service=g.get("service"),
                             **{"Seq No": g.get("seq"), "Piping Spec": g.get("spec"),
                                "Insulation": g.get("insul") or ""}, Context=ctx(m)))

        def nearest(pos):
            if not line_pos:
                return "", "", ""
            p = min(line_pos, key=lambda x: abs(x[0] - pos))
            if abs(p[0] - pos) > 400:
                return "", "", ""
            return p[1], p[2] + '"', p[3]

        # ---- valves
        for m in rx["valve"].finditer(t):
            if not free(m):
                continue
            g = m.groupdict()
            claimed.append(m.span())
            n, sz, sp = nearest(m.start())
            rows.append(dict(base, Category="Valve", Tag=f'{g["type"]}-{g["num"]}{g.get("sfx") or ""}',
                             Type=VALVE_MAP.get(g["type"], g["type"]), Prefix=g["type"], Number=g["num"],
                             Suffix=g.get("sfx") or "", **{"Nearest Line No": n, "Size (from line)": sz,
                                                           "Piping Spec (from line)": sp}, Context=ctx(m)))
        # ---- equipment
        for m in rx["equipment"].finditer(t):
            if not free(m):
                continue
            g = m.groupdict()
            claimed.append(m.span())
            rows.append(dict(base, Category="Equipment", Tag=f'{g["type"]}-{g["num"]}{g.get("sfx") or ""}',
                             Type=EQUIP_MAP.get(g["type"], g["type"]), Prefix=g["type"], Number=g["num"],
                             Suffix=g.get("sfx") or "", Context=ctx(m)))
        # ---- instruments (ISA-5.1 validated)
        for m in rx["instrument"].finditer(t):
            if not free(m):
                continue
            g = m.groupdict()
            fn = g["func"]
            if fn in INSTR_EXCLUDE or fn[0] not in VAR_MAP or fn[-1] not in INSTR_LAST:
                continue
            claimed.append(m.span())
            n, sz, sp = nearest(m.start())
            rows.append(dict(base, Category="Instrument", Tag=f'{fn}-{g["loop"]}{g.get("sfx") or ""}',
                             **{"Measured Variable": VAR_MAP[fn[0]],
                                "Function": " / ".join(FUNC_MAP.get(c, c) for c in fn[1:]),
                                "Loop No": g["loop"]},
                             Prefix=fn, Number=g["loop"], Suffix=g.get("sfx") or "",
                             **{"Nearest Line No": n}, Context=ctx(m)))
        # ---- user-defined lists
        for name, cre in custom:
            for m in cre.finditer(t):
                row = dict(base, Category=name, Tag=m.group(0).strip(), Context=ctx(m))
                row.update({k: v for k, v in m.groupdict().items() if v})
                rows.append(row)
    return pd.DataFrame(rows)


def loc_join(s):
    items = sorted(set(s), key=natural_key)
    return ", ".join(items[:15]) + (" ..." if len(items) > 15 else "")


def aggregate(raw, category, tag_label):
    d = raw[raw["Category"] == category].copy()
    if d.empty:
        return pd.DataFrame()
    d["Qty"] = 1
    static = [c for c in d.columns if c not in BASE_KEYS | {"Qty"}]
    agg = {c: "first" for c in static}
    agg.update({"Source": lambda s: ", ".join(sorted(set(s))), "Loc": loc_join,
                "Context": "first", "Qty": "sum"})
    out = d.groupby("Tag", sort=False).agg(agg).reset_index()
    out = out.loc[:, [c for c in out.columns if out[c].replace("", pd.NA).notna().any() or c == "Tag"]]
    out = out.fillna("")
    out = out.sort_values("Tag", key=lambda s: s.map(natural_str)).reset_index(drop=True)
    out = out.rename(columns={"Tag": tag_label, "Source": "Source File(s)", "Loc": "Page/Sheet",
                              "Qty": "Qty (occurrences)"})
    # tidy column order: tag first, then technical fields, then source info
    tail = ["Source File(s)", "Page/Sheet", "Qty (occurrences)", "Context"]
    cols = [tag_label] + [c for c in out.columns if c not in tail + [tag_label]] + [c for c in tail if c in out.columns]
    out = out[cols]
    if "P&ID Number" not in out.columns:
        out.insert(1, "P&ID Number", "")
    blank = out["P&ID Number"].astype(str).str.strip() == ""
    out.loc[blank, "P&ID Number"] = out.loc[blank, "Source File(s)"].astype(str).str.split(",").str[0] \
        .map(lambda x: os.path.splitext(x.strip())[0])
    out["Remarks"] = ""
    return out



# ---- derived valve fields (paste above build_lists) --------------------------
TAG_PARSE = re.compile(r'^(?:(?P<unit>[A-Z]{2,4}\d{2,4})-)?(?:(?P<sys>[A-Z]{2,3})-)?(?P<fn>[A-Z]{2,4})-?'
                       r'(?P<num>\d{2,5})(?P<sfx>[A-Z]?)$')


def add_valve_fields(v):
    """Plant Unit / System Code / Valve Code / Sheet No, split out of tags like BWN007-NB-VL-6400 and 'p3'."""
    unit, sysc, code, sheet = [], [], [], []
    prefs = v["Prefix"] if "Prefix" in v.columns else [""] * len(v)
    locs = v["Page/Sheet"] if "Page/Sheet" in v.columns else [""] * len(v)
    for tag, pref, loc in zip(v["Tag No"], prefs, locs):
        m = TAG_PARSE.match(str(tag).strip().upper())
        unit.append((m.group("unit") or "") if m else "")
        sysc.append((m.group("sys") or "") if m else "")
        code.append(pref or ((m.group("fn") or "") if m else ""))
        sheet.append(", ".join(re.findall(r"(?<![A-Za-z])p(\d+)", str(loc))))
    v["Plant Unit"], v["System Code"], v["Valve Code"], v["Sheet No"] = unit, sysc, code, sheet
    return v


def build_lists(raw):
    L = {}
    if raw.empty:
        return L
    v = aggregate(raw, "Valve", "Tag No")
    if not v.empty:
        v = v.rename(columns={"Type": "Valve Type"})
        v = add_valve_fields(v)
        pref = ["Tag No", "P&ID Number", "Description", "Valve Type", "Valve Size", "Service", "Operator",
                "Fail Position", "Accessories", "BOM Item", "BOM Description", "Size Source", "Reading Check"]
        v = v[[c for c in pref if c in v.columns] + [c for c in v.columns if c not in pref]]
    L["Valve List"] = v
    L["Equipment List"] = aggregate(raw, "Equipment", "Tag No")
    L["Line List"] = aggregate(raw, "Line", "Line No")
    L["Instrument List"] = aggregate(raw, "Instrument", "Tag No")
    tags = raw[raw["Category"].isin(["Valve", "Equipment", "Instrument"])].copy()
    if not tags.empty:
        desc = tags["Type"].fillna("") if "Type" in tags else pd.Series("", index=tags.index)
        if "Function" in tags:
            desc = desc.where(desc != "", tags["Measured Variable"].fillna("") + " " + tags["Function"].fillna(""))
        tags["Description"] = desc
        tags["Qty"] = 1
        g = tags.groupby("Tag", sort=False).agg(
            Category=("Category", "first"), Description=("Description", "first"),
            **({"Nearest Line No": ("Nearest Line No", "first")} if "Nearest Line No" in tags else {}),
            **({"Valve Size": ("Valve Size", "first")} if "Valve Size" in tags else {}),
            **({"Service": ("Service", "first")} if "Service" in tags else {}),
            Source=("Source", lambda s: ", ".join(sorted(set(s)))), Loc=("Loc", loc_join),
            Qty=("Qty", "sum")).reset_index().fillna("")
        g = g.sort_values(["Category", "Tag"], key=lambda s: s.map(natural_str)).reset_index(drop=True)
        g = g.rename(columns={"Tag": "Tag No", "Source": "Source File(s)", "Loc": "Page/Sheet",
                              "Qty": "Qty (occurrences)"})
        g["Remarks"] = ""
        L["Tag List"] = g
    for name in sorted(set(raw["Category"]) - {"Valve", "Equipment", "Line", "Instrument"}):
        L[name] = aggregate(raw, name, "Tag No")
    return {k: v for k, v in L.items() if v is not None and not v.empty}


# =============================================================================
#  FORMAT / TEMPLATE HELPERS
# =============================================================================
def build_tpl_df(name, fields):
    loaded = st.session_state["tpl_loaded"].get(name, {})
    rows = []
    for i, f in enumerate(fields):
        if f in loaded:
            c = loaded[f]
            rows.append([c.get("order", i + 1), c.get("include", True), f, c.get("header", f)])
        else:
            rows.append([i + 1 + (100 if loaded else 0), f not in HIDE_BY_DEFAULT, f, f])
    return pd.DataFrame(rows, columns=["Order", "Include", "Field", "Header"])


def apply_template(df, tdf, add_sno):
    t = tdf[tdf["Include"]].sort_values("Order")
    out, used = pd.DataFrame(index=df.index), set()
    for _, r in t.iterrows():
        h = str(r["Header"] or r["Field"])
        while h in used:
            h += " "
        used.add(h)
        out[h] = df[r["Field"]] if r["Field"] in df.columns else ""
    out = out.reset_index(drop=True)
    if add_sno:
        out.insert(0, "S.No", range(1, len(out) + 1))
    return out


def to_excel_bytes(sheets):
    bio = io.BytesIO()
    with pd.ExcelWriter(bio, engine="openpyxl") as xw:
        for n, d in sheets.items():
            d.to_excel(xw, sheet_name=n[:31], index=False)
            ws = xw.sheets[n[:31]]
            for col in ws.columns:
                ws.column_dimensions[col[0].column_letter].width = min(
                    45, max(10, max(len(str(c.value or "")) for c in col) + 2))
    return bio.getvalue()


# =============================================================================
#  CUSTOMER FORMAT UPLOAD  (header row of an uploaded Excel/CSV/Word file = output layout)
# =============================================================================
SYN = {
    "Tag No": ["tag", "tag no", "tag number", "tag name", "valve tag", "valve no", "valve number", "equipment no",
               "equipment tag", "instrument tag", "item no", "tag nr", "tag id"],
    "Line No": ["line no", "line number", "line tag", "pipe no", "line id", "line"],
    "Description": ["description", "valve description", "item description", "equipment description",
                    "tag description"],
    "Valve Type": ["valve type", "type of valve", "type"],
    "Type": ["type", "equipment type"],
    "Valve Size": ["size", "valve size", "nps", "dn", "nominal size", "line size"],
    "Size": ["size", "line size", "nps", "dn"],
    "Service": ["service", "fluid", "fluid service", "process service", "line service",
                "p&id description", "service description"],
    "Operator": ["operator", "actuator", "actuation", "operation", "valve operator"],
    "Fail Position": ["fail", "fail position", "failure position", "fail safe", "fail action"],
    "Accessories": ["accessories", "accessory", "actuator accessories"],
    "P&ID Number": ["p&id number", "p&id no", "p&id", "pid no", "pid number", "pid", "drawing no", "drawing number",
                    "dwg no", "reference drawing", "ref drawing", "vendor p&id no"],
    "Source File(s)": ["source", "source file", "file name"],
    "Page/Sheet": ["sheet", "page"],
    "Sheet No": ["sht no", "sheet no", "sheet number"],
    "Plant Unit": ["plant unit", "unit"],
    "System Code": ["system code", "system"],
    "Valve Code": ["valve code"],
    "Number": ["number", "loop code"],
    "Suffix": ["suffix"],
    "Qty (occurrences)": ["qty", "quantity", "no off", "count", "nos"],
    "Remarks": ["remarks", "remark", "notes", "comments", "note"],
    "Measured Variable": ["measured variable", "variable", "parameter"],
    "Function": ["function", "instrument type", "instrument function"],
    "Piping Spec": ["spec", "piping spec", "piping class", "pipe class", "class", "line class"],
    "Insulation": ["insulation", "insulation type"],
    "Loop No": ["loop", "loop no", "loop number"],
    "Associated Valve": ["associated valve", "related valve"],
    "Category": ["category", "tag type"],
    "Nearest Line No": ["line no", "line number", "line tag"],
}
SNO_WORDS = {"s no", "sno", "sr no", "sl no", "serial no", "serial number", "sr", "sl", "#", "no", "item",
             "item no", "s n", "sn"}
REV_WORDS = {"revision", "rev", "rev no", "rev."}
TECH_FIELDS = {"Prefix", "Number", "Suffix", "Context", "Size Source", "Page/Sheet"}   # not used for loose matching
NO_AUTO = {"Prefix", "Context", "Size Source"}                                       # never auto-mapped
SPECIAL_TARGETS = ["(blank)", "(S.No)", "(Rev from sidebar)"]
GUIDE_WORDS = ("guide", "note", "instruction", "readme", "legend", "lookup", "help", "cover", "revision hist")

# Known customer forms: header -> (field, fixed value).  Matched by header names, so it works for any file name.
# "Fixed value" is written when the mapped field is empty (or "(blank)" is chosen).
PRESETS = [dict(
    name="Plant Bowen - manual valve list (Southern Company)",
    signature={"valve mark", "vendor p&id no", "plant unit", "system code"},
    cols={
        "Tag Name": ("Tag No", ""),
        "P&ID Description": ("Service", ""),
        "Valve Mark": ("(blank)", "--"),
        "Plant Unit": ("Plant Unit", ""),
        "System Code": ("System Code", ""),
        "Size": ("Valve Size", ""),
        "Number": ("Number", ""),
        "Suffix": ("Suffix", ""),
        "Valve Type": ("Valve Code", ""),
        "Insulation Type": ("(blank)", ""),
        "End Connection": ("(blank)", ""),
        "Rating": ("(blank)", ""),
        "Body Material": ("(blank)", ""),
        "Manufacturer": ("(blank)", ""),
        "Model Number": ("(blank)", ""),
        "Design Pressure [PSIG]": ("(blank)", ""),
        "Design Temperature [F]": ("(blank)", ""),
        "Purchased By": ("(blank)", "BV"),
        "Supplied By": ("(blank)", "UCC"),
        "Vendor Revision": ("(blank)", ""),
        "Vendor P&ID No": ("P&ID Number", ""),
        "Vendor Tag": ("Tag No", ""),
        "Purchase Order": ("(blank)", ""),
        "Skid Reference": ("(blank)", ""),
        "Skid Description": ("(blank)", ""),
        "PID No": ("(blank)", ""),
        "Sht No": ("Sheet No", ""),
        "Project ID": ("(blank)", "422049"),
        "Project Name": ("(blank)", "PLANT BOWEN"),
        "Project Description": ("(blank)", "BW PROJECT"),
        "Operating Company": ("(blank)", "SOUTHERN COMPANY"),
    })]


def _norm(h):
    return re.sub(r"[^a-z0-9&#]+", " ", str(h).lower()).strip()


def _hkey(h):
    return _norm(re.sub(r"\(\d+\)", "", str(h)))             # drop footnote marks like (1)


def is_guide_sheet(n):
    return any(w in str(n).lower() for w in GUIDE_WORDS)


def preset_for(headers):
    hn = {_hkey(h) for h in headers}
    for p in PRESETS:
        if p["signature"] <= hn:
            return {_hkey(k): v for k, v in p["cols"].items()}
    return {}


def format_sheets(data, ext):
    if ext in ("xlsx", "xlsm") and openpyxl:
        return openpyxl.load_workbook(io.BytesIO(data)).sheetnames
    return [None]


def _limit_to_table(ws, hrow, vals):
    """Only the columns of the form itself: the Excel Table on the header row, else the first contiguous run.
    (Stops the look-up lists / helper columns that sit beside the table being treated as output columns.)"""
    from openpyxl.utils import range_boundaries
    for t in ws.tables.values():
        c0, r0, c1, _ = range_boundaries(t.ref)
        if r0 == hrow:
            inside = [(c, h) for c, h in vals if c0 <= c <= c1]
            if len(inside) >= 2:
                return inside
    run = [vals[0]]
    for c, h in vals[1:]:
        if c - run[-1][0] > 1:
            break
        run.append((c, h))
    return run if len(run) >= 2 else vals


def read_format(data, ext, sheet=None):
    """Return dict(kind, cols=[(excel_col_index, header)], hrow, sheet, data) or None."""
    if ext in ("xlsx", "xlsm") and openpyxl:
        wb = openpyxl.load_workbook(io.BytesIO(data))
        ws = wb[sheet] if sheet in wb.sheetnames else wb.worksheets[0]
        for r in range(1, min(ws.max_row, 60) + 1):
            vals = [(c.column, str(c.value).strip()) for c in ws[r] if c.value not in (None, "")]
            if len(vals) >= 2 and sum(1 for _, v in vals if not re.fullmatch(r"[\d.]+", v)) >= 2:
                vals = _limit_to_table(ws, r, vals)
                # a "group" form has a drawing-number sub-header row (drawing no. alone on its row);
                # a normal example data row that merely contains a drawing number is NOT a group row
                nxt = [c for c in ws[r + 1] if c.value not in (None, "")]
                grp = [c.column for c in nxt if isinstance(c.value, str) and DWG_FULL.match(c.value.strip())] \
                    if len(nxt) <= 2 else []
                return dict(kind="xlsx", data=data, sheet=ws.title, hrow=r, cols=vals,
                            group_col=grp[0] if grp else None)
        return None
    if ext == "csv":
        df = pd.read_csv(io.BytesIO(data), nrows=0, encoding_errors="ignore")
        return dict(kind="csv", cols=list(enumerate(df.columns, 1)), hrow=1, sheet=None, data=data)
    if ext == "docx" and docx:
        d = docx.Document(io.BytesIO(data))
        if d.tables:
            return dict(kind="docx", cols=[(i + 1, c.text.strip()) for i, c in enumerate(d.tables[0].rows[0].cells)
                                           if c.text.strip()], hrow=1, sheet=None, data=data)
    return None


def auto_map(headers, fields):
    out = []
    exact = [f for f in fields if f not in NO_AUTO]
    loose = [f for f in fields if f not in TECH_FIELDS]
    for h in headers:
        n = _hkey(h)
        pick = "(blank)"
        if n in SNO_WORDS:
            pick = "(S.No)"
        elif n in REV_WORDS:
            pick = "(Rev from sidebar)"
        else:
            for f in exact:                       # exact / synonym match, first field in table order wins
                if n == _norm(f) or n in SYN.get(f, []):
                    pick = f
                    break
            else:
                for f in loose:                   # looser: header contains the field name / a long synonym
                    if (len(_norm(f)) >= 6 and _norm(f) in n) or any(len(x) >= 6 and x in n for x in SYN.get(f, [])):
                        pick = f
                        break
        out.append(pick)
    return out


def default_mapping(headers, fields):
    """-> (targets, fixed_values) : known-customer preset first, then automatic matching."""
    targets = auto_map(headers, fields)
    fixed = [""] * len(headers)
    pre = preset_for(headers)
    for i, h in enumerate(headers):
        k = _hkey(h)
        if k in pre:
            t, fx = pre[k]
            targets[i] = t if (t in fields or t in SPECIAL_TARGETS) else "(blank)"
            fixed[i] = fx
    return targets, fixed


def _isblank(v):
    return v is None or (isinstance(v, float) and v != v) or str(v).strip() in ("", "None", "nan")


def _num(v):
    """numeric text -> int (so Excel stores 6400 / 422049 as numbers, like the form's example row)"""
    s = str(v).strip()
    return int(s) if re.fullmatch(r"[1-9]\d*", s) else v


def apply_format(df, fmt, mapdf):
    out, used = pd.DataFrame(index=range(len(df))), set()
    df = df.reset_index(drop=True)
    has_fx = "Fixed value" in mapdf.columns
    for (_, hdr), (_, m) in zip(fmt["cols"], mapdf.iterrows()):
        h = re.sub(r"\s*\n\s*", " ", str(hdr)).strip()
        while h in used:
            h += " "
        used.add(h)
        tgt = m["Map to field"]
        fx = "" if (not has_fx or _isblank(m["Fixed value"])) else _num(m["Fixed value"])
        if tgt == "(S.No)":
            out[h] = list(range(1, len(df) + 1))
        elif tgt == "(Rev from sidebar)":
            out[h] = st.session_state.get("meta_rev", "")
        elif tgt in df.columns:
            vals = [("" if _isblank(v) else v) for v in df[tgt].values]
            if tgt in ("Number", "Sheet No"):
                vals = [_num(v) for v in vals]
            if fx != "":
                vals = [fx if v == "" else v for v in vals]
            out[h] = vals
        else:
            out[h] = [fx] * len(df) if fx != "" else [""] * len(df)
    return out


def _meta_tokens(list_name):
    s = st.session_state
    return {"{{customer}}": s["meta_customer"], "{{project}}": s["meta_project"], "{{docno}}": s["meta_docno"],
            "{{rev}}": s["meta_rev"], "{{date}}": datetime.date.today().strftime("%d-%b-%Y"),
            "{{prepared_by}}": s["meta_by"], "{{list}}": list_name}


def _extend_validations(ws, old_last, new_last):
    """drop-down ranges that ran to the form's last row follow the table when it grows"""
    from openpyxl.worksheet.cell_range import MultiCellRange
    for dv in ws.data_validations.dataValidation:
        rngs = []
        for rg in str(dv.sqref).split():
            m = re.fullmatch(r"([A-Z]+)(\d+):([A-Z]+)(\d+)", rg)
            if m and int(m.group(4)) == old_last:
                rg = f"{m.group(1)}{m.group(2)}:{m.group(3)}{new_last}"
            rngs.append(rg)
        dv.sqref = MultiCellRange(" ".join(rngs))


def fill_sheet(wb, sheet, hrow, cols, out, list_name, groups=None, group_col=None):
    from openpyxl.styles import Alignment, Border, Side
    from openpyxl.utils import get_column_letter, range_boundaries
    ws = wb[sheet]
    tokens = _meta_tokens(list_name)
    for row in ws.iter_rows(min_row=1, max_row=max(hrow, 1)):        # title block placeholders
        for c in row:
            if isinstance(c.value, str) and "{{" in c.value:
                v = c.value
                for k, t in tokens.items():
                    v = re.sub(re.escape(k), str(t), v, flags=re.I)
                c.value = v
    colidx = [c for c, _ in cols]
    c0, c1 = min(colidx), max(colidx)
    # Excel Table / drop-downs / look-up lists beside the table => never delete rows or touch the other columns
    tbl, tbl_end, tbl_c0, tbl_c1 = None, None, None, None
    for t in ws.tables.values():
        a, b, c, d = range_boundaries(t.ref)
        if b == hrow and a <= c0 and c1 <= c:
            tbl, tbl_end, tbl_c0, tbl_c1 = t, d, a, c
            break
    keep = bool(tbl or ws.data_validations.dataValidation)
    grouped = bool(group_col and groups)
    norm_row = hrow + 2 if grouped else hrow + 1
    styles = {col: copy(ws.cell(row=norm_row, column=col)._style) for col in range(c0, c1 + 1)}
    gstyles = {col: copy(ws.cell(row=hrow + 1, column=col)._style) for col in range(c0, c1 + 1)}
    norm_h = ws.row_dimensions[norm_row].height
    grp_h = ws.row_dimensions[hrow + 1].height
    g_dim = copy(ws.row_dimensions[hrow + 1]._style)
    n_dim = copy(ws.row_dimensions[norm_row]._style)
    grp_rows = set()
    for r in range(hrow + 1, ws.max_row + 1):                        # clear the form's sample/stale content
        for col in range(c0, c1 + 1):
            try:
                ws.cell(row=r, column=col).value = None
            except Exception:
                pass
        for col in range(1, c0):                                     # "RE- EXAMPLE" style markers left of the table
            cc = ws.cell(row=r, column=col)
            if isinstance(cc.value, str) and "EXAMPLE" in cc.value.upper():
                cc.value = None
    r, prev = hrow + 1, None
    for i, row in enumerate(out.itertuples(index=False)):
        g = groups[i] if grouped else None
        if grouped and g and g != prev:                              # drawing sub-header row
            for col in range(c0, c1 + 1):
                ws.cell(row=r, column=col)._style = copy(gstyles[col])
            ws.cell(row=r, column=group_col).value = g
            ws.row_dimensions[r].height = grp_h
            grp_rows.add(r)
            prev = g
            r += 1
        for col in range(c0, c1 + 1):
            cell = ws.cell(row=r, column=col)
            cell._style = copy(styles[col])
            al = cell.alignment
            cell.alignment = Alignment(horizontal=al.horizontal, vertical="center", wrap_text=True)
        for (col, _), val in zip(cols, row):
            ws.cell(row=r, column=col).value = val.item() if hasattr(val, "item") else val
        ws.row_dimensions[r].height = norm_h
        r += 1
    last = max(r - 1, hrow)
    for rr in range(hrow + 1, last + 1):                             # row-level styles of the old template rows
        ws.row_dimensions[rr]._style = copy(g_dim if rr in grp_rows else n_dim)
    if keep:
        if tbl is not None and last > tbl_end:                       # more rows than the form holds: grow the table
            ref = f"{get_column_letter(tbl_c0)}{hrow}:{get_column_letter(tbl_c1)}{last}"
            tbl.ref = ref
            if tbl.autoFilter is not None:
                tbl.autoFilter.ref = ref
            _extend_validations(ws, tbl_end, last)
        ws.print_area = f"A1:{get_column_letter(c1)}{max(last, tbl_end or last)}"
    else:
        for rr in range(hrow + 1, last + 1):                         # clear stray template fills beside the table
            for col in range(c1 + 1, ws.max_column + 1):
                ws.cell(row=rr, column=col)._style = copy(ws.cell(row=ws.max_row, column=col)._style)
        if ws.max_row > last:                                        # drop unused pre-formatted blank rows
            ws.delete_rows(last + 1, ws.max_row - last)
        med = Side(style="medium")
        for col in range(c0, c1 + 1):                                # close the table box
            c = ws.cell(row=last, column=col)
            b = copy(c.border)
            c.border = Border(left=b.left, right=b.right, top=b.top, bottom=med)
        ws.print_area = f"A1:{get_column_letter(c1 + 1)}{last}"
    try:
        if not (ws.oddFooter.center.text or ws.oddFooter.right.text or ws.oddFooter.left.text):
            ws.oddFooter.center.text = "Page &P of &N"
    except Exception:
        pass


def fill_template_xlsx(fmt, out, list_name="", groups=None):
    """One list -> customer's own workbook (other DATA sheets removed; guide / notes sheets are kept)."""
    wb = openpyxl.load_workbook(io.BytesIO(fmt["data"]))
    fill_sheet(wb, fmt["sheet"], fmt["hrow"], fmt["cols"], out, list_name, groups, fmt.get("group_col"))
    for n in list(wb.sheetnames):
        if n != fmt["sheet"] and not is_guide_sheet(n):
            wb.remove(wb[n])
    bio = io.BytesIO()
    wb.save(bio)
    return bio.getvalue()


def fill_book_xlsx():
    """All lists -> customer's workbook, every assigned sheet filled."""
    s = st.session_state
    book = s["fmtbook"]
    wb = openpyxl.load_workbook(io.BytesIO(book["data"]))
    done = 0
    for sheet, info in book["sheets"].items():
        ln = s.get(f"fmtassign_{sheet}", "(ignore)")
        if ln in s["final"] and s["final_fmt"].get(ln):
            fill_sheet(wb, sheet, info["hrow"], info["cols"], s["final"][ln], ln,
                       s["final_grp"].get(ln), info.get("group_col"))
            done += 1
    bio = io.BytesIO()
    wb.save(bio)
    return bio.getvalue(), done


FORMAT_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "customer_formats")
SHEET_KW = {"Valve List": ["valve"], "Equipment List": ["equip"], "Line List": ["line", "pipe"],
            "Instrument List": ["instr"], "Tag List": ["tag"]}


def safe_name(n):
    return re.sub(r"[^A-Za-z0-9 _.-]", "", n).strip()


def saved_formats():
    try:
        return sorted(f[:-5] for f in os.listdir(FORMAT_DIR) if f.endswith(".xlsx") and not f.endswith("_mapping.xlsx"))
    except Exception:
        return []


def _guess_list(*texts):
    for text in texts:
        hit = next((ln for ln, kws in SHEET_KW.items() if any(k in text.lower() for k in kws)), None)
        if hit:
            return hit
    return "(ignore)"


def build_book(data, name, assign=None, maps=None, fixed=None):
    s = st.session_state
    wb = openpyxl.load_workbook(io.BytesIO(data))
    sheets, guesses = {}, {}
    for sh in wb.sheetnames:
        if is_guide_sheet(sh):                             # guide / notes sheets are not data sheets
            continue
        info = read_format(data, "xlsx", sh)
        if not info:
            continue
        sheets[sh] = dict(hrow=info["hrow"], cols=info["cols"], group_col=info.get("group_col"))
        wsx = wb[sh]
        title = " ".join(str(c.value) for row in wsx.iter_rows(min_row=1, max_row=info["hrow"] - 1)
                         for c in row if c.value)
        guesses[sh] = _guess_list(sh, title)               # sheet name, then title text above the headers
    for sh, g in guesses.items():
        if g == "(ignore)" and len(sheets) == 1:           # single data sheet: the file name may tell the list
            g = _guess_list(name)
        s[f"fmtassign_{sh}"] = (assign or {}).get(sh, g)
    s["fmtbook"] = dict(name=name, data=data, sheets=sheets, maps=dict(maps or {}), maps0=dict(maps or {}),
                        fixed=dict(fixed or {}), fixed0=dict(fixed or {}))


def load_uploaded_format():
    f = st.session_state.get("fmt_upload")
    if f is not None and openpyxl:
        build_book(f.getvalue(), os.path.splitext(f.name)[0])
        st.session_state["fmt_save_name"] = os.path.splitext(f.name)[0]


def load_saved_format():
    ch = st.session_state.get("fmt_choice")
    if not ch or ch == "(none)":
        st.session_state["fmtbook"] = None
        return
    try:
        with open(os.path.join(FORMAT_DIR, ch + ".xlsx"), "rb") as fh:
            data = fh.read()
        assign, maps, fixed = {}, {}, {}
        mp = os.path.join(FORMAT_DIR, ch + "_mapping.xlsx")
        jp = os.path.join(FORMAT_DIR, ch + ".json")
        if os.path.exists(mp):
            sh = pd.read_excel(mp, sheet_name=None, dtype=object)
            for _, r in sh.get("Sheet assignment", pd.DataFrame(columns=["Sheet", "List"])).iterrows():
                assign[str(r["Sheet"])] = str(r["List"])
            for _, r in sh.get("Column mapping", pd.DataFrame(columns=["List", "Customer column", "Map to field"])).iterrows():
                maps.setdefault(str(r["List"]), {})[str(r["Customer column"])] = str(_cell(r["Map to field"]))
                fixed.setdefault(str(r["List"]), {})[str(r["Customer column"])] = str(_cell(r.get("Fixed value", "")))
        elif os.path.exists(jp):                       # older saved formats
            with open(jp) as fh:
                meta = json.load(fh)
            assign, maps = meta.get("assign", {}), meta.get("maps", {})
        build_book(data, ch, assign, maps, fixed)
        st.session_state["fmt_save_name"] = ch
    except Exception as ex:
        st.session_state["fmt_msg"] = f"Could not load format: {ex}"


def save_format():
    s = st.session_state
    book, nm = s["fmtbook"], safe_name(s.get("fmt_save_name", ""))
    if not book or not nm:
        s["fmt_msg"] = "Upload a format and enter a customer name first."
        return
    os.makedirs(FORMAT_DIR, exist_ok=True)
    with open(os.path.join(FORMAT_DIR, nm + ".xlsx"), "wb") as fh:
        fh.write(book["data"])
    assign = [(sh, s.get(f"fmtassign_{sh}", "(ignore)")) for sh in book["sheets"]]
    maps = [(ln, c, t, book.get("fixed", {}).get(ln, {}).get(c, "")) for ln, d in book["maps"].items()
            for c, t in d.items()]
    with pd.ExcelWriter(os.path.join(FORMAT_DIR, nm + "_mapping.xlsx"), engine="openpyxl") as xw:
        pd.DataFrame(assign, columns=["Sheet", "List"]).to_excel(xw, sheet_name="Sheet assignment", index=False)
        pd.DataFrame(maps, columns=["List", "Customer column", "Map to field", "Fixed value"]) \
            .to_excel(xw, sheet_name="Column mapping", index=False)
        for ws in xw.book.worksheets:
            for col in ws.columns:
                ws.column_dimensions[col[0].column_letter].width = min(
                    60, max(14, max(len(str(c.value or "")) for c in col) + 2))
    s["fmt_msg"] = f"Saved customer format '{nm}' (form + mapping Excel files in the customer_formats folder)."


def delete_format():
    nm = st.session_state.get("fmt_choice")
    for ext in (".xlsx", "_mapping.xlsx", ".json"):
        try:
            os.remove(os.path.join(FORMAT_DIR, nm + ext))
        except Exception:
            pass
    st.session_state["fmtbook"] = None
    st.session_state["fmt_choice"] = "(none)"


def active_fmt(list_name):
    s = st.session_state
    book = s["fmtbook"]
    if not book:
        return None
    for sh, info in book["sheets"].items():
        if s.get(f"fmtassign_{sh}") == list_name:
            return dict(kind="xlsx", data=book["data"], sheet=sh, hrow=info["hrow"], cols=info["cols"],
                        group_col=info.get("group_col"))
    return None


# =============================================================================
#  PDF OUTPUT
# =============================================================================
def build_pdf(sections, meta, style, logo):
    from reportlab.lib import colors
    from reportlab.lib.pagesizes import A3, A4, landscape, portrait
    from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
    from reportlab.lib.units import mm
    from reportlab.lib.utils import ImageReader
    from reportlab.platypus import PageBreak, Paragraph, SimpleDocTemplate, Spacer, Table, TableStyle

    base = A3 if style["size"] == "A3" else A4
    ps = landscape(base) if style["orient"] == "Landscape" else portrait(base)
    W, H = ps
    margin = 10 * mm
    hdr_col = colors.HexColor(style["color"])
    fs = style["font"]
    today = datetime.date.today().strftime("%d-%b-%Y")
    ss = getSampleStyleSheet()
    cell = ParagraphStyle("c", parent=ss["Normal"], fontSize=fs, leading=fs + 2)
    head = ParagraphStyle("h", parent=cell, textColor=colors.white, fontName="Helvetica-Bold")
    h1 = ParagraphStyle("h1", parent=ss["Heading2"], textColor=hdr_col)
    logo_reader = ImageReader(io.BytesIO(logo)) if logo else None

    def decorate(c, doc):
        c.saveState()
        c.setFillColor(hdr_col)
        c.rect(0, H - 20 * mm, W, 20 * mm, stroke=0, fill=1)
        c.setFillColor(colors.white)
        c.setFont("Helvetica-Bold", 14)
        c.drawString(margin, H - 9 * mm, meta.get("project") or "Project List")
        c.setFont("Helvetica", 8)
        c.drawString(margin, H - 15 * mm,
                     f"Doc No: {meta.get('docno', '')}   |   Rev: {meta.get('rev', '')}   |   "
                     f"Prepared by: {meta.get('by', '')}   |   Date: {today}")
        right = W - margin
        if logo_reader:
            c.drawImage(logo_reader, W - margin - 28 * mm, H - 18.5 * mm, width=28 * mm, height=17 * mm,
                        preserveAspectRatio=True, mask="auto")
            right -= 31 * mm
        c.setFont("Helvetica-Bold", 10)
        c.drawRightString(right, H - 9 * mm, meta.get("customer", ""))
        c.setFillColor(colors.grey)
        c.setFont("Helvetica", 7)
        c.drawString(margin, 6 * mm, "Generated by List Developer")
        c.drawRightString(W - margin, 6 * mm, f"Page {doc.page}")
        c.restoreState()

    story = []
    if len(sections) > 1:
        story.append(Paragraph("Summary of Lists", h1))
        data = [[Paragraph("List", head), Paragraph("No. of entries", head)]] + \
               [[Paragraph(escape(n), cell), Paragraph(str(len(d)), cell)] for n, d in sections.items()]
        t = Table(data, colWidths=[90 * mm, 40 * mm])
        t.setStyle(TableStyle([("BACKGROUND", (0, 0), (-1, 0), hdr_col),
                               ("GRID", (0, 0), (-1, -1), .25, colors.grey)]))
        story += [t, PageBreak()]
    avail = W - 2 * margin
    for i, (name, df) in enumerate(sections.items()):
        story.append(Paragraph(f"{escape(name)}  ({len(df)} entries)", h1))
        cols = list(df.columns)
        weights = []
        for c in cols:
            ml = max([len(str(x)) for x in df[c].head(300)] + [len(str(c))])
            weights.append(min(max(ml, 6), 38))
        widths = [avail * w / sum(weights) for w in weights]
        data = [[Paragraph(escape(str(c)).replace("\n", "<br/>"), head) for c in cols]]
        for row in df.itertuples(index=False):
            data.append([Paragraph(escape(str(v)), cell) for v in row])
        tb = Table(data, colWidths=widths, repeatRows=1)
        tb.setStyle(TableStyle([
            ("BACKGROUND", (0, 0), (-1, 0), hdr_col),
            ("GRID", (0, 0), (-1, -1), .25, colors.HexColor("#94a3b8")),
            ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.white, colors.HexColor("#ecfdf5")]),
            ("VALIGN", (0, 0), (-1, -1), "TOP"),
            ("LEFTPADDING", (0, 0), (-1, -1), 3), ("RIGHTPADDING", (0, 0), (-1, -1), 3),
        ]))
        story += [tb, Spacer(1, 6)]
        if i < len(sections) - 1:
            story.append(PageBreak())
    bio = io.BytesIO()
    SimpleDocTemplate(bio, pagesize=ps, leftMargin=margin, rightMargin=margin,
                      topMargin=24 * mm, bottomMargin=12 * mm,
                      title="List Developer output").build(story, onFirstPage=decorate, onLaterPages=decorate)
    return bio.getvalue()


def current_meta():
    s = st.session_state
    return dict(customer=s["meta_customer"], project=s["meta_project"], docno=s["meta_docno"],
                rev=s["meta_rev"], by=s["meta_by"])


def current_style():
    s = st.session_state
    return dict(size=s["sty_size"], orient=s["sty_orient"], color=s["sty_color"], font=s["sty_font"])


# =============================================================================
#  PROFILE (customer format) SAVE / LOAD
# =============================================================================
PID_KEYS = ("pid_on", "pid_prefix", "pid_sv", "pid_right", "pid_svc", "pid_vmap", "pid_ocr", "pid_ocr_text", "pid_tess")
BOOL_KEYS = {"sno", "pid_on", "pid_sv", "pid_ocr", "pid_ocr_text"}
INT_KEYS = {"font", "pid_right"}


def profile_dict():
    s = st.session_state
    tpl = {}
    for n, d in s["tpl_out"].items():
        tpl[n] = {r.Field: dict(order=int(r.Order), include=bool(r.Include), header=r.Header)
                  for r in d.itertuples()}
    return dict(meta=current_meta(), style=current_style(), sno=s["sty_sno"],
                patterns={k: s["pat_" + k] for k in DEFAULT_PATTERNS},
                custom_rules=s["custom_rules"], extras=s["extras"], templates=tpl,
                pid={k: s[k] for k in PID_KEYS})


def make_profile():
    """Customer profile as an Excel workbook (editable in Excel)."""
    p = profile_dict()
    settings = [("Project", k, v) for k, v in p["meta"].items()] + \
               [("PDF style", k, v) for k, v in p["style"].items()] + \
               [("PDF style", "sno", p["sno"])] + \
               [("P&ID reader", k, v) for k, v in p["pid"].items()]
    cols = [(ln, f, c["include"], c["order"], c["header"]) for ln, d in p["templates"].items() for f, c in d.items()]
    extras = [(ln, c) for ln, cs in p["extras"].items() for c in cs]
    bio = io.BytesIO()
    with pd.ExcelWriter(bio, engine="openpyxl") as xw:
        pd.DataFrame(settings, columns=["Section", "Key", "Value"]).to_excel(xw, sheet_name="Settings", index=False)
        pd.DataFrame(list(p["patterns"].items()), columns=["Name", "Pattern"]).to_excel(xw, sheet_name="Detection rules", index=False)
        pd.DataFrame(p["custom_rules"] or [], columns=["Name", "Pattern"]).to_excel(xw, sheet_name="Custom lists", index=False)
        pd.DataFrame(cols, columns=["List", "Field", "Include", "Order", "Header"]).to_excel(xw, sheet_name="Columns", index=False)
        pd.DataFrame(extras, columns=["List", "Extra column"]).to_excel(xw, sheet_name="Extra columns", index=False)
        for ws in xw.book.worksheets:
            for col in ws.columns:
                ws.column_dimensions[col[0].column_letter].width = min(
                    70, max(12, max(len(str(c.value or "")) for c in col) + 2))
    return bio.getvalue()


def _cell(v):
    return "" if (v is None or (isinstance(v, float) and v != v)) else v


def xlsx_to_profile(data):
    sheets = pd.read_excel(io.BytesIO(data), sheet_name=None, dtype=object)
    p = dict(meta={}, style={}, pid={}, patterns={}, custom_rules=[], extras={}, templates={})
    for _, r in sheets.get("Settings", pd.DataFrame(columns=["Section", "Key", "Value"])).iterrows():
        sec, k, v = str(r["Section"]), str(r["Key"]), _cell(r["Value"])
        if k in BOOL_KEYS:
            v = str(v).strip().lower() in ("true", "1", "yes")
        elif k in INT_KEYS:
            v = int(float(v)) if str(v) != "" else 0
        else:
            v = str(v)
        if k == "sno":
            p["sno"] = v
        elif sec == "Project":
            p["meta"][k] = v
        elif sec == "PDF style":
            p["style"][k] = v
        else:
            p["pid"][k] = v
    for _, r in sheets.get("Detection rules", pd.DataFrame(columns=["Name", "Pattern"])).iterrows():
        p["patterns"][str(r["Name"])] = str(_cell(r["Pattern"]))
    for _, r in sheets.get("Custom lists", pd.DataFrame(columns=["Name", "Pattern"])).iterrows():
        if _cell(r["Name"]) and _cell(r["Pattern"]):
            p["custom_rules"].append({"Name": str(r["Name"]), "Pattern": str(r["Pattern"])})
    for _, r in sheets.get("Columns", pd.DataFrame(columns=["List", "Field", "Include", "Order", "Header"])).iterrows():
        inc = _cell(r["Include"])
        p["templates"].setdefault(str(r["List"]), {})[str(r["Field"])] = dict(
            include=(str(inc).strip().lower() in ("true", "1", "yes")), order=int(float(_cell(r["Order"]) or 0)),
            header=str(_cell(r["Header"])))
    for _, r in sheets.get("Extra columns", pd.DataFrame(columns=["List", "Extra column"])).iterrows():
        if _cell(r["Extra column"]):
            p["extras"].setdefault(str(r["List"]), []).append(str(r["Extra column"]))
    return p


def apply_profile():
    f = st.session_state.get("profile_upload")
    if not f:
        return
    try:
        raw_ = f.getvalue()
        p = xlsx_to_profile(raw_) if f.name.lower().endswith((".xlsx", ".xlsm")) else json.loads(raw_.decode("utf-8"))
    except Exception as ex:
        st.session_state["profile_msg"] = f"Could not read the profile file ({ex})."
        return
    s = st.session_state
    m, sty = p.get("meta", {}), p.get("style", {})
    for k in ("customer", "project", "docno", "rev", "by"):
        if k in m:
            s["meta_" + k] = m[k]
    for k in ("size", "orient", "color", "font"):
        if k in sty:
            s["sty_" + k] = sty[k]
    s["sty_sno"] = p.get("sno", True)
    for k, v in p.get("patterns", {}).items():
        if k in DEFAULT_PATTERNS:
            s["pat_" + k] = v
    for k, v in p.get("pid", {}).items():
        if k in s:
            s[k] = v
    s["custom_rules"] = p.get("custom_rules", [])
    s["extras"] = p.get("extras", {})
    s["tpl_loaded"] = p.get("templates", {})
    s["tpl_ver"] += 1
    s["profile_msg"] = "Customer profile loaded. Click 'Extract Lists' again if the rules changed."


def add_extra(name):
    key = f"newcol_{name}"
    v = st.session_state.get(key, "").strip()
    if v:
        st.session_state["extras"].setdefault(name, [])
        if v not in st.session_state["extras"][name]:
            st.session_state["extras"][name].append(v)
    st.session_state[key] = ""


# =============================================================================
#  SIDEBAR
# =============================================================================
with st.sidebar:
    st.markdown("### 📁 1. Upload documents")
    files = st.file_uploader(
        "P&IDs, specs, drawings, data sheets",
        type=["pdf", "dxf", "dwg", "xlsx", "xlsm", "xls", "docx", "doc", "csv", "txt"],
        accept_multiple_files=True)
    use_ocr = st.checkbox("OCR scanned PDF pages (needs Tesseract)", value=False)
    st.checkbox("Read P&ID PDFs by drawing geometry (valve size / service / type)", key="pid_on")
    st.text_input("Project tag prefix (optional, e.g. BWN007-)", key="pid_prefix")
    st.checkbox("List solenoid valves (SV) as valves (otherwise in Instrument List)", key="pid_sv")
    st.checkbox("OCR drawings whose text is outlined / scanned (needs Tesseract; ~1 min per sheet)", key="pid_ocr")
    st.checkbox("OCR free text too (line service, pipe size) - slower", key="pid_ocr_text")
    if st.button("🚀 Extract Lists", type="primary"):
        if not files:
            st.warning("Upload at least one file.")
        else:
            with st.spinner("Reading documents (P&ID geometry reading can take a few seconds per sheet)..."):
                blocks, log = read_files(files, use_ocr)
                raw = scan_blocks(blocks)
                spatial_rows, spatial_src, hints = [], set(), set()
                if st.session_state["pid_on"]:
                    cfg = pid_cfg()
                    _ph = st.empty()
                    for f in files:
                        if f.name.lower().endswith(".pdf"):
                            r, h = pid_extract(f.name, f.getvalue(), cfg, log, lambda m: _ph.info(m))
                            if r:
                                spatial_rows += r
                                spatial_src.add(f.name)
                                hints.update(h)
                if spatial_rows:
                    if not raw.empty:
                        drop = raw["Category"].isin(["Valve", "Instrument", "Equipment"]) & raw["Source"].isin(spatial_src)
                        raw = raw[~drop]
                    raw = pd.concat([raw, pd.DataFrame(spatial_rows)], ignore_index=True)
            if st.session_state["pid_on"]:
                _ph.empty()
            st.session_state.update(raw=raw, log=log, lists=build_lists(raw), prefix_hints=sorted(hints),
                                    run_id=st.session_state["run_id"] + 1, pdf_bytes={})
            st.success(f"Done: {len(raw)} raw matches found.")
            for _l in log:
                if re.search(r"missing|Error|install|outlined|enable OCR", str(_l[3]), re.I):
                    st.error(f"{_l[0]}: {_l[3]}")
            if hints and not st.session_state["pid_prefix"]:
                st.info("Tag prefixes found on drawing: " + ", ".join(sorted(hints)) +
                        ". Enter one in 'Project tag prefix' and extract again to get full tags.")
    st.divider()
    st.markdown("### 🧾 2. Project / customer details")
    st.text_input("Customer", key="meta_customer")
    st.text_input("Project", key="meta_project")
    st.text_input("Document No.", key="meta_docno")
    c1, c2 = st.columns(2)
    c1.text_input("Rev", key="meta_rev")
    c2.text_input("Prepared by", key="meta_by")
    lg = st.file_uploader("Customer logo (optional)", type=["png", "jpg", "jpeg"])
    if lg:
        st.session_state["logo"] = lg.getvalue()
    st.divider()
    st.markdown("### 📐 4. Customer Excel format")
    if not openpyxl:
        st.warning("Install openpyxl (pip install openpyxl)")
    st.selectbox("Saved customer formats", ["(none)"] + saved_formats(), key="fmt_choice",
                 on_change=load_saved_format)
    st.file_uploader("Or upload a customer's Excel form (.xlsx)", type=["xlsx", "xlsm"], key="fmt_upload",
                     on_change=load_uploaded_format)
    _book = st.session_state["fmtbook"]
    if _book:
        st.caption(f"Active format: **{_book['name']}**. Assign each sheet to a list:")
        _opts = ["(ignore)"] + list(dict.fromkeys(LIST_NAMES + list(st.session_state["lists"])))
        for _sh in _book["sheets"]:
            st.selectbox(f"Sheet '{_sh}' →", _opts, key=f"fmtassign_{_sh}")
        st.text_input("Customer name (to save this format)", key="fmt_save_name")
        _c1, _c2 = st.columns(2)
        _c1.button("💾 Save customer format", on_click=save_format)
        _c2.button("🗑 Delete saved", on_click=delete_format)
        st.caption("Tip: put {{customer}}, {{project}}, {{docno}}, {{rev}}, {{date}}, {{prepared_by}} in the "
                   "Excel form's title area and they are filled automatically.")
    if st.session_state.get("fmt_msg"):
        st.info(st.session_state.pop("fmt_msg"))
    st.divider()
    st.markdown("### 🎨 5. PDF style")
    c1, c2 = st.columns(2)
    c1.selectbox("Page", ["A4", "A3"], key="sty_size")
    c2.selectbox("Orientation", ["Landscape", "Portrait"], key="sty_orient")
    st.color_picker("Header colour", key="sty_color")
    st.slider("Table font size", 5, 11, key="sty_font")
    st.checkbox("Add S.No column", key="sty_sno")

# =============================================================================
#  MAIN
# =============================================================================
st.markdown(
    "<div class='hero'><h1>📋 List Developer</h1>"
    "<p>Valve • Equipment • Line • Instrument • Tag • Custom lists — extracted from P&amp;IDs, specifications, "
    "AutoCAD (DXF), PDF, Excel &amp; Word</p></div>", unsafe_allow_html=True)

tab_names = ["📊 Dashboard"] + ["🔧 Valve List", "🏭 Equipment List", "🧵 Line List",
                                "🎛️ Instrument List", "🏷️ Tag List", "➕ Custom Lists", "⚙️ Settings & Formats"]
tabs = st.tabs(tab_names)


def render_list(name):
    s = st.session_state
    base = s["lists"].get(name)
    if base is None or base.empty:
        st.info(f"No {name} data yet. Upload documents and click **Extract Lists**, "
                "or check the detection rules in Settings & Formats.")
        return
    extras = s["extras"].get(name, [])
    df_full = base.copy()
    for e in extras:
        if e not in df_full.columns:
            df_full[e] = ""
    st.caption("✏️ Edit cells, add or delete rows - changes flow into the PDF/Excel output.")
    edited = st.data_editor(df_full, num_rows="dynamic", key=f"ed_{name}_{s['run_id']}_{len(extras)}")

    cc1, cc2 = st.columns([3, 1])
    cc1.text_input("➕ Add a customer-specific blank column (e.g. Vendor, Rating, Remarks 2)",
                   key=f"newcol_{name}")
    cc2.write("")
    cc2.write("")
    cc2.button("Add column", key=f"btn_{name}", on_click=add_extra, args=(name,))

    with st.expander("🛠 Customize format for this customer (select, rename, re-order columns)", expanded=False):
        tdf_in = build_tpl_df(name, list(df_full.columns))
        tdf = st.data_editor(
            tdf_in, disabled=["Field"], hide_index=True,
            key=f"tpl_{name}_{s['tpl_ver']}_{len(extras)}_{s['run_id']}")
    s["tpl_out"][name] = tdf if "tdf" in locals() else build_tpl_df(name, list(df_full.columns))

    fmt, mapdf = active_fmt(name), None
    with st.expander("📐 Customer Excel format - column mapping", expanded=bool(fmt)):
        if not fmt:
            st.caption("No customer format sheet is assigned to this list. Upload / select a customer Excel "
                       "format in the sidebar (section 4) and assign one of its sheets to this list.")
        else:
            book = s["fmtbook"]
            headers = [h for _, h in fmt["cols"]]
            fields = list(df_full.columns)
            opts = SPECIAL_TARGETS + fields
            auto, auto_fx = default_mapping(headers, fields)
            saved = book["maps0"].get(name, {})
            saved_fx = book["fixed0"].get(name, {})
            guess = [saved[h] if saved.get(h) in opts else a for h, a in zip(headers, auto)]
            gfx = [saved_fx.get(h, "") if h in saved else f for h, f in zip(headers, auto_fx)]
            st.success(f"Using sheet '{fmt['sheet']}' of '{book['name']}' (header row {fmt['hrow']}).")
            st.caption("Check the matching; change any 'Map to field' cell. '(blank)' leaves the column empty "
                       "for the customer. 'Fixed value' is written whenever the mapped field is empty "
                       "(e.g. Purchased By = BV). Press 'Save customer format' in the sidebar to remember your corrections.")
            mapdf = st.data_editor(
                pd.DataFrame({"Customer column": headers, "Map to field": guess, "Fixed value": gfx}),
                hide_index=True, disabled=["Customer column"],
                column_config={"Map to field": st.column_config.SelectboxColumn(
                                   "Map to field", options=opts, required=True),
                               "Fixed value": st.column_config.TextColumn("Fixed value (used when field is empty)")},
                key=f"fmtmap2_{name}_{book['name']}_{fmt['sheet']}_{s['run_id']}_{len(extras)}")
            book["maps"][name] = dict(zip(headers, mapdf["Map to field"]))
            book["fixed"][name] = dict(zip(headers, mapdf["Fixed value"].fillna("")))
            st.checkbox("Use this customer format for the output", value=True, key=f"usefmt_{name}")
    use_fmt = bool(fmt and mapdf is not None and s.get(f"usefmt_{name}", True))
    grp = None
    if use_fmt and fmt.get("group_col") and "P&ID Number" in edited.columns:   # form groups valves per P&ID
        edited = edited.assign(_k=edited["P&ID Number"].map(natural_str)).sort_values("_k", kind="stable") \
            .drop(columns="_k").reset_index(drop=True)
        grp = [str(x) for x in edited["P&ID Number"]]
    s["final_grp"][name] = grp
    out = apply_format(edited, fmt, mapdf) if use_fmt else apply_template(edited, s["tpl_out"][name], s["sty_sno"])
    s["final"][name] = out
    s["final_fmt"][name] = use_fmt
    st.markdown(f"**Output preview - {len(out)} rows**")
    st.dataframe(out)

    b1, b2, b3 = st.columns(3)
    b1.download_button("⬇ CSV", out.to_csv(index=False).encode("utf-8-sig"), f"{name.replace(' ', '_')}.csv")
    if use_fmt and fmt["kind"] == "xlsx":
        xbytes = fill_template_xlsx(fmt, out, name, grp)   # customer's own workbook, filled in
    else:
        xbytes = to_excel_bytes({name: out})
    b2.download_button("⬇ Excel", xbytes, f"{name.replace(' ', '_')}.xlsx")
    if b3.button("📄 Generate PDF", key=f"pdfbtn_{name}"):
        s["pdf_bytes"][name] = build_pdf({name: out}, current_meta(), current_style(), s["logo"])
    if name in s["pdf_bytes"]:
        st.download_button("⬇ Download PDF", s["pdf_bytes"][name], f"{name.replace(' ', '_')}.pdf",
                           "application/pdf", key=f"pdfdl_{name}")


for tab, nm in zip(tabs[1:6], LIST_NAMES):
    with tab:
        render_list(nm)

# ---------------- custom lists
with tabs[6]:
    st.subheader("Custom lists")
    st.write("Define your own list with a regular expression in **Settings & Formats → Custom list rules**. "
             "Each named group `(?P<name>...)` becomes a column.")
    customs = [n for n in st.session_state["lists"] if n not in LIST_NAMES]
    if not customs:
        st.info("No custom list extracted yet.")
    else:
        pick = st.selectbox("Custom list", customs)
        render_list(pick)

# ---------------- settings
with tabs[7]:
    st.subheader("Customer format profile")
    st.write("Save everything (column layout, headers, detection rules, PDF style, custom columns) as a "
             "profile and re-use it for the same customer next time.")
    c1, c2 = st.columns(2)
    c1.download_button("💾 Save customer profile (Excel)", make_profile(), "customer_profile.xlsx",
                       "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    c2.file_uploader("Load customer profile (Excel; old JSON also works)", type=["xlsx", "json"], key="profile_upload")
    c2.button("Apply profile", on_click=apply_profile)
    if st.session_state.get("profile_msg"):
        st.info(st.session_state["profile_msg"])
    st.divider()
    st.subheader("Detection rules (regular expressions)")
    st.caption("Adjust to match the customer's tag/line numbering standard, then press Extract Lists again.")
    for k in DEFAULT_PATTERNS:
        st.text_input(f"{k.capitalize()} pattern", key="pat_" + k)
        try:
            re.compile(st.session_state["pat_" + k])
        except re.error as ex:
            st.error(f"{k}: invalid regex - default will be used ({ex})")
    if st.button("Reset rules to default"):
        for k, v in DEFAULT_PATTERNS.items():
            st.session_state["pat_" + k] = v
        st.rerun()
    st.divider()
    st.subheader("P&ID geometry reader (valve size / service / type)")
    st.caption("Used for PDF P&IDs. Size and service are taken from the pipe the valve sits on; entries marked "
               "'nearest - verify' are a proximity guess.")
    st.text_area("Service keywords (comma separated) - labels containing these words are treated as line service",
                 key="pid_svc", height=110)
    st.text_area("Valve code → type override (one per line, e.g. VL=Ball Valve)", key="pid_vmap", height=90)
    st.slider("Ignore right-hand strip of sheet (% width) - title block / notes", 0, 30, key="pid_right")
    st.text_input("Tesseract OCR path (only if not found automatically, e.g. C:\\Program Files\\Tesseract-OCR\\tesseract.exe)",
                  key="pid_tess")
    st.caption("OCR status: " + ("Tesseract found ✔ (" + str(pytesseract.pytesseract.tesseract_cmd) + ")" if tess_ready()
                                 else "Tesseract NOT found - check the path above (file must exist)"))
    if st.session_state.get("prefix_hints"):
        st.info("Tag prefixes found on drawing: " + ", ".join(st.session_state["prefix_hints"]))
    st.divider()
    st.subheader("Custom list rules")
    st.caption('Example: Name = Nozzle List, Pattern = N\\d{1,2}[A-Z]?-(?P<size>\\d+)"  (matched on upper-case text)')
    cr = st.data_editor(pd.DataFrame(st.session_state["custom_rules"] or [{"Name": "", "Pattern": ""}],
                                      columns=["Name", "Pattern"]),
                        num_rows="dynamic", key=f"crules_{st.session_state['tpl_ver']}")
    st.session_state["custom_rules"] = [r for r in cr.fillna("").to_dict("records") if r["Name"] and r["Pattern"]]
    st.divider()
    st.subheader("Notes")
    st.markdown(
        "- **Text-based** PDFs, DXF, Excel, Word are read fully. Scanned drawings need OCR.\n"
        "- **DWG** files must be saved as **DXF** first (AutoCAD *SAVEAS* or ODA File Converter).\n"
        "- *Nearest Line No / Size / Spec* on valves & instruments is a proximity guess - verify against the P&ID.\n"
        "- Quantities count occurrences of the tag in the documents, not physical items.")

# ---------------- dashboard (rendered last so it sees the latest edits)
with tabs[0]:
    s = st.session_state
    L, raw = s["lists"], s["raw"]
    if not L:
        st.info("👈 Upload your P&IDs / specs / drawings in the sidebar and click **Extract Lists**.")
    else:
        cards = [("Files", len({r[0] for r in s["log"]})), ("Lines", len(L.get("Line List", []))),
                 ("Valves", len(L.get("Valve List", []))), ("Equipment", len(L.get("Equipment List", []))),
                 ("Instruments", len(L.get("Instrument List", []))), ("Total Tags", len(L.get("Tag List", [])))]
        cols = st.columns(len(cards))
        for c, (t, n) in zip(cols, cards):
            c.markdown(f"<div class='kpi'><div class='n'>{n}</div><div class='t'>{t}</div></div>",
                       unsafe_allow_html=True)
        st.write("")
        g1, g2 = st.columns(2)
        if "Valve List" in L and "Type" in L["Valve List"]:
            g1.markdown("**Valves by type**")
            g1.bar_chart(L["Valve List"]["Type"].value_counts())
        if "Instrument List" in L and "Measured Variable" in L["Instrument List"]:
            g2.markdown("**Instruments by measured variable**")
            g2.bar_chart(L["Instrument List"]["Measured Variable"].value_counts())
        g3, g4 = st.columns(2)
        if "Equipment List" in L and "Type" in L["Equipment List"]:
            g3.markdown("**Equipment by type**")
            g3.bar_chart(L["Equipment List"]["Type"].value_counts())
        if not raw.empty:
            g4.markdown("**Matches per source file**")
            g4.bar_chart(raw.groupby("Source").size())
        if "Line List" in L and "Piping Spec" in L["Line List"]:
            st.markdown("**Lines by piping spec**")
            st.bar_chart(L["Line List"]["Piping Spec"].value_counts())
        with st.expander("📑 Extraction log"):
            st.dataframe(pd.DataFrame(s["log"], columns=["File", "Type", "Characters read", "Notes"]))

        st.divider()
        st.subheader("📄 Combined PDF output")
        avail = [n for n in s["final"] if n in L]
        chosen = st.multiselect("Lists to include", avail, default=avail)
        if st.button("Generate combined PDF"):
            if chosen:
                s["pdf_bytes"]["__all__"] = build_pdf({n: s["final"][n] for n in chosen}, current_meta(),
                                                       current_style(), s["logo"])
        if "__all__" in s["pdf_bytes"]:
            st.download_button("⬇ Download combined PDF", s["pdf_bytes"]["__all__"], "List_Developer_Output.pdf",
                               "application/pdf")
        if s["fmtbook"] and openpyxl and any(s["final_fmt"].get(n) for n in avail):
            _xb, _n = fill_book_xlsx()
            st.download_button(f"⬇ Customer-format Excel workbook ({_n} sheet(s) filled)", _xb,
                               f"{safe_name(s['fmtbook']['name']) or 'customer'}_lists.xlsx")
        if avail:
            st.download_button("⬇ Download all lists (Excel workbook)",
                               to_excel_bytes({n: s["final"][n] for n in avail}), "List_Developer_All.xlsx")