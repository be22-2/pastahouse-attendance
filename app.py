
import re
import os
import json
import base64
from datetime import datetime, timezone
import requests
from pathlib import Path
import numpy as np
import pandas as pd
import openpyxl
import streamlit as st
import plotly.graph_objects as go
from plotly.subplots import make_subplots

# =========================================================
# Page Config
# =========================================================
st.set_page_config(
    page_title="لوحة الحضور والانضباط | Executive",
    page_icon="📌",
    layout="wide",
    initial_sidebar_state="expanded",
)

# =========================================================
# CSS (Executive Dark + Arabic RTL + Pro Inputs)
# =========================================================
st.markdown(
    """
<style>
:root{
  --bg0:#071A14;
  --card:#06110E;
  --border: rgba(52,211,153,0.22);
  --text:#F9FAFB;
  --muted:#CBD5E1;
  --accent:#34D399;

  --inputBg: rgba(5,10,22,0.78);
  --inputBd: rgba(255,255,255,0.16);
  --popBg:  #071A14;
  --popBd:  rgba(52,211,153,0.22);
}

/* Base background */
html, body, [data-testid="stAppViewContainer"]{
  background: var(--bg0) !important;
}

/* Main background gradient */
[data-testid="stAppViewContainer"] > .main{
  background:
    radial-gradient(1200px 700px at 80% 30%, rgba(30,107,75,0.65), transparent 55%),
    radial-gradient(900px 650px at 30% 70%, rgba(8,30,24,0.85), transparent 60%),
    linear-gradient(120deg, #050A16 0%, var(--bg0) 35%, rgba(30,107,75,0.70) 100%) !important;
}

/* Container spacing */
.block-container{
  padding-top: 1.1rem !important;
  padding-bottom: 1.5rem !important;
  max-width: 1500px;
}

/* Header transparent */
[data-testid="stHeader"]{ background: transparent !important; }

/* RTL (content only) */
[data-testid="stAppViewContainer"] .main{ direction: rtl !important; }
[data-testid="stAppViewContainer"] .block-container{
  direction: rtl !important;
  text-align: right !important;
}
[data-testid="stAppViewContainer"] h1,
[data-testid="stAppViewContainer"] h2,
[data-testid="stAppViewContainer"] h3,
[data-testid="stAppViewContainer"] h4,
[data-testid="stAppViewContainer"] p,
[data-testid="stAppViewContainer"] li,
[data-testid="stAppViewContainer"] label{
  text-align: right !important;
}

/* Exceptions (keep LTR) */
[data-testid="stHeader"],
[data-testid="stToolbar"],
[data-testid="collapsedControl"]{ direction:ltr !important; }
[data-testid="collapsedControl"] *{ direction:ltr !important; }

/* Plotly LTR */
.js-plotly-plot, .js-plotly-plot *{
  direction: ltr !important;
  text-align: left !important;
}

/* Sidebar (LEFT ثابت) */
section[data-testid="stSidebar"]{
  background: linear-gradient(180deg, #050A16 0%, #06110E 100%) !important;
  left: 0 !important;
  right: auto !important;
  border-right: 1px solid rgba(255,255,255,0.06) !important;
  border-left: 0 !important;
}
section[data-testid="stSidebar"][aria-expanded="false"]{
  width: 0 !important; min-width: 0 !important;
}
section[data-testid="stSidebar"][aria-expanded="false"] > div{ display:none !important; }
[data-testid="stAppViewContainer"] .main{ margin:0 !important; }

/* Text */
h1, h2, h3, h4, p, li, label, span, div{ color: var(--text) !important; }
.stCaption{ color: var(--muted) !important; }
hr{ border-color: rgba(255,255,255,0.10) !important; }

div[data-testid="stHorizontalBlock"]{ gap: 1.10rem !important; }
div[data-testid="column"]{ padding-top: 0.35rem !important; padding-bottom: 0.35rem !important; }

/* KPI Card */
.kpi-card{
  background:
    radial-gradient(900px 520px at 85% 20%, rgba(52,211,153,0.18), transparent 48%),
    radial-gradient(900px 520px at 15% 80%, rgba(34,197,94,0.10), transparent 55%),
    linear-gradient(145deg, rgba(5,10,22,0.90) 0%, rgba(6,17,14,0.92) 60%, rgba(10,40,30,0.95) 100%);
  border: 1px solid var(--border);
  border-radius: 16px;
  padding: 14px;
  box-shadow: 0 10px 26px rgba(0,0,0,0.45);
  position: relative;
  overflow: hidden;

  height: 110px;
  display: flex;
  flex-direction: column;
  justify-content: center;
  gap: 6px;

  transition: transform 160ms ease, box-shadow 160ms ease, border-color 160ms ease, filter 160ms ease;
  will-change: transform;
  cursor: default;
}
.kpi-card:hover{
  transform: translateY(-5px) scale(1.02);
  box-shadow: 0 18px 46px rgba(0,0,0,0.60);
  border-color: rgba(52,211,153,0.55);
  filter: brightness(1.05);
}
.kpi-card:before{
  content:"";
  position:absolute;
  inset:0;
  border-radius:16px;
  padding:1px;
  background: linear-gradient(135deg,
    rgba(52,211,153,0.60),
    rgba(34,197,94,0.20),
    rgba(255,255,255,0.10)
  );
  -webkit-mask: linear-gradient(#000 0 0) content-box, linear-gradient(#000 0 0);
  -webkit-mask-composite: xor;
  mask-composite: exclude;
  pointer-events:none;
}
.kpi-label{
  color: rgba(249,250,251,0.92) !important;
  font-size: 0.92rem;
  line-height: 1.2;
  margin: 0;
}
.kpi-value{
  color:#FFFFFF !important;
  font-size: clamp(1.20rem, 1.55vw, 1.65rem);
  font-weight: 900;
  line-height: 1.05;

  white-space: nowrap;
  overflow: hidden;
  text-overflow: ellipsis;
  direction: ltr !important;
  text-align: left !important;
}
.kpi-sub{
  margin: 0;
  font-size: 0.82rem;
  color: rgba(203,213,225,0.95) !important;
}

/* Dataframes */
[data-testid="stDataFrame"]{ direction: rtl; }

/* Inputs (Select/Date) */
div[data-baseweb="select"] > div {
    background-color: rgba(255, 255, 255, 0.03) !important;
    border: 1px solid rgba(52, 211, 153, 0.15) !important;
    border-radius: 12px !important;
    padding: 2px 8px !important;
    transition: all 0.3s ease !important;
    box-shadow: inset 0 1px 1px rgba(255,255,255,0.05) !important;
}
div[data-baseweb="select"] > div:hover {
    border-color: var(--accent) !important;
    background-color: rgba(52, 211, 153, 0.05) !important;
}
div[data-baseweb="popover"] > div{
    background-color: #0c1e19 !important;
    border: 1px solid rgba(52, 211, 153, 0.2) !important;
    border-radius: 14px !important;
    box-shadow: 0 15px 35px rgba(0, 0, 0, 0.6) !important;
    padding: 8px !important;
    margin-top: 4px !important;
}
li[role="option"] {
    border-radius: 8px !important;
    margin-bottom: 3px !important;
    padding: 12px 16px !important;
    color: var(--text) !important;
}
li[role="option"]:hover {
    background-color: rgba(52, 211, 153, 0.12) !important;
    color: var(--accent) !important;
    padding-right: 20px !important;
}
li[aria-selected="true"] {
    background-color: var(--accent) !important;
    color: #071A14 !important;
    font-weight: 700 !important;
}
[data-testid="stFileUploaderDropzone"]{
  background: rgba(5,10,22,0.70) !important;
  border: 1px dashed rgba(52,211,153,0.38) !important;
  border-radius: 14px !important;
}
[data-testid="stFileUploaderDropzone"] *{ color: var(--text) !important; }
</style>
""",
    unsafe_allow_html=True,
)

# =========================================================
# Plotly Theme
# =========================================================
COLORS = {
    "green": "#34D399",
    "cyan":  "#22D3EE",
    "orange":"#FB923C",
    "red":   "#F87171",
    "purple":"#A78BFA",
    "grid":  "rgba(255,255,255,0.08)",
}

PLOTLY_LAYOUT = dict(
    template="plotly_dark",
    paper_bgcolor="rgba(0,0,0,0)",
    plot_bgcolor="rgba(0,0,0,0)",
    font=dict(color="#E5E7EB"),
    margin=dict(l=10, r=10, t=60, b=10),
)

def apply_plotly_theme(fig, height=380, title=None, showlegend=True, hovermode="x unified"):
    fig.update_layout(**PLOTLY_LAYOUT, height=height, showlegend=showlegend, hovermode=hovermode)
    if title:
        fig.update_layout(title=dict(text=title, x=1.0, xanchor="right", font=dict(size=20, color="#FFFFFF")))
    fig.update_yaxes(showgrid=True, gridcolor=COLORS["grid"], zeroline=False)
    fig.update_xaxes(showgrid=False)
    fig.update_layout(
        hoverlabel=dict(
            bgcolor="rgba(2, 6, 23, 0.95)",
            bordercolor="rgba(52,211,153,0.45)",
            font=dict(color="#FFFFFF", size=14),
            align="left",
            namelength=-1
        ),
        legend=dict(
            bgcolor="rgba(2, 6, 23, 0.55)",
            bordercolor="rgba(255,255,255,0.14)",
            borderwidth=1,
            font=dict(size=13, color="#FFFFFF"),
            orientation="h",
            yanchor="bottom",
            y=1.02,
            xanchor="right",
            x=1
        )
    )
    return fig

def kpi_card(label: str, value: str, sub: str = ""):
    sub_html = f'<div class="kpi-sub">{sub}</div>' if sub else ""
    st.markdown(
        f"""
        <div class="kpi-card">
          <div class="kpi-label">{label}</div>
          <div class="kpi-value">{value}</div>
          {sub_html}
        </div>
        """,
        unsafe_allow_html=True,
    )

# =========================================================
# Helpers (Loader من شغلك + Fix الاسم)
# =========================================================
AR_NUM_MAP = str.maketrans("٠١٢٣٤٥٦٧٨٩", "0123456789")
# =========================================================
# Arabic Names Map (by employee number)
# =========================================================
ARABIC_NAME_MAP = {
    "1": "حورمت علي",
    "3": "ناصر باخشوين",
    "4": "محمد دماج",
    "6": "شوباش",
    "7": "محمد سانولا انصاري",
    "8": "صالح باياسين",
    "9": "ايمن كامل",
    "10": "شاسان",
    "11": "محمد احمد سيف",
    "13": "محمد نور الدين",
    "14": "نياز عالم",
    "15": "مشبال علام",
    "16": "براديب",
    "17": "جبل حسين",
    "19": "راجان",
    "20": "معين الدين",
    "21": "شير رام",
    "22": "طارق اسلام",
    "23": "عثمان فردوس",
    "24": "مد رقيب",
    "25": "جانجير حسين",
    "26": "حسين مفلح",
    "27": "عبدالله محمد",
    "28": "أحمد بديان",
    "29": "أحمد أنصاري",
    "30": "سالم باحمدين",
    "31": "بدر",
    "32": "محسن ابو زيدان",
    "33": "جبل",
    "35": "عبدو",
    "37": "إفح",
    "47": "سفيان",
    "48": "أحمد الصالحي",
    "49": "حسن",
    "50": "حمود",
    "51": "نديم",
    "52": "مهيوب",
}


def _norm_txt(x):
    if x is None:
        return ""
    s = str(x).strip()
    return s.replace("\u200f", "").replace("\u200e", "").replace("：", ":")

def _to_ascii_digits(s: str) -> str:
    return _norm_txt(s).translate(AR_NUM_MAP)

def _parse_hhmm_to_minutes(v):
    s = _to_ascii_digits(v)
    if not s or s.lower() in {"nan", "none"}:
        return 0
    m = re.match(r"^(\d{1,2}):(\d{2})$", s)
    if not m:
        return 0
    return int(m.group(1)) * 60 + int(m.group(2))

def _parse_date(v):
    s = _to_ascii_digits(v).replace(".", "/").replace("-", "/")
    for fmt in ("%d/%m/%y", "%d/%m/%Y"):
        try:
            return pd.to_datetime(s, format=fmt).date()
        except Exception:
            pass
    try:
        return pd.to_datetime(s, dayfirst=True).date()
    except Exception:
        return None

def _build_merge_lookup(ws):
    lookup = {}
    for r in ws.merged_cells.ranges:
        anchor = (r.min_row, r.min_col)
        for rr in range(r.min_row, r.max_row + 1):
            for cc in range(r.min_col, r.max_col + 1):
                lookup[(rr, cc)] = anchor
    return lookup

def _get_cell_value(ws, merge_lookup, r, c):
    anchor = merge_lookup.get((r, c), (r, c))
    return ws.cell(anchor[0], anchor[1]).value

def _looks_like_name(s: str) -> bool:
    s = _norm_txt(s)
    if not s:
        return False
    if re.search(r"(?:ال|الإ)?اسم\s*(?:الكامل|الموظف)|رقم\s*الموظف", s):
        return False
    if not re.search(r"[\u0600-\u06FF]", s):
        return False
    if re.fullmatch(r"\d+|\d{1,2}:\d{2}|\d{4}-\d{2}-\d{2}.*", _to_ascii_digits(s)):
        return False
    return len(s) >= 2

def _best_name_from_row(row, pivot_idx=None):
    candidates = []
    for idx, cell in enumerate(row):
        t = _norm_txt(cell)
        if _looks_like_name(t):
            dist = abs(idx - pivot_idx) if pivot_idx is not None else 999
            score = (1000 - dist) + min(len(t), 30)
            candidates.append((score, idx, t))
    if candidates:
        candidates.sort(reverse=True)
        return candidates[0][2].strip()
    texts = [_norm_txt(x) for x in row if _norm_txt(x)]
    texts = [t for t in texts if re.search(r"[\u0600-\u06FF]", t) and not re.search(r"(?:ال|الإ)?اسم\s*(?:الكامل|الموظف)|رقم\s*الموظف", t)]
    if texts:
        return max(texts, key=len).strip()
    return ""

def _best_empid_from_row(row):
    for idx, cell in enumerate(row):
        t = _to_ascii_digits(_norm_txt(cell))
        if "رقم الموظف" in t:
            m = re.search(r"رقم\s*الموظف\s*[:：]?\s*([0-9]+)", t)
            if m:
                return m.group(1)
            for k in range(idx + 1, min(idx + 8, len(row))):
                t2 = _to_ascii_digits(_norm_txt(row[k]))
                m2 = re.search(r"\b([0-9]{1,10})\b", t2)
                if m2:
                    return m2.group(1)
    return None

def normalize_emp_id(x):
    if pd.isna(x):
        return ""

    s = str(x).translate(AR_NUM_MAP)
    s = s.replace("\u200f", "").replace("\u200e", "").strip()
    s = re.sub(r"\.0+$", "", s)       # 48.0 -> 48
    s = re.sub(r"[^\d]", "", s)       # يحتفظ بالأرقام فقط

    return s

@st.cache_data(show_spinner=False)
def load_attendance_excel(uploaded_file, sheet_name=None):
    # --------------------------------------------
    # Read tidy Excel (XLSX preferred)
    # NOTE: .xls requires xlrd OR convert to xlsx
    # --------------------------------------------
    if sheet_name is None:
        df_raw = pd.read_excel(uploaded_file)  # default: أول شيت
    else:
        df_raw = pd.read_excel(uploaded_file, sheet_name=sheet_name)
    # Clean column names
    df_raw.columns = [
        str(c).strip().replace("\u200f", "").replace("\u200e", "")
        for c in df_raw.columns
    ]

    # Support "الاسم" vs "الإسم"
    if "الإسم" not in df_raw.columns and "الاسم" in df_raw.columns:
        df_raw = df_raw.rename(columns={"الاسم": "الإسم"})

    # Required columns for your dashboard
    required = ["رقم الموظف", "الإسم", "التاريخ"]
    missing = [c for c in required if c not in df_raw.columns]
    if missing:
        raise ValueError(f"الملف ناقص أعمدة أساسية: {missing}")

    df = df_raw.copy()

    # تنظيف الاسم الإنجليزي الأصلي
    df["اسم الموظف"] = df["الإسم"].astype(str).str.strip()

    # تنظيف رقم الموظف
    df["رقم الموظف"] = df["رقم الموظف"].apply(normalize_emp_id)

    # اسم عربي من القاموس حسب الرقم
    df["اسم الموظف عربي"] = df["رقم الموظف"].map(ARABIC_NAME_MAP)

    # ---------------------------------------------------------
    # fallback بالأسماء الإنجليزية بعد تنظيفها
    # ---------------------------------------------------------
    def normalize_eng_name(x):
        if pd.isna(x):
            return ""
        s = str(x).strip().lower()
        s = s.replace("\u200f", "").replace("\u200e", "")
        s = re.sub(r"[^a-z]", "", s)   # يشيل الفواصل والمسافات والرموز
        return s

    fallback_map = {
        "essalhiahmed": "أحمد الصالحي",
        "hassan": "حسن",
        "hamood": "حمود",
        "soufiane": "سفيان",
        "afh": "إفح",
        "abdullahmohammed": "عبدالله محمد",
        "ahmedbadyan": "أحمد بديان",
        "gabal": "جبل",
        "hussbainmuflih": "حسين مفلح",
        "mahyob": "مهيوب",
        "nadeem": "نديم",
        "abdo": "عبدو",

        # الأسماء اللي ظاهرة عندك في الصور
        "mohmmedali": "محمد علي",
        "mzeeshan": "محمد زيشان",
        "nage": "ماغينو",
        "saif": "محمد أحمد سيف",
    }

    df["اسم انجليزي منظم"] = df["اسم الموظف"].apply(normalize_eng_name)
    df["اسم fallback عربي"] = df["اسم انجليزي منظم"].map(fallback_map)

    df["اسم الموظف عرض"] = np.where(
        df["اسم الموظف عربي"].notna() & (df["اسم الموظف عربي"].astype(str).str.strip() != ""),
        df["اسم الموظف عربي"],
        np.where(
            df["اسم fallback عربي"].notna() & (df["اسم fallback عربي"].astype(str).str.strip() != ""),
            df["اسم fallback عربي"],
            df["اسم الموظف"]
        )
    )

    # Date
    df["التاريخ"] = pd.to_datetime(df["التاريخ"], errors="coerce", dayfirst=True)
    df = df[df["التاريخ"].notna()].copy()

    # Helper: convert HH:MM to minutes (also accepts numbers)
    def hhmm_to_minutes(x):
        if pd.isna(x):
            return 0
        s = str(x).strip().translate(AR_NUM_MAP).replace(".", ":")
        m = re.match(r"^(\d{1,2}):(\d{2})$", s)
        if m:
            return int(m.group(1)) * 60 + int(m.group(2))
        try:
            v = float(s)
            return int(v) if v > 0 else 0
        except Exception:
            return 0

    # Late minutes
    if "تأخير" in df.columns:
        df["دقائق التأخير"] = df["تأخير"].apply(hhmm_to_minutes)
    else:
        df["دقائق التأخير"] = 0

    # Early leave minutes
    if "إنصراف مبكر" in df.columns:
        df["دقائق الانصراف المبكر"] = df["إنصراف مبكر"].apply(hhmm_to_minutes)
    else:
        df["دقائق الانصراف المبكر"] = 0

    # Absence -> 0/1
    # ✅ الغياب عندك:
    # 1) عمود "غياب" قيمته true
    # 2) أو عمود "الحضور الفعلي" فاضي

    def _is_true(v):
        if pd.isna(v):
            return False
        s = str(v).strip().lower()
        return s in {"true", "1", "yes", "y", "t", "صحيح"}

    def _is_blank(v):
        if pd.isna(v):
            return True
        s = str(v).strip()
        return s == "" or s.lower() in {"nan", "none"}

    abs1 = pd.Series([0] * len(df), index=df.index)  # من عمود غياب
    abs2 = pd.Series([0] * len(df), index=df.index)  # من الحضور الفعلي الفاضي

    # (1) from "غياب"
    if "غياب" in df.columns:
        abs1 = df["غياب"].apply(_is_true).astype(int)

    # (2) from "الحضور الفعلي" blank
    # اسم العمود عندك غالبًا "الحضور الفعلي" — لو مختلف غيّره هنا
    if "الحضور الفعلي" in df.columns:
        abs2 = df["الحضور الفعلي"].apply(_is_blank).astype(int)

    # ---------------------------------------------------------
    # Leave detection (approved leave must not be counted as absence)
    # Supports Arabic/English labels across common status columns.
    # ---------------------------------------------------------
    leave_keywords = {
        "إجازة", "اجازة", "إجازه", "اجازه", "إجازة سنوية", "اجازة سنوية",
        "إجازة مرضية", "اجازة مرضية", "مرضي", "مرضية", "مأمورية", "مهمة عمل",
        "leave", "annual leave", "vacation", "sick leave", "medical leave",
        "business trip", "official leave", "approved leave"
    }

    def _normalize_status(v):
        if pd.isna(v):
            return ""
        s = str(v).strip().lower().translate(AR_NUM_MAP)
        s = s.replace("أ", "ا").replace("إ", "ا").replace("آ", "ا").replace("ة", "ه")
        s = re.sub(r"\s+", " ", s)
        return s

    normalized_leave_keywords = {_normalize_status(x) for x in leave_keywords}

    # Search only likely status/type/notes columns to avoid false positives in names.
    likely_leave_columns = [
        c for c in df.columns
        if any(k in str(c).lower() for k in [
            "حالة", "الحاله", "نوع", "ملاح", "سبب", "دوام", "اجاز", "إجاز",
            "استثنا", "إستثنا", "استثناء", "إستثناء",
            "status", "type", "remark", "note", "reason", "leave", "exception"
        ])
    ]

    leave_mask = pd.Series(False, index=df.index)
    for col in likely_leave_columns:
        values = df[col].apply(_normalize_status)
        col_mask = values.apply(
            lambda x: any(k and (x == k or k in x) for k in normalized_leave_keywords)
        )
        leave_mask = leave_mask | col_mask

    # Dedicated boolean leave columns, when exported by the attendance system.
    for leave_col in ["إجازة", "اجازة", "إجازه", "اجازه", "Leave", "On Leave"]:
        if leave_col in df.columns:
            leave_mask = leave_mask | df[leave_col].apply(_is_true)

    df["إجازة"] = leave_mask.astype(int)

    # Merge absence conditions, then explicitly exclude approved leave.
    df["غياب"] = (((abs1 == 1) | (abs2 == 1)) & (~leave_mask)).astype(int)

    # Attendance is neither absent nor on approved leave.
    df["حضور"] = ((df["غياب"] == 0) & (df["إجازة"] == 0)).astype(int)
    df["يوم عمل محسوب"] = (df["إجازة"] == 0).astype(int)

    # Late flag (leave days cannot be late)
    df["تأخر"] = ((df["دقائق التأخير"] > 0) & (df["إجازة"] == 0)).astype(int)
    df.loc[df["إجازة"] == 1, ["دقائق التأخير", "دقائق الانصراف المبكر"]] = 0

    # Weekday + Month (same as your dashboard)
    df["اليوم_الاسبوعي"] = df["التاريخ"].dt.day_name()
    day_map = {
        "Sunday":"الأحد","Monday":"الإثنين","Tuesday":"الثلاثاء","Wednesday":"الأربعاء",
        "Thursday":"الخميس","Friday":"الجمعة","Saturday":"السبت"
    }
    df["اليوم_الاسبوعي"] = df["اليوم_الاسبوعي"].map(day_map).fillna(df["اليوم_الاسبوعي"])
    df["شهر"] = df["التاريخ"].dt.to_period("M").astype(str)

    return df

# =========================================================
# Login & Roles
# =========================================================
# في Streamlit Cloud ضعي القيم الحقيقية داخل Secrets بدل تغييرها هنا.
# القيم المؤقتة أدناه مخصصة للتجربة المحلية فقط.
DEFAULT_USERS = {
    "bayan": {"password": "Bayan@2026", "role": "owner", "display_name": "بيان"},
    "manager": {"password": "Manager@2026", "role": "viewer", "display_name": "المدير"},
}


def get_users_config():
    """قراءة المستخدمين من st.secrets عند توفرها، وإلا استخدام حسابات التجربة."""
    try:
        configured = st.secrets.get("users", {})
        if configured:
            users = {}
            for username, info in configured.items():
                users[str(username)] = {
                    "password": str(info.get("password", "")),
                    "role": str(info.get("role", "viewer")),
                    "display_name": str(info.get("display_name", username)),
                }
            return users, False
    except Exception:
        pass
    return DEFAULT_USERS, True


def render_login():
    users, using_defaults = get_users_config()

    if st.session_state.get("authenticated"):
        return users, using_defaults

    st.markdown(
        """
        <div style="max-width:520px;margin:8vh auto 1rem auto;padding:28px;
                    border:1px solid rgba(52,211,153,.28);border-radius:18px;
                    background:rgba(5,10,22,.82);box-shadow:0 18px 50px rgba(0,0,0,.45)">
          <div style="font-size:30px;font-weight:900;text-align:center">تسجيل الدخول</div>
          <div style="margin-top:8px;color:#CBD5E1!important;text-align:center">
            لوحة الحضور والانضباط
          </div>
        </div>
        """,
        unsafe_allow_html=True,
    )

    _, form_col, _ = st.columns([1, 1.25, 1])
    with form_col:
        with st.form("login_form", clear_on_submit=False):
            username = st.text_input("اسم المستخدم")
            password = st.text_input("كلمة المرور", type="password")
            submitted = st.form_submit_button("دخول", use_container_width=True)

        if submitted:
            account = users.get(username.strip())
            if account and password == account.get("password"):
                st.session_state["authenticated"] = True
                st.session_state["username"] = username.strip()
                st.session_state["role"] = account.get("role", "viewer")
                st.session_state["display_name"] = account.get("display_name", username.strip())
                st.rerun()
            else:
                st.error("اسم المستخدم أو كلمة المرور غير صحيحة.")

        if using_defaults:
            st.caption("تنبيه: حسابات التجربة المؤقتة مفعلة. غيّريها من Streamlit Secrets قبل مشاركة الرابط.")
    st.stop()


USERS_CONFIG, USING_DEFAULT_ACCOUNTS = render_login()
CURRENT_ROLE = st.session_state.get("role", "viewer")
IS_OWNER = CURRENT_ROLE == "owner"

# ملفات البصمة التي يتم حفظها تلقائيًا داخل مستودع GitHub.
ATTENDANCE_FILES = [
    Path(__file__).with_name("attendance_current.xls"),
    Path(__file__).with_name("attendance_current.xlsx"),
]
ATTENDANCE_META_FILE = Path(__file__).with_name("attendance_meta.json")

def find_saved_attendance_file():
    for path in ATTENDANCE_FILES:
        if path.exists():
            return path
    return None

def github_settings():
    try:
        cfg = st.secrets.get("github", {})
        token = str(cfg.get("token", "")).strip()
        repo = str(cfg.get("repo", "")).strip()
        branch = str(cfg.get("branch", "main")).strip() or "main"
        if token and repo:
            return token, repo, branch
    except Exception:
        pass
    return None

def _github_headers(token):
    return {
        "Authorization": f"Bearer {token}",
        "Accept": "application/vnd.github+json",
        "X-GitHub-Api-Version": "2022-11-28",
    }

def _github_get_sha(token, repo, branch, path):
    url = f"https://api.github.com/repos/{repo}/contents/{path}"
    response = requests.get(url, headers=_github_headers(token), params={"ref": branch}, timeout=30)
    if response.status_code == 404:
        return None
    response.raise_for_status()
    return response.json().get("sha")

def _github_put_file(token, repo, branch, path, content_bytes, message):
    url = f"https://api.github.com/repos/{repo}/contents/{path}"
    payload = {
        "message": message,
        "content": base64.b64encode(content_bytes).decode("ascii"),
        "branch": branch,
    }
    sha = _github_get_sha(token, repo, branch, path)
    if sha:
        payload["sha"] = sha
    response = requests.put(url, headers=_github_headers(token), json=payload, timeout=60)
    response.raise_for_status()

def _github_delete_file(token, repo, branch, path, message):
    sha = _github_get_sha(token, repo, branch, path)
    if not sha:
        return
    url = f"https://api.github.com/repos/{repo}/contents/{path}"
    response = requests.delete(
        url,
        headers=_github_headers(token),
        json={"message": message, "sha": sha, "branch": branch},
        timeout=30,
    )
    response.raise_for_status()

def save_attendance_to_github(uploaded_file):
    settings = github_settings()
    if not settings:
        raise RuntimeError("إعداد GitHub غير موجود في Secrets")
    token, repo, branch = settings
    suffix = Path(uploaded_file.name).suffix.lower()
    if suffix not in {".xls", ".xlsx"}:
        raise ValueError("الملف يجب أن يكون Excel بصيغة xls أو xlsx")
    target = f"attendance_current{suffix}"
    other = "attendance_current.xlsx" if suffix == ".xls" else "attendance_current.xls"
    content = uploaded_file.getvalue()
    now = datetime.now(timezone.utc).astimezone().strftime("%Y-%m-%d %H:%M")
    _github_put_file(token, repo, branch, target, content, f"Update attendance file - {now}")
    _github_delete_file(token, repo, branch, other, f"Remove old attendance file - {now}")
    meta = {"original_name": uploaded_file.name, "updated_at": now, "updated_by": st.session_state.get("display_name", "بيان")}
    _github_put_file(token, repo, branch, "attendance_meta.json", json.dumps(meta, ensure_ascii=False, indent=2).encode("utf-8"), f"Update attendance metadata - {now}")
    return target

def read_attendance_meta():
    try:
        if ATTENDANCE_META_FILE.exists():
            return json.loads(ATTENDANCE_META_FILE.read_text(encoding="utf-8"))
    except Exception:
        pass
    return {}

# =========================================================
# Header
# =========================================================
st.markdown(
    """
    <div style="line-height:1.15; text-align:right;">
      <div style="font-size:34px; font-weight:900; color:#E5E7EB;">لوحة الحضور والانضباط</div>
      <div style="margin-top:6px; color:#94A3B8; font-size:14px;">
        
      </div>
    </div>
    """,
    unsafe_allow_html=True,
)
st.markdown("---")

# =========================================================
# Data Source + Filters
# =========================================================
data_source = find_saved_attendance_file()
meta = read_attendance_meta()

if IS_OWNER:
    st.subheader("تحديث بيانات الشهر")
    st.caption("ارفعي ملف البصمة الجديد ثم اضغطي حفظ. بعد الحفظ يراه المدير تلقائيًا من نفس الرابط.")
    uploaded = st.file_uploader("اختاري ملف Excel الخاص بالبصمة", type=["xlsx", "xls"], key="monthly_attendance_upload")
    if uploaded is not None:
        col_save, col_preview = st.columns([1, 2])
        with col_save:
            if st.button("حفظ وتحديث الداشبورد", type="primary", use_container_width=True):
                try:
                    with st.spinner("جاري حفظ ملف الشهر..."):
                        save_attendance_to_github(uploaded)
                    st.success("تم حفظ ملف الشهر بنجاح. سيُحدّث التطبيق تلقائيًا خلال دقيقة أو دقيقتين.")
                    st.info("بعد التحديث، المدير يفتح نفس الرابط ويشاهد البيانات مباشرة.")
                except Exception as exc:
                    st.error(f"تعذر حفظ الملف: {exc}")
        with col_preview:
            st.caption("يمكنكِ معاينة الملف الآن قبل اكتمال تحديث الموقع.")
            data_source = uploaded
    elif data_source is not None:
        st.success("البيانات المحفوظة جاهزة للعرض.")
else:
    if data_source is None:
        st.info("لا توجد بيانات شهر محفوظة حتى الآن. ستظهر هنا تلقائيًا بعد أن ترفع المسؤولة ملف البصمة.")
        st.stop()

if meta:
    st.caption(f"آخر تحديث: {meta.get('updated_at', 'غير محدد')} — الملف: {meta.get('original_name', '')}")

if data_source is None:
    st.info("ارفعي ملف البصمة ثم اضغطي زر حفظ وتحديث الداشبورد.")
    st.stop()

try:
    df = load_attendance_excel(data_source)
except Exception as exc:
    st.error(f"تعذر قراءة ملف البصمة: {exc}")
    st.stop()

# فحص الأسماء بعد التحويل
name_check = (
    df[[
        "رقم الموظف",
        "اسم الموظف",
        "اسم انجليزي منظم",
        "اسم الموظف عربي",
        "اسم fallback عربي",
        "اسم الموظف عرض"
    ]]
    .drop_duplicates()
    .sort_values("رقم الموظف")
)

with st.sidebar:
    st.markdown(f"**مرحبًا، {st.session_state.get('display_name', '')}**")
    st.caption("صلاحية كاملة" if IS_OWNER else "مشاهدة فقط")
    if st.button("تسجيل الخروج", use_container_width=True):
        for key in ["authenticated", "username", "role", "display_name"]:
            st.session_state.pop(key, None)
        st.rerun()
    st.markdown("---")
    st.success(" جاهز للتحليل")
    st.caption("Executive Attendance Analytics")
    st.markdown("---")
    st.subheader(" الفلاتر")

    emps = sorted(df["اسم الموظف عرض"].dropna().astype(str).str.strip().unique().tolist())
    emp_options = ["الكل"] + emps

    selected_emps_raw = st.multiselect(
        "اختر الموظفين",
        options=emp_options,
        default=["الكل"],
        placeholder="ابحث واختر الموظفين"
    )

    if "الكل" in selected_emps_raw:
        selected_emps = emps
    else:
        selected_emps = selected_emps_raw

    min_date = df["التاريخ"].min().date()
    max_date = df["التاريخ"].max().date()
    date_range = st.date_input("نطاق التاريخ", (min_date, max_date))

    show_raw = st.toggle("عرض البيانات الخام", value=False)

if not selected_emps:
    st.warning("اختر موظف واحد على الأقل.")
    st.stop()

dff = df[df["اسم الموظف عرض"].isin(selected_emps)].copy()

if isinstance(date_range, (list, tuple)) and len(date_range) == 2:
    dff = dff[(dff["التاريخ"].dt.date >= date_range[0]) & (dff["التاريخ"].dt.date <= date_range[1])]

if len(dff) == 0:
    st.warning("ما فيه بيانات بعد الفلاتر الحالية.")
    st.stop()

# =========================================================
# KPIs (Executive)
# =========================================================
total_rows = len(dff)
absent_days = int(dff["غياب"].sum())
leave_days = int(dff["إجازة"].sum())
present_days = int(dff["حضور"].sum())
counted_workdays = int(dff["يوم عمل محسوب"].sum())
late_days = int(dff["تأخر"].sum())
late_minutes = int(dff["دقائق التأخير"].sum())
early_minutes = int(dff["دقائق الانصراف المبكر"].sum())
attendance_rate = 0 if counted_workdays == 0 else (present_days / counted_workdays) * 100

k1, k2, k3, k4, k5 = st.columns(5, gap="small")
with k1: kpi_card("نسبة الحضور", f"{attendance_rate:.1f}%")
with k2: kpi_card("مجموع دقائق التأخير", f"{late_minutes:,}")
with k3: kpi_card("أيام التأخير", f"{late_days:,}")
with k4: kpi_card("أيام الغياب", f"{absent_days:,}")
with k5: kpi_card("أيام الإجازة", f"{leave_days:,}")
#with k5: kpi_card("عدد السجلات", f"{total_rows:,}")

st.markdown("---")

# =========================================================
# Aggregations
# =========================================================

# =========================================================
# Department Analytics
# =========================================================

if "الإداره" in dff.columns:

    dept_stats = (
        dff.groupby("الإداره", as_index=False)
        .agg(
            عدد_السجلات=("رقم الموظف","count"),
            أيام_العمل=("يوم عمل محسوب","sum"),
            الحضور=("حضور","sum"),
            الغياب=("غياب","sum"),
            الإجازات=("إجازة","sum"),
            التأخير_أيام=("تأخر","sum"),
            دقائق_التأخير=("دقائق التأخير","sum")
        )
    )

    dept_stats["نسبة_الحضور"] = (
        dept_stats["الحضور"] / dept_stats["أيام_العمل"].replace(0, np.nan)
    ).mul(100).fillna(0)

else:
    dept_stats = pd.DataFrame()


per_emp = (
    dff.groupby(["رقم الموظف", "اسم الموظف عرض"], as_index=False)
    .agg(
        أيام_الحضور=("حضور", "sum"),
        أيام_الغياب=("غياب", "sum"),
        أيام_الإجازة=("إجازة", "sum"),
        أيام_العمل=("يوم عمل محسوب", "sum"),
        أيام_التأخير=("تأخر", "sum"),
        دقائق_التأخير=("دقائق التأخير", "sum"),
        دقائق_انصراف_مبكر=("دقائق الانصراف المبكر", "sum"),
        سجلات=("التاريخ", "count"),
    )
)

# نرجع اسم العمود للاسم القديم حتى يشتغل باقي الكود بسهولة
per_emp = per_emp.rename(columns={"اسم الموظف عرض": "اسم الموظف"})

per_emp["نسبة_الحضور"] = (per_emp["أيام_الحضور"] / per_emp["أيام_العمل"].replace(0, np.nan) * 100).fillna(0)
per_emp["متوسط_تأخير_اليوم"] = np.where(
    per_emp["أيام_التأخير"] > 0,
    per_emp["دقائق_التأخير"] / per_emp["أيام_التأخير"],
    0
)
# =========================================================
# Discipline Score
# =========================================================

per_emp["Discipline_Score"] = (
    100
    - (per_emp["دقائق_التأخير"] / 10)
    - (per_emp["أيام_الغياب"] * 5)
)

per_emp["Discipline_Score"] = per_emp["Discipline_Score"].clip(lower=0)
daily = (
    dff.groupby("التاريخ", as_index=False)
    .agg(
        سجلات=("رقم الموظف","count"),
        أيام_العمل=("يوم عمل محسوب","sum"),
        حضور=("حضور","sum"),
        إجازات=("إجازة","sum"),
        غياب=("غياب","sum"),
        تأخير_دقائق=("دقائق التأخير","sum"),
        تأخير_أيام=("تأخر","sum"),
    )
    .sort_values("التاريخ")
)
daily["نسبة_الحضور"] = (daily["حضور"] / daily["أيام_العمل"].replace(0, np.nan) * 100).fillna(0)

monthly = (
    dff.groupby("شهر", as_index=False)
    .agg(
        سجلات=("رقم الموظف","count"),
        أيام_العمل=("يوم عمل محسوب","sum"),
        حضور=("حضور","sum"),
        إجازات=("إجازة","sum"),
        غياب=("غياب","sum"),
        تأخير_دقائق=("دقائق التأخير","sum"),
        تأخير_أيام=("تأخر","sum"),
    )
)
monthly["نسبة_الحضور"] = (monthly["حضور"] / monthly["أيام_العمل"].replace(0, np.nan) * 100).fillna(0)

weekday = (
    dff.groupby("اليوم_الاسبوعي", as_index=False)
    .agg(
        سجلات=("رقم الموظف","count"),
        أيام_العمل=("يوم عمل محسوب","sum"),
        حضور=("حضور","sum"),
        إجازات=("إجازة","sum"),
        غياب=("غياب","sum"),
        تأخير_دقائق=("دقائق التأخير","sum"),
        تأخير_أيام=("تأخر","sum"),
    )
)
weekday["نسبة_الحضور"] = (weekday["حضور"] / weekday["أيام_العمل"].replace(0, np.nan) * 100).fillna(0)

# ترتيب أيام الأسبوع للعربي
order_days = ["الأحد","الإثنين","الثلاثاء","الأربعاء","الخميس","الجمعة","السبت"]
weekday["ترتيب"] = weekday["اليوم_الاسبوعي"].apply(lambda x: order_days.index(x) if x in order_days else 99)
weekday = weekday.sort_values("ترتيب").drop(columns=["ترتيب"])




# =========================================================
# Historical Insights (Narrative)
# =========================================================
def _safe_pct_change(curr, prev):
    if prev is None or pd.isna(prev) or prev == 0:
        return np.nan
    return ((curr - prev) / prev) * 100

# Best/Worst month attendance
best_month = monthly.sort_values("نسبة_الحضور", ascending=False).head(1)
worst_month = monthly.sort_values("نسبة_الحضور", ascending=True).head(1)

# Month over month change (last two months)
monthly_sorted = monthly.sort_values("شهر")
mom_txt = "—"
if len(monthly_sorted) >= 2:
    last = monthly_sorted.iloc[-1]
    prev = monthly_sorted.iloc[-2]
    delta = _safe_pct_change(last["نسبة_الحضور"], prev["نسبة_الحضور"])
    if not pd.isna(delta):
        arrow = "" if delta >= 0 else ""
        mom_txt = f"{arrow} تغيّر شهري: {delta:+.1f}% (آخر شهر مقارنة باللي قبله)"

# Outlier days (high absence)
z = (daily["غياب"] - daily["غياب"].mean()) / (daily["غياب"].std(ddof=0) if daily["غياب"].std(ddof=0) else 1)
daily["Z"] = z
outliers = daily.sort_values("Z", ascending=False).head(3)

# Top weekday lateness
top_late_day = weekday.sort_values("تأخير_دقائق", ascending=False).head(1)
# =========================================================
# Forecast / Predictive-lite
# =========================================================

# 1) توقع الغياب للشهر القادم بناءً على آخر 3 شهور
monthly_forecast = monthly.sort_values("شهر").copy()
monthly_forecast["متوسط_غياب_3_شهور"] = monthly_forecast["غياب"].rolling(3, min_periods=1).mean()

forecast_next_absence = None
forecast_next_attendance = None
next_month_label = None

if len(monthly_forecast) > 0:
    last_month_period = pd.Period(monthly_forecast["شهر"].iloc[-1], freq="M")
    next_month_label = str(last_month_period + 1)
    forecast_next_absence = monthly_forecast["متوسط_غياب_3_شهور"].iloc[-1]

    avg_records_per_month = monthly_forecast["سجلات"].tail(3).mean()
    if avg_records_per_month and avg_records_per_month > 0:
        forecast_next_attendance = max(
            0,
            (1 - (forecast_next_absence / avg_records_per_month)) * 100
        )

# 2) تحليل أيام الأسبوع: هل الويكند/بعض الأيام أعلى غياب؟
weekday_forecast = (
    dff.groupby("اليوم_الاسبوعي", as_index=False)
    .agg(
        متوسط_الغياب=("غياب", "mean"),
        متوسط_التأخير=("دقائق التأخير", "mean"),
        عدد_السجلات=("رقم الموظف", "count")
    )
)

weekday_forecast["متوسط_الغياب"] = weekday_forecast["متوسط_الغياب"] * 100
weekday_forecast["ترتيب"] = weekday_forecast["اليوم_الاسبوعي"].apply(
    lambda x: order_days.index(x) if x in order_days else 99
)
weekday_forecast = weekday_forecast.sort_values("ترتيب").drop(columns=["ترتيب"])

# 3) هل الغياب يرتفع شهرياً؟
monthly_trend = monthly.sort_values("شهر").copy()
monthly_trend["متوسط_غياب_3_شهور"] = monthly_trend["غياب"].rolling(3, min_periods=1).mean()
# =========================================================
# Tabs
# =========================================================
tabs = st.tabs([
    "نظرة تنفيذية",
    "تحليلات تاريخية",
    "ترندات",
    "أداء الموظفين",
    "الجداول"
])

# =========================================================
# Tab 1: Executive Overview
# =========================================================
with tabs[0]:
    st.markdown("##  ملخص تنفيذي سريع")
    st.markdown("---")

    bm = best_month.iloc[0] if len(best_month) else None
    wm = worst_month.iloc[0] if len(worst_month) else None
    tl = top_late_day.iloc[0] if len(top_late_day) else None

    insight_rows = []

    if bm is not None:
        insight_rows.append(
            ("", f"أفضل شهر حضور", f"{bm['شهر']} • {bm['نسبة_الحضور']:.1f}%")
        )

    if wm is not None:
        insight_rows.append(
            ("", f"أسوأ شهر حضور", f"{wm['شهر']} • {wm['نسبة_الحضور']:.1f}%")
        )

    if mom_txt != "—":
        insight_rows.append(
            ("", "التغير الشهري", mom_txt.replace("", "").replace("", "").strip())
        )

    if tl is not None:
        insight_rows.append(
            ("", "أعلى يوم تأخير", f"{tl['اليوم_الاسبوعي']} • {int(tl['تأخير_دقائق']):,} دقيقة")
        )

    if len(outliers):
        d0 = outliers.iloc[0]
        insight_rows.append(
            ("", "أعلى يوم غياب شاذ", f"{d0['التاريخ'].date()} • {int(d0['غياب'])} حالة")
        )

    if len(weekday_forecast) > 0:
        worst_day = weekday_forecast.sort_values("متوسط_الغياب", ascending=False).iloc[0]
        insight_rows.append(
            ("", "اليوم الأعلى احتمالًا للغياب", f"{worst_day['اليوم_الاسبوعي']} • {worst_day['متوسط_الغياب']:.1f}%")
        )

    if forecast_next_attendance is not None and next_month_label is not None:
        insight_rows.append(
            ("", "التوقع المبدئي للشهر القادم", f"{next_month_label} • حضور متوقع {forecast_next_attendance:.1f}%")
        )

    if len(monthly_trend) >= 2:
        last_abs = monthly_trend["غياب"].iloc[-1]
        prev_abs = monthly_trend["غياب"].iloc[-2]

        if last_abs > prev_abs:
            trend_note = "الغياب في ارتفاع مقارنة بالشهر السابق"
            trend_icon = ""
        elif last_abs < prev_abs:
            trend_note = "الغياب في تحسن مقارنة بالشهر السابق"
            trend_icon = ""
        else:
            trend_note = "الغياب مستقر تقريبًا مقارنة بالشهر السابق"
            trend_icon = ""

        insight_rows.append((trend_icon, "اتجاه الغياب", trend_note))

    left_box, center_box, right_box = st.columns([1.2, 3.6, 1.2])

    with center_box:
        rows_html = ""
        for icon, label, value in insight_rows:
            rows_html += f"""
            <div style="
                display:flex;
                align-items:flex-start;
                gap:12px;
                padding:12px 0;
                border-bottom:1px solid rgba(255,255,255,0.07);
            ">
                <div style="
                    font-size:20px;
                    line-height:1.2;
                    width:28px;
                    text-align:center;
                    flex-shrink:0;
                ">{icon}</div>
                <div style="flex:1; text-align: center;">
                    <div style="
                        color:#E5E7EB;
                        font-weight:700;
                        font-size:15px;
                        margin-bottom:4px;
                    ">{label}</div>
                    <div style="
                        color:#CBD5E1;
                        font-size:14px;
                        line-height:1.8;
                    ">{value}</div>
                </div>
            </div>
            """

        st.markdown(
            f"""
            <div style="text-align:center; margin-bottom:14px;">
                <div style="font-size:28px; font-weight:800; color:#F8FAFC;">
                     تحليلات تنفيذية
                </div>
                <div style="font-size:13px; color:#94A3B8; margin-top:4px;">
                    أهم المؤشرات والاستنتاجات التنفيذية من بيانات الحضور والانضباط
                </div>
            </div>

            <div style="
                background: linear-gradient(180deg, rgba(255,255,255,0.025), rgba(255,255,255,0.015));
                border: 1px solid rgba(52,211,153,0.16);
                border-radius: 18px;
                padding: 18px 22px;
                box-shadow: 0 10px 30px rgba(0,0,0,0.22);
            ">
                {rows_html}
            </div>
            """,
            unsafe_allow_html=True
        )

    st.markdown("<div style='height:10px'></div>", unsafe_allow_html=True)
    # Quick Top charts
    t1, t2 = st.columns(2, gap="small")

    with t1:
        st.subheader(" أعلى 10 غياب")
        top_abs = (
            per_emp.sort_values(["أيام_الغياب", "دقائق_التأخير"], ascending=[False, False])
            .head(10)
            .copy()
        )
        fig_abs = go.Figure()
        fig_abs.add_trace(go.Bar(
            x=top_abs["اسم الموظف"],
            y=top_abs["أيام_الغياب"],
            name="أيام الغياب",
            marker_color=COLORS["red"],
            hovertemplate="<b>%{x}</b><br>أيام الغياب: %{y}<extra></extra>",
        ))
        apply_plotly_theme(fig_abs, height=360, title=None, showlegend=False, hovermode="closest")
        fig_abs.update_yaxes(title="أيام")
        fig_abs.update_xaxes(title="", tickangle=-35)
        st.plotly_chart(fig_abs, use_container_width=True, key="chart_top_abs")

    with t2:
        st.subheader(" أعلى 10 متاخرين ")
        top_late = (
            per_emp.sort_values(["دقائق_التأخير", "أيام_التأخير"], ascending=[False, False])
            .head(10)
            .copy()
        )
        fig_late = go.Figure()
        fig_late.add_trace(go.Bar(
            x=top_late["اسم الموظف"],
            y=top_late["دقائق_التأخير"],
            name="دقائق التأخير",
            marker_color=COLORS["orange"],
            hovertemplate="<b>%{x}</b><br>دقائق التأخير: %{y:,.0f}<extra></extra>",
        ))
        apply_plotly_theme(fig_late, height=360, title=None, showlegend=False, hovermode="closest")
        fig_late.update_yaxes(title="دقائق")
        fig_late.update_xaxes(title="", tickangle=-35)
        st.plotly_chart(fig_late, use_container_width=True, key="chart_top_late")


    # شارتات إضافية جديدة
    x1, x2 = st.columns(2, gap="small")

    with x1:
        st.subheader(" أعلى 10 انصراف مبكر")
        top_early = (
            per_emp.sort_values(["دقائق_انصراف_مبكر", "أيام_الغياب"], ascending=[False, False])
            .head(10)
            .copy()
        )

        fig_early = go.Figure()
        fig_early.add_trace(go.Bar(
            x=top_early["اسم الموظف"],
            y=top_early["دقائق_انصراف_مبكر"],
            name="دقائق الانصراف المبكر",
            marker_color=COLORS["purple"],
            hovertemplate="<b>%{x}</b><br>دقائق الانصراف المبكر: %{y:,.0f}<extra></extra>",
        ))
        apply_plotly_theme(fig_early, height=360, title=None, showlegend=False, hovermode="closest")
        fig_early.update_yaxes(title="دقائق")
        fig_early.update_xaxes(title="", tickangle=-35)
        st.plotly_chart(fig_early, use_container_width=True, key="chart_top_early")

    with x2:
        st.subheader(" أفضل 10 موظفين انضباطًا")
        best_emp = (
            per_emp.sort_values(
                ["نسبة_الحضور", "دقائق_التأخير", "أيام_الغياب"],
                ascending=[False, True, True]
            )
            .head(10)
            .copy()
        )

        fig_best = go.Figure()
        fig_best.add_trace(go.Bar(
            x=best_emp["اسم الموظف"],
            y=best_emp["نسبة_الحضور"],
            name="نسبة الحضور",
            marker_color=COLORS["green"],
            customdata=np.stack([
                best_emp["أيام_الحضور"].to_numpy(),
                best_emp["أيام_الغياب"].to_numpy(),
                best_emp["دقائق_التأخير"].to_numpy(),
            ], axis=-1),
            hovertemplate=(
                "<b>%{x}</b><br>"
                "نسبة الحضور: %{y:.1f}%<br>"
                "أيام الحضور: %{customdata[0]:.0f}<br>"
                "أيام الغياب: %{customdata[1]:.0f}<br>"
                "دقائق التأخير: %{customdata[2]:.0f}"
                "<extra></extra>"
            ),
        ))
        apply_plotly_theme(fig_best, height=360, title=None, showlegend=False, hovermode="closest")
        fig_best.update_yaxes(title="%")
        fig_best.update_xaxes(title="", tickangle=-35)
        st.plotly_chart(fig_best, use_container_width=True, key="chart_best_emp")

        st.markdown("---")
    # تحليل الإدارات
    if "الإداره" in dff.columns:
        st.markdown("###  مقارنة الإدارات")

        dept_df = dff.copy()
        dept_df["الإداره"] = dept_df["الإداره"].astype(str).fillna("غير محدد").str.strip()
        dept_df["الإداره"] = dept_df["الإداره"].replace({"": "غير محدد", "nan": "غير محدد", "None": "غير محدد"})

        dept_summary = (
            dept_df.groupby("الإداره", as_index=False)
            .agg(
                سجلات=("رقم الموظف", "count"),
                غياب=("غياب", "sum"),
                تأخير_دقائق=("دقائق التأخير", "sum"),
                انصراف_مبكر_دقائق=("دقائق الانصراف المبكر", "sum"),
            )
        )

        dept_summary["نسبة_الحضور"] = (
            1 - (dept_summary["غياب"] / dept_summary["سجلات"].replace(0, np.nan))
        ) * 100
        dept_summary["نسبة_الحضور"] = dept_summary["نسبة_الحضور"].fillna(0)
        dept_summary = dept_summary.sort_values("غياب", ascending=False)

        # ✅ هذا يخلي البلوك كله في المنتصف
        outer_left, center_col, outer_right = st.columns([1, 5, 1])

        with center_col:
            d1, d2 = st.columns(2, gap="medium")

            with d1:
                fig_dept_abs = go.Figure()
                fig_dept_abs.add_trace(go.Bar(
                    x=dept_summary["الإداره"],
                    y=dept_summary["غياب"],
                    name="غياب",
                    marker_color=COLORS["red"],
                    customdata=np.stack([
                        dept_summary["سجلات"].to_numpy(),
                        dept_summary["نسبة_الحضور"].to_numpy(),
                    ], axis=-1),
                    hovertemplate=(
                        "<b>%{x}</b><br>"
                        "الغياب: %{y:.0f}<br>"
                        "السجلات: %{customdata[0]:.0f}<br>"
                        "نسبة الحضور: %{customdata[1]:.1f}%"
                        "<extra></extra>"
                    ),
                ))
                apply_plotly_theme(
                    fig_dept_abs,
                    height=380,
                    title="الغياب حسب الإدارة",
                    showlegend=False,
                    hovermode="closest"
                )
                fig_dept_abs.update_yaxes(title="عدد حالات الغياب")
                fig_dept_abs.update_xaxes(title="", tickangle=-20)
                st.plotly_chart(fig_dept_abs, use_container_width=True, key="dept_abs_chart")

            with d2:
                fig_dept_late = go.Figure()
                fig_dept_late.add_trace(go.Bar(
                    x=dept_summary["الإداره"],
                    y=dept_summary["تأخير_دقائق"],
                    name="دقائق التأخير",
                    marker_color=COLORS["orange"],
                    customdata=np.stack([
                        dept_summary["انصراف_مبكر_دقائق"].to_numpy(),
                        dept_summary["نسبة_الحضور"].to_numpy(),
                    ], axis=-1),
                    hovertemplate=(
                        "<b>%{x}</b><br>"
                        "دقائق التأخير: %{y:,.0f}<br>"
                        "دقائق الانصراف المبكر: %{customdata[0]:,.0f}<br>"
                        "نسبة الحضور: %{customdata[1]:.1f}%"
                        "<extra></extra>"
                    ),
                ))
                apply_plotly_theme(
                    fig_dept_late,
                    height=380,
                    title="التأخير حسب الإدارة",
                    showlegend=False,
                    hovermode="closest"
                )
                fig_dept_late.update_yaxes(title="دقائق")
                fig_dept_late.update_xaxes(title="", tickangle=-20)
                st.plotly_chart(fig_dept_late, use_container_width=True, key="dept_late_chart")

    # =========================================================
# Tab 2: Historical Insights
# =========================================================
with tabs[1]:
    st.markdown("##  تحليلات تاريخية")

    c1, c2 = st.columns(2, gap="large")

    with c1:
        st.subheader("مؤشر الحضور ومتوسط التأخير")
        figm = go.Figure()
        figm.add_trace(go.Bar(
            x=monthly["شهر"],
            y=monthly["نسبة_الحضور"],
            name="نسبة الحضور",
            marker_color=COLORS["green"],
            hovertemplate="<b>%{x}</b><br>نسبة الحضور: %{y:.1f}%<extra></extra>"
        ))
        figm.add_trace(go.Scatter(
            x=monthly["شهر"],
            y=(monthly["تأخير_دقائق"] / monthly["سجلات"].replace(0, np.nan)).fillna(0),
            mode="lines+markers",
            name="متوسط التأخير/سجل",
            line=dict(color=COLORS["cyan"], width=3),
            hovertemplate="<b>%{x}</b><br>متوسط التأخير: %{y:.2f} دقيقة<extra></extra>"
        ))
        apply_plotly_theme(figm, height=420, showlegend=True)
        figm.update_yaxes(title="%", rangemode="tozero")
        st.plotly_chart(figm, use_container_width=True, key="hist_month")

    with c2:
        st.subheader("مقارنة الغياب والتأخير الأسبوعية")
        figw = go.Figure()
        figw.add_trace(go.Bar(
            x=weekday["اليوم_الاسبوعي"],
            y=weekday["غياب"],
            name="غياب",
            marker_color=COLORS["red"],
            hovertemplate="<b>%{x}</b><br>غياب: %{y}<extra></extra>"
        ))
        figw.add_trace(go.Bar(
            x=weekday["اليوم_الاسبوعي"],
            y=weekday["تأخير_أيام"],
            name="أيام التأخير",
            marker_color=COLORS["orange"],
            hovertemplate="<b>%{x}</b><br>أيام التأخير: %{y}<extra></extra>"
        ))
        apply_plotly_theme(figw, height=420, showlegend=True, hovermode="x unified")
        figw.update_layout(barmode="group")
        st.plotly_chart(figw, use_container_width=True, key="hist_week")

    st.markdown("---")
    st.markdown("### الأيام الأكثر غياباً ")

    od = outliers.copy()
    if len(od):
        od["التاريخ"] = od["التاريخ"].dt.strftime("%Y-%m-%d")
        st.dataframe(
            od[["التاريخ","غياب","تأخير_دقائق","تأخير_أيام","Z"]].rename(columns={
                "التاريخ":"التاريخ",
                "غياب":"عدد الغياب",
                "تأخير_دقائق":"دقائق التأخير",
                "تأخير_أيام":"أيام التأخير",
                "Z":"مؤشر الانحراف الاستثنائي"
            }),
            use_container_width=True,
            hide_index=True
        )
        st.caption(" اقتراح: اربط هذه الأيام بأحداث (إجازات/طقس/مناسبات/ضغط عمل) عشان تصير Insight قوية للإدارة.")

    else:
        st.info("ما قدرت أحدد Outliers (البيانات قليلة أو التباين ضعيف).")
        
# =========================================================
# Tab 3: Trends
# =========================================================
with tabs[2]:
    st.markdown("##  ترندات وتحليلات عميقة")

    c1, c2 = st.columns(2, gap="large")

    with c1:
        st.markdown("### التحليل الزمني لاتجاهات الغياب ")
        fig = go.Figure()
        fig.add_trace(go.Scatter(
            x=daily["التاريخ"], y=daily["غياب"],
            mode="lines",
            name="غياب يومي",
            line=dict(color=COLORS["red"], width=2),
            hovertemplate="<b>%{x|%Y-%m-%d}</b><br>غياب: %{y}<extra></extra>"
        ))
        # Rolling average 7 days
        roll = daily["غياب"].rolling(7, min_periods=1).mean()
        fig.add_trace(go.Scatter(
            x=daily["التاريخ"], y=roll,
            mode="lines",
            name="متوسط متحرك 7 أيام",
            line=dict(color=COLORS["purple"], width=3),
            hovertemplate="<b>%{x|%Y-%m-%d}</b><br>متوسط 7 أيام: %{y:.2f}<extra></extra>"
        ))
        apply_plotly_theme(fig, height=420,  showlegend=True)
        fig.update_yaxes(title="عدد الحالات")
        st.plotly_chart(fig, use_container_width=True, key="trend_abs_rolling")

    with c2:
        st.markdown("### توزيع دقائق التأخير")
        # Distribution of late minutes
        fig = go.Figure()
        late_vals = dff["دقائق التأخير"].clip(lower=0)
        fig.add_trace(go.Histogram(
            x=late_vals[late_vals > 0],
            nbinsx=25,
            marker_color=COLORS["orange"],
            hovertemplate="دقائق: %{x}<br>عدد: %{y}<extra></extra>"
        ))
        apply_plotly_theme(fig, height=420,  showlegend=False, hovermode="closest")
        fig.update_xaxes(title="دقائق التأخير")
        fig.update_yaxes(title="عدد السجلات")
        st.plotly_chart(fig, use_container_width=True, key="trend_late_dist")

    st.markdown("---")
    st.markdown("###  الخريطة الحرارية لكثافة الحضور")

    # Heatmap: weekday x (month) attendance rate
    heat = (
        dff.groupby(["شهر","اليوم_الاسبوعي"], as_index=False)
        .agg(سجلات=("رقم الموظف","count"), غياب=("غياب","sum"))
    )
    heat["نسبة_الحضور"] = (1 - (heat["غياب"] / heat["سجلات"].replace(0, np.nan))) * 100
    heat["نسبة_الحضور"] = heat["نسبة_الحضور"].fillna(0)

    pivot = heat.pivot(index="اليوم_الاسبوعي", columns="شهر", values="نسبة_الحضور").reindex(order_days)
    fig = go.Figure(data=go.Heatmap(
        z=pivot.values,
        x=list(pivot.columns),
        y=list(pivot.index),
        colorscale="Viridis",
        hovertemplate="شهر: %{x}<br>اليوم: %{y}<br>نسبة الحضور: %{z:.1f}%<extra></extra>"
    ))
    apply_plotly_theme(fig, height=420, showlegend=False, hovermode="closest")
    st.plotly_chart(fig, use_container_width=True, key="heat_att")

# =========================================================
# Tab 4: Employees Performance
# =========================================================
with tabs[3]:
    st.markdown("##  أداء الموظفين ")

    left, right = st.columns([1.15, 1], gap="large")

    with left:
        st.subheader(" ملخص شامل للموظفين")
        view = per_emp.sort_values(
            ["نسبة_الحضور", "أيام_الغياب", "دقائق_التأخير"],
            ascending=[True, False, False]
        ).copy()
        st.dataframe(view, use_container_width=True, hide_index=True)

    with right:
        st.subheader(" أفضل 15 حضورًا")
        top_att = per_emp.sort_values(
            ["نسبة_الحضور", "دقائق_التأخير"],
            ascending=[False, True]
        ).head(15)

        fig_top = go.Figure()
        fig_top.add_trace(go.Bar(
            x=top_att["اسم الموظف"],
            y=top_att["نسبة_الحضور"],
            name="نسبة الحضور",
            marker_color=COLORS["green"],
            hovertemplate="<b>%{x}</b><br>حضور: %{y:.1f}%<extra></extra>"
        ))
        apply_plotly_theme(fig_top, height=420,  showlegend=False, hovermode="closest")
        fig_top.update_yaxes(title="%")
        fig_top.update_xaxes(title="", tickangle=-35)
        st.plotly_chart(fig_top, use_container_width=True, key="emp_top_att")

    st.markdown("---")
    st.subheader(" أقل 15 انضباطًا")




    # =========================================================
    # Non-Compliance (%): غياب + تأخير (>20 دقيقة) / إجمالي الأيام
    # =========================================================
    # =========================================================
    # Non-Compliance (%): غياب + تأخير (>20 دقيقة) / إجمالي الأيام
    # =========================================================
    late20 = (
        dff.assign(تأخير20=(dff["دقائق التأخير"] > 20).astype(int))
        .groupby(["رقم الموظف", "اسم الموظف عرض"], as_index=False)
        .agg(أيام_تأخير_20=("تأخير20", "sum"))
        .rename(columns={"اسم الموظف عرض": "اسم الموظف"})
    )

    bad = per_emp.merge(late20, on=["رقم الموظف", "اسم الموظف"], how="left")
    bad["أيام_تأخير_20"] = bad["أيام_تأخير_20"].fillna(0)

    bad["أيام_عدم_الانضباط"] = bad["أيام_الغياب"] + bad["أيام_تأخير_20"]
    bad["نسبة_عدم_الانضباط"] = np.where(
        bad["سجلات"] > 0,
        (bad["أيام_عدم_الانضباط"] / bad["سجلات"]) * 100,
        0
    )

    bad15 = bad.sort_values("نسبة_عدم_الانضباط", ascending=False).head(15)

    fig_bad = go.Figure()
    fig_bad.add_trace(go.Bar(
        x=bad15["اسم الموظف"],
        y=bad15["نسبة_عدم_الانضباط"],
        name="عدم الانضباط (%)",
        marker_color=COLORS["red"],
        customdata=np.stack([
            bad15["أيام_عدم_الانضباط"].to_numpy(),
            bad15["سجلات"].to_numpy(),
            bad15["أيام_الغياب"].to_numpy(),
            bad15["أيام_تأخير_20"].to_numpy(),
        ], axis=-1),
        hovertemplate=(
            "<b>%{x}</b><br>"
            "عدم الانضباط: <b>%{y:.1f}%</b><br>"
            "أيام عدم الانضباط: %{customdata[0]:.0f}<br>"
            "إجمالي الأيام: %{customdata[1]:.0f}<br>"
            "غياب: %{customdata[2]:.0f}<br>"
            "تأخير >20 دقيقة: %{customdata[3]:.0f}"
            "<extra></extra>"
        )
    ))

    apply_plotly_theme(
        fig_bad,
        height=620,
        #title=" أعلى 15 عدم انضباط (%)",
        showlegend=False,
        hovermode="closest"
    )
    fig_bad.update_yaxes(title="نسبة عدم الانضباط %", rangemode="tozero")
    fig_bad.update_xaxes(title="", tickangle=-35)

    st.plotly_chart(fig_bad, use_container_width=True, key="emp_noncompliance_top15")

    st.caption(" المؤشر يعتمد على الغياب + التأخير فوق 20 دقيقة مقارنة بإجمالي الأيام.")
# =========================================================
# Tab 5: Tables
# =========================================================
with tabs[4]:
    st.markdown("##  الجداول")
    st.subheader("ملخص الموظفين")
    st.dataframe(per_emp, use_container_width=True, hide_index=True)

    st.markdown("---")
    st.subheader("ملخص يومي")
    st.dataframe(daily.drop(columns=["Z"]), use_container_width=True, hide_index=True)

    if show_raw:
        st.markdown("---")
        st.subheader("البيانات الخام (بعد الفلاتر)")
        st.dataframe(dff, use_container_width=True, hide_index=True)

    st.markdown("---")
    st.subheader("فحص تحويل الأسماء")
    st.dataframe(name_check, use_container_width=True, hide_index=True)