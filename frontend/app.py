import sys
import os
import io
import re
import zipfile
import xml.etree.ElementTree as ET
from datetime import datetime

# ============================================================
# PROJECT ROOT
# ============================================================

PROJECT_ROOT = os.path.dirname(
    os.path.dirname(
        os.path.abspath(__file__)
    )
)

if PROJECT_ROOT not in sys.path:
    sys.path.insert(0, PROJECT_ROOT)


# ============================================================
# IMPORTS
# ============================================================

import streamlit as st
import pandas as pd
import numpy as np
import plotly.graph_objects as go
from openpyxl import Workbook
from openpyxl.drawing.image import Image as XLImage
from openpyxl.styles import Font as XLFont, PatternFill, Alignment, Border, Side

from core.backend_service import BackendService


# ============================================================
# PAGE CONFIGURATION
# ============================================================

st.set_page_config(
    page_title="Helicopter Analysis System",
    page_icon="🚁",
    layout="wide",
    initial_sidebar_state="expanded"
)


# ============================================================
# SESSION STATE
# ============================================================

if "service" not in st.session_state:
    st.session_state.service = BackendService()

if "file_loaded" not in st.session_state:
    st.session_state.file_loaded = False

if "analysis_result" not in st.session_state:
    st.session_state.analysis_result = None

if "selected_attributes" not in st.session_state:
    st.session_state.selected_attributes = []

if "graph_type" not in st.session_state:
    st.session_state.graph_type = "Line"

if "converted_file" not in st.session_state:
    st.session_state.converted_file = None

if "uploaded_filename" not in st.session_state:
    st.session_state.uploaded_filename = None

if "uploaded_file_signature" not in st.session_state:
    st.session_state.uploaded_file_signature = None

# Multiple-file workspace state. The uploader can hold several flight
# datasets, while exactly one dataset is active for analysis at a time.
if "uploaded_files_signatures" not in st.session_state:
    st.session_state.uploaded_files_signatures = []

if "uploaded_files_metadata" not in st.session_state:
    st.session_state.uploaded_files_metadata = []

if "selected_file_signature" not in st.session_state:
    st.session_state.selected_file_signature = None

if "active_tab" not in st.session_state:
    st.session_state.active_tab = "Data Input"

if "x_axis" not in st.session_state:
    st.session_state.x_axis = None

if "y_axis" not in st.session_state:
    st.session_state.y_axis = []

if "graph_generation_mode" not in st.session_state:
    st.session_state.graph_generation_mode = "selected"

if "graph_ready" not in st.session_state:
    st.session_state.graph_ready = False

if "report_data" not in st.session_state:
    st.session_state.report_data = {
        "file_information": {},
        "graphs": [],
        "dataset_issues": [],
        "summary": ""
    }

if "anomaly_cache_key" not in st.session_state:
    st.session_state.anomaly_cache_key = None

if "anomaly_report" not in st.session_state:
    st.session_state.anomaly_report = []

if "anomaly_issues" not in st.session_state:
    st.session_state.anomaly_issues = []

if "graph_data" not in st.session_state:
    st.session_state.graph_data = None

if "workspace_by_signature" not in st.session_state:
    st.session_state.workspace_by_signature = {}

if "known_file_registry" not in st.session_state:
    st.session_state.known_file_registry = {}

if "removed_file_signatures" not in st.session_state:
    st.session_state.removed_file_signatures = set()

if "file_header_by_signature" not in st.session_state:
    st.session_state.file_header_by_signature = {}

if "row_cleanup_done_by_signature" not in st.session_state:
    st.session_state.row_cleanup_done_by_signature = {}

if "cleaned_file_by_signature" not in st.session_state:
    st.session_state.cleaned_file_by_signature = {}

if "cleaned_file_download_by_signature" not in st.session_state:
    st.session_state.cleaned_file_download_by_signature = {}

if "anomaly_report" not in st.session_state:
    st.session_state.anomaly_report = []

if "anomaly_issues" not in st.session_state:
    st.session_state.anomaly_issues = []

if "anomaly_summary" not in st.session_state:
    st.session_state.anomaly_summary = ""

if "anomaly_cache_key" not in st.session_state:
    st.session_state.anomaly_cache_key = None

if "x_axis" not in st.session_state:
    st.session_state.x_axis = None


service = st.session_state.service


# ============================================================
# NAVIGATION FUNCTION
# ============================================================

def change_tab(tab_name):
    st.session_state.active_tab = tab_name


# ============================================================
# CSS
# ============================================================

st.markdown("""<style>
:root{--blue:#0646A8;--navy:#19324D;--red:#E21F2F;--bg:#F5F8FC;--line:#DCE5EF;--muted:#61748A;--green:#11823B;}
.stApp{background:var(--bg)!important;} [data-testid="stHeader"]{background:transparent!important;}
[data-testid="stMainBlockContainer"]{max-width:1500px!important;padding-top:.2rem!important;padding-bottom:2rem!important;}
section[data-testid="stSidebar"]{background:#fff!important;border-right:1px solid #DCE4ED!important;min-width:245px!important;max-width:285px!important;}
section[data-testid="stSidebar"]>div:first-child{padding:.35rem .75rem 1rem!important;}
section[data-testid="stSidebar"] p,section[data-testid="stSidebar"] span,section[data-testid="stSidebar"] label{color:var(--navy)!important;}
.hal-sidebar-brand{text-align:center;padding:0 2px 14px;border-bottom:1px solid #E4EAF1;margin-bottom:15px;}
.hal-real-logo{display:block;width:165px;height:auto;margin:0 auto 7px;object-fit:contain;}
.hal-company{font-size:10px;font-weight:800;color:var(--blue)!important;letter-spacing:.2px;}
.hal-company-sub{font-size:8px;color:#8290A0!important;letter-spacing:.8px;margin-top:3px;}
.side-heading{font-size:10px;font-weight:900;color:var(--blue)!important;letter-spacing:1px;margin:14px 7px 8px;}
section[data-testid="stSidebar"] div.stButton>button{min-height:42px!important;border-radius:8px!important;border:1px solid transparent!important;background:#fff!important;color:var(--navy)!important;text-align:left!important;font-size:13px!important;font-weight:650!important;box-shadow:none!important;padding:0 12px!important;margin:2px 0!important;white-space:normal!important;overflow:visible!important;overflow-wrap:anywhere!important;word-break:normal!important;}
section[data-testid="stSidebar"] div.stButton>button:hover{background:#F2F7FD!important;border-color:#D8E6F5!important;color:var(--blue)!important;}
section[data-testid="stSidebar"] div.stButton>button[kind="primary"]{background:#0755B8!important;border-color:#0755B8!important;color:#fff!important;box-shadow:0 4px 10px rgba(7,85,184,.16)!important;}
section[data-testid="stSidebar"] div.stButton>button[kind="primary"] p,section[data-testid="stSidebar"] div.stButton>button[kind="primary"] span{color:#fff!important;}
.side-status{margin:10px 7px;padding:13px;border:1px solid var(--line);border-radius:9px;background:#FAFCFE;}
.side-status-title{color:var(--green)!important;font-weight:850;font-size:11px;margin-bottom:10px;} .side-status-row{display:flex;justify-content:space-between;font-size:10px;margin:8px 0;color:#607286!important;} .side-status-row strong{color:#08762F!important;}
.side-aircraft{text-align:center;margin:20px auto 4px;} .side-aircraft img{width:205px;max-width:90%;height:auto;} .side-motto{text-align:center;color:#0C4D9C!important;font-size:8px;font-weight:850;line-height:1.5;margin-top:3px;} .side-motto span{color:var(--red)!important;}
.top-header{background:linear-gradient(100deg,#063D8D,#0758BA 65%,#0B68D1);color:#fff;padding:12px 18px;border-bottom:3px solid var(--red);display:flex;align-items:center;justify-content:space-between;margin:0 0 18px;min-height:50px;gap:15px;flex-wrap:wrap;}
.top-header-left{display:flex;align-items:center;gap:12px;min-width:0;} .top-heli-icon{font-size:24px;line-height:1;filter:grayscale(1) brightness(0) invert(1);} .top-title{color:#fff!important;font-size:18px;font-weight:850;letter-spacing:.25px;white-space:normal;overflow-wrap:anywhere;} .top-actions{display:flex;align-items:center;gap:12px;flex-wrap:wrap;} .top-ready{color:#fff!important;font-size:10px;font-weight:850;background:rgba(255,255,255,.08);padding:7px 12px;border:1px solid rgba(255,255,255,.28);border-radius:20px;white-space:normal;overflow-wrap:anywhere;} .ready-dot{color:#28D466!important;} .top-date{color:#fff!important;font-size:10px;white-space:normal;overflow-wrap:anywhere;}
.page-kicker{font-size:11px;font-weight:900;color:var(--blue)!important;letter-spacing:1.35px;margin:2px 0 5px;} .main-title{font-size:30px;line-height:1.12;font-weight:850;color:var(--navy)!important;margin:0 0 7px;letter-spacing:-.5px;} .main-subtitle{font-size:13px;color:var(--muted)!important;margin:0 0 19px;} .title-rule{height:1px;background:#DCE5EF;margin:0 0 29px;}
.upload-card-head{display:flex;align-items:center;justify-content:space-between;gap:20px;background:#FFFFFF;border:1px solid #DCE5EF;border-radius:9px;padding:15px 18px;margin:0 0 10px;box-shadow:0 2px 8px rgba(35,63,89,.035);}
.upload-kicker,.section-label{font-size:11px;font-weight:900;color:var(--blue)!important;letter-spacing:1.35px;margin:2px 0 5px;}
.upload-title{font-size:18px;font-weight:850;color:var(--navy)!important;line-height:1.25;margin-bottom:4px;}
.upload-subtitle{font-size:12px;color:var(--muted)!important;line-height:1.5;}
.upload-format{font-size:10px;font-weight:800;color:#4F6982!important;white-space:nowrap;text-align:right;}
.info-strip{background:#EEF5FF;border:1px solid #D6E5F5;border-radius:8px;padding:11px 14px;color:#294664!important;font-size:11px;line-height:1.5;margin:5px 0 18px;}
.info-strip b{color:var(--blue)!important;font-weight:900;}
.axis-help{background:#F0F6FF;border:1px solid #D9E7F5;border-radius:8px;padding:11px 14px;color:#314B68!important;font-size:11px;line-height:1.6;margin-top:10px;}
div[data-testid="stMainBlockContainer"] .stMarkdown p{color:var(--navy)!important;}
div[data-testid="stMainBlockContainer"] div[data-testid="stCaptionContainer"] p{color:var(--muted)!important;}
div[data-testid="stMainBlockContainer"] .section-label{color:var(--blue)!important;}
div[data-testid="stMainBlockContainer"] .main-title{color:var(--navy)!important;}
div[data-testid="stMainBlockContainer"] .main-subtitle{color:var(--muted)!important;}

.file-card{background:#fff;border:1px solid var(--line);border-radius:9px;padding:14px 18px;margin:3px 0 22px;box-shadow:0 2px 8px rgba(35,63,89,.035);} .section-card{background:#fff;border:1px solid var(--line);border-radius:9px;padding:22px 24px;box-shadow:0 2px 8px rgba(35,63,89,.03);}
.axis-label{font-size:11px;font-weight:900;color:var(--blue)!important;letter-spacing:.5px;margin-bottom:7px;} .axis-note{font-size:11px;color:var(--navy)!important;line-height:1.65;margin-top:7px;} .help-card{background:#F0F6FF;border-radius:8px;padding:15px 16px;min-height:115px;} .help-title{font-size:12px;font-weight:850;color:var(--blue)!important;margin-bottom:7px;} .help-text{font-size:11px;color:#314B68!important;line-height:1.6;} .info-card{background:#EEF5FF;border-radius:8px;padding:15px 17px;margin-top:20px;color:#294664!important;} .info-title{font-size:12px;font-weight:850;color:var(--blue)!important;margin-bottom:6px;} .info-text{font-size:11px;color:#36516E!important;}
h1,h2,h3{color:var(--navy)!important;} h1{font-size:29px!important;font-weight:850!important;} h2{font-size:21px!important;font-weight:800!important;} h3{font-size:17px!important;font-weight:800!important;}
div.stButton>button{border-radius:8px!important;min-height:40px!important;font-weight:750!important;border:1px solid #CFDCE8!important;color:var(--navy)!important;background:#fff!important;} div.stButton>button[kind="primary"]{background:#0755B8!important;border-color:#0755B8!important;color:#fff!important;box-shadow:0 4px 10px rgba(7,85,184,.16)!important;} div.stButton>button[kind="primary"] p,div.stButton>button[kind="primary"] span{color:#fff!important;}
div[data-baseweb="select"]>div,div[data-baseweb="input"]>div{background:#fff!important;border-color:#C9D7E5!important;border-radius:7px!important;} div[data-baseweb="select"] span,div[data-baseweb="select"] input{color:var(--navy)!important;}
div[data-testid="stMainBlockContainer"] label{color:var(--navy)!important;}
div[data-testid="stMainBlockContainer"] [data-testid="stFileUploader"] label{color:var(--navy)!important;}
div[data-testid="stMainBlockContainer"] [data-testid="stFileUploader"] small{color:var(--muted)!important;}

div[data-testid="stMetric"]{background:#fff!important;border:1px solid var(--line)!important;border-radius:8px!important;padding:11px 13px!important;} div[data-testid="stMetricLabel"] p{color:var(--muted)!important;font-size:10px!important;} div[data-testid="stMetricValue"]{color:var(--navy)!important;font-weight:850!important;} div[data-testid="stDataFrame"]{border:1px solid var(--line);border-radius:8px;overflow:hidden;} hr{border-color:#DCE5EF!important;margin:1rem 0!important;} div[data-testid="stPlotlyChart"]{background:#fff;border:1px solid var(--line);border-radius:9px;padding:5px;}

/* ===== FINAL TEXT VISIBILITY OVERRIDE =====
   Keep the existing UI/background colors, but make all readable text black.
   White is retained only where it is intentionally used inside dark/blue UI
   surfaces so contrast remains good.
*/
div[data-testid="stAppViewContainer"] *,
div[data-testid="stMainBlockContainer"] *,
div[data-testid="stSidebar"] * {
    text-shadow: none !important;
}

/* Main page: all normal text black */
div[data-testid="stMainBlockContainer"] p,
div[data-testid="stMainBlockContainer"] span,
div[data-testid="stMainBlockContainer"] label,
div[data-testid="stMainBlockContainer"] small,
div[data-testid="stMainBlockContainer"] div,
div[data-testid="stMainBlockContainer"] li,
div[data-testid="stMainBlockContainer"] td,
div[data-testid="stMainBlockContainer"] th,
div[data-testid="stMainBlockContainer"] h1,
div[data-testid="stMainBlockContainer"] h2,
div[data-testid="stMainBlockContainer"] h3,
div[data-testid="stMainBlockContainer"] h4,
div[data-testid="stMainBlockContainer"] h5,
div[data-testid="stMainBlockContainer"] h6,
div[data-testid="stMainBlockContainer"] strong,
div[data-testid="stMainBlockContainer"] b {
    color: #000000 !important;
}

/* Sidebar text: black */
div[data-testid="stSidebar"] p,
div[data-testid="stSidebar"] span,
div[data-testid="stSidebar"] label,
div[data-testid="stSidebar"] small,
div[data-testid="stSidebar"] li,
div[data-testid="stSidebar"] td,
div[data-testid="stSidebar"] th,
div[data-testid="stSidebar"] h1,
div[data-testid="stSidebar"] h2,
div[data-testid="stSidebar"] h3,
div[data-testid="stSidebar"] h4,
div[data-testid="stSidebar"] h5,
div[data-testid="stSidebar"] h6,
div[data-testid="stSidebar"] strong,
div[data-testid="stSidebar"] b {
    color: #000000 !important;
}

/* Inputs and selection controls: black text */
div[data-testid="stMainBlockContainer"] input,
div[data-testid="stMainBlockContainer"] textarea,
div[data-testid="stMainBlockContainer"] [role="combobox"],
div[data-testid="stMainBlockContainer"] [role="option"],
div[data-testid="stMainBlockContainer"] [role="listbox"],
div[data-testid="stMainBlockContainer"] [data-baseweb="select"],
div[data-testid="stMainBlockContainer"] [data-baseweb="select"] *,
div[data-testid="stMainBlockContainer"] [data-baseweb="input"] *,
div[data-testid="stMainBlockContainer"] [data-testid="stSelectbox"] *,
div[data-testid="stMainBlockContainer"] [data-testid="stMultiSelect"] * {
    color: #000000 !important;
}

/* File uploader text: black */
div[data-testid="stFileUploader"],
div[data-testid="stFileUploader"] *,
div[data-testid="stFileUploaderDropzone"],
div[data-testid="stFileUploaderDropzone"] * {
    color: #000000 !important;
}

/* Dataset/file selection cards and all custom UI text: black */
.file-card *,
.dataset-card *,
.dataset-row *,
.selection-card *,
.info-strip *,
.axis-help *,
.help-text *,
.info-text *,
.upload-card *,
.upload-card-head *,
.upload-kicker *,
.upload-title *,
.upload-subtitle *,
.upload-format *,
.section-label *,
.main-title *,
.main-subtitle *,
.page-kicker * {
    color: #000000 !important;
}

/* Keep the existing colored backgrounds, borders, icons and buttons.
   Primary blue buttons use white text for contrast. */
div[data-testid="stMainBlockContainer"] button[kind="primary"],
div[data-testid="stMainBlockContainer"] button[kind="primary"] *,
div[data-testid="stSidebar"] button[kind="primary"],
div[data-testid="stSidebar"] button[kind="primary"] * {
    color: #ffffff !important;
}

/* Blue header/top-bar intentionally keeps white text */
.top-header,
.top-header *,
.topbar,
.topbar *,
.header,
.header * {
    color: #ffffff !important;
}


/* ===== FILE UPLOADER ONLY: WHITE TEXT =====
   The rest of the website keeps the black-text rule.
*/
div[data-testid="stFileUploader"],
div[data-testid="stFileUploader"] * {
    color: #ffffff !important;
}

div[data-testid="stFileUploaderDropzone"],
div[data-testid="stFileUploaderDropzone"] * {
    color: #ffffff !important;
}

div[data-testid="stFileUploader"] label,
div[data-testid="stFileUploader"] label *,
div[data-testid="stFileUploader"] small,
div[data-testid="stFileUploader"] span,
div[data-testid="stFileUploader"] p {
    color: #ffffff !important;
}

/* Uploaded file cards inside the dark uploader */
div[data-testid="stFileUploader"] [data-testid="stFileUploaderFile"],
div[data-testid="stFileUploader"] [data-testid="stFileUploaderFile"] *,
div[data-testid="stFileUploader"] [data-testid="stFileUploaderFileName"],
div[data-testid="stFileUploader"] [data-testid="stFileUploaderFileSize"] {
    color: #ffffff !important;
}

/* Keep the upload button itself readable */
div[data-testid="stFileUploader"] button,
div[data-testid="stFileUploader"] button * {
    color: #000000 !important;
}


/* Upload button inside the dark file-uploader only */
div[data-testid="stFileUploader"] button,
div[data-testid="stFileUploader"] button *,
div[data-testid="stFileUploaderDropzone"] button,
div[data-testid="stFileUploaderDropzone"] button * {
    color: #FFFFFF !important;
}

/* Keep the upload icon visible as well */
div[data-testid="stFileUploader"] button svg,
div[data-testid="stFileUploaderDropzone"] button svg {
    color: #FFFFFF !important;
    fill: #FFFFFF !important;
    stroke: #FFFFFF !important;
}


/* =========================================================
   FILE UPLOADER SECTION — WHITE TEXT ONLY
   Everything inside the dark uploader is white.
   ========================================================= */

/* Uploader container and every nested text element */
div[data-testid="stFileUploader"],
div[data-testid="stFileUploader"] * {
    color: #FFFFFF !important;
}

/* Uploader label: "Choose helicopter flight-data files" */
div[data-testid="stFileUploader"] label,
div[data-testid="stFileUploader"] label *,
div[data-testid="stFileUploader"] label p,
div[data-testid="stFileUploader"] label span {
    color: #FFFFFF !important;
}

/* Dark dropzone */
div[data-testid="stFileUploader"] [data-testid="stFileUploaderDropzone"],
div[data-testid="stFileUploader"] [data-testid="stFileUploaderDropzone"] * {
    color: #FFFFFF !important;
}

/* Upload button */
div[data-testid="stFileUploader"] button,
div[data-testid="stFileUploader"] button *,
div[data-testid="stFileUploader"] button p,
div[data-testid="stFileUploader"] button span {
    color: #FFFFFF !important;
}

/* Uploaded file chips/cards */
div[data-testid="stFileUploader"] [data-testid="stFileUploaderFile"],
div[data-testid="stFileUploader"] [data-testid="stFileUploaderFile"] *,
div[data-testid="stFileUploader"] [data-testid="stFileUploaderFile"] span,
div[data-testid="stFileUploader"] [data-testid="stFileUploaderFile"] small {
    color: #FFFFFF !important;
}

/* File names and file sizes */
div[data-testid="stFileUploader"] [data-testid="stFileUploaderFileName"],
div[data-testid="stFileUploader"] [data-testid="stFileUploaderFileName"] *,
div[data-testid="stFileUploader"] [data-testid="stFileUploaderFileSize"],
div[data-testid="stFileUploader"] [data-testid="stFileUploaderFileSize"] * {
    color: #FFFFFF !important;
}

/* Remove/close X */
div[data-testid="stFileUploader"] button[aria-label*="Remove"],
div[data-testid="stFileUploader"] button[aria-label*="remove"],
div[data-testid="stFileUploader"] button[aria-label*="Delete"],
div[data-testid="stFileUploader"] button[aria-label*="delete"],
div[data-testid="stFileUploader"] button[title*="Remove"],
div[data-testid="stFileUploader"] button[title*="remove"] {
    color: #FFFFFF !important;
}

/* SVG icons inside uploader */
div[data-testid="stFileUploader"] svg {
    color: #FFFFFF !important;
    fill: #FFFFFF !important;
    stroke: #FFFFFF !important;
}

/* Any uploader links / small helper text */
div[data-testid="stFileUploader"] a,
div[data-testid="stFileUploader"] small,
div[data-testid="stFileUploader"] p,
div[data-testid="stFileUploader"] span {
    color: #FFFFFF !important;
}

/* ========================================================= */


/* JSON / code-style output: white background with black text */
div[data-testid="stJson"],
div[data-testid="stJson"] > div,
div[data-testid="stJson"] pre,
div[data-testid="stJson"] code,
div[data-testid="stJson"] * {
    background: #FFFFFF !important;
    color: #000000 !important;
}

/* Code blocks shown by Streamlit */
div[data-testid="stCodeBlock"],
div[data-testid="stCodeBlock"] > div,
div[data-testid="stCodeBlock"] pre,
div[data-testid="stCodeBlock"] code,
div[data-testid="stCodeBlock"] * {
    background: #FFFFFF !important;
    color: #000000 !important;
}

/* =========================================================
   DEPLOYED UPLOADER — BLACK & WHITE
   ========================================================= */
div[data-testid="stFileUploader"] {
    background: #FFFFFF !important;
    border: 1px solid #111111 !important;
    border-radius: 10px !important;
}

div[data-testid="stFileUploader"] [data-testid="stFileUploaderDropzone"] {
    background: #111111 !important;
    border: 1px dashed #FFFFFF !important;
    border-radius: 8px !important;
}

div[data-testid="stFileUploader"] [data-testid="stFileUploaderDropzone"] * {
    color: #FFFFFF !important;
}

div[data-testid="stFileUploader"] button {
    background: #FFFFFF !important;
    color: #111111 !important;
    border: 1px solid #FFFFFF !important;
}

div[data-testid="stFileUploader"] button * {
    color: #111111 !important;
    fill: #111111 !important;
    stroke: #111111 !important;
}

div[data-testid="stFileUploader"] [data-testid="stFileUploaderFile"] {
    background: #FFFFFF !important;
    border: 1px solid #111111 !important;
}

div[data-testid="stFileUploader"] [data-testid="stFileUploaderFile"] *,
div[data-testid="stFileUploader"] [data-testid="stFileUploaderFileName"],
div[data-testid="stFileUploader"] [data-testid="stFileUploaderFileSize"] {
    color: #111111 !important;
}

div[data-testid="stFileUploader"] [data-testid="stFileUploaderFile"] button,
div[data-testid="stFileUploader"] [data-testid="stFileUploaderFile"] button * {
    background: #FFFFFF !important;
    color: #111111 !important;
    fill: #111111 !important;
    stroke: #111111 !important;
}


/* ===== FINAL UPLOAD / CLEAN-DOWNLOAD BUTTON COLORS =====
   Change ONLY the clickable controls requested by the user.
   The surrounding uploader/dropzone is intentionally untouched.
*/
/* Upload-file/Browse button: black */
div[data-testid="stFileUploader"] button:not([aria-label*="Remove"]):not([aria-label*="remove"]):not([aria-label*="Delete"]):not([aria-label*="delete"]):not([title*="Remove"]):not([title*="remove"]) {
    background:#000000 !important;
    border-color:#000000 !important;
    color:#FFFFFF !important;
}
div[data-testid="stFileUploader"] button:not([aria-label*="Remove"]):not([aria-label*="remove"]):not([aria-label*="Delete"]):not([aria-label*="delete"]):not([title*="Remove"]):not([title*="remove"]) * {
    color:#FFFFFF !important;
}
/* Cleaned-Excel download button: white */
div[data-testid="stDownloadButton"] button {
    background:#FFFFFF !important;
    border:1px solid #CFDCE8 !important;
    color:#000000 !important;
}
div[data-testid="stDownloadButton"] button * {
    color:#000000 !important;
}


/* ===== LIGHT INPUT CONTROLS =====
   Keep the page light: selectboxes, multiselects and threshold input
   must use a white field instead of the dark Streamlit theme surface.
*/
div[data-testid="stMainBlockContainer"] [data-testid="stSelectbox"] [data-baseweb="select"] > div,
div[data-testid="stMainBlockContainer"] [data-testid="stMultiSelect"] [data-baseweb="select"] > div,
div[data-testid="stMainBlockContainer"] [data-testid="stNumberInput"] > div,
div[data-testid="stMainBlockContainer"] [data-testid="stNumberInput"] [data-baseweb="input"] > div {
    background-color:#FFFFFF !important;
    background:#FFFFFF !important;
    border:1px solid #C9D7E5 !important;
    color:#000000 !important;
}

div[data-testid="stMainBlockContainer"] [data-testid="stSelectbox"] [data-baseweb="select"] > div:hover,
div[data-testid="stMainBlockContainer"] [data-testid="stMultiSelect"] [data-baseweb="select"] > div:hover,
div[data-testid="stMainBlockContainer"] [data-testid="stNumberInput"] > div:hover {
    background-color:#FFFFFF !important;
}

div[data-testid="stMainBlockContainer"] [data-testid="stSelectbox"] [data-baseweb="select"] span,
div[data-testid="stMainBlockContainer"] [data-testid="stSelectbox"] [data-baseweb="select"] input,
div[data-testid="stMainBlockContainer"] [data-testid="stMultiSelect"] [data-baseweb="select"] span,
div[data-testid="stMainBlockContainer"] [data-testid="stMultiSelect"] [data-baseweb="select"] input,
div[data-testid="stMainBlockContainer"] [data-testid="stNumberInput"] input {
    color:#000000 !important;
    -webkit-text-fill-color:#000000 !important;
    background:transparent !important;
}

div[data-testid="stMainBlockContainer"] [data-testid="stSelectbox"] [data-baseweb="select"] svg,
div[data-testid="stMainBlockContainer"] [data-testid="stMultiSelect"] [data-baseweb="select"] svg,
div[data-testid="stMainBlockContainer"] [data-testid="stNumberInput"] svg {
    color:#000000 !important;
    fill:#000000 !important;
    stroke:#000000 !important;
}

/* Dropdown menus opened from the controls */
div[data-baseweb="popover"],
div[data-baseweb="popover"] > div,
ul[role="listbox"],
div[role="listbox"] {
    background:#FFFFFF !important;
    color:#000000 !important;
}
div[data-baseweb="popover"] [role="option"],
div[data-baseweb="popover"] [role="option"] * ,
ul[role="listbox"] [role="option"],
div[role="listbox"] [role="option"] {
    background:#FFFFFF !important;
    color:#000000 !important;
}
div[data-baseweb="popover"] [role="option"]:hover,
div[data-baseweb="popover"] [aria-selected="true"],
ul[role="listbox"] [aria-selected="true"] {
    background:#EEF5FF !important;
    color:#000000 !important;
}

/* Number-input +/- controls remain light and readable */
div[data-testid="stMainBlockContainer"] [data-testid="stNumberInput"] button {
    background:#FFFFFF !important;
    color:#000000 !important;
    border-color:#C9D7E5 !important;
}
div[data-testid="stMainBlockContainer"] [data-testid="stNumberInput"] button * {
    color:#000000 !important;
    fill:#000000 !important;
    stroke:#000000 !important;
}

</style>""",unsafe_allow_html=True)

# ============================================================

def _analysis_number(value):
    try:
        return round(float(value), 2)
    except (TypeError, ValueError):
        return value


def build_graph_analysis(df, x_axis, y_axis):
    """Create a simple-language description of one selected graph."""
    work = df[[x_axis, y_axis]].copy()

    missing_x = int(work[x_axis].isna().sum())
    missing_y = int(work[y_axis].isna().sum())

    work[y_axis] = pd.to_numeric(
        work[y_axis],
        errors="coerce"
    )

    valid = work.dropna(
        subset=[x_axis, y_axis]
    ).copy()

    if valid.empty:
        return {
            "mean": None,
            "median": None,
            "minimum": None,
            "maximum": None,
            "analysis_points": [
                f"The X-axis is {x_axis}.",
                f"The Y-axis is {y_axis}.",
                "There is not enough valid numeric data to describe this graph."
            ],
            "missing_x": missing_x,
            "missing_y": missing_y,
            "min_x": None,
            "max_x": None
        }

    values = valid[y_axis]

    mean_value = values.mean()
    median_value = values.median()
    minimum = values.min()
    maximum = values.max()

    min_index = values.idxmin()
    max_index = values.idxmax()

    min_x = valid.loc[min_index, x_axis]
    max_x = valid.loc[max_index, x_axis]

    first_value = float(values.iloc[0])
    last_value = float(values.iloc[-1])

    if abs(first_value) > 1e-12:
        change_percent = (
            (last_value - first_value)
            / abs(first_value)
        ) * 100
    else:
        change_percent = None

    if change_percent is None or abs(change_percent) < 2:
        trend = (
            "The parameter is broadly stable from the beginning "
            "to the end of the recorded data."
        )
    elif change_percent > 0:
        trend = (
            f"The parameter shows an overall increase of about "
            f"{abs(change_percent):.1f}% from the first valid reading "
            f"to the last."
        )
    else:
        trend = (
            f"The parameter shows an overall decrease of about "
            f"{abs(change_percent):.1f}% from the first valid reading "
            f"to the last."
        )

    value_range = maximum - minimum

    if abs(mean_value) > 1e-12:
        variation = (value_range / abs(mean_value)) * 100
    else:
        variation = None

    if variation is not None and variation < 5:
        variation_text = "The values show only small variation."
    elif variation is not None and variation < 20:
        variation_text = "The values show moderate variation."
    else:
        variation_text = "The values show noticeable variation."

    points = [
        f"The X-axis represents {x_axis}.",
        f"The Y-axis represents {y_axis}.",
        f"The average value is {_analysis_number(mean_value)} "
        f"and the middle value is {_analysis_number(median_value)}.",
        f"The lowest value is {_analysis_number(minimum)} "
        f"at {x_axis} = {_analysis_number(min_x)}.",
        f"The highest value is {_analysis_number(maximum)} "
        f"at {x_axis} = {_analysis_number(max_x)}.",
        trend,
        variation_text
    ]

    if missing_x or missing_y:
        points.append(
            f"There are {missing_x} missing X value(s) and "
            f"{missing_y} missing Y value(s)."
        )

    return {
        "mean": _analysis_number(mean_value),
        "median": _analysis_number(median_value),
        "minimum": _analysis_number(minimum),
        "maximum": _analysis_number(maximum),
        "analysis_points": points,
        "missing_x": missing_x,
        "missing_y": missing_y,
        "min_x": min_x,
        "max_x": max_x
    }


# ============================================================
# ============================================================
# SIDEBAR
# ============================================================

with st.sidebar:
    st.markdown(f"""<div class="hal-sidebar-brand"><img class="hal-real-logo" src="data:image/png;base64,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"/><div class="hal-company">HINDUSTAN AERONAUTICS LIMITED</div><div class="hal-company-sub">AEROSPACE • DEFENCE • ENGINEERING</div></div><div class="side-heading">ANALYSIS MODULE</div>""",unsafe_allow_html=True)
    side_tabs=[("Data Input","01   Data Input"),("Visualization","02   Visualization"),("Analysis","03   Analysis"),("Anomaly Detection","04   Anomaly Detection"),("Export","05   Export Report")]
    for tab_key,label in side_tabs:
        if st.session_state.active_tab==tab_key: st.button(label,use_container_width=True,type="primary",key=f"side_active_{tab_key}")
        elif st.button(label,use_container_width=True,key=f"side_nav_{tab_key}"):
            change_tab(tab_key); st.rerun()
    st.markdown(f"""<div class="side-heading" style="margin-top:22px;">SYSTEM STATUS</div><div class="side-status"><div class="side-status-title">● SYSTEM READY</div><div class="side-status-row"><span>Backend</span><strong>Online</strong></div><div class="side-status-row"><span>Data File</span><strong>{'Loaded' if st.session_state.file_loaded else 'Waiting'}</strong></div></div><div class="side-aircraft"><img src="data:image/png;base64,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"/></div><div class="side-motto">DESIGN • DEVELOP • DELIVER<br>FOR A <span>SELF-RELIANT</span> INDIA</div>""",unsafe_allow_html=True)


# ============================================================
# MAIN HEADER
# ============================================================

st.markdown("""<div class="top-header"><div class="top-header-left"><div class="top-heli-icon">🚁</div><div class="top-title">HAL • HELICOPTER FLIGHT DATA ANALYSIS SYSTEM</div></div><div class="top-actions"><div class="top-ready"><span class="ready-dot">●</span>&nbsp; SYSTEM READY</div><div class="top-date">03 Sep 2026&nbsp;&nbsp;|&nbsp;&nbsp;23:50</div></div></div>""",unsafe_allow_html=True)

def _safe_text(value):
    """Convert report values to clean printable text."""
    if value is None:
        return "Not available"
    try:
        if pd.isna(value):
            return "Not available"
    except Exception:
        pass
    return str(value)


def _fmt_number(value):
    if value is None:
        return "Not available"
    try:
        if pd.isna(value):
            return "Not available"
        return f"{float(value):,.4f}".rstrip("0").rstrip(".")
    except Exception:
        return str(value)


def _prepare_export_report():
    """Prepare only graph image metadata and Maximum/Minimum values for export."""
    graph_df = st.session_state.get("graph_data")
    if graph_df is None or graph_df.empty:
        return None

    x_axis = st.session_state.get("x_axis") or "Row Serial Number"
    y_axes = st.session_state.get("y_axis") or st.session_state.get("selected_attributes", [])

    if x_axis not in graph_df.columns:
        x_axis = "Row Serial Number"

    y_axes = [y for y in y_axes if y in graph_df.columns]
    if not y_axes:
        return None

    graphs = []
    for number, y_axis in enumerate(y_axes, start=1):
        analysis = build_graph_analysis(graph_df, x_axis, y_axis)
        graphs.append({
            "graph_number": number,
            "x_axis": x_axis,
            "y_axis": y_axis,
            "maximum": analysis.get("maximum"),
            "minimum": analysis.get("minimum")
        })

    return {"graphs": graphs}


def _prepare_report_time_axis(series, x_axis_name):
    """Convert Time values only for report display; never modify source data."""
    if str(x_axis_name).strip().lower() != "time":
        return series

    if pd.api.types.is_datetime64_any_dtype(series):
        return series

    # Excel stores times as fractions of a day (e.g. 0.41669 = 10:00:02).
    numeric = pd.to_numeric(series, errors="coerce")
    numeric_ratio = numeric.notna().sum() / max(1, series.notna().sum())
    if numeric_ratio >= 0.80:
        return pd.Timestamp("1970-01-01") + pd.to_timedelta(numeric, unit="D")

    # Handle text times such as 10:00:02.
    parsed = pd.to_datetime(series.astype(str).str.strip(), errors="coerce")
    parsed_ratio = parsed.notna().sum() / max(1, series.notna().sum())
    if parsed_ratio >= 0.80:
        return parsed

    return series


def _make_graph_image(report_graph, graph_df, output_dir, prefix="graph"):
    """Create the report graph from the exact same X/Y data preparation used by Visualization."""
    x_axis = report_graph["x_axis"]
    y_axis = report_graph["y_axis"]
    graph_type = st.session_state.get("graph_type", "Line")

    # Use the same X conversion and selected X/Y columns as Visualization.
    plot_df = graph_df[[x_axis, y_axis]].copy()
    plot_df[x_axis] = _prepare_plot_x(plot_df[x_axis], x_axis)
    plot_df[y_axis] = pd.to_numeric(plot_df[y_axis], errors="coerce")
    plot_df = plot_df.dropna(subset=[x_axis, y_axis])

    # Use the same browser-facing downsampling routine as Visualization so
    # the report follows the same plotted readings rather than rebuilding a
    # different graph from a separate sample.
    display_graph = _downsample_for_plot(plot_df) if not plot_df.empty else plot_df

    import matplotlib.pyplot as plt
    from matplotlib.ticker import MaxNLocator
    fig, ax = plt.subplots(figsize=(10.8, 5.8), dpi=150)
    if display_graph.empty:
        ax.text(0.5, 0.5, "No plottable numeric data", ha="center", va="center")
        ax.set_axis_off()
    else:
        x_values = display_graph[x_axis]
        y_values = pd.to_numeric(display_graph[y_axis], errors="coerce")

        if graph_type == "Bar":
            if pd.api.types.is_datetime64_any_dtype(x_values):
                ax.bar(x_values, y_values, width=0.0014)  # about 2 minutes
            else:
                ax.bar(x_values, y_values, width=0.8)
        elif graph_type == "Scatter":
            ax.scatter(x_values, y_values, s=18)
        else:
            ax.plot(x_values, y_values, linewidth=1.8)

        ax.set_title(f"{y_axis} vs {x_axis}", loc="left", fontsize=15, fontweight="bold")
        ax.set_xlabel(x_axis, fontsize=11)
        ax.set_ylabel(y_axis, fontsize=11)
        ax.tick_params(axis="both", labelsize=9)
        ax.grid(True, alpha=0.25)

        # Match Visualization: Y-axis starts at zero and uses the maximum
        # of the same displayed readings.
        y_max = y_values.max()
        ax.set_ylim(0, max(1.0, float(y_max) * 1.05) if pd.notna(y_max) else 1.0)

        # Match Visualization X-axis tick spacing.
        if pd.api.types.is_datetime64_any_dtype(x_values):
            import matplotlib.dates as mdates
            # Keep the requested 2-minute tick interval, but rotate the
            # labels so adjacent HH:MM:SS values do not overlap in reports.
            ax.xaxis.set_major_locator(mdates.MinuteLocator(interval=2))
            ax.xaxis.set_major_formatter(mdates.DateFormatter("%H:%M:%S"))
            ax.tick_params(axis="x", labelsize=7, pad=2)
            plt.setp(ax.get_xticklabels(), rotation=45, ha="right")
        elif pd.api.types.is_numeric_dtype(x_values):
            ax.xaxis.set_major_locator(MaxNLocator(nbins=8, steps=[2]))

        fig.tight_layout()

    path = os.path.join(output_dir, f"{prefix}_{report_graph['graph_number']}.png")
    fig.savefig(path, bbox_inches="tight", facecolor="white")
    plt.close(fig)
    return path

def _make_anomaly_image(anomaly_item, graph_df, output_dir):
    x_axis = anomaly_item["x_axis"]
    y_axis = anomaly_item["y_axis"]
    details = anomaly_item.get("anomalies", [])

    plot_df = graph_df[[x_axis, y_axis]].copy()
    plot_df[y_axis] = pd.to_numeric(plot_df[y_axis], errors="coerce")
    plot_df = plot_df.dropna(subset=[x_axis, y_axis])

    import matplotlib.pyplot as plt
    from matplotlib.ticker import MaxNLocator
    fig, ax = plt.subplots(figsize=(10.8, 5.8), dpi=150)
    if not plot_df.empty:
        ax.plot(plot_df[x_axis], plot_df[y_axis], linewidth=1.8, marker="o", markersize=2.5, label="Normal readings")
    if details:
        adf = pd.DataFrame(details)
        adf["value"] = pd.to_numeric(adf["value"], errors="coerce")
        adf = adf.dropna(subset=["value"])
        if not adf.empty:
            ax.scatter(adf["x"], adf["value"], s=65, facecolors="white", edgecolors="#DC2626", linewidths=2, label="Anomaly", zorder=5)
    ax.set_title(f"Anomaly Detection — {y_axis} vs {x_axis}", loc="left", fontsize=15, fontweight="bold")
    ax.set_xlabel(x_axis, fontsize=11, color="#111827")
    ax.set_ylabel(y_axis, fontsize=11, color="#111827")
    ax.tick_params(axis="both", labelsize=9, colors="#1F2937")
    ax.grid(True, alpha=0.25)
    ax.legend()
    fig.tight_layout()
    path = os.path.join(output_dir, f"anomaly_{anomaly_item['graph_number']}.png")
    fig.savefig(path, bbox_inches="tight", facecolor="white")
    plt.close(fig)
    return path


def _build_export_images(report):
    graph_df = st.session_state.graph_data.copy()
    if "Row Serial Number" not in graph_df.columns:
        graph_df.insert(0, "Row Serial Number", range(1, len(graph_df) + 1))
    output_dir = os.path.join(PROJECT_ROOT, "output", "report_images")
    os.makedirs(output_dir, exist_ok=True)

    graph_images = []
    for item in report["graphs"]:
        graph_images.append(_make_graph_image(item, graph_df, output_dir, "graph"))

    anomaly_images = []
    for item in report.get("anomalies", []):
        anomaly_images.append(_make_anomaly_image(item, graph_df, output_dir))
    return graph_images, anomaly_images


def _add_docx_heading(document, text, level=1):
    p = document.add_heading(text, level=level)
    if p.runs:
        p.runs[0].font.name = "Aptos"
    return p


def _generate_docx(report, graph_images, anomaly_images):
    """Word report: graph + Maximum + Minimum for every graph, and nothing else."""
    from docx import Document
    from docx.shared import Inches, Pt
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    from docx.enum.table import WD_TABLE_ALIGNMENT
    doc = Document()
    section = doc.sections[0]
    section.top_margin = Inches(0.6)
    section.bottom_margin = Inches(0.6)
    section.left_margin = Inches(0.7)
    section.right_margin = Inches(0.7)

    title = doc.add_paragraph()
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = title.add_run("HELICOPTER ANALYSIS SYSTEM")
    r.bold = True
    r.font.size = Pt(20)

    for item, image_path in zip(report.get("graphs", []), graph_images):
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        r = p.add_run(
            f"Graph {item.get('graph_number', '')} — "
            f"{item.get('y_axis', '')} vs {item.get('x_axis', '')}"
        )
        r.bold = True
        r.font.size = Pt(13)

        doc.add_picture(image_path, width=Inches(6.8))
        doc.paragraphs[-1].alignment = WD_ALIGN_PARAGRAPH.CENTER

        table = doc.add_table(rows=1, cols=2)
        table.alignment = WD_TABLE_ALIGNMENT.CENTER
        table.style = "Table Grid"
        table.rows[0].cells[0].text = f"Maximum\n{_fmt_number(item.get('maximum'))}"
        table.rows[0].cells[1].text = f"Minimum\n{_fmt_number(item.get('minimum'))}"
        for cell in table.rows[0].cells:
            for run in cell.paragraphs[0].runs:
                run.bold = True

        doc.add_paragraph()

    out = io.BytesIO()
    doc.save(out)
    out.seek(0)
    return out.getvalue()



def _generate_pdf(report, graph_images, anomaly_images):
    """PDF report: graph + Maximum + Minimum for every graph, and nothing else."""
    out = io.BytesIO()
    from reportlab.lib import colors
    from reportlab.lib.enums import TA_CENTER
    from reportlab.lib.pagesizes import A4
    from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
    from reportlab.lib.units import inch
    from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Image, Table, TableStyle, PageBreak
    doc = SimpleDocTemplate(
        out, pagesize=A4,
        rightMargin=36, leftMargin=36, topMargin=36, bottomMargin=36
    )
    styles = getSampleStyleSheet()
    styles.add(ParagraphStyle(
        name="GraphReportTitle",
        parent=styles["Title"],
        alignment=TA_CENTER,
        fontSize=18,
        leading=22,
        spaceAfter=16
    ))

    story = [Paragraph("HELICOPTER ANALYSIS SYSTEM", styles["GraphReportTitle"])]

    for item, image_path in zip(report.get("graphs", []), graph_images):
        story.append(Paragraph(
            f"Graph {item.get('graph_number', '')} — "
            f"{_safe_text(item.get('y_axis'))} vs "
            f"{_safe_text(item.get('x_axis'))}",
            styles["Heading2"]
        ))
        story.append(Image(image_path, width=6.7 * inch, height=3.55 * inch))
        story.append(Spacer(1, 7))

        stats = Table(
            [["Maximum", "Minimum"],
             [_fmt_number(item.get("maximum")), _fmt_number(item.get("minimum"))]],
            colWidths=[3.25 * inch, 3.25 * inch]
        )
        stats.setStyle(TableStyle([
            ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#EAF2F8")),
            ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
            ("FONTNAME", (0, 1), (-1, 1), "Helvetica-Bold"),
            ("GRID", (0, 0), (-1, -1), .4, colors.HexColor("#CBD5E1")),
            ("ALIGN", (0, 0), (-1, -1), "CENTER"),
            ("FONTSIZE", (0, 0), (-1, -1), 9),
        ]))
        story += [stats, Spacer(1, 18)]

    doc.build(story)
    out.seek(0)
    return out.getvalue()



def _generate_xlsx(report, graph_images, anomaly_images):
    """Excel report: graph + Maximum + Minimum for every graph, and nothing else."""
    wb = Workbook()
    ws = wb.active
    ws.title = "Graph Report"

    ws["A1"] = "HELICOPTER ANALYSIS SYSTEM"
    ws["A1"].font = XLFont(size=18, bold=True)
    ws["A2"] = "Flight Data Graph Report"
    ws["A2"].font = XLFont(size=12, italic=True)

    row = 4
    for item, image_path in zip(report.get("graphs", []), graph_images):
        ws.cell(
            row, 1,
            f"Graph {item.get('graph_number', '')} — "
            f"{item.get('y_axis', '')} vs {item.get('x_axis', '')}"
        ).font = XLFont(bold=True, size=12)

        ws.cell(row + 1, 1, "Maximum").font = XLFont(bold=True)
        ws.cell(row + 1, 2, item.get("maximum"))
        ws.cell(row + 2, 1, "Minimum").font = XLFont(bold=True)
        ws.cell(row + 2, 2, item.get("minimum"))

        img = XLImage(image_path)
        img.width = 720
        img.height = 385
        ws.add_image(img, f"D{row}")

        row += 27

    ws.column_dimensions["A"].width = 18
    ws.column_dimensions["B"].width = 18

    out = io.BytesIO()
    wb.save(out)
    out.seek(0)
    return out.getvalue()



def _blank_workflow_state():
    return {
        "file_loaded": False,
        "uploaded_filename": None,
        "uploaded_file_signature": None,
        "converted_file": None,
        "analysis_result": None,
        "selected_attributes": [],
        "graph_ready": False,
        "graph_data": None,
        "x_axis": None,
        "y_axis": [],
        "graph_type": "Line",
        "report_data": {
            "file_information": {},
            "graphs": [],
            "dataset_issues": [],
            "summary": ""
        },
        "anomaly_report": [],
        "anomaly_issues": [],
        "anomaly_summary": "",
        "export_docx": None,
        "export_pdf": None,
        "export_xlsx": None
    }


def _save_active_workspace():
    signature = st.session_state.get("uploaded_file_signature")
    if signature is None:
        return

    st.session_state.workspace_by_signature[signature] = {
        key: st.session_state.get(key)
        for key in (
            "file_loaded",
            "uploaded_filename",
            "uploaded_file_signature",
            "converted_file",
            "analysis_result",
            "selected_attributes",
            "graph_ready",
            "graph_data",
            "x_axis",
            "y_axis",
            "graph_type",
            "report_data",
            "anomaly_report",
            "anomaly_issues",
            "anomaly_summary",
            "export_docx",
            "export_pdf",
            "export_xlsx"
        )
    }


def _restore_workspace(signature, filename=None):
    saved = st.session_state.workspace_by_signature.get(signature)

    if saved is None:
        saved = _blank_workflow_state()

    for key, value in saved.items():
        st.session_state[key] = value

    st.session_state.file_loaded = True
    st.session_state.uploaded_filename = filename or st.session_state.get(
        "uploaded_filename"
    )
    st.session_state.uploaded_file_signature = signature


def _clear_active_workspace():
    signature = st.session_state.get("uploaded_file_signature")

    if signature is not None:
        st.session_state.workspace_by_signature.pop(signature, None)

    blank = _blank_workflow_state()
    for key, value in blank.items():
        st.session_state[key] = value


# ============================================================
# LARGE DATASET LOADER
# ============================================================
# The supplied HAL test file is ~59 MB and contains 180,000
# telemetry rows. Some Excel readers spend a very long time
# trying to infer worksheet dimensions when the XML has no
# <dimension> tag. For .xlsx files we therefore read the
# worksheet XML directly in streaming mode.
#
# This keeps the application responsive and avoids loading
# the workbook repeatedly on every Streamlit rerun.
# ============================================================

_XLSX_MAIN_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
_XLSX_REL_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
_PACKAGE_REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"


def _excel_col_index(cell_ref):
    letters = re.match(r"[A-Z]+", cell_ref or "")
    if not letters:
        return None

    index = 0
    for char in letters.group(0):
        index = index * 26 + (ord(char) - 64)

    return index - 1


def _xlsx_sheet_paths(xlsx_path):
    """Return workbook sheet names mapped to worksheet XML paths."""
    with zipfile.ZipFile(xlsx_path, "r") as archive:
        workbook_root = ET.fromstring(archive.read("xl/workbook.xml"))
        rels_root = ET.fromstring(archive.read("xl/_rels/workbook.xml.rels"))

        relationships = {}
        for rel in rels_root:
            rel_id = rel.attrib.get("Id")
            target = rel.attrib.get("Target", "")
            if rel_id:
                relationships[rel_id] = target

        result = []

        for sheet in workbook_root.find(
            f"{{{_XLSX_MAIN_NS}}}sheets"
        ):
            name = sheet.attrib.get("name", "")
            rel_id = sheet.attrib.get(
                f"{{{_XLSX_REL_NS}}}id"
            )

            target = relationships.get(rel_id, "")
            if target.startswith("/"):
                xml_path = target.lstrip("/")
            else:
                xml_path = "xl/" + target.lstrip("./")

            xml_path = xml_path.replace("xl/xl/", "xl/")
            result.append((name, xml_path))

        return result


def _read_xlsx_streaming(xlsx_path):
    """
    Read the first substantial worksheet from an XLSX file using
    XML streaming instead of pandas/openpyxl workbook discovery.

    Designed for large telemetry files with numeric and inline-string
    cells, including files where worksheet dimensions are not declared.
    """
    with zipfile.ZipFile(xlsx_path, "r") as archive:
        sheet_paths = _xlsx_sheet_paths(xlsx_path)

        if not sheet_paths:
            raise ValueError("The Excel workbook contains no worksheets.")

        # Prefer a telemetry/data sheet. Otherwise choose the first sheet.
        preferred = None
        for sheet_name, xml_path in sheet_paths:
            name_lower = sheet_name.lower()
            if any(
                token in name_lower
                for token in ("telemetry", "data", "flight", "test")
            ):
                preferred = (sheet_name, xml_path)
                break

        candidates = [preferred] if preferred else []
        candidates.extend(
            item for item in sheet_paths if item != preferred
        )

        selected_name = None
        selected_path = None

        # Read candidates until one has a real header row and data.
        for sheet_name, xml_path in candidates:
            try:
                with archive.open(xml_path, "r") as stream:
                    header_values = None
                    data_rows = []
                    max_col = 0

                    for _, row_element in ET.iterparse(
                        stream,
                        events=("end",)
                    ):
                        if row_element.tag != f"{{{_XLSX_MAIN_NS}}}row":
                            continue

                        row_values = {}

                        for cell in row_element:
                            if cell.tag != f"{{{_XLSX_MAIN_NS}}}c":
                                continue

                            ref = cell.attrib.get("r", "")
                            col_index = _excel_col_index(ref)

                            if col_index is None:
                                continue

                            max_col = max(max_col, col_index + 1)

                            value_node = cell.find(
                                f"{{{_XLSX_MAIN_NS}}}v"
                            )
                            inline_node = cell.find(
                                f"{{{_XLSX_MAIN_NS}}}is"
                            )

                            if inline_node is not None:
                                text_node = inline_node.find(
                                    f"{{{_XLSX_MAIN_NS}}}t"
                                )
                                value = (
                                    text_node.text
                                    if text_node is not None
                                    else ""
                                )
                            elif value_node is not None:
                                value = value_node.text

                                cell_type = cell.attrib.get("t")

                                if cell_type == "b":
                                    value = value == "1"
                                elif cell_type == "n" or cell_type is None:
                                    try:
                                        number = float(value)
                                        value = (
                                            int(number)
                                            if number.is_integer()
                                            else number
                                        )
                                    except (TypeError, ValueError):
                                        pass
                            else:
                                value = None

                            row_values[col_index] = value

                        if header_values is None and row_values:
                            header_values = [
                                row_values.get(i)
                                for i in range(max_col)
                            ]

                            # If this is an obviously non-data sheet,
                            # continue to the next candidate.
                            header_text = " ".join(
                                str(v).lower()
                                for v in header_values
                                if v is not None
                            )

                            if not any(
                                token in header_text
                                for token in (
                                    "time",
                                    "date",
                                    "sample",
                                    "rpm",
                                    "pressure",
                                    "temp",
                                    "fuel",
                                    "ng",
                                )
                            ):
                                break

                        elif header_values is not None:
                            values = [
                                row_values.get(i)
                                for i in range(len(header_values))
                            ]
                            data_rows.append(values)

                            # A sheet with at least one data row is enough.
                            if len(data_rows) >= 1:
                                selected_name = sheet_name
                                selected_path = xml_path
                                break

                        row_element.clear()

                if selected_path:
                    break

            except KeyError:
                continue

        if not selected_path or not header_values:
            raise ValueError(
                "No usable telemetry/data worksheet was found in the Excel file."
            )

        # Re-open the selected worksheet and read the complete dataset.
        # This second pass is deliberate: it avoids retaining rows from
        # candidate sheets and keeps the logic deterministic.
        rows = []

        with archive.open(selected_path, "r") as stream:
            headers = None

            for _, row_element in ET.iterparse(
                stream,
                events=("end",)
            ):
                if row_element.tag != f"{{{_XLSX_MAIN_NS}}}row":
                    continue

                row_values = {}

                for cell in row_element:
                    if cell.tag != f"{{{_XLSX_MAIN_NS}}}c":
                        continue

                    col_index = _excel_col_index(
                        cell.attrib.get("r", "")
                    )

                    if col_index is None:
                        continue

                    value_node = cell.find(
                        f"{{{_XLSX_MAIN_NS}}}v"
                    )
                    inline_node = cell.find(
                        f"{{{_XLSX_MAIN_NS}}}is"
                    )

                    if inline_node is not None:
                        text_node = inline_node.find(
                            f"{{{_XLSX_MAIN_NS}}}t"
                        )
                        value = (
                            text_node.text
                            if text_node is not None
                            else ""
                        )
                    elif value_node is not None:
                        value = value_node.text
                        cell_type = cell.attrib.get("t")

                        if cell_type == "b":
                            value = value == "1"
                        else:
                            try:
                                number = float(value)
                                value = (
                                    int(number)
                                    if number.is_integer()
                                    else number
                                )
                            except (TypeError, ValueError):
                                pass
                    else:
                        value = None

                    row_values[col_index] = value

                if headers is None and row_values:
                    max_col = max(row_values.keys()) + 1
                    headers = [
                        row_values.get(i)
                        for i in range(max_col)
                    ]

                    # Make column names safe and unique.
                    safe_headers = []
                    used = {}

                    for index, header in enumerate(headers):
                        name = str(header).strip() if header is not None else ""
                        if not name:
                            name = f"Column_{index + 1}"

                        if name in used:
                            used[name] += 1
                            name = f"{name}_{used[name]}"
                        else:
                            used[name] = 1

                        safe_headers.append(name)

                    headers = safe_headers

                elif headers is not None:
                    rows.append([
                        row_values.get(i)
                        for i in range(len(headers))
                    ])

                row_element.clear()

        dataframe = pd.DataFrame(rows, columns=headers)

        return dataframe, selected_name



# ============================================================
# FAST / LAZY XLSX WORKFLOW
# ============================================================
# IMPORTANT:
# Uploading a 60 MB XLSX should NOT mean reading all 180,000
# rows immediately.
#
# Step A: on upload -> read only workbook metadata + header.
#          This makes the attributes/X/Y options appear quickly.
#
# Step B: on Generate Graph -> read ONLY the selected X/Y
#          columns from the worksheet.
#
# The original full-data reader is kept below for compatibility,
# but the normal UI path uses these faster functions.
# ============================================================

_XLSX_ROW_END = re.compile(rb"</row>")


def _safe_column_name(name, index, used):
    value = str(name).strip() if name is not None else ""
    if not value:
        value = f"Column_{index + 1}"

    if value in used:
        used[value] += 1
        value = f"{value}_{used[value]}"
    else:
        used[value] = 1

    return value


def _read_xlsx_header_fast(xlsx_path):
    """
    Read only the first worksheet header.

    For the supplied 60 MB HAL workbook this avoids parsing the
    remaining ~180,000 data rows just to populate the UI.
    """
    with zipfile.ZipFile(xlsx_path, "r") as archive:
        sheet_paths = _xlsx_sheet_paths(xlsx_path)

        if not sheet_paths:
            raise ValueError("The Excel workbook contains no worksheets.")

        selected_name = None
        selected_path = None

        for sheet_name, xml_path in sheet_paths:
            try:
                with archive.open(xml_path, "r") as stream:
                    buffer = b""

                    while len(buffer) < 2 * 1024 * 1024:
                        chunk = stream.read(256 * 1024)
                        if not chunk:
                            break

                        buffer += chunk
                        end = buffer.find(b"</row>")

                        if end == -1:
                            continue

                        first_row = buffer[:end + len(b"</row>")]
                        match = re.search(
                            rb"<row[^>]*>(.*?)</row>",
                            first_row,
                            flags=re.DOTALL
                        )

                        if not match:
                            break

                        row_xml = match.group(1)

                        # Check that this looks like a data sheet.
                        header_text = row_xml.decode(
                            "utf-8",
                            errors="ignore"
                        ).lower()

                        if any(
                            token in header_text
                            for token in (
                                "time",
                                "date",
                                "sample",
                                "rpm",
                                "pressure",
                                "temp",
                                "fuel",
                                "ng"
                            )
                        ):
                            selected_name = sheet_name
                            selected_path = xml_path

                        break

                if selected_path:
                    break

            except Exception:
                continue

        if not selected_path:
            raise ValueError(
                "No usable telemetry/data worksheet was found."
            )

        with archive.open(selected_path, "r") as stream:
            buffer = b""

            while True:
                chunk = stream.read(256 * 1024)

                if not chunk:
                    break

                buffer += chunk
                end = buffer.find(b"</row>")

                if end != -1:
                    first_row = buffer[:end + len(b"</row>")]
                    break

                if len(buffer) > 4 * 1024 * 1024:
                    raise ValueError(
                        "The worksheet header could not be located."
                    )

        cells = re.findall(
            rb'<c r="[A-Z]+1"[^>]*>(.*?)</c>',
            first_row,
            flags=re.DOTALL
        )

        headers = []

        for cell_xml in cells:
            inline = re.search(
                rb"<t>(.*?)</t>",
                cell_xml,
                flags=re.DOTALL
            )

            if inline:
                value = inline.group(1).decode(
                    "utf-8",
                    errors="replace"
                )
            else:
                value_node = re.search(
                    rb"<v>(.*?)</v>",
                    cell_xml,
                    flags=re.DOTALL
                )
                value = (
                    value_node.group(1).decode(
                        "utf-8",
                        errors="replace"
                    )
                    if value_node
                    else ""
                )

            headers.append(value)

        used = {}
        headers = [
            _safe_column_name(
                value,
                index,
                used
            )
            for index, value in enumerate(headers)
        ]

        # Exact row count is intentionally not calculated here.
        # Counting all XML rows would defeat the fast-upload goal.
        return {
            "columns": headers,
            "worksheet": selected_name,
            "rows": None
        }


def _decode_xlsx_cell(cell_xml):
    """Decode a basic XLSX inline-string/numeric cell."""
    inline = re.search(
        rb"<is>.*?<t[^>]*>(.*?)</t>.*?</is>",
        cell_xml,
        flags=re.DOTALL
    )

    if inline:
        return inline.group(1).decode(
            "utf-8",
            errors="replace"
        )

    value_node = re.search(
        rb"<v>(.*?)</v>",
        cell_xml,
        flags=re.DOTALL
    )

    if not value_node:
        return None

    raw = value_node.group(1).decode(
        "utf-8",
        errors="replace"
    )

    try:
        number = float(raw)

        if number.is_integer():
            return int(number)

        return number

    except ValueError:
        return raw


@st.cache_data(
    show_spinner=False,
    max_entries=10
)
def _load_xlsx_selected_columns(
    xlsx_path,
    file_signature,
    selected_columns
):
    """Load only requested XLSX columns using a targeted XML scan.

    The regex is restricted to the selected Excel columns, avoiding Python
    work for the other ~50 telemetry columns in a wide workbook. The result
    is cached, so generating the same graph again is effectively instant.
    """
    selected_columns = tuple(selected_columns)
    header_info = _read_xlsx_header_fast(xlsx_path)
    headers = header_info["columns"]
    worksheet = header_info["worksheet"]

    wanted_indices = []
    for column in selected_columns:
        if column not in headers:
            raise ValueError(f"Column '{column}' was not found.")
        wanted_indices.append(headers.index(column))

    # Convert zero-based indices to Excel letters once and build a targeted
    # cell regex. This avoids matching every cell in a wide telemetry row.
    def _excel_col_letters(index):
        n = int(index) + 1
        out = ""
        while n:
            n, rem = divmod(n - 1, 26)
            out = chr(65 + rem) + out
        return out

    wanted_by_letters = {
        _excel_col_letters(idx): pos
        for pos, idx in enumerate(wanted_indices)
    }
    letters_pattern = "|".join(
        re.escape(v) for v in wanted_by_letters.keys()
    )
    cell_pattern = re.compile(
        rb'<c r="(' + letters_pattern.encode("ascii") + rb')([0-9]+)"[^>]*>(.*?)</c>',
        flags=re.DOTALL
    )

    rows = []
    with zipfile.ZipFile(xlsx_path, "r") as archive:
        sheet_path = next(
            path for name, path in _xlsx_sheet_paths(xlsx_path)
            if name == worksheet
        )
        with archive.open(sheet_path, "r") as stream:
            buffer = b""
            header_seen = False
            while True:
                chunk = stream.read(8 * 1024 * 1024)
                if not chunk:
                    break
                buffer += chunk
                parts = buffer.split(b"</row>")
                buffer = parts.pop()

                for row_part in parts:
                    if b"<row" not in row_part:
                        continue
                    if not header_seen:
                        header_seen = True
                        continue

                    values = [None] * len(wanted_indices)
                    for match in cell_pattern.finditer(row_part):
                        col_letters = match.group(1).decode("ascii", errors="ignore")
                        target_position = wanted_by_letters.get(col_letters)
                        if target_position is not None:
                            values[target_position] = _decode_xlsx_cell(match.group(3))
                    rows.append(values)

    return pd.DataFrame(rows, columns=list(selected_columns)), worksheet


def _convert_003_to_xlsx(source_path):
    """Convert a text-based .003 telemetry file to XLSX."""
    source_path = os.path.abspath(source_path)
    output_path = os.path.splitext(source_path)[0] + "_converted.xlsx"

    if os.path.exists(output_path) and os.path.getmtime(output_path) >= os.path.getmtime(source_path):
        return output_path

    encodings = ["utf-8-sig", "utf-8", "cp1252", "latin1"]
    last_error = None
    dataframe = None

    for encoding in encodings:
        try:
            with open(source_path, "r", encoding=encoding, errors="strict") as fh:
                sample = fh.read(65536)

            # Remove empty lines and common BOM/control characters.
            sample = sample.replace("\\x00", "")
            lines = [line.strip() for line in sample.splitlines() if line.strip()]
            if not lines:
                raise ValueError("The .003 file is empty.")

            import csv
            delimiter = None
            try:
                dialect = csv.Sniffer().sniff("\\n".join(lines[:50]), delimiters=",;\\t|:")
                delimiter = dialect.delimiter
            except Exception:
                pass

            if delimiter:
                dataframe = pd.read_csv(
                    source_path,
                    sep=delimiter,
                    engine="python",
                    encoding=encoding,
                    on_bad_lines="warn"
                )
            else:
                # Handles whitespace-separated telemetry exports.
                dataframe = pd.read_csv(
                    source_path,
                    sep=r"\\s+",
                    engine="python",
                    encoding=encoding,
                    on_bad_lines="warn"
                )

            if dataframe is None or dataframe.empty or len(dataframe.columns) < 1:
                raise ValueError("The .003 file could not be converted to a table.")

            # Clean unnamed/blank headers while preserving all columns.
            cleaned = []
            for i, col in enumerate(dataframe.columns):
                name = str(col).replace("\\ufeff", "").strip()
                cleaned.append(name if name else f"Column_{i + 1}")
            dataframe.columns = cleaned
            break
        except Exception as exc:
            last_error = exc
            dataframe = None

    if dataframe is None:
        raise ValueError(
            "The .003 file could not be converted to a table. "
            f"Parser error: {last_error}"
        )

    dataframe.to_excel(output_path, index=False, engine="openpyxl")
    return output_path


def _prepare_003_upload(uploaded_file, input_directory):
    """Save an uploaded .003 file and return its converted XLSX path."""
    raw_path = os.path.join(input_directory, uploaded_file.name)
    with open(raw_path, "wb") as out:
        out.write(uploaded_file.getvalue())
    return _convert_003_to_xlsx(raw_path)


def _convert_any_tabular_to_xlsx(source_path, output_path):
    """Convert a non-XLSX tabular file to XLSX while preserving raw rows.

    The conversion intentionally keeps the first physical row as row 1 rather
    than treating it as a header. That is important because the first rows
    may contain unwanted/hash metadata that the user wants to delete first.
    """
    import csv as _csv

    ext = os.path.splitext(source_path)[1].lower()

    # Legacy Excel (.xls): preserve every physical row, including the first.
    if ext == ".xls" or ext == ".xlsb":
        try:
            engine = "pyxlsb" if ext == ".xlsb" else None
            df = pd.read_excel(source_path, header=None, engine=engine)
            wb = Workbook(write_only=True)
            ws = wb.create_sheet("Telemetry_Data")
            for row in df.itertuples(index=False, name=None):
                ws.append([None if pd.isna(v) else v for v in row])
            wb.save(output_path)
            return output_path
        except Exception as exc:
            raise ValueError(f"Could not convert {ext} file to Excel: {exc}")

    # JSON is common for machine-generated telemetry exports.
    if ext == ".json" or ext == ".jsonl" or ext == ".ndjson":
        try:
            if ext in (".jsonl", ".ndjson"):
                df = pd.read_json(source_path, lines=True)
            else:
                df = pd.read_json(source_path)
            df.to_excel(output_path, index=False, header=True, engine="openpyxl")
            return output_path
        except Exception as exc:
            raise ValueError(f"Could not convert JSON file to Excel: {exc}")

    # Parquet, when available, is also directly tabular.
    if ext == ".parquet":
        try:
            df = pd.read_parquet(source_path)
            df.to_excel(output_path, index=False, header=True, engine="openpyxl")
            return output_path
        except Exception as exc:
            raise ValueError(f"Could not convert Parquet file to Excel: {exc}")

    # Everything else is treated as a text/delimited telemetry export.
    # Read line-by-line so even .00.out/.out/.dat/.log and unknown extensions
    # can enter the same cleanup -> Excel -> analysis workflow.
    encodings = ("utf-8-sig", "utf-8", "cp1252", "latin1")
    last_error = None
    for encoding in encodings:
        try:
            with open(source_path, "r", encoding=encoding, errors="strict", newline="") as fh:
                sample = fh.read(65536)
                fh.seek(0)
                lines = sample.splitlines()
                if not lines:
                    raise ValueError("The uploaded file is empty.")

                try:
                    dialect = _csv.Sniffer().sniff(
                        "\n".join(lines[:50]),
                        delimiters=",;\\t|:\\x00"
                    )
                    delimiter = dialect.delimiter
                except Exception:
                    # Whitespace-separated telemetry is very common.
                    delimiter = None

                wb = Workbook(write_only=True)
                ws = wb.create_sheet("Telemetry_Data")
                reader = (
                    _csv.reader(fh, dialect)
                    if delimiter is not None
                    else _csv.reader(fh, delimiter=" ", skipinitialspace=True)
                )
                for row in reader:
                    if delimiter is None and len(row) == 1:
                        row = row[0].split()
                    ws.append(row)
                wb.save(output_path)
                return output_path
        except Exception as exc:
            last_error = exc

    raise ValueError(
        f"The uploaded file could not be converted to Excel. Parser error: {last_error}"
    )


def _normalize_upload_to_xlsx(raw_path, input_directory, signature):
    """Return an XLSX processing path for any supported/recognizable upload."""
    ext = os.path.splitext(raw_path)[1].lower()

    # Real XLSX files stay untouched and keep their original workbook.
    if ext == ".xlsx":
        return raw_path, False

    base = os.path.splitext(os.path.basename(raw_path))[0]
    safe_sig = re.sub(r"[^A-Za-z0-9_-]+", "_", str(signature))[:50]
    out_path = os.path.join(input_directory, f"{base}_{safe_sig}_converted.xlsx")

    if os.path.exists(out_path) and os.path.getsize(out_path) > 0:
        return out_path, True

    # Detect files whose extension is wrong but whose content is already XLSX.
    try:
        with open(raw_path, "rb") as fh:
            magic = fh.read(8)
        if magic[:2] == b"PK":
            import zipfile as _zipfile
            with _zipfile.ZipFile(raw_path, "r") as zf:
                if "[Content_Types].xml" in zf.namelist():
                    import shutil
                    shutil.copyfile(raw_path, out_path)
                    return out_path, True
    except Exception:
        pass

    _convert_any_tabular_to_xlsx(raw_path, out_path)
    return out_path, True


@st.cache_data(
    show_spinner=False,
    max_entries=10
)
def _read_file_header_cached(
    file_path,
    file_signature
):
    """
    Fast metadata-only read used immediately after upload.
    """
    extension = os.path.splitext(
        file_path
    )[1].lower()

    if extension == ".003":
        file_path = _convert_003_to_xlsx(file_path)
        extension = ".xlsx"

    if extension == ".xlsx":
        return _read_xlsx_header_fast(
            file_path
        )

    if extension == ".csv":
        # CSV headers can be obtained without reading the full file.
        header = pd.read_csv(
            file_path,
            nrows=0
        )

        return {
            "columns": list(header.columns),
            "worksheet": "CSV",
            "rows": None
        }

    if extension == ".xls":
        dataframe = pd.read_excel(
            file_path,
            nrows=0
        )

        return {
            "columns": list(dataframe.columns),
            "worksheet": "Excel",
            "rows": None
        }

    if extension == ".txt":
        try:
            header = pd.read_csv(
                file_path,
                sep=None,
                engine="python",
                nrows=0
            )
        except Exception:
            header = pd.read_csv(
                file_path,
                sep=r"\s+",
                engine="python",
                nrows=0
            )

        return {
            "columns": list(header.columns),
            "worksheet": "TXT",
            "rows": None
        }

    raise ValueError(
        f"Unsupported file type: {extension}"
    )


def _load_active_columns(
    file_path,
    file_signature,
    filename,
    selected_columns
):
    """
    Load only columns needed for the current graph.

    IMPORTANT:
    For .003 uploads, file_path is already the converted .xlsx path.
    Therefore the physical file extension must be used here, not the
    original uploaded filename. Otherwise the converted XLSX would be
    sent through the .003 converter a second time and binary XLSX bytes
    (starting with PK) would be treated as text.
    """
    extension = os.path.splitext(file_path)[1].lower()

    # Safety fallback: if a caller supplies the original .003 path,
    # convert it exactly once.
    if extension == ".003":
        file_path = _convert_003_to_xlsx(file_path)
        extension = ".xlsx"

    if extension == ".xlsx":
        dataframe, worksheet = _load_xlsx_selected_columns(
            file_path,
            file_signature,
            tuple(selected_columns)
        )
    elif extension == ".csv":
        dataframe = pd.read_csv(
            file_path,
            usecols=list(selected_columns),
            low_memory=False
        )
        worksheet = "CSV"
    elif extension == ".xls":
        dataframe = pd.read_excel(
            file_path,
            usecols=list(selected_columns)
        )
        worksheet = "Excel"
    elif extension == ".txt":
        try:
            dataframe = pd.read_csv(
                file_path,
                sep=None,
                engine="python",
                usecols=list(selected_columns)
            )
        except Exception:
            dataframe = pd.read_csv(
                file_path,
                sep=r"\s+",
                engine="python",
                usecols=list(selected_columns)
            )

        worksheet = "TXT"
    else:
        raise ValueError(
            f"Unsupported file type: {extension}"
        )

    return dataframe, worksheet



@st.cache_data(
    show_spinner=False,
    max_entries=5
)
def _load_dataset_cached(file_path, file_signature):
    """
    Cached loader keyed by the uploaded file signature.

    Returns:
        DataFrame, worksheet name
    """
    extension = os.path.splitext(file_path)[1].lower()

    if extension == ".003":
        file_path = _convert_003_to_xlsx(file_path)
        extension = ".xlsx"

    if extension == ".xlsx":
        return _read_xlsx_streaming(file_path)

    if extension == ".csv":
        dataframe = pd.read_csv(
            file_path,
            low_memory=False
        )
        return dataframe, "CSV"

    if extension == ".xls":
        dataframe = pd.read_excel(
            file_path
        )
        return dataframe, "Excel"

    if extension == ".txt":
        # Keep TXT support simple and compatible with the existing workflow.
        try:
            dataframe = pd.read_csv(
                file_path,
                sep=None,
                engine="python",
            )
        except Exception:
            dataframe = pd.read_csv(
                file_path,
                sep=r"\s+",
                engine="python",
            )

        return dataframe, "TXT"

    raise ValueError(
        f"Unsupported file type: {extension}"
    )


def _load_dataset_into_service(file_path, file_signature, filename):
    """
    Load the active file without depending on BackendService.load_file().
    The frontend already performs graph/analysis/anomaly work, so the
    service is used as a lightweight shared data container.
    """
    dataframe, worksheet_name = _load_dataset_cached(
        file_path,
        file_signature
    )

    if dataframe is None or dataframe.empty:
        raise ValueError("The selected dataset contains no data rows.")

    dataframe.columns = [
        str(column).strip()
        if str(column).strip()
        else f"Column_{index + 1}"
        for index, column in enumerate(dataframe.columns)
    ]

    # Keep the original row order. Add no physical serial-number column;
    # Row Serial Number remains a virtual graph-only option.
    service.data = dataframe
    service.attributes = list(dataframe.columns)
    service.file_type = os.path.splitext(filename)[1].upper().replace(".", "")
    service.file_info = {
        "rows": int(len(dataframe)),
        "columns": int(len(dataframe.columns)),
        "worksheet": worksheet_name,
        "file_name": filename,
        "file_size_mb": round(
            os.path.getsize(file_path) / (1024 * 1024),
            2
        )
    }

    return {
        "success": True,
        "message": (
            f"Loaded {len(dataframe):,} rows × "
            f"{len(dataframe.columns):,} columns "
            f"from {worksheet_name}."
        ),
        "converted_file": None
    }




# ============================================================
# RAW FIRST-SIX-ROW CLEANUP
# ============================================================

def _preferred_xlsx_sheet_name(xlsx_path):
    """Choose the same kind of telemetry worksheet used by the main loader."""
    from openpyxl import load_workbook
    wb = load_workbook(xlsx_path, read_only=True, data_only=True)
    names = list(wb.sheetnames)
    wb.close()
    for name in names:
        low = str(name).lower()
        if any(token in low for token in ("telemetry", "data", "flight", "test")):
            return name
    return names[0] if names else None


@st.cache_data(show_spinner=False, max_entries=20)
def _read_first_six_raw_rows(file_path, file_signature=None, file_size=None, file_mtime_ns=None):
    """Read only the first six physical rows with a lightweight XLSX parser.

    The preview is cached because Streamlit reruns on widget interaction.
    XLSX preview uses worksheet XML directly instead of openpyxl, so a large
    workbook does not need to be fully opened just to show six rows.
    """
    extension = os.path.splitext(file_path)[1].lower()

    if extension == ".xlsx":
        sheet_paths = _xlsx_sheet_paths(file_path)
        if not sheet_paths:
            raise ValueError("The Excel workbook contains no worksheets.")

        preferred = None
        for name, path in sheet_paths:
            if any(token in str(name).lower() for token in ("telemetry", "data", "flight", "test")):
                preferred = (name, path)
                break
        sheet_name, sheet_path = preferred or sheet_paths[0]

        rows = []
        cell_pattern = re.compile(rb'<c r="([A-Z]+)([0-9]+)"[^>]*>(.*?)</c>', flags=re.DOTALL)

        with zipfile.ZipFile(file_path, "r") as archive:
            with archive.open(sheet_path, "r") as stream:
                buffer = b""
                while len(rows) < 6:
                    chunk = stream.read(256 * 1024)
                    if not chunk:
                        break
                    buffer += chunk
                    parts = buffer.split(b"</row>")
                    buffer = parts.pop()
                    for row_part in parts:
                        if b"<row" not in row_part:
                            continue
                        values = {}
                        for match in cell_pattern.finditer(row_part):
                            col_letters = match.group(1).decode("ascii", errors="ignore")
                            col_index = _excel_col_index(col_letters + "1")
                            values[col_index] = _decode_xlsx_cell(match.group(3))
                        if values:
                            max_col = max(values) + 1
                            row = [values.get(i) for i in range(max_col)]
                            rows.append(row)
                        else:
                            rows.append([])
                        if len(rows) >= 6:
                            break

    elif extension in (".csv", ".txt"):
        import csv
        rows = []
        if extension == ".csv":
            with open(file_path, "r", encoding="utf-8-sig", errors="replace", newline="") as fh:
                sample = fh.read(8192)
                fh.seek(0)
                try:
                    dialect = csv.Sniffer().sniff(sample, delimiters=",;\t|")
                    reader = csv.reader(fh, dialect)
                except Exception:
                    reader = csv.reader(fh)
                for row in reader:
                    rows.append(row)
                    if len(rows) >= 6:
                        break
        else:
            with open(file_path, "r", encoding="utf-8-sig", errors="replace") as fh:
                for line in fh:
                    if line.strip():
                        rows.append(re.split(r"\s+", line.strip()))
                        if len(rows) >= 6:
                            break
    elif extension == ".xls":
        df = pd.read_excel(file_path, header=None, nrows=6)
        rows = df.where(pd.notna(df), "").values.tolist()
    else:
        raise ValueError(f"Unsupported file type: {extension}")

    if not rows:
        return pd.DataFrame()
    width = max(len(row) for row in rows)
    padded = [list(row) + [""] * (width - len(row)) for row in rows]
    columns = [f"Column {i + 1}" for i in range(width)]
    return pd.DataFrame(padded, columns=columns)


def _create_cleaned_xlsx(source_path, delete_rows_1based, output_path):
    """Create a cleaned XLSX using low-memory streaming XML editing."""
    extension = os.path.splitext(source_path)[1].lower()
    delete_set = {int(x) for x in delete_rows_1based if 1 <= int(x) <= 6}

    if extension == ".xlsx":
        from lxml import etree
        NS_MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
        NS_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
        NS_PKG_REL = "http://schemas.openxmlformats.org/package/2006/relationships"

        def shift_row_number(old_row):
            return old_row - sum(1 for n in delete_set if n < old_row)

        def shift_cell_ref(ref):
            if not ref:
                return ref
            m = re.match(r"^([A-Za-z]+)(\d+)(.*)$", str(ref))
            if not m:
                return ref
            return f"{m.group(1)}{shift_row_number(int(m.group(2)))}{m.group(3)}"

        def shift_range(value):
            if not value:
                return value
            return ":".join(shift_cell_ref(part) for part in str(value).split(":"))

        os.makedirs(os.path.dirname(output_path), exist_ok=True)
        temp_xml = output_path + ".sheet.tmp"

        with zipfile.ZipFile(source_path, "r") as zin:
            workbook_root = etree.fromstring(zin.read("xl/workbook.xml"))
            rels_root = etree.fromstring(zin.read("xl/_rels/workbook.xml.rels"))
            sheets = workbook_root.findall(f"{{{NS_MAIN}}}sheets/{{{NS_MAIN}}}sheet")
            target_sheet = next(
                (sh for sh in sheets if any(t in str(sh.get("name", "")).lower() for t in ("telemetry", "data", "flight", "test"))),
                sheets[0] if sheets else None
            )
            if target_sheet is None:
                raise ValueError("The Excel workbook contains no worksheets.")
            rid = target_sheet.get(f"{{{NS_REL}}}id")
            target_part = None
            for rel in rels_root.findall(f"{{{NS_PKG_REL}}}Relationship"):
                if rel.get("Id") == rid:
                    target_part = rel.get("Target", "").lstrip("/")
                    if not target_part.startswith("xl/"):
                        target_part = "xl/" + target_part
                    break
            if not target_part or target_part not in zin.namelist():
                raise ValueError("Could not locate the selected worksheet XML.")

            with zin.open(target_part, "r") as src, open(temp_xml, "wb") as dst:
                context = etree.iterparse(src, events=("start", "end"), huge_tree=True)
                writer = None
                root_ctx = None
                sheetdata_ctx = None
                root_seen = False

                for event, elem in context:
                    local = etree.QName(elem).localname
                    if event == "start":
                        if not root_seen:
                            root_seen = True
                            xf = etree.xmlfile(dst, encoding="UTF-8")
                            writer = xf.__enter__()
                            root_ctx = writer.element(elem.tag, dict(elem.attrib), nsmap=elem.nsmap)
                            root_ctx.__enter__()
                        elif local == "sheetData" and etree.QName(elem.getparent()).localname == "worksheet":
                            sheetdata_ctx = writer.element(elem.tag, dict(elem.attrib))
                            sheetdata_ctx.__enter__()
                        continue

                    # End events.
                    if local == "row":
                        old_r = int(elem.get("r", "0") or 0)
                        if old_r in delete_set:
                            elem.clear()
                        else:
                            elem.set("r", str(shift_row_number(old_r)))
                            for cell in elem:
                                if etree.QName(cell).localname == "c" and cell.get("r"):
                                    cell.set("r", shift_cell_ref(cell.get("r")))
                            writer.write(elem)
                            elem.clear()
                    elif local == "sheetData" and sheetdata_ctx is not None:
                        sheetdata_ctx.__exit__(None, None, None)
                        sheetdata_ctx = None
                        elem.clear()
                    elif local == "dimension" and elem.get("ref"):
                        elem.set("ref", shift_range(elem.get("ref")))
                        writer.write(elem)
                        elem.clear()
                    elif root_seen and local != "worksheet":
                        parent = elem.getparent()
                        if parent is not None and etree.QName(parent).localname == "worksheet":
                            writer.write(elem)
                            elem.clear()

                if root_ctx is not None:
                    root_ctx.__exit__(None, None, None)
                if writer is not None:
                    xf.__exit__(None, None, None)

        with zipfile.ZipFile(source_path, "r") as zin, zipfile.ZipFile(
            output_path, "w", compression=zipfile.ZIP_DEFLATED, compresslevel=1
        ) as zout:
            for info in zin.infolist():
                if info.filename == target_part:
                    with open(temp_xml, "rb") as fh:
                        zout.writestr(info, fh.read())
                else:
                    zout.writestr(info, zin.read(info.filename))
        try:
            os.remove(temp_xml)
        except OSError:
            pass
        return output_path

    # Non-XLSX inputs retain the existing conversion behavior.
    if extension == ".csv":
        import csv
        with open(source_path, "r", encoding="utf-8-sig", errors="replace", newline="") as fh:
            sample = fh.read(8192)
        try:
            dialect = csv.Sniffer().sniff(sample, delimiters=",;\t|")
            sep = dialect.delimiter
        except Exception:
            sep = ","
        df = pd.read_csv(source_path, header=None, sep=sep, engine="python")
    elif extension == ".xls":
        df = pd.read_excel(source_path, header=None)
    elif extension == ".txt":
        try:
            df = pd.read_csv(source_path, header=None, sep=None, engine="python")
        except Exception:
            df = pd.read_csv(source_path, header=None, sep=r"\s+", engine="python")
    else:
        raise ValueError(f"Unsupported file type: {extension}")
    drop_zero = [r - 1 for r in delete_set if 1 <= r <= len(df)]
    if drop_zero:
        df = df.drop(index=drop_zero).reset_index(drop=True)
    df.to_excel(output_path, index=False, header=False, engine="openpyxl")
    return output_path

def _safe_display_value(value):
    if value is None or (isinstance(value, float) and np.isnan(value)):
        return ""
    return str(value)


# ------------------------------------------------------------
# Build the current multi-file list.
# The uploader is only the input mechanism. Once a file is copied to
# input/, the local registry/session state becomes the source of truth.
# This prevents tab navigation from making the uploaded dataset vanish.
# ------------------------------------------------------------
input_directory = os.path.join(PROJECT_ROOT, "input")
os.makedirs(input_directory, exist_ok=True)

# Persistent uploader value. The local input/ registry below remains the source of truth during navigation.
_persistent_uploaded_files = st.session_state.get("persistent_uploaded_files", [])

# Read the uploader only when it is rendered in Data Input.
if st.session_state.active_tab == "Data Input":
    _uploaded_now = st.file_uploader(
        "Upload helicopter dataset",
        type=None,
        accept_multiple_files=True,
        key="helicopter_file_uploader"
    )
    if _uploaded_now:
        _persistent_uploaded_files = _uploaded_now
        st.session_state.persistent_uploaded_files = _uploaded_now

if _persistent_uploaded_files is None:
    _persistent_uploaded_files = []

# Streamlit uploader state. Keep this defined before the persistent-file registry.

_current_files = []

for _file in _persistent_uploaded_files:
    _file_id = getattr(_file, "file_id", None)
    _signature = (
        _file_id
        if _file_id is not None
        else (
            _file.name,
            int(getattr(_file, "size", 0))
        )
    )

    if _signature in st.session_state.removed_file_signatures:
        continue

    _raw_path = os.path.join(input_directory, _file.name)
    _original_extension = os.path.splitext(_file.name)[1].lower()
    _path = _raw_path
    _existing_registry = st.session_state.known_file_registry.get(_signature, {})

    # Write only a genuinely new upload.
    if (
        _signature not in st.session_state.known_file_registry
        or not os.path.exists(_raw_path)
    ):
        with open(_raw_path, "wb") as _out:
            _out.write(_file.getvalue())

    # Every non-XLSX upload is normalized to XLSX before cleanup/reading.
    # This includes .003, .005, .00.out, .out, .dat, .log, CSV, TXT, JSON, etc.
    # The converted workbook preserves the physical rows so the user can
    # delete unwanted leading rows before the normal Excel pipeline starts.
    try:
        _path, _was_converted = _normalize_upload_to_xlsx(
            _raw_path, input_directory, _signature
        )
    except Exception as exc:
        st.error(f"Unable to convert {_file.name} to Excel: {exc}")
        continue

    _size = os.path.getsize(_raw_path)

    if _existing_registry.get("cleaned") and os.path.exists(_existing_registry.get("path", "")):
        _path = _existing_registry["path"]

    _current_files.append({
        "file": _file,
        "signature": _signature,
        "name": _file.name,
        "size": _size,
        "path": _path,
        "source_path": _raw_path,
        "is_003": _original_extension == ".003",
        "cleaned": bool(_existing_registry.get("cleaned", False)),
    })

    st.session_state.known_file_registry[_signature] = {
        "signature": _signature,
        "name": _file.name,
        "size": _size,
        "path": _path if not _existing_registry.get("cleaned") else _existing_registry.get("path", _path),
        "source_path": _raw_path,
        "is_003": _original_extension == ".003",
        "cleaned": bool(_existing_registry.get("cleaned", False)),
    }
    if _existing_registry.get("cleaned"):
        st.session_state.row_cleanup_done_by_signature[_signature] = True

# Merge the local registry with the uploader value. This means that
# previously uploaded datasets remain selectable even when Streamlit
# temporarily reports only the newest uploader value after navigation.
_seen_signatures = {
    item["signature"] for item in _current_files
}

for _signature, _meta in st.session_state.known_file_registry.items():
    if _signature in _seen_signatures:
        continue

    _path = _meta.get("path")
    if _path and os.path.exists(_path):
        _current_files.append({
            "file": None,
            "signature": _signature,
            "name": _meta["name"],
            "size": int(_meta.get("size", os.path.getsize(_path))),
            "path": _path,
            "source_path": _meta.get("source_path", _path),
            "is_003": bool(_meta.get("is_003", False)),
        })

_current_signatures = [item["signature"] for item in _current_files]

_selected_signature = st.session_state.get("selected_file_signature")

if _current_signatures:
    if _selected_signature not in _current_signatures:
        _selected_signature = _current_signatures[0]
        st.session_state.selected_file_signature = _selected_signature

    # If the user changes the active dataset, preserve the complete state
    # of the previous dataset before restoring the new one.
    old_signature = st.session_state.get("uploaded_file_signature")
    if old_signature is not None and old_signature != _selected_signature:
        _save_active_workspace()
else:
    _selected_signature = None
# ------------------------------------------------------------
# File selection panel — shown on Data Input only.
# ------------------------------------------------------------
if st.session_state.active_tab == "Data Input" and _current_files:
    st.markdown(
        '<div class="section-label">UPLOADED DATASETS</div>',
        unsafe_allow_html=True
    )

    _labels = []
    _sig_to_label = {}
    for _i, _item in enumerate(_current_files, start=1):
        _size_mb = _item["size"] / (1024 * 1024)
        _label = f"{_i:02d}  {_item['name']}  •  {_size_mb:.2f} MB"
        _labels.append(_label)
        _sig_to_label[_item["signature"]] = _label

    _selected_label = st.radio(
        "Select dataset to analyze",
        _labels,
        index=max(0, _current_signatures.index(_selected_signature))
        if _selected_signature in _current_signatures else 0,
        key="active_dataset_selector",
        label_visibility="collapsed"
    )
    _label_to_sig = {v: k for k, v in _sig_to_label.items()}
    _new_selected_signature = _label_to_sig[_selected_label]

    if _new_selected_signature != st.session_state.get("selected_file_signature"):
        _save_active_workspace()
        st.session_state.selected_file_signature = _new_selected_signature
        _selected_signature = _new_selected_signature

    if st.button(
        "✕  Remove Selected Dataset",
        use_container_width=False,
        key="remove_selected_dataset"
    ):
        st.session_state.removed_file_signatures.add(
            _selected_signature
        )
        st.session_state.known_file_registry.pop(
            _selected_signature,
            None
        )
        st.session_state.workspace_by_signature.pop(
            _selected_signature,
            None
        )

        if st.session_state.get("uploaded_file_signature") == _selected_signature:
            _clear_active_workspace()

        st.session_state.selected_file_signature = None
        st.rerun()

# ------------------------------------------------------------
# Load the selected dataset into the existing BackendService.
# Header metadata is read once. Full data is loaded only when the user
# generates a graph, and selected columns are cached.
# ------------------------------------------------------------
if _selected_signature is not None:
    _selected_item = next(
        (
            item for item in _current_files
            if item["signature"] == _selected_signature
        ),
        None
    )

    if _selected_item is not None:
        _selected_path = _selected_item["path"]
        _selected_filename = _selected_item["name"]
        if _selected_item.get("cleaned"):
            st.session_state.row_cleanup_done_by_signature[_selected_signature] = True

        # --------------------------------------------------------
        # FIRST SIX RAW ROWS — USER CLEANUP
        # --------------------------------------------------------
        # Show this only until the user confirms which of the first six
        # physical rows should be removed. This happens before header
        # detection, so metadata/hash rows can be deleted safely.
        _cleanup_done = st.session_state.row_cleanup_done_by_signature.get(
            _selected_signature, False
        )

        if not _cleanup_done:
            try:
                _raw_preview = _read_first_six_raw_rows(
                    _selected_path,
                    _selected_signature,
                    _selected_item.get("size", 0),
                    os.path.getmtime(_selected_path) if os.path.exists(_selected_path) else 0
                )
            except Exception as _preview_exc:
                _raw_preview = pd.DataFrame()
                st.warning(f"Could not preview the first six rows: {_preview_exc}")

            if not _raw_preview.empty:
                st.divider()
                st.markdown(
                    '<div class="section-label">DATA CLEANUP</div>',
                    unsafe_allow_html=True
                )
                st.subheader("Review the first 6 rows")
                st.caption(
                    "Some flight-data files contain hash values, metadata, or unwanted rows before the real attribute/header row. "
                    "Select the rows you want to DELETE. Your original uploaded file is not modified."
                )

                # Keep all six selections inside a form. Checking a row no
                # longer reruns the entire app, which removes the white/frozen
                # pause on large workbooks. The expensive delete operation only
                # runs when the user presses the submit button.
                with st.form(key=f"cleanup_form_{_selected_signature}", clear_on_submit=False):
                    _delete_flags = []
                    for _row_number in range(1, 7):
                        if _row_number <= len(_raw_preview):
                            _row_values = [
                                _safe_display_value(v)
                                for v in _raw_preview.iloc[_row_number - 1].tolist()
                            ]
                            _preview_text = " | ".join(_row_values[:8])
                            if len(_preview_text) > 120:
                                _preview_text = _preview_text[:117] + "..."
                        else:
                            _preview_text = "<empty row>"

                        _delete_flags.append(
                            st.checkbox(
                                f"Delete Row {_row_number}: {_preview_text}",
                                value=False,
                                key=f"cleanup_delete_{_selected_signature}_{_row_number}"
                            )
                        )

                    _selected_delete_rows = [
                        i + 1 for i, flag in enumerate(_delete_flags) if flag
                    ]

                    _submit_cleanup = st.form_submit_button(
                        "🗑️ Delete Selected Rows & Generate Clean Excel",
                        type="primary"
                    )

                if _submit_cleanup:
                    if not _selected_delete_rows:
                        st.warning("Select at least one row to delete.")
                    elif len(_selected_delete_rows) >= 6 and len(_raw_preview) <= 6:
                        st.error("Please keep at least one of the first six rows so the file has a header/data starting point.")
                    else:
                        _clean_dir = os.path.join(PROJECT_ROOT, "input", "cleaned")
                        os.makedirs(_clean_dir, exist_ok=True)
                        _base_name = os.path.splitext(_selected_filename)[0]
                        _cleaned_path = os.path.join(
                            _clean_dir,
                            f"{_base_name}_cleaned.xlsx"
                        )
                        try:
                            with st.spinner("Removing selected rows and generating the cleaned Excel file..."):
                                _create_cleaned_xlsx(
                                    _selected_path,
                                    _selected_delete_rows,
                                    _cleaned_path
                                )

                            _selected_item["path"] = _cleaned_path
                            _selected_item["cleaned"] = True
                            st.session_state.cleaned_file_by_signature[_selected_signature] = _cleaned_path
                            st.session_state.cleaned_file_download_by_signature[_selected_signature] = _cleaned_path
                            st.session_state.known_file_registry[_selected_signature]["path"] = _cleaned_path
                            st.session_state.known_file_registry[_selected_signature]["cleaned"] = True

                            _clear_active_workspace()
                            st.session_state.selected_file_signature = _selected_signature
                            st.session_state.row_cleanup_done_by_signature[_selected_signature] = True
                            st.session_state.file_header_by_signature.pop(_selected_signature, None)
                            st.success("Selected rows deleted. Clean Excel generated and set as the active dataset.")
                            st.rerun()
                        except Exception as _clean_exc:
                            st.error(f"Could not generate the cleaned Excel file: {_clean_exc}")

        # The cleaned-Excel download belongs ONLY on the Data Input tab.
        # Do not carry this control into Visualization, Analysis, Anomaly
        # Detection, or Export.
        if st.session_state.active_tab == "Data Input":
            _existing_cleaned = st.session_state.cleaned_file_download_by_signature.get(_selected_signature)
            if _existing_cleaned and os.path.exists(_existing_cleaned):
                st.success("Clean Excel file is ready for download.")
                _download_size = os.path.getsize(_existing_cleaned)
                if _download_size <= 80 * 1024 * 1024:
                    @st.cache_data(show_spinner=False, max_entries=10)
                    def _cached_clean_file_bytes(path, mtime_ns):
                        with open(path, "rb") as fh:
                            return fh.read()
                    _clean_bytes = _cached_clean_file_bytes(
                        _existing_cleaned,
                        os.stat(_existing_cleaned).st_mtime_ns
                    )
                    st.download_button(
                        "⬇️ Download Cleaned Excel",
                        _clean_bytes,
                        file_name=os.path.basename(_existing_cleaned),
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        key=f"download_cleaned_excel_final_{_selected_signature}"
                    )
                else:
                    st.caption("The cleaned Excel is ready. Use the file from the cleaned input folder if the browser limits very large downloads.")

        # The rest of the original file-loading pipeline continues unchanged.
        if st.session_state.get("uploaded_file_signature") != _selected_signature:
            try:
                header_info = _read_file_header_cached(
                    _selected_path,
                    _selected_signature
                )
                st.session_state.file_header_by_signature[
                    _selected_signature
                ] = header_info
            except Exception as exc:
                st.error(f"Unable to read {_selected_filename}: {exc}")
                st.stop()

            _restore_workspace(
                _selected_signature,
                _selected_filename
            )

            service.attributes = list(header_info["columns"])
            service.file_type = os.path.splitext(
                _selected_filename
            )[1].upper().replace(".", "")

            if _selected_item.get("is_003"):
                st.session_state.converted_file = _selected_item["path"]

            service.file_info = {
                "rows": header_info.get("rows"),
                "columns": len(header_info["columns"]),
                "worksheet": header_info.get("worksheet"),
                "file_name": _selected_filename,
                "file_size_mb": round(
                    _selected_item["size"] / (1024 * 1024),
                    2
                )
            }

            # If this file already has a graph, restore it immediately.
            graph_df = st.session_state.get("graph_data")
            if graph_df is not None:
                real_cols = [
                    c for c in graph_df.columns
                    if c != "Row Serial Number"
                ]
                service.data = graph_df[real_cols]
                # Keep the full worksheet attribute list available for
                # future X/Y selections.
                service.attributes = list(header_info["columns"])
                service.file_info["rows"] = len(graph_df)

        else:
            # Navigation rerun: never touch the workbook again.
            graph_df = st.session_state.get("graph_data")
            header_info = st.session_state.file_header_by_signature.get(
                _selected_signature
            )

            if header_info is None:
                header_info = _read_file_header_cached(
                    _selected_path,
                    _selected_signature
                )
                st.session_state.file_header_by_signature[
                    _selected_signature
                ] = header_info

            # Navigation must never shrink the attribute list to the
            # columns used by the last graph.
            service.attributes = list(header_info["columns"])
            service.file_type = os.path.splitext(
                _selected_filename
            )[1].upper().replace(".", "")

            if graph_df is not None:
                real_cols = [
                    c for c in graph_df.columns
                    if c != "Row Serial Number"
                ]
                service.data = graph_df[real_cols]
                service.file_info = {
                    **getattr(service, "file_info", {}),
                    "rows": len(graph_df),
                    "columns": len(real_cols),
                    "file_name": _selected_filename,
                    "file_size_mb": round(
                        _selected_item["size"] / (1024 * 1024),
                        2
                    )
                }


# ============================================================
# LARGE DATASET VISUALIZATION HELPERS
# ============================================================

MAX_PLOT_POINTS = 12000


def _downsample_for_plot(dataframe, max_points=MAX_PLOT_POINTS):
    """Downsample only browser plotting data; keep full data for analysis."""
    if len(dataframe) <= max_points:
        return dataframe

    step = max(1, int(np.ceil(len(dataframe) / max_points)))
    sampled = dataframe.iloc[::step].copy()

    if len(sampled) and sampled.index[-1] != dataframe.index[-1]:
        sampled = pd.concat(
            [sampled, dataframe.iloc[[-1]]],
            ignore_index=False
        )

    return sampled


def _prepare_x_for_plot(series):
    """Keep ordinary X values unchanged; parse date-like strings when reliable."""
    if pd.api.types.is_numeric_dtype(series):
        return series

    parsed = pd.to_datetime(series, errors="coerce")
    non_null = series.notna().sum()

    if non_null and (parsed.notna().sum() / non_null) >= 0.80:
        return parsed

    return series


def _prepare_time_for_display(series):
    """Convert common CSV/Excel Time representations to HH:MM:SS display values.

    The source data is never changed; this is only the plotting representation.
    Supports Excel fractions such as 0.41669 as well as HH:MM:SS text.
    """
    s = series.copy()

    # Already datetime: keep it, but make sure Plotly formats it as a clock.
    if pd.api.types.is_datetime64_any_dtype(s):
        return s

    numeric = pd.to_numeric(s, errors="coerce")
    numeric_non_null = int(numeric.notna().sum())

    # Excel stores a time-of-day as fraction of one day.
    if numeric_non_null and numeric_non_null >= max(1, int(s.notna().sum() * 0.80)):
        vals = numeric.to_numpy(dtype="float64")
        valid = np.isfinite(vals)
        if valid.any() and np.nanmin(vals[valid]) >= 0 and np.nanmax(vals[valid]) < 1.0:
            return pd.Timestamp("1899-12-30") + pd.to_timedelta(numeric, unit="D")

    # Text time such as 10:00:02 or 10:00:02.500.
    text = s.astype("string")
    td = pd.to_timedelta(text, errors="coerce")
    non_null = int(s.notna().sum())
    if non_null and int(td.notna().sum()) >= max(1, int(non_null * 0.80)):
        return pd.Timestamp("1899-12-30") + td

    return s


def _prepare_plot_x(series, axis_name):
    """Prepare only the displayed X values; source dataframe remains untouched."""
    if str(axis_name).strip().lower() == "time":
        return _prepare_time_for_display(series)
    return _prepare_x_for_plot(series)



# ============================================================
# TAB 1 — DATA INPUT
# ============================================================

if st.session_state.active_tab == "Data Input":

    st.markdown(
        """<div class="section-label">01 — DATA INPUT</div>
        <div class="main-title">Flight Data Input</div>
        <div class="main-subtitle">Load helicopter flight-test data and prepare parameters for engineering analysis.</div>""",
        unsafe_allow_html=True
    )

    st.markdown(
        """<div class="info-strip"><b>DATA INTAKE</b>&nbsp;&nbsp; Maximum file size: 200 MB &nbsp;•&nbsp; Supported formats: XLSX, XLS, CSV, TXT, 003, 005, 00.OUT, OUT, DAT, LOG, JSON, JSONL, NDJSON, PARQUET &nbsp;•&nbsp; Other text/delimited telemetry files supported</div>""",
        unsafe_allow_html=True
    )


    # ========================================================
    # FILE INFORMATION
    # ========================================================

    if st.session_state.file_loaded:

        st.divider()

        st.markdown(
            '<div class="section-label">'
            'FILE INFORMATION'
            '</div>',
            unsafe_allow_html=True
        )


        info = service.file_info


        col1, col2, col3 = st.columns(3)


        with col1:

            st.metric(
                "File Type",
                service.file_type
            )


        with col2:

            st.metric(
                "Rows",
                (
                    f"{info['rows']:,}"
                    if info.get("rows") is not None
                    else "On demand"
                )
            )


        with col3:

            st.metric(
                "Columns",
                info["columns"]
            )


        st.write(
            "**Available Attributes**"
        )


        st.write(
            service.attributes
        )


        # ====================================================
        # TEXT/003 → EXCEL
        # ====================================================

        if st.session_state.converted_file:

            st.divider()

            st.subheader(
                "File → Excel Conversion"
            )


            st.success(
                "The uploaded .003 file has been converted to Excel."
            )


            converted_path = (
                st.session_state.converted_file
            )


            try:

                with open(
                    converted_path,
                    "rb"
                ) as file:

                    converted_data = file.read()


                st.download_button(

                    label=
                        "⬇️ Download Converted Excel",

                    data=converted_data,

                    file_name=
                        "converted_data.xlsx",

                    mime=(
                        "application/vnd.openxmlformats-"
                        "officedocument.spreadsheetml.sheet"
                    ),

                    use_container_width=True,

                    key=
                        "download_converted_excel"

                )


            except FileNotFoundError:

                st.error(
                    "Converted Excel file could not be found."
                )


        # ====================================================
        # AXIS SELECTION
        # ====================================================

        st.divider()

        st.markdown(
            '<div class="section-label">'
            '02 — AXIS SELECTION'
            '</div>',
            unsafe_allow_html=True
        )

        st.header("Select X-axis and Y-axis")

        attributes = list(service.attributes)

        # Row Serial Number is a virtual column created by the app.
        axis_options = ["Row Serial Number"] + attributes

        st.caption(
            "Select any available attribute for the X-axis and one or more "
            "attributes for the Y-axis."
        )

        axis_col1, axis_col2 = st.columns(2)

        with axis_col1:
            x_axis_selection = st.selectbox(
                "X-axis",
                options=["-- Select X-axis --"] + axis_options,
                index=0,
                key=f"x_axis_selector_{_selected_signature}"
            )

        with axis_col2:
            y_axis_selection = st.multiselect(
                "Y-axis",
                options=axis_options,
                default=[],
                key=f"y_axis_selector_{_selected_signature}"
            )

        st.markdown(
            '<div class="axis-help">'
            'X-axis: choose the variable to plot horizontally. '
            'Y-axis: choose one or more variables to plot vertically.'
            '</div>',
            unsafe_allow_html=True
        )

        # Keep the old session-state name for compatibility with other modules.
        st.session_state.x_axis = None

        # ====================================================
        # GRAPH TYPE
        # ====================================================

        st.divider()

        st.markdown(
            '<div class="section-label">'
            '03 — GRAPH CONFIGURATION'
            '</div>',
            unsafe_allow_html=True
        )


        st.header(
            "Choose Graph Type"
        )


        graph_type = st.selectbox(
            "Graph type",
            [
                "Line",
                "Bar",
                "Scatter"
            ],
            key=f"selected_graph_type_{_selected_signature}"
        )


        st.session_state.graph_type = (
            graph_type
        )


        # ====================================================
        # GENERATE GRAPH
        # ====================================================

        st.write("")

        if st.button(
            "📈 Generate Graph",
            type="primary",
            use_container_width=True,
            key="generate_graph_button"
        ):

            # --------------------------------------------
            # DEFAULT AXES
            # --------------------------------------------

            x_missing = (
                x_axis_selection == "-- Select X-axis --"
            )

            y_missing = (
                len(y_axis_selection) == 0
            )

            # If an axis is not selected, use Row Serial Number.
            final_x_axis = (
                "Row Serial Number"
                if x_missing
                else x_axis_selection
            )

            final_y_axes = (
                ["Row Serial Number"]
                if y_missing
                else y_axis_selection.copy()
            )

            # --------------------------------------------
            # "Generate All Graphs" means every available Y-axis parameter
            # was selected together. In that mode the exported report is
            # intentionally graph-only.
            st.session_state.graph_generation_mode = (
                "all"
                if (
                    not y_missing
                    and len(axis_options) > 0
                    and set(final_y_axes) == set(axis_options)
                )
                else "selected"
            )

            # USER-FACING DEFAULT NOTIFICATION
            # --------------------------------------------

            if x_missing:
                message = (
                    "You did not select an X-axis, so "
                    "Row Serial Number is being used as the default X-axis."
                )
                st.warning("⚠️ " + message)
                st.toast("⚠️ " + message)

            if y_missing:
                message = (
                    "You did not select a Y-axis, so "
                    "Row Serial Number is being used as the default Y-axis."
                )
                st.warning("⚠️ " + message)
                st.toast("⚠️ " + message)

            # --------------------------------------------
            # BUILD GRAPH DATA — LAZY COLUMN LOAD
            # --------------------------------------------
            # Do NOT read all 53 columns just because the file has
            # 53 columns. Load only what the selected graph needs.

            real_columns_to_load = [
                column
                for column in (
                    [final_x_axis] + final_y_axes
                )
                if column != "Row Serial Number"
            ]

            try:
                with st.spinner(
                    "Loading selected graph data..."
                ):
                    selected_df, worksheet_name = _load_active_columns(
                        _selected_path,
                        _selected_signature,
                        _selected_filename,
                        list(dict.fromkeys(
                            real_columns_to_load
                        ))
                    )
            except Exception as exc:
                st.error(
                    f"Could not load the selected graph data: {exc}"
                )
                st.stop()

            # Keep a lightweight copy in the service for downstream
            # analysis/anomaly/export modules.
            service.data = selected_df

            # IMPORTANT: do not replace the complete worksheet attributes
            # with only the selected graph columns.
            graph_data = selected_df.copy()

            # Virtual graph-only serial number.
            graph_data.insert(
                0,
                "Row Serial Number",
                range(
                    1,
                    len(graph_data) + 1
                )
            )

            graph_columns = list(
                dict.fromkeys(
                    [final_x_axis] + final_y_axes
                )
            )

            graph_data = graph_data[
                graph_columns
            ].copy()

            service.file_info["rows"] = int(
                len(graph_data)
            )

            # --------------------------------------------
            # SAVE USER SELECTION
            # --------------------------------------------

            st.session_state.x_axis = final_x_axis
            st.session_state.y_axis = final_y_axes
            st.session_state.selected_attributes = final_y_axes
            st.session_state.graph_type = graph_type
            st.session_state.graph_data = graph_data

            # Keep the final graph X-axis exactly as selected/defaulted.
            # Row Serial Number is a valid virtual graph axis.
            st.session_state.x_axis = final_x_axis

            # --------------------------------------------
            # PRE-COMPUTE GRAPH ANALYSIS ONCE
            # --------------------------------------------
            # Analysis is calculated immediately after graph generation and
            # stored in session state. Visiting the Analysis tab later does
            # not recalculate the 180k-row dataset.
            precomputed_graphs = []

            for graph_number, y_axis in enumerate(
                final_y_axes,
                start=1
            ):
                analysis_result = build_graph_analysis(
                    graph_data,
                    final_x_axis,
                    y_axis
                )

                precomputed_graphs.append({
                    "graph_number": graph_number,
                    "x_axis": final_x_axis,
                    "y_axis": y_axis,
                    "mean": analysis_result["mean"],
                    "median": analysis_result["median"],
                    "minimum": analysis_result["minimum"],
                    "maximum": analysis_result["maximum"],
                    "analysis_points": analysis_result["analysis_points"],
                    "missing_x": analysis_result["missing_x"],
                    "missing_y": analysis_result["missing_y"],
                    "min_x": analysis_result["min_x"],
                    "max_x": analysis_result["max_x"]
                })

            st.session_state.report_data["graphs"] = precomputed_graphs
            st.session_state.report_data["file_information"] = {
                "file_name": st.session_state.get(
                    "uploaded_filename",
                    _selected_filename
                ),
                "file_type": service.file_type,
                "rows": len(graph_data),
                "x_axis": final_x_axis,
                "y_axes": final_y_axes
            }

            # --------------------------------------------
            # BACKEND ANALYSIS
            # --------------------------------------------

            # The backend works with real uploaded columns.
            # The virtual Row Serial Number is used only for plotting.
            backend_columns = [
                column
                for column in graph_columns
                if column != "Row Serial Number"
                and column in service.data.columns
            ]

            output_directory = os.path.join(
                PROJECT_ROOT,
                "output"
            )

            os.makedirs(
                output_directory,
                exist_ok=True
            )

            export_path = os.path.join(
                output_directory,
                "final_analysis.xlsx"
            )

            with st.spinner(
                "Generating helicopter analysis..."
            ):

                if backend_columns:
                    # The frontend analysis/visualization pipeline uses
                    # service.data directly. Avoid re-reading a very large
                    # workbook through the backend.
                    result = {
                        "success": True,
                        "message": (
                            "Dataset loaded and ready for visualization."
                        ),
                        "exported_file": None
                    }
                else:
                    result = {
                        "success": True,
                        "message": (
                            "Graph generated using Row Serial Number."
                        ),
                        "exported_file": None
                    }

            if result["success"]:

                st.session_state.analysis_result = result
                st.session_state.graph_ready = True

                # ----------------------------------------
                # AUTOMATIC TAB SWITCH
                # ----------------------------------------

                st.session_state.active_tab = (
                    "Visualization"
                )

                st.rerun()

            else:

                st.error(
                    result.get(
                        "message",
                        "Graph generation failed."
                    )
                )


# ============================================================
# TAB 2 — VISUALIZATION
# ============================================================

elif st.session_state.active_tab == "Visualization":

    st.markdown("""<div class="page-kicker">02 — VISUALIZATION</div><div class="main-title">Select Attributes</div><div class="main-subtitle">Choose the parameters for visualization. Select one attribute for X-axis and one or more for Y-axis.</div><div class="title-rule"></div>""",unsafe_allow_html=True)

    # ========================================================
    # CHECK GRAPH DATA
    # ========================================================

    if not st.session_state.graph_ready:

        st.info(
            "No graph has been generated yet. "
            "Go to Data Input and generate a graph."
        )


    else:

        graph_df = (
            st.session_state.graph_data.copy()
        )

        # Ensure the graph always has the row serial number available.
        if "Row Serial Number" not in graph_df.columns:
            graph_df["Row Serial Number"] = range(
                1,
                len(graph_df) + 1
            )



        selected_attributes = (
            st.session_state.selected_attributes
        )


        graph_type = (
            st.session_state.graph_type
        )


        x_axis = st.session_state.get(
            "x_axis",
            graph_df.columns[0]
        )

        y_axes = st.session_state.get(
            "y_axis",
            selected_attributes
        )

        # Safety fallback for older session state.
        if not x_axis or x_axis not in graph_df.columns:
            x_axis = "Row Serial Number"

        if x_axis not in graph_df.columns:
            graph_df["Row Serial Number"] = range(
                1,
                len(graph_df) + 1
            )

        selected_attributes = [
            col for col in y_axes
            if col in graph_df.columns
        ]

        if not selected_attributes:
            selected_attributes = ["Row Serial Number"]


        st.success(
            f"{len(selected_attributes)} "
            f"parameter graph(s) generated successfully."
        )


        st.caption(
            f"X-axis: {x_axis} • "
            f"Graph type: {graph_type}"
        )


        st.info(
            "💡 Use the mouse wheel to zoom, "
            "drag to pan, and hover over points "
            "to see exact values."
        )


        st.divider()


        # ====================================================
        # PREPARE X-AXIS COLUMN
        # ====================================================

        graph_df[x_axis] = _prepare_plot_x(
            graph_df[x_axis],
            x_axis
        )

        # ====================================================
        # CREATE INDIVIDUAL GRAPH
        # FOR EVERY SELECTED ATTRIBUTE
        # ====================================================

        for attribute in selected_attributes:

            if attribute not in graph_df.columns:

                continue


            st.markdown(
                f'<div class="graph-heading">'
                f'{attribute} vs {x_axis}'
                f'</div>',
                unsafe_allow_html=True
            )


            # ------------------------------------------------
            # GRAPH DATA
            # ------------------------------------------------

            current_graph = graph_df[
                [x_axis, attribute]
            ].copy()


            current_graph[attribute] = pd.to_numeric(
                current_graph[attribute],
                errors="coerce"
            )

            # Never show negative Y values in the visualization. The actual
            # selected data is kept unchanged for analysis/reporting; only
            # the visible Plotly axis is constrained to start at zero.
            current_graph = current_graph.dropna(
                subset=[attribute]
            )

            # Full current_graph is retained for analysis/reporting.
            # Only the browser-facing Plotly data is downsampled.
            display_graph = _downsample_for_plot(
                current_graph
            )

            if current_graph.empty:

                st.warning(
                    f"No numeric data available "
                    f"for {attribute}."
                )

                continue


            # =================================================
            # CREATE FIGURE
            # =================================================

            fig = go.Figure()

            graph_colors = [
                "#1565C0",
                "#00897B",
                "#5E35B1",
                "#EF6C00",
                "#2E7D32",
                "#C62828"
            ]

            color_index = (
                selected_attributes.index(attribute)
                % len(graph_colors)
            )

            current_color = graph_colors[color_index]

            # =================================================
            # ADD GRAPH TRACE
            # =================================================

            if graph_type == "Line":

                fig.add_trace(
                    go.Scatter(
                        x=display_graph[x_axis],
                        y=display_graph[attribute],
                        mode=("lines+markers" if len(display_graph) <= 5000 else "lines"),
                        name=attribute,

                        line=dict(
                            color=current_color,
                            width=2.5
                        ),

                        marker=dict(
                            color=current_color,
                            size=(5 if len(display_graph) <= 5000 else 0)
                        ),

                        hovertemplate=
                            "<b>%{fullData.name}</b>"
                            "<br>X: %{x}"
                            "<br>Value: %{y}"
                            "<extra></extra>"
                    )
                )

            elif graph_type == "Bar":

                fig.add_trace(
                    go.Bar(
                        x=display_graph[x_axis],
                        y=display_graph[attribute],
                        name=attribute,

                        marker=dict(
                            color=current_color
                        ),

                        hovertemplate=
                            "<b>%{fullData.name}</b>"
                            "<br>X: %{x}"
                            "<br>Value: %{y}"
                            "<extra></extra>"
                    )
                )

            elif graph_type == "Scatter":

                fig.add_trace(
                    go.Scatter(
                        x=display_graph[x_axis],
                        y=display_graph[attribute],
                        mode="markers",
                        name=attribute,

                        marker=dict(
                            color=current_color,
                            size=7
                        ),

                        hovertemplate=
                            "<b>%{fullData.name}</b>"
                            "<br>X: %{x}"
                            "<br>Value: %{y}"
                            "<extra></extra>"
                    )
                )

            # =================================================
            # GRAPH LAYOUT
            # =================================================

            fig.update_layout(

                title={
                    "text": f"{attribute} vs {x_axis}",
                    "x": 0.02,
                    "font": {
                        "size": 18,
                        "color": "#17212B"
                    }
                },

                xaxis_title=x_axis,

                yaxis_title=attribute,

                height=430,

                hovermode="x unified",

                margin=dict(
                    l=55,
                    r=18,
                    t=60,
                    b=60
                ),

                plot_bgcolor="#FFFFFF",

                paper_bgcolor="#FFFFFF",

                font=dict(
                    color="#1F2937"
                ),

                xaxis=dict(
                    showgrid=True,
                    tickfont=dict(color="#1F2937", size=10),
                    gridcolor="#D9E2EC",

                    rangeslider=dict(
                        visible=True,
                        bgcolor="#EAF2F8"
                    ),

                    # Keep labels compact. Time uses 2-minute ticks; numeric
                    # X-axes use 2-unit ticks such as 0, 2, 4, 6, 8.
                    **(
                        {
                            "tickformat": "%H:%M:%S",
                            "dtick": 120000
                        }
                        if str(x_axis).strip().lower() == "time"
                        else (
                            {"dtick": 2, "tick0": 0}
                            if pd.api.types.is_numeric_dtype(display_graph[x_axis])
                            else {}
                        )
                    ),

                    fixedrange=False
                ),

                yaxis=dict(
                    showgrid=True,
                    tickfont=dict(color="#1F2937", size=10),
                    gridcolor="#D9E2EC",
                    # Force the visible Y-axis to begin at zero.
                    range=[
                        0,
                        max(1.0, float(pd.to_numeric(display_graph[attribute], errors="coerce").max()) * 1.05)
                    ],
                    rangemode="tozero",
                    fixedrange=False
                )
            )

            # =================================================
            # INTERACTIVE GRAPH CONTROLS
            # =================================================

            config = {
                "displayModeBar": True,
                "displaylogo": False,
                "scrollZoom": True,
                "responsive": True,

                "modeBarButtonsToAdd": [
                    "zoom2d",
                    "pan2d",
                    "select2d",
                    "lasso2d",
                    "autoScale2d",
                    "resetScale2d"
                ]
            }

            # =================================================
            # DISPLAY GRAPH
            # =================================================

            st.plotly_chart(
                fig,
                use_container_width=True,
                config=config
            )

            st.divider()


# ============================================================
# TAB 3 — ANALYSIS
# ============================================================

elif st.session_state.active_tab == "Analysis":

    st.markdown(
        '<div class="section-label">'
        '03 — ANALYSIS'
        '</div>',
        unsafe_allow_html=True
    )

    st.header("Flight Data Analysis")

    if st.session_state.graph_data is None:

        st.info(
            "Generate a graph first to see the analysis."
        )

    else:

        analysis_df = st.session_state.graph_data.copy()

        # Graph-only serial number is always available.
        if "Row Serial Number" not in analysis_df.columns:
            analysis_df["Row Serial Number"] = range(
                1,
                len(analysis_df) + 1
            )

        x_axis = st.session_state.get(
            "x_axis",
            "Row Serial Number"
        )

        if x_axis not in analysis_df.columns:
            x_axis = "Row Serial Number"

        y_axes = st.session_state.get(
            "y_axis",
            st.session_state.get(
                "selected_attributes",
                []
            )
        )

        y_axes = [
            column
            for column in y_axes
            if column in analysis_df.columns
        ]

        if not y_axes:
            y_axes = ["Row Serial Number"]

        st.info(
            f"X-axis: {x_axis}  |  "
            f"Y-axis: {', '.join(y_axes)}"
        )

        st.markdown(
            "The following observations are written from the "
            "actual values in each selected graph."
        )

        st.divider()

        cached_graphs = st.session_state.report_data.get(
            "graphs",
            []
        )

        # If graph settings have not changed, use the analysis captured
        # during Generate Graph. This makes tab navigation instant.
        cached_by_y = {
            item.get("y_axis"): item
            for item in cached_graphs
            if item.get("x_axis") == x_axis
        }

        report_graphs = []

        for graph_number, y_axis in enumerate(
            y_axes,
            start=1
        ):

            st.subheader(
                f"Graph {graph_number} — {y_axis}"
            )

            result = cached_by_y.get(y_axis)

            if result is None:
                result = build_graph_analysis(
                    analysis_df,
                    x_axis,
                    y_axis
                )

            if result["mean"] is not None:

                c1, c2, c3, c4 = st.columns(4)

                with c1:
                    st.metric(
                        "Mean",
                        result["mean"]
                    )

                with c2:
                    st.metric(
                        "Median",
                        result["median"]
                    )

                with c3:
                    st.metric(
                        "Min",
                        result["minimum"]
                    )

                with c4:
                    st.metric(
                        "Max",
                        result["maximum"]
                    )

            st.markdown("#### Graph Analysis")

            for point in result["analysis_points"]:
                st.markdown(
                    f"• {point}"
                )

            report_graphs.append({
                "graph_number": graph_number,
                "x_axis": x_axis,
                "y_axis": y_axis,
                "mean": result["mean"],
                "median": result["median"],
                "minimum": result["minimum"],
                "maximum": result["maximum"],
                "analysis_points": result["analysis_points"],
                "missing_x": result["missing_x"],
                "missing_y": result["missing_y"],
                "min_x": result["min_x"],
                "max_x": result["max_x"]
            })

            st.divider()

        st.session_state.report_data["graphs"] = report_graphs
        st.session_state.report_data["file_information"] = {
            "file_name": st.session_state.get(
                "uploaded_filename",
                "Uploaded flight data"
            ),
            "file_type": getattr(
                service,
                "file_type",
                "Unknown"
            ),
            "rows": len(analysis_df),
            "x_axis": x_axis,
            "y_axes": y_axes
        }

# ============================================================
# TAB 4 — ANOMALY DETECTION
# ============================================================

elif st.session_state.active_tab == "Anomaly Detection":

    st.markdown(
        '<div class="section-label">04 — ANOMALY DETECTION</div>',
        unsafe_allow_html=True
    )
    st.header("Anomaly Detection")

    if st.session_state.graph_data is None:
        st.info("Generate a graph first to run anomaly detection.")
    else:
        # IMPORTANT FOR LARGE DATASETS:
        # Do not copy the complete dataframe on every Streamlit rerun.
        # Threshold changes only need the selected Y column and the X values.
        anomaly_df = st.session_state.graph_data

        x_axis = st.session_state.get("x_axis") or "Row Serial Number"
        if x_axis not in anomaly_df.columns:
            x_axis = "Row Serial Number"

        y_axes = st.session_state.get(
            "y_axis", st.session_state.get("selected_attributes", [])
        )
        y_axes = [column for column in y_axes if column in anomaly_df.columns]
        if not y_axes:
            y_axes = ["Row Serial Number"]

        graph_type = st.session_state.get("graph_type", "Line")

        st.info(
            "Select a value condition for each parameter. The graph focuses on the "
            "matching period while remaining fully interactive for zooming and panning."
        )

        def _first_col(frame, name):
            obj = frame.loc[:, name]
            return obj.iloc[:, 0] if isinstance(obj, pd.DataFrame) else obj

        # Cache the numeric Y conversion in session state. This avoids converting
        # 180,000+ values again when only the threshold/operator changes.
        numeric_cache = st.session_state.setdefault("anomaly_numeric_cache", {})
        x_cache = st.session_state.setdefault("anomaly_x_cache", {})
        result_cache = st.session_state.setdefault("anomaly_result_cache", {})

        def _get_numeric_y(y_col):
            key = (id(anomaly_df), y_col, len(anomaly_df))
            cached = numeric_cache.get(key)
            if cached is None:
                cached = pd.to_numeric(_first_col(anomaly_df, y_col), errors="coerce").to_numpy(dtype="float64")
                numeric_cache[key] = cached
            return cached

        def _get_plot_x(x_col):
            key = (id(anomaly_df), x_col, len(anomaly_df))
            cached = x_cache.get(key)
            if cached is None:
                cached = _prepare_plot_x(_first_col(anomaly_df, x_col), x_col)
                x_cache[key] = cached
            return cached

        def _condition_details_fast(values, x_values, operator, limit, max_details=40):
            """Vectorized threshold detection; never builds a list for every match."""
            valid = np.isfinite(values)
            if operator == "Greater than (>)":
                mask = valid & (values > float(limit))
            elif operator == "Less than (<)":
                mask = valid & (values < float(limit))
            else:
                tolerance = max(1e-9, abs(float(limit)) * 1e-9)
                mask = valid & (np.abs(values - float(limit)) <= tolerance)

            selected_positions = np.flatnonzero(mask)
            matching_count = int(selected_positions.size)
            if matching_count == 0:
                return [], 0, 0, 0, None, None

            # Direction is calculated vectorially against the immediately
            # previous valid Y reading. No Python loop over the full dataset.
            previous = np.empty_like(values)
            previous[:] = np.nan
            if len(values) > 1:
                previous[1:] = values[:-1]
            delta = values - previous

            selected_delta = delta[selected_positions]
            increase_count = int(np.count_nonzero(selected_delta > 0))
            decrease_count = int(np.count_nonzero(selected_delta < 0))
            equal_count = int(np.count_nonzero(selected_delta == 0))

            # Only create annotation records for the first 40 matches.
            # The count still represents ALL matching rows.
            display_positions = selected_positions[:max_details]
            details = []
            for pos in display_positions:
                prev = previous[pos]
                value = float(values[pos])
                if not np.isfinite(prev):
                    direction = "NO PREVIOUS"
                elif value > prev:
                    direction = "UP"
                elif value < prev:
                    direction = "DOWN"
                else:
                    direction = "EQUAL"

                x_value = x_values.iloc[pos] if hasattr(x_values, "iloc") else x_values[pos]
                details.append({
                    "row": int(pos + 1),
                    "x": x_value,
                    "value": value,
                    "direction": direction
                })

            # Return selected positions too, so the caller can focus the graph
            # without constructing details for every matching row.
            return details, increase_count, decrease_count, equal_count, selected_positions, mask

        for graph_number, y_axis in enumerate(y_axes, start=1):
            st.markdown(
                f'<div class="graph-heading">Graph {graph_number} — {y_axis} vs {x_axis}</div>',
                unsafe_allow_html=True
            )

            c1, c2 = st.columns([1, 2])
            with c1:
                operator = st.selectbox(
                    "Condition",
                    ["Greater than (>)", "Less than (<)", "Equal to (=)"],
                    key=f"anomaly_operator_{graph_number}_{y_axis}"
                )
            with c2:
                threshold_key = f"anomaly_threshold_{graph_number}_{y_axis}"
                if threshold_key not in st.session_state:
                    numeric_series = _get_numeric_y(y_axis)
                    finite_values = numeric_series[np.isfinite(numeric_series)]
                    default_value = float(np.median(finite_values)) if finite_values.size else 0.0
                    st.session_state[threshold_key] = default_value

                limit = st.number_input(
                    "Threshold value",
                    step=0.1,
                    format="%.6f",
                    key=threshold_key
                )

            # Cache threshold result so changing unrelated widgets does not
            # repeat the calculation.
            result_key = (id(anomaly_df), y_axis, operator, float(limit), len(anomaly_df))
            cached_result = result_cache.get(result_key)
            if cached_result is None:
                values = _get_numeric_y(y_axis)
                x_values = _get_plot_x(x_axis)
                cached_result = _condition_details_fast(
                    values, x_values, operator, limit
                )
                result_cache[result_key] = cached_result
            else:
                values = _get_numeric_y(y_axis)
                x_values = _get_plot_x(x_axis)

            details, increase_count, decrease_count, equal_count, selected_positions, match_mask = cached_result

            m1, m2, m3 = st.columns(3)
            m1.metric("Matching readings", int(selected_positions.size) if selected_positions is not None else 0)
            m2.metric("Went UP", increase_count)
            m3.metric("Went DOWN", decrease_count)
            if equal_count:
                st.caption(f"Matching readings that were equal to the previous valid reading: {equal_count}")

            # Build only the browser-sized plotting frame. Full telemetry remains
            # in graph_data; Plotly receives at most MAX_PLOT_POINTS rows.
            plot_df = pd.DataFrame({"_x": x_values, "_y": values})
            plot_df = plot_df.loc[np.isfinite(values)]
            if str(x_axis).strip().lower() != "time":
                plot_df = plot_df.loc[pd.notna(plot_df["_x"])]
            display_plot_df = _downsample_for_plot(plot_df)

            fig = go.Figure()
            if not display_plot_df.empty:
                color = "#1565C0"
                if graph_type == "Bar":
                    fig.add_trace(go.Bar(
                        x=display_plot_df["_x"], y=display_plot_df["_y"],
                        name=y_axis, marker=dict(color=color),
                        hovertemplate="<b>%{fullData.name}</b><br>X: %{x}<br>Value: %{y}<extra></extra>"
                    ))
                elif graph_type == "Scatter":
                    fig.add_trace(go.Scatter(
                        x=display_plot_df["_x"], y=display_plot_df["_y"],
                        mode="markers", name=y_axis,
                        marker=dict(color=color, size=7),
                        hovertemplate="<b>%{fullData.name}</b><br>X: %{x}<br>Value: %{y}<extra></extra>"
                    ))
                else:
                    fig.add_trace(go.Scatter(
                        x=display_plot_df["_x"], y=display_plot_df["_y"],
                        mode=("lines+markers" if len(display_plot_df) <= 5000 else "lines"),
                        name=y_axis, line=dict(color=color, width=2.5),
                        marker=dict(color=color, size=(5 if len(display_plot_df) <= 5000 else 0)),
                        hovertemplate="<b>%{fullData.name}</b><br>X: %{x}<br>Value: %{y}<extra></extra>"
                    ))

            # Keep the anomaly graph as the normal full graph. The selected
            # threshold still controls the matching counts above, but it must
            # NOT change the graph zoom or filter the plotted data.
            # Do not place HIGHEST/LOWEST callouts or arrows on the graph.
            # The two extreme readings are described compactly below instead.
            finite_positions = np.flatnonzero(np.isfinite(values))
            full_finite_values = values[finite_positions] if finite_positions.size else np.array([], dtype=float)
            max_y = float(np.max(full_finite_values)) if full_finite_values.size else 1.0
            highest_value = float(np.max(full_finite_values)) if full_finite_values.size else None
            lowest_value = float(np.min(full_finite_values)) if full_finite_values.size else None

            x_layout = dict(
                title={"text": x_axis, "font": {"size": 16, "color": "#111827"}},
                tickfont={"size": 10, "color": "#1F2937"},
                showgrid=True, gridcolor="#D9E2EC",
                rangeslider={"visible": True, "bgcolor": "#EAF2F8"},
                fixedrange=False
            )
            if str(x_axis).strip().lower() == "time":
                # Keep the visible Time labels at exactly 2-minute spacing.
                x_layout.update({"tickformat": "%H:%M:%S", "dtick": 120000})
            elif pd.api.types.is_numeric_dtype(pd.Series(x_values)):
                # Numeric X-axis: compact 2-unit ticks such as 0, 2, 4, 6, 8.
                x_layout.update({"dtick": 2, "tick0": 0})

            # IMPORTANT: do not set x_layout["range"] here. The anomaly graph
            # must stay at the same normal full-data view as the Visualization
            # graph, regardless of the selected threshold. Users can still use
            # Plotly zoom, pan and the range slider themselves.

            fig.update_layout(
                title={"text": f"Anomaly Detection — {y_axis} vs {x_axis}", "x": 0.02,
                        "font": {"size": 18, "color": "#17212B"}},
                xaxis=x_layout,
                yaxis={
                    "title": {"text": y_axis, "font": {"size": 16, "color": "#111827"}},
                    "tickfont": {"size": 10, "color": "#1F2937"},
                    "showgrid": True, "gridcolor": "#D9E2EC",
                    "range": [0, max(1.0, max_y * 1.05)], "rangemode": "tozero",
                    "fixedrange": False
                },
                height=430, hovermode="x unified",
                plot_bgcolor="#FFFFFF", paper_bgcolor="#FFFFFF",
                font={"color": "#1F2937"}, margin=dict(l=55, r=18, t=60, b=60)
            )

            st.plotly_chart(
                fig, use_container_width=True,
                config={
                    "displayModeBar": True, "displaylogo": False,
                    "scrollZoom": True, "responsive": True,
                    "modeBarButtonsToAdd": ["zoom2d", "pan2d", "autoScale2d", "resetScale2d"]
                },
                key=f"anomaly_graph_{graph_number}_{y_axis}"
            )

            matching_count = int(selected_positions.size) if selected_positions is not None else 0
            if highest_value is not None and lowest_value is not None:
                st.caption(
                    f"Highest reading: {highest_value:g} — this is the highest reading among all readings. "
                    f"Lowest reading: {lowest_value:g} — this is the lowest reading among all readings."
                )

            st.session_state.setdefault("anomaly_condition_results", {})[f"{graph_number}_{y_axis}"] = {
                "operator": operator,
                "threshold": float(limit),
                "matching_count": matching_count,
                "up_count": increase_count,
                "down_count": decrease_count,
                "equal_count": equal_count,
                "details": details
            }

        # Preserve the existing compact report payload.
        st.session_state.anomaly_report = []
        for key, result in st.session_state.get("anomaly_condition_results", {}).items():
            st.session_state.anomaly_report.append({
                "graph_number": key.split("_", 1)[0],
                "condition": f'{result["operator"]} {result["threshold"]}',
                "matching_count": result["matching_count"],
                "up_count": result["up_count"],
                "down_count": result["down_count"],
                "equal_count": result["equal_count"],
                "details": result["details"]
            })

elif st.session_state.active_tab == "Export":

    st.markdown(
        '<div class="section-label">05 — EXPORT</div>',
        unsafe_allow_html=True
    )
    st.header("Export Analysis Report")
    st.write("Create a professional report containing the file information, all graphs, graph-by-graph analysis, anomaly findings, dataset issues, and the final summary.")

    report = _prepare_export_report()


    if report is None:
        st.info("Generate a graph first. The export report will be available after flight data has been analyzed.")
    else:
        st.markdown("### Report Contents")
        st.caption("Each graph is exported with only its graph, Maximum, and Minimum.")

        st.divider()
        st.subheader("Choose Export Format")

        c1, c2, c3 = st.columns(3)

        if "export_docx" not in st.session_state:
            st.session_state.export_docx = None
        if "export_pdf" not in st.session_state:
            st.session_state.export_pdf = None
        if "export_xlsx" not in st.session_state:
            st.session_state.export_xlsx = None

        with c1:
            if st.button("📄 Export as Word (.docx)", use_container_width=True, key="make_docx"):
                with st.spinner("Preparing Word report..."):
                    graph_images, anomaly_images = _build_export_images(report)
                    st.session_state.export_docx = _generate_docx(report, graph_images, anomaly_images)
            if st.session_state.export_docx:
                st.download_button("⬇️ Download Word Report", st.session_state.export_docx, "helicopter_analysis_report.docx", "application/vnd.openxmlformats-officedocument.wordprocessingml.document", use_container_width=True, key="download_docx_report")

        with c2:
            if st.button("📕 Export as PDF (.pdf)", use_container_width=True, key="make_pdf"):
                with st.spinner("Preparing PDF report..."):
                    graph_images, anomaly_images = _build_export_images(report)
                    st.session_state.export_pdf = _generate_pdf(report, graph_images, anomaly_images)
            if st.session_state.export_pdf:
                st.download_button("⬇️ Download PDF Report", st.session_state.export_pdf, "helicopter_analysis_report.pdf", "application/pdf", use_container_width=True, key="download_pdf_report")

        with c3:
            if st.button("📊 Export as Excel (.xlsx)", use_container_width=True, key="make_xlsx"):
                with st.spinner("Preparing Excel report..."):
                    graph_images, anomaly_images = _build_export_images(report)
                    st.session_state.export_xlsx = _generate_xlsx(report, graph_images, anomaly_images)
            if st.session_state.export_xlsx:
                st.download_button("⬇️ Download Excel Report", st.session_state.export_xlsx, "helicopter_analysis_report.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True, key="download_xlsx_report")

        st.divider()
        st.caption("The exported reports use the same centralized report data so the Word, PDF, and Excel versions remain consistent.")

# ============================================================
# FOOTER
# ============================================================

st.divider()

st.caption(
    "HELICOPTER ANALYSIS SYSTEM • "
    "Flight Data Analytics Platform • "
    "Module 01"
)
