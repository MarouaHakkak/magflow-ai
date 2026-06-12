"""MagFlow AI v9 — Landing page, colorful UI, offline maintenance Excel, state-safe"""
import streamlit as st
import anthropic
from datetime import datetime
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation
import io
import json
import gspread
from google.oauth2.service_account import Credentials
from magflow_ai import MagFlowAI, ProcessInput
import base64
from github import Github

GITHUB_REPO = "MarouaHakkak/magflow-ai"
EXCEL_FILE_PATH = "Data_Collection_v2.xlsx"

# ============================================================
#  IMAGE URLS  —  REPLACE THE PLACEHOLDERS BELOW
#  (See the chat for step-by-step instructions on getting a
#   direct image URL. It must end in .png / .jpg and load on
#   its own in a browser tab.)
# ============================================================
IMG_JESA_LOGO = "https://i.imgur.com/K0ObAF4.png"
IMG_HERO      = "https://media.licdn.com/dms/image/v2/C561BAQEdzlQbQ6QRaQ/company-background_10000/company-background_10000/0/1583908059030?e=2147483647&v=beta&t=mn-3EDPCiDBT3qiDScVJVLzAlN_XvD2F3rDBL1YinXQ"
IMG_FLOWMETER = "https://upload.wikimedia.org/wikipedia/commons/2/24/%D0%A0%D0%B0%D1%81%D1%85%D0%BE%D0%B4%D0%BE%D0%BC%D0%B5%D1%80_%D1%8D%D0%BB%D0%B5%D0%BA%D1%82%D1%80%D0%BE%D0%BC%D0%B0%D0%B3%D0%BD%D0%B8%D1%82%D0%BD%D1%8B%D0%B9.jpg"  # magnetic flowmeter photo (Wikimedia Commons)

# Vendor website links (used in Layer 3 "Other Vendors" + main 3 cards)
VENDOR_LINKS = {
    "Emerson":          "https://www.emerson.com/en-us/automation/measurement-instrumentation/flow-measurement/magnetic-flow-meters",
    "Endress+Hauser":   "https://www.endress.com/en/field-instruments-overview/flow-measurement-product-overview/electromagnetic-flowmeter",
    "Endress Hauser":   "https://www.endress.com/en/field-instruments-overview/flow-measurement-product-overview/electromagnetic-flowmeter",
    "Krohne":           "https://krohne.com/en/products/flow-measurement/flowmeters/electromagnetic-flowmeters",
    "ABB":              "https://new.abb.com/products/measurement-products/flow/electromagnetic-flowmeters",
    "Siemens":          "https://www.siemens.com/global/en/products/automation/process-instrumentation/flow-measurement/electromagnetic.html",
    "Yokogawa":         "https://www.yokogawa.com/solutions/products-and-services/measurement/flow-meters/magnetic-flow-meters/",
    "Honeywell":        "https://process.honeywell.com/us/en/products/field-instruments/flow-measurement",
    "VEGA":             "https://www.vega.com/en-us/products/product-catalog/flow",
}

def vendor_url(name):
    if not name:
        return None
    key = name.strip()
    if key in VENDOR_LINKS:
        return VENDOR_LINKS[key]
    for k, v in VENDOR_LINKS.items():
        if k.lower() in key.lower() or key.lower() in k.lower():
            return v
    return None

st.set_page_config(page_title="MagFlow AI — JESA", page_icon="🔧", layout="wide")

# ============================================================
#  GLOBAL STYLING  —  colorful theme + spaced tabs
# ============================================================
st.markdown("""
<style>
/* Tab bar: bigger, clearer labels spread across the full page width (note 13) */
.stTabs [data-baseweb="tab-list"] {
    gap: 8px;
    display: flex;
    width: 100%;
}
.stTabs [data-baseweb="tab"] {
    flex: 1 1 0;
    justify-content: center;
    text-align: center;
}
.stTabs [data-baseweb="tab"] p {
    font-size: 16px !important;
    font-weight: 600 !important;
}
/* Soft page background */
.stApp { background: linear-gradient(180deg, #FBFCFF 0%, #F4F8FD 100%); }
/* Buttons */
.stButton button, .stDownloadButton button {
    border-radius: 8px;
    font-weight: 600;
}
/* Landing hero card */
.mf-hero {
    border-radius: 18px;
    overflow: hidden;
    box-shadow: 0 8px 30px rgba(21,42,74,0.18);
    margin-bottom: 8px;
}
.mf-card {
    background: #FFFFFF;
    border-radius: 14px;
    padding: 22px;
    box-shadow: 0 3px 14px rgba(21,42,74,0.08);
    height: 100%;
}
.mf-step {
    background: #FFFFFF;
    border-left: 5px solid #1565C0;
    border-radius: 10px;
