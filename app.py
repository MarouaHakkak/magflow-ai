"""MagFlow AI v9 — Landing page, colorful UI, offline maintenance Excel, state-safe"""
import streamlit as st
import anthropic
from datetime import datetime
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
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
    padding: 16px 18px;
    margin-bottom: 12px;
    box-shadow: 0 2px 10px rgba(21,42,74,0.06);
}
.mf-step b { color: #1565C0; }
</style>
""", unsafe_allow_html=True)

@st.cache_resource
def load_model():
    return MagFlowAI('Data_Collection_v2.xlsx')

ai = load_model()

SHEET_ID = "1jVpJkHPtG808WtlKxKpcIxvNLvLyx8syIi0GlppNTgU"
SCOPES = ["https://www.googleapis.com/auth/spreadsheets",
          "https://www.googleapis.com/auth/drive"]

@st.cache_resource
def get_gsheet():
    try:
        creds_dict = dict(st.secrets["gcp_service_account"])
        creds = Credentials.from_service_account_info(creds_dict, scopes=SCOPES)
        client = gspread.authorize(creds)
        sheet = client.open_by_key(SHEET_ID).sheet1
        return sheet
    except Exception as e:
        return None

def load_history(tag_filter=None):
    sheet = get_gsheet()
    if sheet is None:
        return []
    try:
        records = sheet.get_all_records()
        if tag_filter:
            records = [r for r in records if r.get('Tag Number','').strip() == tag_filter.strip()]
        return records
    except:
        return []

def save_history(date, tag, intervention_type, task, result, technician, notes):
    sheet = get_gsheet()
    if sheet is None:
        return False
    try:
        sheet.append_row([date, tag, intervention_type, task, result, technician, notes])
        return True
    except:
        return False

# ============================================================
#  CHECKLIST PERSISTENCE  (pre-fill the checklist for a tag)
#  Stores filled checklist cells in a Google Sheet tab "Checklist Records"
#  Key per cell: tag :: period :: task :: column_label
# ============================================================
CHECKLIST_WS = "Checklist Records"

def _checklist_key(tag, period, task, col_label):
    return f"{str(tag).strip()}::{str(period).strip()}::{str(task).strip()}::{str(col_label).strip()}"

def _checklist_ws():
    """Return the gspread worksheet for checklist records (create if missing)."""
    try:
        creds_dict = dict(st.secrets["gcp_service_account"])
        creds = Credentials.from_service_account_info(creds_dict, scopes=SCOPES)
        client_gs = gspread.authorize(creds)
        ss = client_gs.open_by_key(SHEET_ID)
        try:
            ws = ss.worksheet(CHECKLIST_WS)
        except gspread.WorksheetNotFound:
            ws = ss.add_worksheet(title=CHECKLIST_WS, rows=2000, cols=6)
            ws.append_row(["Tag", "Period", "Task", "Column", "Value", "Updated"])
        return ws
    except Exception:
        return None

def load_checklist_records(tag):
    """Return {key: value} of saved checklist cells for this tag."""
    ws = _checklist_ws()
    if ws is None:
        return {}
    try:
        out = {}
        for r in ws.get_all_records():
            if str(r.get("Tag", "")).strip() != str(tag).strip():
                continue
            key = _checklist_key(r.get("Tag", ""), r.get("Period", ""), r.get("Task", ""), r.get("Column", ""))
            out[key] = str(r.get("Value", ""))
        return out
    except Exception:
        return {}

def save_checklist_records(tag, cells):
    """cells = {key: value}. Last-write-wins: append rows; latest read wins on load
    because get_all_records keeps order and we overwrite in the dict."""
    ws = _checklist_ws()
    if ws is None:
        return 0
    try:
        now = datetime.now().strftime('%Y-%m-%d %H:%M')
        rows = []
        for key, val in cells.items():
            parts = key.split("::")
            if len(parts) != 4:
                continue
            t, period, task, col = parts
            rows.append([t, period, task, col, val, now])
        if rows:
            ws.append_rows(rows, value_input_option="RAW")
        return len(rows)
    except Exception:
        return 0

def read_checklist_from_excel(uploaded_file, tag):
    """Read the 'Maintenance Checklist' sheet of an uploaded Excel and return
    {key: value} for every filled data cell. Re-derives period from headers."""
    try:
        wb = openpyxl.load_workbook(uploaded_file)
        if "Maintenance Checklist" not in wb.sheetnames:
            return {}
        ws = wb["Maintenance Checklist"]
        out = {}
        current_month = None
        current_period = None
        mode = None  # continuous / dated
        for r in range(1, ws.max_row + 1):
            a = ws.cell(row=r, column=1).value
            a_str = str(a).strip() if a is not None else ""
            if a_str.startswith("📅"):
                toks = a_str.replace("📅", "").strip().split()
                current_month = toks[0] if toks else None
                continue
            if a_str.lstrip().startswith("Week"):
                wk = a_str.strip()
                current_period = f"{current_month}/{wk}"
                mode = "continuous"
                continue
            if "Monthly tasks" in a_str:
                current_period = f"{current_month}/Monthly"
                mode = "dated"
                continue
            if "Quarterly" in a_str and "—" in a_str:
                qlabel = a_str.split("—")[-1].strip()
                current_period = f"{current_month}/Quarterly {qlabel}"
                mode = "dated"
                continue
            if "Semi-annual" in a_str and "—" in a_str:
                slabel = a_str.split("—")[-1].strip()
                current_period = f"{current_month}/Semi-annual {slabel}"
                mode = "dated"
                continue
            if "ANNUAL MAINTENANCE" in a_str:
                current_period = "ANNUAL"
                mode = "dated"
                continue
            if "MULTI-YEAR" in a_str:
                current_period = "MULTI-YEAR"
                mode = "dated"
                continue
            if "COMPONENT REPLACEMENT" in a_str:
                current_period = "REPLACEMENT"
                mode = "dated"
                continue
            task = ws.cell(row=r, column=2).value
            if a_str == "☐" and task and current_period:
                task = str(task).strip()
                if mode == "continuous":
                    cols = [(3,"Mon"),(4,"Tue"),(5,"Wed"),(6,"Thu"),(7,"Fri"),(8,"Sat"),(9,"Sun"),(10,"Status"),(11,"Technician")]
                else:
                    cols = [(10,"Date"),(11,"Technician"),(12,"Work done")]
                for col, lab in cols:
                    val = ws.cell(row=r, column=col).value
                    if val is not None and str(val).strip() != "":
                        out[_checklist_key(tag, current_period, task, lab)] = str(val)
        return out
    except Exception:
        return {}

def import_from_excel(uploaded_file):
    try:
        wb = openpyxl.load_workbook(uploaded_file)
        if "Historical Data" not in wb.sheetnames:
            return 0, "Sheet 'Historical Data' not found in the uploaded file."
        ws = wb["Historical Data"]
        saved = 0
        for row in ws.iter_rows(min_row=4, values_only=True):
            date, tag, itype, task, result, tech, notes = (row[0], row[1], row[2], row[3], row[4], row[5], row[6] if len(row) > 6 else "")
            if not date or not tag or not task:
                continue
            success = save_history(str(date), str(tag), str(itype or ""), str(task), str(result or ""), str(tech or ""), str(notes or ""))
            if success:
                saved += 1
        return saved, None
    except Exception as e:
        return 0, str(e)

def build_context():
    fluids = ", ".join(f['name'] for f in ai.data.fluids[:25]) + f"... ({len(ai.data.fluids)} total)"
    return f"""You are MagFlow AI Assistant for JESA/OCP electromagnetic flowmeters.
Database: {len(ai.data.fluids)} fluids, {len(ai.data.electrodes)} electrodes, {len(ai.data.liners)} liners, {len(ai.data.vendor_models)} models, {len(ai.data.drift_indicators)} drift indicators, {len(ai.data.project_data)} project entries.
Fluids: {fluids}
Rules: PFA default for acids, PTFE for large DN, Soft rubber for rock slurry, Min 5 uS/cm, Max 10 m/s / 3-6 m/s slurries, Grounding ring for FRP/PVC.
Diagnostics: Emerson SMV, E+H Heartbeat, Krohne OPTICHECK.
Answer in same language as question. Source: [JESA Database], [General Knowledge], or [Web Search].
For new tech: USE web_search, provide URLs."""

def search_local(query):
    q = query.lower().replace('-',' '); results = []; words = [w for w in q.split() if len(w)>2]; combined = q.replace(' ','')
    for f in ai.data.fluids:
        fn = f['name'].lower(); fnc = fn.replace(' ','')
        if any(w in fn for w in words) or combined in fnc or fnc in combined:
            results.append(f"Fluid: {f['name']} | Type: {f['type']} | Electrode: {f['electrode']} | Liner: {f['liner']} | Grounding: {f['grounding']}")
    for e in ai.data.electrodes:
        if any(w in e['name'].lower() for w in words): results.append(f"Electrode: {e['name']} | Cost: {e['cost']}")
    for m in ai.data.vendor_models:
        if any(w in (m['model']+' '+m['vendor']).lower() for w in words): results.append(f"Model: {m['vendor']} {m['model']} | Accuracy: {m['accuracy']}")
    for d in ai.data.drift_indicators:
        if any(w in d['indicator'].lower() for w in words): results.append(f"Drift: {d['indicator']} | {d['desc']}")
    for p in ai.data.project_data:
        if any(w in (p.get('fluid','')+' '+p.get('project','')).lower() for w in words): results.append(f"Project: {p['project']} | {p['fluid']} | {p.get('electrode','')}")
    return results[:15]

def ask_claude(question, local_results, history=None):
    if history is None: history = []
    try:
        client = anthropic.Anthropic(api_key=st.secrets["ANTHROPIC_API_KEY"])
        ctx = ""
        if local_results: ctx = "\n\nJESA data:\n" + "\n".join(f"- {r}" for r in local_results)
        conv = ""
        if history:
            conv = "\n\nPrevious conversation:\n"
            for msg in history[-6:]:
                conv += f"{'User' if msg['role']=='user' else 'Assistant'}: {msg['content'][:200]}\n"
        response = client.messages.create(model="claude-haiku-4-5-20251001", max_tokens=1500, system=build_context(),
            tools=[{"type": "web_search_20250305", "name": "web_search"}],
            messages=[{"role": "user", "content": question + ctx + conv}])
        parts = [block.text for block in response.content if hasattr(block, 'text')]
        return "\n".join(parts) if parts else "No response."
    except Exception as e:
        return f"Error: {str(e)}"

def extract_datasheet_with_ai(pdf_bytes):
    """Extract instruments page by page to avoid JSON truncation."""
    try:
        from pypdf import PdfReader
        import io as _io
        from pypdf import PdfWriter

        client = anthropic.Anthropic(api_key=st.secrets["ANTHROPIC_API_KEY"])

        prompt = """This is ONE page from a JESA magnetic flowmeter datasheet.
If this page contains an instrument specification (has a Tag Number and process data), extract it as JSON.
If not (cover page, index, notes), return exactly: []
Return ONLY a JSON array with one object per instrument found:
[{"project":"...","tag":"...","service":"...","fluid":"...","dn":80,"flow_normal":23.0,"flow_max":26.0,"temp_design":105,"pressure_design":10.0,"conductivity":">5 uS/cm","electrode":"Tantalum","liner":"PFA","tube":"316L SS","grounding":"Grounding straps","accuracy":"0.2%","vendor":"VTA","model":"VTA"}]
Use null for missing numeric values. Return ONLY JSON, no markdown."""

        reader = PdfReader(_io.BytesIO(pdf_bytes))
        all_instruments = []

        for page_num, page in enumerate(reader.pages):
            try:
                writer = PdfWriter()
                writer.add_page(page)
                page_buf = _io.BytesIO()
                writer.write(page_buf)
                page_buf.seek(0)
                page_b64 = base64.standard_b64encode(page_buf.read()).decode("utf-8")

                response = client.messages.create(
                    model="claude-haiku-4-5-20251001",
                    max_tokens=1000,
                    messages=[{"role": "user", "content": [
                        {"type": "document", "source": {"type": "base64", "media_type": "application/pdf", "data": page_b64}},
                        {"type": "text", "text": prompt}
                    ]}])

                text = response.content[0].text.strip().replace("```json","").replace("```","").strip()
                start = text.find('[')
                end = text.rfind(']')
                if start != -1 and end != -1:
                    page_instruments = json.loads(text[start:end+1])
                    all_instruments.extend(page_instruments)
            except:
                continue

        return all_instruments, None if all_instruments else ([], "No instruments found in any page of this PDF.")

    except Exception as e:
        return [], str(e)

def get_excel_from_github():
    try:
        token = st.secrets.get("GITHUB_TOKEN", None)
        if token is None:
            try: token = st.secrets["github"]["GITHUB_TOKEN"]
            except: pass
        if token is None:
            raise Exception("GITHUB_TOKEN not found in secrets")
        gh = Github(token)
        repo = gh.get_repo(GITHUB_REPO)
        contents = repo.get_contents(EXCEL_FILE_PATH)
        excel_bytes = base64.b64decode(contents.content)
        wb = openpyxl.load_workbook(io.BytesIO(excel_bytes))
        return wb, contents.sha, None
    except Exception as e:
        return None, None, str(e)

def push_excel_to_github(wb, sha, commit_msg):
    try:
        token = st.secrets.get("GITHUB_TOKEN", None)
        if token is None:
            try: token = st.secrets["github"]["GITHUB_TOKEN"]
            except: pass
        if token is None:
            raise Exception("GITHUB_TOKEN not found in secrets")
        gh = Github(token)
        repo = gh.get_repo(GITHUB_REPO)
        buf = io.BytesIO()
        wb.save(buf)
        buf.seek(0)
        new_content = base64.b64encode(buf.read()).decode("utf-8")
        repo.update_file(EXCEL_FILE_PATH, commit_msg, base64.b64decode(new_content), sha)
        return True, None
    except Exception as e:
        return False, str(e)

def get_existing_materials(wb):
    existing = {"electrodes": set(), "liners": set(), "fluids": set(), "fluid_electrodes": {}}
    try:
        if "Electrode Materials" in wb.sheetnames:
            ws = wb["Electrode Materials"]
            for row in ws.iter_rows(min_row=2, values_only=True):
                if row[0]: existing["electrodes"].add(str(row[0]).strip())
        if "Liner Materials" in wb.sheetnames:
            ws = wb["Liner Materials"]
            for row in ws.iter_rows(min_row=2, values_only=True):
                if row[0]: existing["liners"].add(str(row[0]).strip())
        if "Fluid-Material Matrix" in wb.sheetnames:
            ws = wb["Fluid-Material Matrix"]
            for row in ws.iter_rows(min_row=2, values_only=True):
                if row[0]:
                    existing["fluids"].add(str(row[0]).strip())
                    if row[1]: existing["fluid_electrodes"][str(row[0]).strip()] = str(row[1]).strip()
    except:
        pass
    return existing

def detect_new_materials(instruments, existing):
    new_findings = []
    for inst in instruments:
        electrode = str(inst.get('electrode') or '').strip()
        liner = str(inst.get('liner') or '').strip()
        fluid = str(inst.get('fluid') or '').strip()
        tag = inst.get('tag', '')
        if electrode and electrode not in ('VTA', 'N/A', 'null', '') and electrode not in existing["electrodes"]:
            new_findings.append({"type": "electrode", "value": electrode, "tag": tag, "fluid": fluid, "context": f"Found in {tag} for {fluid}"})
        if liner and liner not in ('VTA', 'N/A', 'null', '') and liner not in existing["liners"]:
            new_findings.append({"type": "liner", "value": liner, "tag": tag, "fluid": fluid, "context": f"Found in {tag} for {fluid}"})
        if fluid and fluid not in ('VTA', 'N/A', 'null', '') and fluid not in existing["fluids"]:
            new_findings.append({"type": "fluid", "value": fluid, "tag": tag, "electrode": electrode, "liner": liner, "context": f"New fluid in {tag}"})
        elif fluid in existing["fluid_electrodes"] and electrode and electrode not in ('VTA', 'N/A', 'null', ''):
            existing_elec = existing["fluid_electrodes"][fluid]
            if electrode not in existing_elec:
                new_findings.append({"type": "fluid_electrode_variant", "value": electrode, "tag": tag, "fluid": fluid, "context": f"New electrode variant for {fluid} (current: {existing_elec})"})
    return new_findings

def apply_updates_to_excel(wb, approved_updates):
    changes = []
    for update in approved_updates:
        utype = update["type"]
        value = update["value"]
        try:
            if utype == "electrode":
                ws = wb["Electrode Materials"]
                ws.cell(row=ws.max_row + 1, column=1, value=value)
                ws.cell(row=ws.max_row, column=2, value="Auto-imported")
                ws.cell(row=ws.max_row, column=3, value=datetime.now().strftime('%Y-%m-%d'))
                changes.append(f"Added electrode: {value}")
            elif utype == "liner":
                ws = wb["Liner Materials"]
                ws.cell(row=ws.max_row + 1, column=1, value=value)
                ws.cell(row=ws.max_row, column=2, value="Auto-imported")
                ws.cell(row=ws.max_row, column=3, value=datetime.now().strftime('%Y-%m-%d'))
                changes.append(f"Added liner: {value}")
            elif utype == "fluid":
                ws = wb["Fluid-Material Matrix"]
                ws.cell(row=ws.max_row + 1, column=1, value=value)
                ws.cell(row=ws.max_row, column=2, value=update.get("electrode", "VTA"))
                ws.cell(row=ws.max_row, column=3, value=update.get("liner", "VTA"))
                changes.append(f"Added fluid: {value}")
            elif utype == "fluid_electrode_variant":
                ws = wb["Fluid-Material Matrix"]
                for row in ws.iter_rows(min_row=2):
                    if row[0].value and str(row[0].value).strip() == update["fluid"]:
                        current = str(row[1].value or "")
                        if value not in current:
                            row[1].value = current + " / " + value if current else value
                            changes.append(f"Added electrode variant {value} to {update['fluid']}")
                        break
        except Exception as e:
            changes.append(f"Error: {str(e)}")
    return wb, changes

def save_project_to_gsheet(instruments, project_name):
    try:
        creds_dict = dict(st.secrets["gcp_service_account"])
        creds = Credentials.from_service_account_info(creds_dict, scopes=["https://www.googleapis.com/auth/spreadsheets"])
        client_gs = gspread.authorize(creds)
        spreadsheet = client_gs.open_by_key(SHEET_ID)
        try:
            ws = spreadsheet.worksheet("Project Data")
        except:
            ws = spreadsheet.add_worksheet(title="Project Data", rows=1000, cols=20)
            ws.append_row(["Project","Tag","Service","Fluid","DN","Flow Normal","Flow Max",
                           "Temp Design","Pressure Design","Conductivity","Electrode","Liner",
                           "Tube","Grounding","Accuracy","Vendor","Model","Import Date"])
        today = datetime.now().strftime('%Y-%m-%d')
        saved = 0
        for inst in instruments:
            ws.append_row([inst.get('project', project_name) or project_name, inst.get('tag',''), inst.get('service',''),
                           inst.get('fluid',''), inst.get('dn',''), inst.get('flow_normal',''), inst.get('flow_max',''),
                           inst.get('temp_design',''), inst.get('pressure_design',''), inst.get('conductivity',''),
                           inst.get('electrode',''), inst.get('liner',''), inst.get('tube',''),
                           inst.get('grounding',''), inst.get('accuracy',''), inst.get('vendor',''),
                           inst.get('model',''), today])
            saved += 1
        return saved, None
    except Exception as e:
        return 0, str(e)

def load_project_data_from_gsheet():
    try:
        creds_dict = dict(st.secrets["gcp_service_account"])
        creds = Credentials.from_service_account_info(creds_dict, scopes=["https://www.googleapis.com/auth/spreadsheets"])
        client_gs = gspread.authorize(creds)
        spreadsheet = client_gs.open_by_key(SHEET_ID)
        ws = spreadsheet.worksheet("Project Data")
        return ws.get_all_records()
    except:
        return []

def generate_maintenance_excel(tag, fluid, cat, mat, vendor_e, vendor_eh, vendor_k, tco, history_records=None, checklist_data=None):
    wb = openpyxl.Workbook()
    vn = vendor_e.model if vendor_e else (vendor_eh.model if vendor_eh else (vendor_k.model if vendor_k else 'N/A'))
    vd = vendor_e.diagnostics if vendor_e else (vendor_eh.diagnostics if vendor_eh else (vendor_k.diagnostics if vendor_k else 'N/A'))
    today = datetime.now().strftime('%Y-%m-%d')
    year = datetime.now().year
    mnt = tco.maintenance
    checklist_data = checklist_data or {}
    def _pf(period, task, col_label):
        """Return saved value for this cell, or None."""
        return checklist_data.get(f"{str(tag).strip()}::{period}::{str(task).strip()}::{col_label}")
    title_font = Font(bold=True, size=14, color="1B2A4A")
    h_font = Font(bold=True, size=11, color="FFFFFF")
    n_font = Font(size=10)
    h_fill = PatternFill(start_color="1B2A4A", end_color="1B2A4A", fill_type="solid")
    g_fill = PatternFill(start_color="E8F5E9", end_color="E8F5E9", fill_type="solid")
    b_fill = PatternFill(start_color="E3F2FD", end_color="E3F2FD", fill_type="solid")
    o_fill = PatternFill(start_color="FFF3E0", end_color="FFF3E0", fill_type="solid")
    p_fill = PatternFill(start_color="F3E5F5", end_color="F3E5F5", fill_type="solid")
    y_fill = PatternFill(start_color="FFF9C4", end_color="FFF9C4", fill_type="solid")
    r_fill = PatternFill(start_color="FFCDD2", end_color="FFCDD2", fill_type="solid")
    q_fill = PatternFill(start_color="FCE4D6", end_color="FCE4D6", fill_type="solid")
    thin = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))
    center = Alignment(horizontal='center', vertical='center', wrap_text=True)
    wrap = Alignment(wrap_text=True, vertical='top')
    ws1 = wb.active; ws1.title = "Guideline"
    ws1.column_dimensions['A'].width = 30; ws1.column_dimensions['B'].width = 55
    ws1.merge_cells('A1:B1'); ws1['A1'] = "🔧 MagFlow AI — Instrument Guideline"; ws1['A1'].font = title_font
    ws1['A2'] = f"Generated: {today}"; ws1['A2'].font = Font(italic=True, color="888888")
    row = 4
    ws1.merge_cells(f'A{row}:B{row}'); ws1[f'A{row}'] = "INSTRUMENT INFORMATION"; ws1[f'A{row}'].font = h_font; ws1[f'A{row}'].fill = h_fill
    row += 1
    info = [("Tag Number", tag), ("Service Fluid", fluid), ("Fluid Category", cat),
        ("Electrode", f"{mat.electrode} (Cost: {mat.electrode_cost})"), ("Liner", f"{mat.liner} (Cost: {mat.liner_cost})"),
        ("Body Material", "SS 304L (coil housing)"), ("Tube Material", "SS 316L"),
        ("Grounding", mat.grounding), ("Penetrant Ring", mat.penetrant or "N/A"), ("O-Ring", mat.o_ring),
        ("Flange Coating", mat.flange_coat), ("Vendor Model", vn), ("Diagnostics", vd[:80]),
        ("CAPEX Score", f"{tco.capex_score}/30"), ("Calibration", f"Every {tco.calib_months} months"),
        ("Liner Life", tco.liner_life), ("Electrode Life", tco.electrode_life)]
    for k, v in info:
        ws1[f'A{row}'] = k; ws1[f'A{row}'].font = Font(bold=True); ws1[f'A{row}'].fill = b_fill; ws1[f'A{row}'].border = thin
        ws1[f'B{row}'] = v; ws1[f'B{row}'].border = thin; row += 1
    row += 1
    ws1.merge_cells(f'A{row}:B{row}'); ws1[f'A{row}'] = "DRIFT RISK ASSESSMENT"; ws1[f'A{row}'].font = h_font; ws1[f'A{row}'].fill = h_fill
    row += 1
    for risk in tco.drift_risks:
        icon = "🔴" if risk.level == "High" else "🟡" if risk.level == "Medium" else "🟢"
        fill = r_fill if risk.level == "High" else y_fill if risk.level == "Medium" else g_fill
        ws1.merge_cells(f'A{row}:B{row}')
        ws1[f'A{row}'] = f"{icon} {risk.indicator} — {risk.level}"; ws1[f'A{row}'].font = Font(bold=True, size=11); ws1[f'A{row}'].fill = fill; row += 1
        ws1[f'A{row}'] = "Description:"; ws1[f'A{row}'].font = Font(bold=True)
        ws1[f'B{row}'] = risk.description; ws1[f'B{row}'].alignment = wrap; row += 1
        ws1[f'A{row}'] = "Diagnostic Steps:"; ws1[f'A{row}'].font = Font(bold=True)
        ws1[f'B{row}'] = "\n".join(risk.steps); ws1[f'B{row}'].alignment = wrap
        ws1.row_dimensions[row].height = max(15 * len(risk.steps), 30); row += 1
        ws1[f'A{row}'] = "Frequency:"; ws1[f'A{row}'].font = Font(bold=True)
        ws1[f'B{row}'] = risk.frequency; row += 2
    ws2 = wb.create_sheet("Maintenance Checklist")
    ws2.column_dimensions['A'].width = 5; ws2.column_dimensions['B'].width = 45
    for col in range(3, 10): ws2.column_dimensions[get_column_letter(col)].width = 10
    ws2.column_dimensions['J'].width = 15; ws2.column_dimensions['K'].width = 20; ws2.column_dimensions['L'].width = 40
    ws2.merge_cells('A1:K1'); ws2['A1'] = f"🔧 Maintenance Checklist — {tag}"; ws2['A1'].font = title_font
    ws2['A2'] = f"Fluid: {fluid} | Category: {cat} | Year: {year}"; ws2['A2'].font = Font(italic=True, color="666666")

    # ---- Instructions table on the right side (cols O+) — note 17 ----
    ws2.column_dimensions['O'].width = 22
    ws2.column_dimensions['P'].width = 50
    instr_title_fill = PatternFill(start_color="1565C0", end_color="1565C0", fill_type="solid")
    instr_hdr_fill = PatternFill(start_color="BBDEFB", end_color="BBDEFB", fill_type="solid")
    ws2.merge_cells('O2:P2')
    ws2['O2'] = "📌 How to Use This Checklist"; ws2['O2'].font = Font(bold=True, size=13, color="FFFFFF")
    ws2['O2'].fill = instr_title_fill; ws2['O2'].alignment = center
    irow = 4
    instr_rows = [
        ("Column / Symbol", "What to do", True),
        ("☐ checkbox", "Tick (✔) once the task is completed for that period.", False),
        ("Mon → Sun", "For continuous tasks, mark the day the check was done.", False),
        ("Status", "Write: OK / Conform, or NC / Non-conform if a problem is found.", False),
        ("Technician", "Enter the full name of the person who did the task.", False),
        ("Date", "For monthly/quarterly/annual tasks, write the date done (YYYY-MM-DD).", False),
        ("Work done", "For monthly/quarterly tasks, describe the action actually performed.", False),
        ("Notes", "Add any observation, anomaly, or action taken.", False),
        ("", "", False),
        ("Cadence", "Filling frequency", True),
        ("Continuous", "Checked every week (transmitter self-test, empty-pipe detection).", False),
        ("Monthly", "Once per month (HART/diagnostic alarms, display check).", False),
        ("Quarterly (Q1–Q4)", "Every 3 months (calibration verification, grounding).", False),
        ("Semi-annual (S1–S2)", "Every 6 months (liner/electrode inspection).", False),
        ("Annual / Multi-year", "Once a year or every 3–5 years (full calibration, replacement).", False),
        ("", "", False),
        ("If Non-conform (NC)", "Log it in the 'Historical Data' sheet and open a corrective action.", True),
    ]
    for label, desc, is_header in instr_rows:
        if label == "" and desc == "":
            irow += 1; continue
        cl = ws2.cell(row=irow, column=15, value=label)  # col O
        cd = ws2.cell(row=irow, column=16, value=desc)    # col P
        cd.alignment = wrap
        if is_header:
            cl.font = Font(bold=True, size=10, color="0D47A1"); cl.fill = instr_hdr_fill
            cd.font = Font(bold=True, size=10, color="0D47A1"); cd.fill = instr_hdr_fill
        else:
            cl.font = Font(bold=True, size=10); cd.font = Font(size=10)
        cl.border = thin; cd.border = thin
        irow += 1

    row = 4
    months_data = [("January",4),("February",4),("March",4),("April",4),("May",4),("June",4),
                   ("July",4),("August",4),("September",4),("October",4),("November",4),("December",4)]
    quarter_months = {3:"Q1",6:"Q2",9:"Q3",12:"Q4"}
    semester_months = {6:"S1",12:"S2"}
    for m_idx, (month_name, weeks) in enumerate(months_data, 1):
        ws2.merge_cells(f'A{row}:K{row}')
        ws2[f'A{row}'] = f"📅 {month_name} {year}"; ws2[f'A{row}'].font = Font(bold=True, size=12, color="1565C0"); ws2[f'A{row}'].fill = b_fill; row += 1
        if mnt.continuous:
            for w in range(1, weeks + 1):
                ws2.merge_cells(f'A{row}:B{row}')
                ws2[f'A{row}'] = f"   Week {w}"; ws2[f'A{row}'].font = Font(bold=True, size=10, color="2E7D32")
                for d, day in enumerate(["Mon","Tue","Wed","Thu","Fri","Sat","Sun"], 3):
                    c = ws2.cell(row=row, column=d, value=day)
                    c.font = Font(bold=True, size=9); c.alignment = center; c.fill = g_fill; c.border = thin
                ws2.cell(row=row, column=10, value="Status").font = Font(bold=True, size=9); ws2.cell(row=row, column=10).fill = g_fill; ws2.cell(row=row, column=10).border = thin
                ws2.cell(row=row, column=11, value="Technician").font = Font(bold=True, size=9); ws2.cell(row=row, column=11).fill = g_fill; ws2.cell(row=row, column=11).border = thin
                row += 1
                for item in mnt.continuous:
                    ws2.cell(row=row, column=1, value="☐").alignment = center
                    ws2.cell(row=row, column=2, value=item).font = n_font; ws2.cell(row=row, column=2).border = thin
                    _period_c = f"{month_name}/Week {w}"
                    _cols_c = {3:"Mon",4:"Tue",5:"Wed",6:"Thu",7:"Fri",8:"Sat",9:"Sun",10:"Status",11:"Technician"}
                    for d in range(3, 12):
                        _sv = _pf(_period_c, item, _cols_c[d])
                        if _sv is not None: ws2.cell(row=row, column=d, value=_sv)
                        ws2.cell(row=row, column=d).border = thin; ws2.cell(row=row, column=d).alignment = center
                    row += 1
                row += 1
        if mnt.monthly:
            ws2.merge_cells(f'A{row}:B{row}')
            ws2[f'A{row}'] = f"   📋 Monthly tasks — {month_name}"; ws2[f'A{row}'].font = Font(bold=True, size=10, color="1565C0"); ws2[f'A{row}'].fill = PatternFill(start_color="BBDEFB", end_color="BBDEFB", fill_type="solid")
            ws2.cell(row=row, column=10, value="Date").font = Font(bold=True, size=9); ws2.cell(row=row, column=10).border = thin
            ws2.cell(row=row, column=11, value="Technician").font = Font(bold=True, size=9); ws2.cell(row=row, column=11).border = thin
            ws2.cell(row=row, column=12, value="Work done").font = Font(bold=True, size=9); ws2.cell(row=row, column=12).border = thin; row += 1
            for item in mnt.monthly:
                ws2.cell(row=row, column=1, value="☐").alignment = center
                ws2.cell(row=row, column=2, value=item).font = n_font; ws2.cell(row=row, column=2).border = thin
                _period_m = f"{month_name}/Monthly"
                for _c,_lab in [(10,"Date"),(11,"Technician"),(12,"Work done")]:
                    _sv = _pf(_period_m, item, _lab)
                    if _sv is not None: ws2.cell(row=row, column=_c, value=_sv)
                ws2.cell(row=row, column=10).border = thin; ws2.cell(row=row, column=11).border = thin
                ws2.cell(row=row, column=12).border = thin; ws2.cell(row=row, column=12).alignment = wrap; row += 1
            row += 1
        if m_idx in quarter_months and mnt.quarterly:
            ws2.merge_cells(f'A{row}:B{row}')
            ws2[f'A{row}'] = f"   🔍 Quarterly — {quarter_months[m_idx]}"; ws2[f'A{row}'].font = Font(bold=True, size=10, color="E65100"); ws2[f'A{row}'].fill = o_fill
            ws2.cell(row=row, column=10, value="Date").font = Font(bold=True, size=9); ws2.cell(row=row, column=10).border = thin
            ws2.cell(row=row, column=11, value="Technician").font = Font(bold=True, size=9); ws2.cell(row=row, column=11).border = thin
            ws2.cell(row=row, column=12, value="Work done").font = Font(bold=True, size=9); ws2.cell(row=row, column=12).border = thin; row += 1
            for item in mnt.quarterly:
                ws2.cell(row=row, column=1, value="☐").alignment = center
                ws2.cell(row=row, column=2, value=item).font = n_font; ws2.cell(row=row, column=2).border = thin
                _period_q = f"{month_name}/Quarterly {quarter_months[m_idx]}"
                for _c,_lab in [(10,"Date"),(11,"Technician"),(12,"Work done")]:
                    _sv = _pf(_period_q, item, _lab)
                    if _sv is not None: ws2.cell(row=row, column=_c, value=_sv)
                ws2.cell(row=row, column=10).border = thin; ws2.cell(row=row, column=11).border = thin
                ws2.cell(row=row, column=12).border = thin; ws2.cell(row=row, column=12).alignment = wrap; row += 1
            row += 1
        if m_idx in semester_months and mnt.semi_annual:
            ws2.merge_cells(f'A{row}:B{row}')
            ws2[f'A{row}'] = f"   ⚙️ Semi-annual — {semester_months[m_idx]}"; ws2[f'A{row}'].font = Font(bold=True, size=10, color="BF360C"); ws2[f'A{row}'].fill = q_fill
            ws2.cell(row=row, column=10, value="Date").font = Font(bold=True, size=9); ws2.cell(row=row, column=10).border = thin
            ws2.cell(row=row, column=11, value="Technician").font = Font(bold=True, size=9); ws2.cell(row=row, column=11).border = thin; row += 1
            for item in mnt.semi_annual:
                ws2.cell(row=row, column=1, value="☐").alignment = center
                ws2.cell(row=row, column=2, value=item).font = n_font; ws2.cell(row=row, column=2).border = thin
                ws2.cell(row=row, column=10).border = thin; ws2.cell(row=row, column=11).border = thin; row += 1
            row += 1
    if mnt.annual:
        ws2.merge_cells(f'A{row}:K{row}')
        ws2[f'A{row}'] = f"🛠️ ANNUAL MAINTENANCE — {year}"; ws2[f'A{row}'].font = Font(bold=True, size=12, color="6A1B9A"); ws2[f'A{row}'].fill = p_fill; row += 1
        for item in mnt.annual:
            ws2.cell(row=row, column=1, value="☐").alignment = center
            ws2.cell(row=row, column=2, value=item).font = n_font; ws2.cell(row=row, column=2).border = thin
            ws2.cell(row=row, column=10).border = thin; ws2.cell(row=row, column=11).border = thin; row += 1
        row += 1
    if mnt.multi_year:
        ws2.merge_cells(f'A{row}:K{row}')
        ws2[f'A{row}'] = "📊 MULTI-YEAR MAINTENANCE (3-5 years)"; ws2[f'A{row}'].font = Font(bold=True, size=12, color="F57F17"); ws2[f'A{row}'].fill = y_fill; row += 1
        for item in mnt.multi_year:
            ws2.cell(row=row, column=1, value="☐").alignment = center
            ws2.cell(row=row, column=2, value=item).font = n_font; ws2.cell(row=row, column=2).border = thin
            ws2.cell(row=row, column=10).border = thin; ws2.cell(row=row, column=11).border = thin; row += 1
        row += 1
    if mnt.replacement:
        ws2.merge_cells(f'A{row}:K{row}')
        ws2[f'A{row}'] = "♻️ COMPONENT REPLACEMENT"; ws2[f'A{row}'].font = Font(bold=True, size=12, color="C62828"); ws2[f'A{row}'].fill = r_fill; row += 1
        for item in mnt.replacement:
            ws2.cell(row=row, column=1, value="☐").alignment = center
            ws2.cell(row=row, column=2, value=item).font = n_font; ws2.cell(row=row, column=2).border = thin
            ws2.cell(row=row, column=10).border = thin; ws2.cell(row=row, column=11).border = thin; row += 1
        row += 1
    ws3 = wb.create_sheet("Historical Data")
    for col, w in [(1,15),(2,20),(3,20),(4,35),(5,15),(6,20),(7,30)]:
        ws3.column_dimensions[get_column_letter(col)].width = w
    ws3.merge_cells('A1:G1')
    ws3['A1'] = f"📋 Maintenance History — {tag}"; ws3['A1'].font = title_font
    headers = ["Date", "Tag Number", "Type", "Task Performed", "Result", "Technician", "Notes"]
    for col, h in enumerate(headers, 1):
        c = ws3.cell(row=3, column=col, value=h)
        c.font = h_font; c.fill = h_fill; c.border = thin; c.alignment = center
    if history_records:
        for r_idx, record in enumerate(history_records, 4):
            for col, val in enumerate([record.get('Date',''), record.get('Tag Number',''), record.get('Type',''),
                                        record.get('Task',''), record.get('Result',''), record.get('Technician',''), record.get('Notes','')], 1):
                c = ws3.cell(row=r_idx, column=col, value=val)
                c.border = thin
                c.fill = b_fill if r_idx % 2 == 0 else PatternFill(fill_type=None)
    else:
        ws3.merge_cells('A4:G4')
        ws3['A4'] = "No interventions recorded yet. Use the 'Maintenance History' tab in MagFlow AI to log interventions."
        ws3['A4'].font = Font(italic=True, color="888888")
    buf = io.BytesIO(); wb.save(buf); buf.seek(0)
    return buf

# ============================================================
#  PAGE ROUTER  —  landing page is NOT a tab; it's shown first
# ============================================================
if "page" not in st.session_state:
    st.session_state.page = "home"

# ------------------------------------------------------------
#  LANDING PAGE
# ------------------------------------------------------------
def render_landing():
    # Top bar: JESA logo + title
    lc1, lc2 = st.columns([1, 4])
    with lc1:
        if IMG_JESA_LOGO and not IMG_JESA_LOGO.startswith("REPLACE_"):
            st.image(IMG_JESA_LOGO, width=160)
    with lc2:
        st.markdown(
            "<h1 style='color:#1B2A4A;margin-bottom:0'>🔧 MagFlow AI</h1>"
            "<p style='font-size:18px;color:#1565C0;margin-top:2px;font-weight:600'>"
            "AI-Based Predictive Maintenance System for Electromagnetic Flowmeters</p>"
            "<p style='font-size:13px;color:#888;margin-top:0'>PFE ENSA — JESA (OCP × Worley) · "
            "Instrumentation & Control · Developed by Maroua Hakkak — 2026</p>",
            unsafe_allow_html=True)

    # Hero image
    if IMG_HERO and not IMG_HERO.startswith("REPLACE_"):
        st.markdown(f"<div class='mf-hero'><img src='{IMG_HERO}' style='width:100%;display:block'></div>",
                    unsafe_allow_html=True)
    else:
        st.info("🖼️ Hero image goes here — set `IMG_HERO` at the top of app.py (see chat for how to get the URL).")

    st.markdown("<br>", unsafe_allow_html=True)

    # Enter button (top, easy to find)
    ec1, ec2, ec3 = st.columns([1, 2, 1])
    with ec2:
        if st.button("🚀 Enter MagFlow AI", type="primary", use_container_width=True):
            st.session_state.page = "app"
            st.rerun()

    st.divider()

    # What it is / what it does
    st.markdown("<h2 style='color:#1B2A4A'>What is MagFlow AI?</h2>", unsafe_allow_html=True)
    st.markdown(
        "<p style='font-size:16px;color:#444'>MagFlow AI is an AI-based predictive-maintenance assistant "
        "<b>developed and validated on electromagnetic flowmeters, with a generic and extensible architecture</b>. "
        "It turns a flowmeter's process conditions into a complete engineering recommendation: the right materials, "
        "compatible vendor models, total cost of ownership, drift risk, and a year-round maintenance plan — "
        "grounded in a database of real JESA/OCP projects.</p>",
        unsafe_allow_html=True)

    st.markdown("<h3 style='color:#1565C0'>What it does</h3>", unsafe_allow_html=True)
    f1, f2, f3, f4 = st.columns(4)
    for col, (icon, title, body, color) in zip(
        [f1, f2, f3, f4],
        [("✅", "Validation", "Checks if a magnetic flowmeter suits your fluid, velocity and conductivity.", "#2E7D32"),
         ("🧪", "Materials", "Selects electrode, liner, grounding, O-ring and coating for the fluid.", "#1565C0"),
         ("🏭", "Vendors", "Matches compatible models from Emerson, Endress+Hauser, Krohne & others.", "#E65100"),
         ("📊", "TCO & Drift", "Estimates cost of ownership, drift risk and a full maintenance plan.", "#6A1B9A")]):
        with col:
            st.markdown(
                f"<div class='mf-card'><div style='font-size:30px'>{icon}</div>"
                f"<p style='font-size:17px;font-weight:700;color:{color};margin:6px 0'>{title}</p>"
                f"<p style='font-size:14px;color:#555'>{body}</p></div>",
                unsafe_allow_html=True)

    st.markdown("<br>", unsafe_allow_html=True)

    # Magnetic flowmeter — photo + description
    st.markdown("<h3 style='color:#1565C0'>The magnetic flowmeter</h3>", unsafe_allow_html=True)
    pc1, pc2 = st.columns([2, 3])
    with pc1:
        if IMG_FLOWMETER and not IMG_FLOWMETER.startswith("REPLACE_"):
            st.image(IMG_FLOWMETER, use_container_width=True, caption="Electromagnetic flowmeter")
        else:
            st.info("🖼️ Flowmeter photo goes here — set `IMG_FLOWMETER` at the top of app.py.")
    with pc2:
        st.markdown(
            "<div class='mf-card'>"
            "<p style='font-size:17px;font-weight:700;color:#1B2A4A;margin:0 0 8px 0'>Magnetic Flowmeter</p>"
            "<p style='font-size:15px;color:#444;line-height:1.6;margin:0'>"
            "A magnetic (electromagnetic) flowmeter measures the flow of conductive liquids using "
            "Faraday's law of induction. As the fluid passes through a magnetic field, it generates a "
            "voltage proportional to its velocity, picked up by two electrodes in the tube wall. "
            "With no moving parts and no obstruction in the flow path, it produces no pressure drop and "
            "needs little maintenance — making it ideal for water, slurries, acids, and abrasive fluids "
            "common in OCP/JESA process lines.</p></div>",
            unsafe_allow_html=True)

    st.markdown("<br>", unsafe_allow_html=True)

    # How to use it (first-time user steps)
    st.markdown("<h2 style='color:#1B2A4A'>How to use it</h2>", unsafe_allow_html=True)
    steps = [
        ("1 · Enter process data", "Open the <b>Recommendation Engine</b>, then fill the fluid, pipe, operating conditions and any special notes."),
        ("2 · Run the engine", "Click <b>Run Recommendation</b>. You'll get all four layers: validation, materials, vendors, and TCO & drift."),
        ("3 · Get your maintenance sheet", "Download the maintenance Excel — Guideline, full month-by-month checklist, and historical data."),
        ("4 · Log interventions over time", "Use <b>Maintenance History</b> to record each intervention by tag — years later, search the tag to find your work."),
        ("5 · Import projects", "Use <b>Project Import</b> to upload a datasheet PDF; the AI extracts instruments and flags new materials."),
    ]
    for title, body in steps:
        st.markdown(f"<div class='mf-step'><b>{title}</b><br>{body}</div>", unsafe_allow_html=True)

    st.divider()
    ec1, ec2, ec3 = st.columns([1, 2, 1])
    with ec2:
        if st.button("🚀 Get Started", type="primary", use_container_width=True, key="enter_bottom"):
            st.session_state.page = "app"
            st.rerun()

if st.session_state.page == "home":
    render_landing()
    st.stop()

# ============================================================
#  SIDEBAR (only inside the app, not on landing)
# ============================================================
with st.sidebar:
    if IMG_JESA_LOGO and not IMG_JESA_LOGO.startswith("REPLACE_"):
        st.image(IMG_JESA_LOGO, width=130)
    if st.button("🏠 Home", use_container_width=True):
        st.session_state.page = "home"
        st.rerun()
    st.divider()
    st.markdown("### 🤖 MagFlow Assistant")
    lang_options = {"🇫🇷 French": "fr-FR", "🇬🇧 English": "en-US", "🇸🇦 Arabic": "ar-SA", "🇪🇸 Spanish": "es-ES"}
    selected_lang = st.selectbox("Voice language", list(lang_options.keys()), label_visibility="collapsed")
    lang_code = lang_options[selected_lang]
    voice_html = f"""<div style="display:flex;gap:8px;align-items:center">
    <button id="voiceBtn" onclick="startVoice()" style="background:#1B2A4A;color:white;border:none;padding:8px 16px;border-radius:6px;cursor:pointer;font-size:13px">🎤 Tap to speak</button>
    <button id="sendBtn" onclick="sendVoice()" style="background:#2E7D32;color:white;border:none;padding:8px 16px;border-radius:6px;cursor:pointer;font-size:13px;display:none">📤 Send</button>
    <span id="status" style="font-size:12px;color:#666"></span></div>
    <script>
    let recognition=null,transcript='';
    function startVoice(){{
        if(!('webkitSpeechRecognition'in window||'SpeechRecognition'in window)){{document.getElementById('status').innerText='Chrome only';return;}}
        const SR=window.SpeechRecognition||window.webkitSpeechRecognition;
        recognition=new SR();recognition.lang='{lang_code}';recognition.continuous=true;recognition.interimResults=true;
        recognition.onstart=()=>{{document.getElementById('voiceBtn').innerText='🔴 Listening...';document.getElementById('status').innerText='';document.getElementById('sendBtn').style.display='inline-block';}};
        recognition.onresult=(e)=>{{transcript=Array.from(e.results).map(r=>r[0].transcript).join('');document.getElementById('status').innerText=transcript.slice(0,50)+'...';}};
        recognition.onerror=()=>{{document.getElementById('voiceBtn').innerText='🎤 Tap to speak';document.getElementById('sendBtn').style.display='none';}};
        recognition.start();
    }}
    function sendVoice(){{
        if(recognition)recognition.stop();
        if(transcript)window.parent.postMessage({{type:'voice_text', text:transcript}}, '*');
        document.getElementById('voiceBtn').innerText='🎤 Tap to speak';
        document.getElementById('sendBtn').style.display='none';
        transcript='';
    }}
    </script>"""
    st.components.v1.html(voice_html, height=80)
    vq = st.query_params.get("voice", None)
    if vq:
        st.query_params.clear()
        st.session_state.voice_input = vq
    st.caption("Or type:")
    if "messages" not in st.session_state:
        st.session_state.messages = [{"role": "assistant", "content": "👋 How can I help you today?"}]
    for msg in st.session_state.messages:
        with st.chat_message(msg["role"]): st.markdown(msg["content"])
    vt = st.session_state.pop("voice_input", None)
    prompt = st.chat_input("Ask a question...", key="ci")
    if vt: prompt = vt
    if prompt:
        st.session_state.messages.append({"role": "user", "content": prompt})
        with st.chat_message("user"): st.markdown(prompt)
        lr = search_local(prompt)
        ch = [{"role": m["role"], "content": m["content"]} for m in st.session_state.messages[1:] if m["role"] in ["user","assistant"]][:-1]
        with st.chat_message("assistant"):
            with st.spinner("Thinking..."): response = ask_claude(prompt, lr, ch)
            st.markdown(response)
        st.session_state.messages.append({"role": "assistant", "content": response})
    if st.button("🗑️ Clear", use_container_width=True):
        st.session_state.messages = [{"role": "assistant", "content": "👋 Cleared."}]

# ============================================================
#  APP HEADER + TABS  (Dashboard moved to LAST)
#  Title/logo is clickable -> returns to Home (note 15)
# ============================================================
# Top row: JESA logo (top-left, bigger) + discreet Home button (top-right)
htop1, htop2 = st.columns([5, 1])
with htop1:
    if IMG_JESA_LOGO and not IMG_JESA_LOGO.startswith("REPLACE_"):
        st.image(IMG_JESA_LOGO, width=160)
with htop2:
    if st.button("🏠 Home", key="title_home"):
        st.session_state.page = "home"
        st.rerun()
# Centered big MagFlow AI title + description
st.markdown(f"""
<style>
@import url('https://fonts.googleapis.com/css2?family=Comfortaa:wght@500;600;700&display=swap');
.mf-header {{ text-align:center; margin-top:4px; }}
.mf-title {{
    font-family:'Comfortaa',cursive;
    font-weight:700;
    font-size:150px !important;
    background:linear-gradient(90deg,#1B2A4A 0%,#1565C0 45%,#028090 100%);
    -webkit-background-clip:text;
    -webkit-text-fill-color:transparent;
    background-clip:text;
    margin:0;
    line-height:1.15;
}}
.mf-sub {{ font-size:19px;color:#444;margin:10px 0 0 0;font-weight:600; }}
.mf-meta {{ font-size:13px;color:#999;margin:4px 0 0 0; }}
</style>
<div class="mf-header">
    <p class="mf-title">🔧 MagFlow AI</p>
    <p class="mf-sub">AI-Based Predictive Maintenance System for Electromagnetic Flowmeters</p>
    <p class="mf-meta">PFE ENSA — JESA (OCP × Worley) | Instrumentation &amp; Control · Developed by Maroua Hakkak — 2026</p>
</div>
""", unsafe_allow_html=True)
st.divider()

# New order: Recommendation Engine → Maintenance History → Project Import → Dashboard
mt_reco, mt_hist, mt_import, mt_dash = st.tabs(
    ["🔧 Recommendation Engine", "📋 Maintenance History", "📂 Project Import", "📊 Dashboard"])

# ------------------------------------------------------------
#  MAINTENANCE HISTORY
# ------------------------------------------------------------
with mt_hist:
    st.header("📋 Maintenance History")
    st.markdown("Log and track all maintenance interventions for each flowmeter across JESA/OCP projects.")
    with st.expander("📥 Import from Maintenance Excel (Historical Data sheet)", expanded=False):
        st.markdown("Upload a MagFlow AI maintenance Excel file — all filled rows from the **Historical Data** sheet will be imported automatically.")
        uploaded_excel = st.file_uploader("Upload Excel file", type=["xlsx"], key="excel_import")
        if uploaded_excel:
            imp_tag = st.text_input("Tag Number for this checklist", placeholder="e.g., 204M-FE/FIT-063M", key="imp_tag")
            if st.button("📤 Import to Database", type="primary"):
                with st.spinner("Reading Excel and saving to database..."):
                    count, error = import_from_excel(uploaded_excel)
                    chk_saved = 0
                    if imp_tag.strip():
                        try:
                            uploaded_excel.seek(0)
                        except Exception:
                            pass
                        cells = read_checklist_from_excel(uploaded_excel, imp_tag.strip())
                        chk_saved = save_checklist_records(imp_tag.strip(), cells)
                if error: st.error(f"❌ Import failed: {error}")
                elif count == 0 and chk_saved == 0: st.warning("⚠️ No valid rows found.")
                else:
                    msg = f"✅ {count} intervention(s) imported!"
                    if chk_saved: msg += f" {chk_saved} checklist cell(s) saved."
                    st.success(msg); st.balloons()
    st.divider()
    col_form, col_view = st.columns([1, 1])
    with col_form:
        st.subheader("➕ Log New Intervention")
        with st.form("history_form"):
            h_tag = st.text_input("Tag Number *", placeholder="e.g., 204M-FE/FIT-063M")
            h_date = st.date_input("Date *", value=datetime.now())
            h_type = st.selectbox("Intervention Type *", ["Calibration","Visual Inspection","Liner Inspection","Electrode Check","Grounding Verification","Zero Reset","Component Replacement","Corrective Maintenance","Other"])
            h_task = st.text_area("Task Performed *", placeholder="Describe what was done...", height=100)
            h_result = st.selectbox("Result *", ["Conform","Non-conform","Replaced","Adjusted","Pending"])
            h_tech = st.text_input("Technician Name *", placeholder="Full name")
            h_notes = st.text_area("Notes", placeholder="Additional observations...", height=80)
            submitted = st.form_submit_button("💾 Save Intervention", use_container_width=True, type="primary")
            if submitted:
                if not h_tag or not h_task or not h_tech: st.error("Please fill in Tag Number, Task, and Technician Name.")
                else:
                    success = save_history(str(h_date), h_tag.strip(), h_type, h_task, h_result, h_tech, h_notes)
                    if success: st.success(f"✅ Intervention saved for {h_tag}"); st.balloons()
                    else: st.error("❌ Error saving. Check Google Sheets connection.")
    with col_view:
        st.subheader("🔍 View History by Tag")
        search_tag = st.text_input("Search Tag Number", placeholder="e.g., 204M-FE/FIT-063M")
        if search_tag:
            with st.spinner("Loading history..."):
                records = load_history(tag_filter=search_tag)
            if records:
                st.success(f"Found {len(records)} intervention(s) for **{search_tag}**")
                for r in records:
                    rc = "🟢" if r.get('Result') == "Conform" else "🔴" if r.get('Result') == "Non-conform" else "🟡"
                    with st.expander(f"{rc} {r.get('Date','')} — {r.get('Type','')}"):
                        st.markdown(f"**Task:** {r.get('Task','')}"); st.markdown(f"**Result:** {r.get('Result','')}"); st.markdown(f"**Technician:** {r.get('Technician','')}")
                        if r.get('Notes'): st.markdown(f"**Notes:** {r.get('Notes','')}")
            else: st.info(f"No history found for tag **{search_tag}**")
        st.divider()
        st.subheader("📊 All Interventions by Tag")
        with st.spinner("Loading..."):
            all_records = load_history()
        if all_records:
            st.caption(f"Total interventions recorded: **{len(all_records)}**")
            tags = {}
            for r in all_records:
                tag = r.get('Tag Number', 'Unknown')
                if tag not in tags: tags[tag] = []
                tags[tag].append(r)
            for tag, records in sorted(tags.items()):
                cc = sum(1 for r in records if r.get('Result') == 'Conform')
                nc = sum(1 for r in records if r.get('Result') == 'Non-conform')
                oc = len(records) - cc - nc
                summary = (f"🟢 {cc}" if cc else "") + (f"  🔴 {nc}" if nc else "") + (f"  🟡 {oc}" if oc else "")
                with st.expander(f"📁 **{tag}** — {len(records)} intervention(s)  {summary}"):
                    for r in sorted(records, key=lambda x: x.get('Date',''), reverse=True):
                        rc = "🟢" if r.get('Result') == "Conform" else "🔴" if r.get('Result') == "Non-conform" else "🟡"
                        st.markdown(f"{rc} **{r.get('Date','')}** | {r.get('Type','')} | {r.get('Technician','')}")
                        if r.get('Task'): st.caption(f"↳ {r.get('Task','')[:80]}{'...' if len(r.get('Task','')) > 80 else ''}")
        else: st.info("No interventions recorded yet.")

# ------------------------------------------------------------
#  DASHBOARD  (now the LAST tab)
# ------------------------------------------------------------
with mt_dash:
    st.header("📊 JESA Project Analytics")
    pc = {}
    for p in ai.data.project_data:
        pn = p.get('project','')
        if pn and not any(k in pn for k in ['LEGEND','NOTE','Original','UPDATED']): pc[pn]=pc.get(pn,0)+1
    for pn,c in sorted(pc.items(), key=lambda x:-x[1]):
        st.markdown(f"**{pn}** — {c}"); st.progress(c/max(pc.values()))
    d1,d2 = st.columns(2)
    with d1:
        st.subheader("Electrodes")
        ec = {}
        for p in ai.data.project_data:
            e = p.get('electrode','').strip()
            if e and e!='N/A': ec[e]=ec.get(e,0)+1
        for en,c in sorted(ec.items(),key=lambda x:-x[1])[:10]:
            st.markdown(f"**{en}** — {c}"); st.progress(c/max(ec.values()))
    with d2:
        st.subheader("Liners")
        lc = {}
        for p in ai.data.project_data:
            l = p.get('liner','').strip()
            if l and l!='N/A': lc[l]=lc.get(l,0)+1
        for ln,c in sorted(lc.items(),key=lambda x:-x[1])[:10]:
            st.markdown(f"**{ln}** — {c}"); st.progress(c/max(lc.values()))
    s1,s2,s3,s4,s5,s6 = st.columns(6)
    with s1: st.metric("Fluids", len(ai.data.fluids))
    with s2: st.metric("Electrodes", len(ai.data.electrodes))
    with s3: st.metric("Liners", len(ai.data.liners))
    with s4: st.metric("Models", len(ai.data.vendor_models))
    with s5: st.metric("Drift", len(ai.data.drift_indicators))
    with s6: st.metric("Projects", len(ai.data.project_data))

# ------------------------------------------------------------
#  RECOMMENDATION ENGINE
# ------------------------------------------------------------
with mt_reco:
    tab1,tab2,tab3,tab4 = st.tabs(["🧪 Process Fluid","📐 Pipe Dimensions","🌡️ Operating Conditions","📝 Special Notes"])
    with tab1:
        c1,c2 = st.columns(2)
        with c1:
            fluid = st.selectbox("Service Fluid", ai.get_fluid_names())
            tag_number = st.text_input("🏷️ Tag Number", placeholder="e.g., 204M-FE/FIT-063M")
        with c2:
            cat = ai.get_fluid_category(fluid)
            colors = {'Corrosive':'🔴','Abrasive':'🟠','Charged':'🟡','Clean':'🟢','Unknown':'⚪'}
            st.markdown(f"<p style='font-size:17px;font-weight:600;color:#1B2A4A;margin-bottom:4px'>Fluid Category: {colors.get(cat,'')} {cat}</p>", unsafe_allow_html=True)
            conductivity = st.number_input("Conductivity (µS/cm)", value=5000.0, min_value=0.0, step=100.0)
    with tab2:
        c1,c2,c3 = st.columns(3)
        with c1: dn = st.number_input("DN (mm)", value=80, min_value=3, max_value=3000, step=10)
        with c2: pipe_mat = st.selectbox("Pipe Material", ["FRP","Stainless Steel","Carbon Steel","ZeCor","PVC","HDPE"])
        with c3: pipe_liner = st.selectbox("Pipe Liner", ["Not applicable","HDPE","Rubber lined","Epoxy coated","Glass lined"])
        pipe_thick = st.number_input("Pipe Thickness (mm)", value=0.0, min_value=0.0, step=0.5)
    with tab3:
        st.subheader("Flow")
        fc1,fc2 = st.columns(2)
        with fc1: flow_normal = st.number_input("Normal Flow (m³/h)", value=23.0, min_value=0.0, step=1.0)
        with fc2: flow_max = st.number_input("Max Flow (m³/h)", value=26.0, min_value=0.0, step=1.0)
        st.subheader("Temperature (°C)")
        tc1,tc2,tc3,tc4 = st.columns(4)
        with tc1: t_min = st.number_input("Min Temp", value=0.0, step=1.0)
        with tc2: t_norm = st.number_input("Normal Temp", value=82.0, step=1.0)
        with tc3: t_max = st.number_input("Max Temp", value=95.0, step=1.0)
        with tc4: t_des = st.number_input("Design Temp", value=105.0, step=1.0)
        st.subheader("Pressure (bar)")
        pc1,pc2,pc3,pc4,pc5 = st.columns(5)
        with pc1: p_min = st.number_input("Min P", value=0.0, step=0.5)
        with pc2: p_norm = st.number_input("Normal P", value=5.0, step=0.5)
        with pc3: p_max = st.number_input("Max P", value=8.0, step=0.5)
        with pc4: p_des = st.number_input("Design P", value=10.0, step=0.5)
        with pc5: p_drop = st.number_input("ΔP Allow.", value=0.0, step=0.1)
        st.subheader("Fluid Properties")
        fp1,fp2 = st.columns(2)
        with fp1: visc = st.number_input("Viscosity (cP)", value=16.0, min_value=0.0)
        with fp2: dens = st.number_input("Density (kg/m³)", value=1730.0, min_value=0.0)
    with tab4:
        special = st.text_area("Special Conditions", placeholder="ATEX, SIL 2...", height=80)
        notes = st.text_area("User Notes", placeholder="Prefers Emerson...", height=80)
    st.divider()

    # ---- Run: compute and STORE results in session_state (so downloads/lang don't wipe them) ----
    if st.button("🚀 Run Recommendation", type="primary", use_container_width=True):
        inp = ProcessInput(fluid_name=fluid, pipe_material=pipe_mat, pipe_thickness=pipe_thick,
            pipe_liner=pipe_liner, tube_material='SS 316L', dn=dn, flow_normal=flow_normal, flow_max=flow_max,
            temp_min=t_min, temp_normal=t_norm, temp_max=t_max, temp_design=t_des,
            pressure_min=p_min, pressure_normal=p_norm, pressure_max=p_max, pressure_design=p_des,
            pressure_drop=p_drop, conductivity=conductivity, viscosity=visc, density=dens,
            special_conditions=special, user_notes=notes)
        result = ai.recommend(inp)
        st.session_state["reco_result"] = result
        st.session_state["reco_meta"] = {
            "fluid": fluid, "cat": cat,
            "tag": tag_number if tag_number else "NEW-INSTRUMENT",
        }

    # ---- Render results from session_state (survives any re-run) ----
    if "reco_result" in st.session_state:
        result = st.session_state["reco_result"]
        meta = st.session_state["reco_meta"]
        fluid = meta["fluid"]; cat = meta["cat"]
        v,m,vendors,tco = result['validation'],result['materials'],result['vendors'],result['tco']

        with st.container(border=True):
            st.markdown("<div style='background-color:#E8F5E9;padding:15px;border-radius:10px;border-left:5px solid #2E7D32'><h2 style='color:#2E7D32;margin:0'>Layer 1 — Input Validation</h2></div>", unsafe_allow_html=True)
            st.write("")
            if v.is_valid:
                st.success(f"✅ Magflow suitable — Velocity: {v.velocity} m/s — Category: {v.fluid_category}")
                if v.recommended_dn > 0: st.caption(f"💡 DN {v.recommended_dn} mm recommended")
            else:
                st.error("❌ Magflow NOT suitable")
                for e in v.errors: st.error(e)
            for w in v.warnings: st.warning(w)

        if not v.is_valid:
            st.stop()

        with st.container(border=True):
            st.markdown("<div style='background-color:#E3F2FD;padding:15px;border-radius:10px;border-left:5px solid #1565C0'><h2 style='color:#1565C0;margin:0'>Layer 2 — Material Selection</h2></div>", unsafe_allow_html=True)
            st.write("")
            c1,c2,c3,c4 = st.columns(4)
            with c1:
                ab = " ⚠️ SO" if 'special' in m.electrode_avail.lower() else ""
                st.markdown(f"<p style='font-size:18px;font-weight:bold;margin-bottom:2px'>Electrode</p><p style='font-size:15px;color:#333'>{m.electrode}{ab}</p><p style='font-size:12px;color:#2E7D32'>Cost: {m.electrode_cost}</p>", unsafe_allow_html=True)
                if m.electrode_alt: st.caption(f"Alt: {', '.join(m.electrode_alt)}")
            with c2:
                ab2 = " ⚠️ SO" if 'special' in m.liner_avail.lower() else ""
                st.markdown(f"<p style='font-size:18px;font-weight:bold;margin-bottom:2px'>Liner</p><p style='font-size:15px;color:#333'>{m.liner}{ab2}</p><p style='font-size:12px;color:#2E7D32'>Cost: {m.liner_cost}</p>", unsafe_allow_html=True)
                if m.liner_alt: st.caption(f"Alt: {', '.join(m.liner_alt)}")
            with c3: st.markdown(f"<p style='font-size:18px;font-weight:bold;margin-bottom:2px'>Body Material</p><p style='font-size:15px;color:#333'>SS 304L</p><p style='font-size:12px;color:#666'>Coil housing</p>", unsafe_allow_html=True)
            with c4: st.markdown(f"<p style='font-size:18px;font-weight:bold;margin-bottom:2px'>Grounding</p><p style='font-size:15px;color:#333'>{m.grounding}</p>", unsafe_allow_html=True)
            c5,c6,c7 = st.columns(3)
            with c5: st.markdown(f"<p style='font-size:18px;font-weight:bold;margin-bottom:2px'>Penetrant Ring</p><p style='font-size:15px;color:#333'>{m.penetrant or 'N/A'}</p>", unsafe_allow_html=True)
            with c6: st.markdown(f"<p style='font-size:18px;font-weight:bold;margin-bottom:2px'>O-Ring</p><p style='font-size:15px;color:#333'>{m.o_ring}</p>", unsafe_allow_html=True)
            with c7: st.markdown(f"<p style='font-size:18px;font-weight:bold;margin-bottom:2px'>Flange Coating</p><p style='font-size:15px;color:#333'>{m.flange_coat}</p>", unsafe_allow_html=True)
            if m.liner_dn_warning: st.warning(f"📐 {m.liner_dn_warning}")
            if m.tube_warning: st.info(f"🔧 {m.tube_warning}")
            for w in m.warnings: st.warning(w)
            if m.remarks: st.info(f"📝 {m.remarks}")

        with st.container(border=True):
            st.markdown("<div style='background-color:#FFF3E0;padding:15px;border-radius:10px;border-left:5px solid #E65100'><h2 style='color:#E65100;margin:0'>Layer 3 — Vendor Recommendations</h2></div>", unsafe_allow_html=True)
            st.write("")
            st.caption(f"Compatible: {len(vendors['all'])} / {len(ai.data.vendor_models)} models")
            vc1,vc2,vc3 = st.columns(3)
            for cw,label,key,vname in [(vc1,"🔴 EMERSON","emerson","Emerson"),(vc2,"🔵 ENDRESS+HAUSER","eh","Endress+Hauser"),(vc3,"🟢 KROHNE","krohne","Krohne")]:
                with cw:
                    url = vendor_url(vname)
                    head = f"<a href='{url}' target='_blank' style='text-decoration:none;color:#1565C0'>{label} 🔗</a>" if url else label
                    st.markdown(f"<p style='font-size:18px;font-weight:bold'>{head}</p>", unsafe_allow_html=True)
                    rec = vendors[key]
                    if rec:
                        so = " ⚠️ SO" if rec.special_order else ""
                        st.markdown(f"<p style='font-size:16px;font-weight:bold;color:#333'>{rec.model}</p><p style='font-size:12px;color:#666'>{rec.model_type}{so}</p>", unsafe_allow_html=True)
                        st.markdown(f"**Accuracy:** {rec.accuracy}"); st.markdown(f"**Pressure:** {rec.pressure}")
                        st.markdown(f"**DN:** {rec.pipe_sizes}"); st.markdown(f"**Excitation:** {rec.excitation}")
                        st.markdown(f"**IP:** {rec.ip}"); st.markdown(f"**Protocols:** {rec.protocols}"); st.markdown(f"**Diagnostics:** {rec.diagnostics}")
                    else: st.warning("No compatible model")
            with st.expander("🌍 Other Vendors"):
                for ev in vendors.get('extra',[]):
                    url = vendor_url(ev['vendor'])
                    name = f"<a href='{url}' target='_blank'>{ev['vendor']}</a>" if url else ev['vendor']
                    st.markdown(f"**{name}** — {ev['model']} ({ev['accuracy']}) — {ev['notes']}", unsafe_allow_html=True)
            for w in vendors['warnings']: st.warning(w)

        with st.container(border=True):
            st.markdown("<div style='background-color:#F3E5F5;padding:15px;border-radius:10px;border-left:5px solid #6A1B9A'><h2 style='color:#6A1B9A;margin:0'>Layer 4 — TCO & Drift Prediction</h2></div>", unsafe_allow_html=True)
            st.write("")
            # Labels bold + larger than values; values not bold + smaller
            def metric_block(label, value):
                return (f"<p style='font-size:20px;font-weight:800;color:#1B2A4A;margin:0 0 2px 0'>{label}</p>"
                        f"<p style='font-size:15px;font-weight:400;color:#444;margin:0'>{value}</p>")
            t1,t2,t3,t4 = st.columns(4)
            with t1: st.markdown(metric_block("CAPEX Score", f"{tco.capex_score}/30"), unsafe_allow_html=True)
            with t2: st.markdown(metric_block("Calibration", f"Every {tco.calib_months} mo"), unsafe_allow_html=True)
            with t3: st.markdown(metric_block("Liner Life", tco.liner_life), unsafe_allow_html=True)
            with t4: st.markdown(metric_block("Electrode Life", tco.electrode_life), unsafe_allow_html=True)
            if tco.pressure_check: st.info(f"🔒 {tco.pressure_check}")
            st.markdown("<p style='font-size:18px;font-weight:bold'>CAPEX Breakdown</p>", unsafe_allow_html=True)
            bc = st.columns(len(tco.breakdown))
            for i,(k,vs) in enumerate(tco.breakdown.items()):
                with bc[i]: st.markdown(f"**{k}:** {vs}/6 {'🟩'*vs}{'⬜'*(6-vs)}")
            st.markdown("<p style='font-size:18px;font-weight:bold'>Drift Risk Assessment</p>", unsafe_allow_html=True)
            for risk in tco.drift_risks:
                ic = "🔴" if risk.level=="High" else "🟡" if risk.level=="Medium" else "🟢"
                with st.expander(f"{ic} {risk.indicator} — {risk.level}"):
                    st.markdown(f"**Description:** {risk.description}")
                    st.markdown("**Diagnostic Steps:**")
                    for step in risk.steps: st.markdown(f"  {step}")
                    st.markdown("**Recommended Tools:**")
                    for vn,tool in risk.tools.items(): st.markdown(f"  • **{vn}:** {tool}")
                    st.markdown(f"**Frequency:** {risk.frequency}")
            st.markdown("<p style='font-size:18px;font-weight:bold'>📋 Maintenance Plan</p>", unsafe_allow_html=True)
            mnt = tco.maintenance
            for title,items,icon,bg,tc2 in [
                ("Continuous Monitoring",mnt.continuous,"🔄","#E8F5E9","#2E7D32"),
                ("Monthly Checks",mnt.monthly,"📅","#E3F2FD","#1565C0"),
                ("Quarterly Inspection",mnt.quarterly,"🔍","#FFF3E0","#E65100"),
                ("Semi-Annual",mnt.semi_annual,"⚙️","#FCE4D6","#BF360C"),
                ("Annual Maintenance",mnt.annual,"🛠️","#F3E5F5","#6A1B9A"),
                ("Multi-Year (3-5yr)",mnt.multi_year,"📊","#FFF9C4","#F57F17"),
                ("Component Lifespan",mnt.replacement,"♻️","#FFCDD2","#C62828")]:
                if items:
                    st.markdown(f"<div style='background:{bg};padding:12px 15px;border-radius:8px;border-left:4px solid {tc2};margin:8px 0'><p style='font-size:16px;font-weight:bold;color:{tc2};margin:0'>{icon} {title}</p></div>", unsafe_allow_html=True)
                    for item in items: st.markdown(f"&nbsp;&nbsp;&nbsp;&nbsp;• {item}")
            st.markdown("---")

            tag = meta["tag"]

            # ---- Download Excel maintenance sheet ----
            history_for_export = load_history(tag_filter=tag) if tag != "NEW-INSTRUMENT" else []
            checklist_saved = load_checklist_records(tag) if tag != "NEW-INSTRUMENT" else {}
            excel_buf = generate_maintenance_excel(tag, fluid, cat, m, vendors['emerson'], vendors['eh'], vendors['krohne'], tco, history_for_export, checklist_saved)
            st.download_button(label="📄 Download Maintenance Excel", data=excel_buf,
                file_name=f"MagFlow_Maintenance_{tag.replace('/','_')}_{datetime.now().strftime('%Y%m%d')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True)
            st.caption("💡 The Excel sheet has 3 tabs: Guideline, Maintenance Checklist (month-by-month, "
                       "with a Work done column for monthly & quarterly tasks), and Historical Data.")

            st.markdown("<p style='font-size:18px;font-weight:bold'>🔍 Validation vs JESA Projects</p>", unsafe_allow_html=True)
            if tco.validation:
                for match in tco.validation:
                    ic = "✅" if "MATCH" in match.match_type else ("⚠️" if "PARTIAL" in match.match_type else "❌")
                    with st.expander(f"{ic} {match.match_type} — {match.project} ({match.fluid}) — {match.confidence}%"):
                        for k,vd in match.details.items(): st.markdown(f"**{k}:** {vd}")
                        if match.explanation: st.markdown(f"**Why:** {match.explanation}")
                        if match.cost_impact: st.markdown(f"**Cost:** {match.cost_impact}")
                total = len(tco.validation); matches = sum(1 for mt in tco.validation if mt.confidence>=80)
                st.caption(f"{matches}/{total} matches. {'Validated ✓' if matches>total/2 else 'Review ⚠'}")
            else: st.info("No matching project.")
            if tco.notes_response:
                st.markdown("<p style='font-size:18px;font-weight:bold'>💡 AI Response</p>", unsafe_allow_html=True)
                st.markdown(tco.notes_response)
        st.divider()
        st.caption("Source: JESA Internal DB, JESA Flow App 2024, Vendor datasheets")

# ------------------------------------------------------------
#  PROJECT IMPORT
# ------------------------------------------------------------
with mt_import:
    st.header("📂 Project Import — Datasheet Analyzer")
    st.markdown("Upload a JESA flowmeter datasheet PDF. The AI extracts all instrument data, **detects new materials**, and updates the database automatically.")
    col_upload, col_db = st.columns([1, 1])
    with col_upload:
        st.subheader("📤 Upload Datasheet PDF")
        project_name_override = st.text_input("Project Name (optional)", placeholder="e.g. Central Axis Program / SAFI")
        uploaded_pdf = st.file_uploader("Upload JESA Datasheet PDF", type=["pdf"], key="pdf_import")
        if uploaded_pdf:
            st.success(f"✅ **{uploaded_pdf.name}** ready")
            if st.button("🤖 Analyze & Detect New Materials", type="primary", use_container_width=True):
                with st.spinner("Step 1/3 — AI reading datasheet..."):
                    pdf_bytes = uploaded_pdf.read()
                    instruments, error = extract_datasheet_with_ai(pdf_bytes)
                if error or not instruments:
                    st.error(f"❌ Extraction failed: {error or 'No instruments found'}")
                else:
                    with st.spinner("Step 2/3 — Loading current database from GitHub..."):
                        wb, sha, err = get_excel_from_github()
                    if err:
                        st.error(f"❌ Cannot load database: {err}")
                    else:
                        with st.spinner("Step 3/3 — Detecting new materials..."):
                            existing = get_existing_materials(wb)
                            new_findings = detect_new_materials(instruments, existing)
                        st.session_state['import_instruments'] = instruments
                        st.session_state['import_wb'] = wb
                        st.session_state['import_sha'] = sha
                        st.session_state['import_new'] = new_findings
                        st.session_state['import_project'] = project_name_override or (instruments[0].get('project','') if instruments else '')
                        st.success(f"✅ Found **{len(instruments)}** instrument(s) — **{len(new_findings)}** new material(s) detected")
        if 'import_instruments' in st.session_state:
            instruments = st.session_state['import_instruments']
            new_findings = st.session_state['import_new']
            pname = st.session_state['import_project']
            st.divider()
            st.subheader("📋 Extracted Instruments")
            for inst in instruments:
                with st.expander(f"🔧 {inst.get('tag','N/A')} — {inst.get('fluid','N/A')}"):
                    c1, c2 = st.columns(2)
                    with c1:
                        st.markdown(f"**Project:** {inst.get('project', pname)}"); st.markdown(f"**Fluid:** {inst.get('fluid','')}")
                        st.markdown(f"**DN:** {inst.get('dn','')} mm"); st.markdown(f"**Flow:** {inst.get('flow_normal','')} → {inst.get('flow_max','')} m³/h")
                        st.markdown(f"**Temp design:** {inst.get('temp_design','')} °C"); st.markdown(f"**Pressure design:** {inst.get('pressure_design','')} bar-g")
                    with c2:
                        st.markdown(f"**Electrode:** {inst.get('electrode','')}"); st.markdown(f"**Liner:** {inst.get('liner','')}")
                        st.markdown(f"**Tube:** {inst.get('tube','')}"); st.markdown(f"**Grounding:** {inst.get('grounding','')}")
                        st.markdown(f"**Accuracy:** {inst.get('accuracy','')}"); st.markdown(f"**Vendor/Model:** {inst.get('vendor','')} / {inst.get('model','')}")
            if new_findings:
                st.divider()
                st.subheader("🆕 New Materials Detected")
                st.warning(f"**{len(new_findings)} new item(s)** not in current database:")
                approved = []
                for i, finding in enumerate(new_findings):
                    icons = {"electrode": "⚡", "liner": "🟡", "fluid": "💧", "fluid_electrode_variant": "🔀"}
                    labels = {"electrode": "New Electrode", "liner": "New Liner", "fluid": "New Fluid", "fluid_electrode_variant": "New Electrode Variant"}
                    checked = st.checkbox(f"{icons.get(finding['type'],'🆕')} **{labels.get(finding['type'],'New')}:** `{finding['value']}` — {finding['context']}", value=True, key=f"approve_{i}")
                    if checked: approved.append(finding)
                st.session_state['import_approved'] = approved
            else:
                st.info("✅ No new materials detected.")
                st.session_state['import_approved'] = []
            st.divider()
            col_save1, col_save2 = st.columns(2)
            with col_save1:
                if st.button("💾 Save to Project Database", use_container_width=True):
                    with st.spinner("Saving to Google Sheets..."):
                        saved, err = save_project_to_gsheet(instruments, pname)
                    if err: st.error(f"❌ {err}")
                    else: st.success(f"✅ {saved} instrument(s) saved!")
            with col_save2:
                approved_updates = st.session_state.get('import_approved', [])
                if approved_updates:
                    st.info(f"💡 **{len(approved_updates)} new material(s) detected**\n\nThese materials have been flagged for review. To integrate them into the database, an I&C engineer must complete their full specification (liner, o-ring, grounding, conductivity limits) in Data_Collection_v2.xlsx before adding them.")
                    for finding in approved_updates:
                        icons = {"electrode": "⚡", "liner": "🟡", "fluid": "💧", "fluid_electrode_variant": "🔀"}
                        st.markdown(f"{icons.get(finding['type'],'🆕')} **{finding['value']}** — {finding['context']}")
                else:
                    st.success("✅ All materials already in database.")
    with col_db:
        st.subheader("📊 Imported Projects Database")
        with st.spinner("Loading..."):
            project_records = load_project_data_from_gsheet()
        if project_records:
            st.caption(f"Total instruments imported: **{len(project_records)}**")
            projects = list(set(r.get('Project','') for r in project_records if r.get('Project','')))
            for proj in sorted(projects):
                proj_insts = [r for r in project_records if r.get('Project','') == proj]
                with st.expander(f"📁 {proj} — {len(proj_insts)} instrument(s)"):
                    for r in proj_insts:
                        st.markdown(f"**{r.get('Tag','')}** | {r.get('Fluid','')} | DN{r.get('DN','')} | ⚡{r.get('Electrode','')} / 🟡{r.get('Liner','')} | _{r.get('Import Date','')}_")
        else: st.info("No projects imported yet. Upload a datasheet to get started.")
