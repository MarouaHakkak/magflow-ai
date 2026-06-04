"""MagFlow AI v8 — Project Import + Auto DB Update via GitHub"""
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

st.set_page_config(page_title="MagFlow AI — JESA", page_icon="🔧", layout="wide")

@st.cache_resource
def load_model():
    return MagFlowAI('Data_Collection_v2.xlsx')

ai = load_model()

# ===== GOOGLE SHEETS CONNECTION =====
SHEET_ID = "1jVpJkHPtG808WtlKxKpcIxvNLvLyx8syIi0GlppNTgU"
SCOPES = ["https://www.googleapis.com/auth/spreadsheets"]

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

def import_from_excel(uploaded_file):
    """Read Historical Data sheet from uploaded Excel and save rows to Google Sheets."""
    try:
        wb = openpyxl.load_workbook(uploaded_file)
        if "Historical Data" not in wb.sheetnames:
            return 0, "Sheet 'Historical Data' not found in the uploaded file."
        ws = wb["Historical Data"]
        saved = 0
        skipped = 0
        for row in ws.iter_rows(min_row=4, values_only=True):
            date, tag, itype, task, result, tech, notes = (row[0], row[1], row[2], row[3], row[4], row[5], row[6] if len(row) > 6 else "")
            if not date or not tag or not task:
                skipped += 1
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

def generate_maintenance_excel(tag, fluid, cat, mat, vendor_e, vendor_eh, vendor_k, tco, history_records=None):
    wb = openpyxl.Workbook()
    vn = vendor_e.model if vendor_e else (vendor_eh.model if vendor_eh else (vendor_k.model if vendor_k else 'N/A'))
    vd = vendor_e.diagnostics if vendor_e else (vendor_eh.diagnostics if vendor_eh else (vendor_k.diagnostics if vendor_k else 'N/A'))
    today = datetime.now().strftime('%Y-%m-%d')
    year = datetime.now().year
    mnt = tco.maintenance

    # Styles
    title_font = Font(bold=True, size=14, color="1B2A4A")
    h_font = Font(bold=True, size=11, color="FFFFFF")
    s_font = Font(bold=True, size=11, color="1B2A4A")
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

    # ===== SHEET 1: GUIDELINE =====
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
        ws1[f'B{row}'] = v; ws1[f'B{row}'].border = thin
        row += 1

    row += 1
    ws1.merge_cells(f'A{row}:B{row}'); ws1[f'A{row}'] = "DRIFT RISK ASSESSMENT"; ws1[f'A{row}'].font = h_font; ws1[f'A{row}'].fill = h_fill
    row += 1
    for risk in tco.drift_risks:
        icon = "🔴" if risk.level == "High" else "🟡" if risk.level == "Medium" else "🟢"
        fill = r_fill if risk.level == "High" else y_fill if risk.level == "Medium" else g_fill
        ws1.merge_cells(f'A{row}:B{row}')
        ws1[f'A{row}'] = f"{icon} {risk.indicator} — {risk.level}"; ws1[f'A{row}'].font = Font(bold=True, size=11); ws1[f'A{row}'].fill = fill
        row += 1
        ws1[f'A{row}'] = "Description:"; ws1[f'A{row}'].font = Font(bold=True)
        ws1[f'B{row}'] = risk.description; ws1[f'B{row}'].alignment = wrap
        row += 1
        ws1[f'A{row}'] = "Diagnostic Steps:"; ws1[f'A{row}'].font = Font(bold=True)
        steps_text = "\n".join(risk.steps)
        ws1[f'B{row}'] = steps_text; ws1[f'B{row}'].alignment = wrap
        ws1.row_dimensions[row].height = max(15 * len(risk.steps), 30)
        row += 1
        ws1[f'A{row}'] = "Frequency:"; ws1[f'A{row}'].font = Font(bold=True)
        ws1[f'B{row}'] = risk.frequency
        row += 2

    # ===== SHEET 2: MAINTENANCE CHECKLIST =====
    ws2 = wb.create_sheet("Maintenance Checklist")
    ws2.column_dimensions['A'].width = 5; ws2.column_dimensions['B'].width = 45
    for col in range(3, 10): ws2.column_dimensions[get_column_letter(col)].width = 10
    ws2.column_dimensions['J'].width = 15; ws2.column_dimensions['K'].width = 20

    ws2.merge_cells('A1:K1'); ws2['A1'] = f"🔧 Maintenance Checklist — {tag}"; ws2['A1'].font = title_font
    ws2['A2'] = f"Fluid: {fluid} | Category: {cat} | Year: {year}"; ws2['A2'].font = Font(italic=True, color="666666")

    row = 4
    months_data = [
        ("January", 4), ("February", 4), ("March", 4),
        ("April", 4), ("May", 4), ("June", 4),
        ("July", 4), ("August", 4), ("September", 4),
        ("October", 4), ("November", 4), ("December", 4)
    ]
    quarter_months = {3: "Q1", 6: "Q2", 9: "Q3", 12: "Q4"}
    semester_months = {6: "S1", 12: "S2"}

    for m_idx, (month_name, weeks) in enumerate(months_data, 1):
        ws2.merge_cells(f'A{row}:K{row}')
        ws2[f'A{row}'] = f"📅 {month_name} {year}"; ws2[f'A{row}'].font = Font(bold=True, size=12, color="1565C0"); ws2[f'A{row}'].fill = b_fill
        row += 1

        if mnt.continuous:
            for w in range(1, weeks + 1):
                ws2.merge_cells(f'A{row}:B{row}')
                ws2[f'A{row}'] = f"   Week {w}"; ws2[f'A{row}'].font = Font(bold=True, size=10, color="2E7D32")
                days = ["Mon", "Tue", "Wed", "Thu", "Fri", "Sat", "Sun"]
                for d, day in enumerate(days, 3):
                    c = ws2.cell(row=row, column=d, value=day)
                    c.font = Font(bold=True, size=9); c.alignment = center; c.fill = g_fill; c.border = thin
                ws2.cell(row=row, column=10, value="Status").font = Font(bold=True, size=9); ws2.cell(row=row, column=10).fill = g_fill; ws2.cell(row=row, column=10).border = thin
                ws2.cell(row=row, column=11, value="Technician").font = Font(bold=True, size=9); ws2.cell(row=row, column=11).fill = g_fill; ws2.cell(row=row, column=11).border = thin
                row += 1
                for item in mnt.continuous:
                    ws2.cell(row=row, column=1, value="☐").alignment = center
                    ws2.cell(row=row, column=2, value=item).font = n_font; ws2.cell(row=row, column=2).border = thin
                    for d in range(3, 12):
                        ws2.cell(row=row, column=d).border = thin; ws2.cell(row=row, column=d).alignment = center
                    row += 1
                row += 1

        if mnt.monthly:
            ws2.merge_cells(f'A{row}:B{row}')
            ws2[f'A{row}'] = f"   📋 Monthly tasks — {month_name}"; ws2[f'A{row}'].font = Font(bold=True, size=10, color="1565C0"); ws2[f'A{row}'].fill = PatternFill(start_color="BBDEFB", end_color="BBDEFB", fill_type="solid")
            ws2.cell(row=row, column=10, value="Date").font = Font(bold=True, size=9); ws2.cell(row=row, column=10).border = thin
            ws2.cell(row=row, column=11, value="Technician").font = Font(bold=True, size=9); ws2.cell(row=row, column=11).border = thin
            row += 1
            for item in mnt.monthly:
                ws2.cell(row=row, column=1, value="☐").alignment = center
                ws2.cell(row=row, column=2, value=item).font = n_font; ws2.cell(row=row, column=2).border = thin
                ws2.cell(row=row, column=10).border = thin; ws2.cell(row=row, column=11).border = thin
                row += 1
            row += 1

        if m_idx in quarter_months and mnt.quarterly:
            q_name = quarter_months[m_idx]
            ws2.merge_cells(f'A{row}:B{row}')
            ws2[f'A{row}'] = f"   🔍 Quarterly tasks — {q_name}"; ws2[f'A{row}'].font = Font(bold=True, size=10, color="E65100"); ws2[f'A{row}'].fill = o_fill
            ws2.cell(row=row, column=10, value="Date").font = Font(bold=True, size=9); ws2.cell(row=row, column=10).border = thin
            ws2.cell(row=row, column=11, value="Technician").font = Font(bold=True, size=9); ws2.cell(row=row, column=11).border = thin
            row += 1
            for item in mnt.quarterly:
                ws2.cell(row=row, column=1, value="☐").alignment = center
                ws2.cell(row=row, column=2, value=item).font = n_font; ws2.cell(row=row, column=2).border = thin
                ws2.cell(row=row, column=10).border = thin; ws2.cell(row=row, column=11).border = thin
                row += 1
            row += 1

        if m_idx in semester_months and mnt.semi_annual:
            s_name = semester_months[m_idx]
            ws2.merge_cells(f'A{row}:B{row}')
            ws2[f'A{row}'] = f"   ⚙️ Semi-annual tasks — {s_name}"; ws2[f'A{row}'].font = Font(bold=True, size=10, color="BF360C"); ws2[f'A{row}'].fill = q_fill
            ws2.cell(row=row, column=10, value="Date").font = Font(bold=True, size=9); ws2.cell(row=row, column=10).border = thin
            ws2.cell(row=row, column=11, value="Technician").font = Font(bold=True, size=9); ws2.cell(row=row, column=11).border = thin
            row += 1
            for item in mnt.semi_annual:
                ws2.cell(row=row, column=1, value="☐").alignment = center
                ws2.cell(row=row, column=2, value=item).font = n_font; ws2.cell(row=row, column=2).border = thin
                ws2.cell(row=row, column=10).border = thin; ws2.cell(row=row, column=11).border = thin
                row += 1
            row += 1

    if mnt.annual:
        ws2.merge_cells(f'A{row}:K{row}')
        ws2[f'A{row}'] = f"🛠️ ANNUAL MAINTENANCE — {year}"; ws2[f'A{row}'].font = Font(bold=True, size=12, color="6A1B9A"); ws2[f'A{row}'].fill = p_fill
        row += 1
        for item in mnt.annual:
            ws2.cell(row=row, column=1, value="☐").alignment = center
            ws2.cell(row=row, column=2, value=item).font = n_font; ws2.cell(row=row, column=2).border = thin
            ws2.cell(row=row, column=10, value="").border = thin
            ws2.cell(row=row, column=11, value="").border = thin
            row += 1
        row += 1

    if mnt.multi_year:
        ws2.merge_cells(f'A{row}:K{row}')
        ws2[f'A{row}'] = f"📊 MULTI-YEAR MAINTENANCE (3-5 years)"; ws2[f'A{row}'].font = Font(bold=True, size=12, color="F57F17"); ws2[f'A{row}'].fill = y_fill
        row += 1
        for item in mnt.multi_year:
            ws2.cell(row=row, column=1, value="☐").alignment = center
            ws2.cell(row=row, column=2, value=item).font = n_font; ws2.cell(row=row, column=2).border = thin
            ws2.cell(row=row, column=10, value="").border = thin
            ws2.cell(row=row, column=11, value="").border = thin
            row += 1
        row += 1

    if mnt.replacement:
        ws2.merge_cells(f'A{row}:K{row}')
        ws2[f'A{row}'] = "♻️ COMPONENT REPLACEMENT (based on lifespan)"; ws2[f'A{row}'].font = Font(bold=True, size=12, color="C62828"); ws2[f'A{row}'].fill = r_fill
        row += 1
        for item in mnt.replacement:
            ws2.cell(row=row, column=1, value="☐").alignment = center
            ws2.cell(row=row, column=2, value=item).font = n_font; ws2.cell(row=row, column=2).border = thin
            ws2.cell(row=row, column=10, value="").border = thin
            ws2.cell(row=row, column=11, value="").border = thin
            row += 1

    # ===== SHEET 3: HISTORICAL DATA (pre-filled from Google Sheets) =====
    ws3 = wb.create_sheet("Historical Data")
    ws3.column_dimensions['A'].width = 15
    ws3.column_dimensions['B'].width = 20
    ws3.column_dimensions['C'].width = 20
    ws3.column_dimensions['D'].width = 35
    ws3.column_dimensions['E'].width = 15
    ws3.column_dimensions['F'].width = 20
    ws3.column_dimensions['G'].width = 30

    ws3.merge_cells('A1:G1')
    ws3['A1'] = f"📋 Maintenance History — {tag}"
    ws3['A1'].font = title_font

    headers = ["Date", "Tag Number", "Type", "Task Performed", "Result", "Technician", "Notes"]
    for col, h in enumerate(headers, 1):
        c = ws3.cell(row=3, column=col, value=h)
        c.font = h_font; c.fill = h_fill; c.border = thin; c.alignment = center

    if history_records:
        for r_idx, record in enumerate(history_records, 4):
            row_data = [
                record.get('Date', ''),
                record.get('Tag Number', ''),
                record.get('Type', ''),
                record.get('Task', ''),
                record.get('Result', ''),
                record.get('Technician', ''),
                record.get('Notes', '')
            ]
            for col, val in enumerate(row_data, 1):
                c = ws3.cell(row=r_idx, column=col, value=val)
                c.border = thin
                c.fill = b_fill if r_idx % 2 == 0 else PatternFill(fill_type=None)
    else:
        ws3.merge_cells('A4:G4')
        ws3['A4'] = "No interventions recorded yet. Use the 'Maintenance History' tab in MagFlow AI to log interventions."
        ws3['A4'].font = Font(italic=True, color="888888")

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf

# ===== SIDEBAR CHATBOT =====
with st.sidebar:
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
        if(transcript)window.location.href='?voice='+encodeURIComponent(transcript);
        document.getElementById('voiceBtn').innerText='🎤 Tap to speak';
        document.getElementById('sendBtn').style.display='none';
        transcript='';
    }}
    </script>"""
    st.components.v1.html(voice_html, height=80)
    vq = st.query_params.get("voice", None)
    if vq: st.query_params.clear(); st.session_state.voice_input = vq
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
        st.session_state.messages = [{"role": "assistant", "content": "👋 Cleared."}]; st.rerun()

# ===== MAIN =====
st.markdown("""<div style='text-align:center;padding:10px 0'><h1 style='color:#1B2A4A;margin:0'>🔧 MagFlow AI</h1><p style='font-size:16px;color:#555'>AI-Based Predictive Maintenance System for Electromagnetic Flowmeters</p><p style='font-size:12px;color:#999'>PFE ENSA — JESA (OCP × Worley) | Instrumentation & Control</p><p style='font-size:11px;color:#bbb'>Developed by Maroua Hakkak — 2026</p></div>""", unsafe_allow_html=True)
st.divider()

mt1, mt2, mt3, mt4 = st.tabs(["🔧 Recommendation Engine", "📊 Dashboard", "📋 Maintenance History", "📂 Project Import"])

# ===== TAB 3: MAINTENANCE HISTORY =====
with mt3:
    st.header("📋 Maintenance History")
    st.markdown("Log and track all maintenance interventions for each flowmeter across JESA/OCP projects.")

    # ===== EXCEL IMPORT =====
    with st.expander("📥 Import from Maintenance Excel (Historical Data sheet)", expanded=False):
        st.markdown("Upload a MagFlow AI maintenance Excel file — all filled rows from the **Historical Data** sheet will be imported automatically.")
        uploaded_excel = st.file_uploader("Upload Excel file", type=["xlsx"], key="excel_import")
        if uploaded_excel:
            if st.button("📤 Import to Database", type="primary"):
                with st.spinner("Reading Excel and saving to database..."):
                    count, error = import_from_excel(uploaded_excel)
                if error:
                    st.error(f"❌ Import failed: {error}")
                elif count == 0:
                    st.warning("⚠️ No valid rows found. Make sure rows start at line 4 and have Date, Tag and Task filled.")
                else:
                    st.success(f"✅ {count} intervention(s) imported successfully!")
                    st.balloons()

    st.divider()
    col_form, col_view = st.columns([1, 1])

    with col_form:
        st.subheader("➕ Log New Intervention")
        with st.form("history_form"):
            h_tag = st.text_input("Tag Number *", placeholder="e.g., 204M-FE/FIT-063M")
            h_date = st.date_input("Date *", value=datetime.now())
            h_type = st.selectbox("Intervention Type *", [
                "Calibration", "Visual Inspection", "Liner Inspection",
                "Electrode Check", "Grounding Verification", "Zero Reset",
                "Component Replacement", "Corrective Maintenance", "Other"
            ])
            h_task = st.text_area("Task Performed *", placeholder="Describe what was done...", height=100)
            h_result = st.selectbox("Result *", ["Conform", "Non-conform", "Replaced", "Adjusted", "Pending"])
            h_tech = st.text_input("Technician Name *", placeholder="Full name")
            h_notes = st.text_area("Notes", placeholder="Additional observations...", height=80)
            submitted = st.form_submit_button("💾 Save Intervention", use_container_width=True, type="primary")

            if submitted:
                if not h_tag or not h_task or not h_tech:
                    st.error("Please fill in Tag Number, Task, and Technician Name.")
                else:
                    success = save_history(
                        str(h_date), h_tag.strip(), h_type,
                        h_task, h_result, h_tech, h_notes
                    )
                    if success:
                        st.success(f"✅ Intervention saved for {h_tag}")
                        st.balloons()
                    else:
                        st.error("❌ Error saving. Check Google Sheets connection.")

    with col_view:
        st.subheader("🔍 View History by Tag")
        search_tag = st.text_input("Search Tag Number", placeholder="e.g., 204M-FE/FIT-063M")

        if search_tag:
            with st.spinner("Loading history..."):
                records = load_history(tag_filter=search_tag)
            if records:
                st.success(f"Found {len(records)} intervention(s) for **{search_tag}**")
                for r in records:
                    result_color = "🟢" if r.get('Result') == "Conform" else "🔴" if r.get('Result') == "Non-conform" else "🟡"
                    with st.expander(f"{result_color} {r.get('Date','')} — {r.get('Type','')}"):
                        st.markdown(f"**Task:** {r.get('Task','')}")
                        st.markdown(f"**Result:** {r.get('Result','')}")
                        st.markdown(f"**Technician:** {r.get('Technician','')}")
                        if r.get('Notes'): st.markdown(f"**Notes:** {r.get('Notes','')}")
            else:
                st.info(f"No history found for tag **{search_tag}**")

        st.divider()
        st.subheader("📊 All Interventions by Tag")
        with st.spinner("Loading..."):
            all_records = load_history()
        if all_records:
            st.caption(f"Total interventions recorded: **{len(all_records)}**")
            tags = {}
            for r in all_records:
                tag = r.get('Tag Number', 'Unknown')
                if tag not in tags:
                    tags[tag] = []
                tags[tag].append(r)
            for tag, records in sorted(tags.items()):
                conform_count = sum(1 for r in records if r.get('Result') == 'Conform')
                nonconform_count = sum(1 for r in records if r.get('Result') == 'Non-conform')
                other_count = len(records) - conform_count - nonconform_count
                summary = f"🟢 {conform_count}" if conform_count else ""
                if nonconform_count: summary += f"  🔴 {nonconform_count}"
                if other_count: summary += f"  🟡 {other_count}"
                with st.expander(f"📁 **{tag}** — {len(records)} intervention(s)  {summary}"):
                    for r in sorted(records, key=lambda x: x.get('Date',''), reverse=True):
                        result_color = "🟢" if r.get('Result') == "Conform" else "🔴" if r.get('Result') == "Non-conform" else "🟡"
                        st.markdown(f"{result_color} **{r.get('Date','')}** | {r.get('Type','')} | {r.get('Technician','')}")
                        if r.get('Task'):
                            st.caption(f"↳ {r.get('Task','')[:80]}{'...' if len(r.get('Task','')) > 80 else ''}")
        else:
            st.info("No interventions recorded yet.")

# ===== TAB 2: DASHBOARD =====
with mt2:
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

# ===== TAB 1: RECOMMENDATION ENGINE =====
with mt1:
    tab1,tab2,tab3,tab4 = st.tabs(["🧪 Process Fluid","📐 Pipe Dimensions","🌡️ Operating Conditions","📝 Special Notes"])
    with tab1:
        c1,c2 = st.columns(2)
        with c1:
            fluid = st.selectbox("Service Fluid", ai.get_fluid_names())
            tag_number = st.text_input("🏷️ Tag Number", placeholder="e.g., 204M-FE/FIT-063M")
        with c2:
            cat = ai.get_fluid_category(fluid)
            colors = {'Corrosive':'🔴','Abrasive':'🟠','Charged':'🟡','Clean':'🟢','Unknown':'⚪'}
            st.markdown(f"### Fluid Category: {colors.get(cat,'')} **{cat}**")
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

    if st.button("🚀 Run Recommendation", type="primary", use_container_width=True):
        inp = ProcessInput(fluid_name=fluid, pipe_material=pipe_mat, pipe_thickness=pipe_thick,
            pipe_liner=pipe_liner, tube_material='SS 316L', dn=dn, flow_normal=flow_normal, flow_max=flow_max,
            temp_min=t_min, temp_normal=t_norm, temp_max=t_max, temp_design=t_des,
            pressure_min=p_min, pressure_normal=p_norm, pressure_max=p_max, pressure_design=p_des,
            pressure_drop=p_drop, conductivity=conductivity, viscosity=visc, density=dens,
            special_conditions=special, user_notes=notes)
        result = ai.recommend(inp)
        v,m,vendors,tco = result['validation'],result['materials'],result['vendors'],result['tco']

        # LAYER 1
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
        if not v.is_valid: st.stop()

        # LAYER 2
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
            with c3:
                st.markdown(f"<p style='font-size:18px;font-weight:bold;margin-bottom:2px'>Body Material</p><p style='font-size:15px;color:#333'>SS 304L</p><p style='font-size:12px;color:#666'>Coil housing</p>", unsafe_allow_html=True)
            with c4:
                st.markdown(f"<p style='font-size:18px;font-weight:bold;margin-bottom:2px'>Grounding</p><p style='font-size:15px;color:#333'>{m.grounding}</p>", unsafe_allow_html=True)
            c5,c6,c7 = st.columns(3)
            with c5: st.markdown(f"<p style='font-size:18px;font-weight:bold;margin-bottom:2px'>Penetrant Ring</p><p style='font-size:15px;color:#333'>{m.penetrant or 'N/A'}</p>", unsafe_allow_html=True)
            with c6: st.markdown(f"<p style='font-size:18px;font-weight:bold;margin-bottom:2px'>O-Ring</p><p style='font-size:15px;color:#333'>{m.o_ring}</p>", unsafe_allow_html=True)
            with c7: st.markdown(f"<p style='font-size:18px;font-weight:bold;margin-bottom:2px'>Flange Coating</p><p style='font-size:15px;color:#333'>{m.flange_coat}</p>", unsafe_allow_html=True)
            if m.liner_dn_warning: st.warning(f"📐 {m.liner_dn_warning}")
            if m.tube_warning: st.info(f"🔧 {m.tube_warning}")
            for w in m.warnings: st.warning(w)
            if m.remarks: st.info(f"📝 {m.remarks}")

        # LAYER 3
        with st.container(border=True):
            st.markdown("<div style='background-color:#FFF3E0;padding:15px;border-radius:10px;border-left:5px solid #E65100'><h2 style='color:#E65100;margin:0'>Layer 3 — Vendor Recommendations</h2></div>", unsafe_allow_html=True)
            st.write("")
            st.caption(f"Compatible: {len(vendors['all'])} / {len(ai.data.vendor_models)} models")
            vc1,vc2,vc3 = st.columns(3)
            for cw,label,key in [(vc1,"🔴 EMERSON","emerson"),(vc2,"🔵 ENDRESS+HAUSER","eh"),(vc3,"🟢 KROHNE","krohne")]:
                with cw:
                    rec = vendors[key]
                    st.markdown(f"<p style='font-size:18px;font-weight:bold'>{label}</p>", unsafe_allow_html=True)
                    if rec:
                        so = " ⚠️ SO" if rec.special_order else ""
                        st.markdown(f"<p style='font-size:16px;font-weight:bold;color:#333'>{rec.model}</p><p style='font-size:12px;color:#666'>{rec.model_type}{so}</p>", unsafe_allow_html=True)
                        st.markdown(f"**Accuracy:** {rec.accuracy}")
                        st.markdown(f"**Pressure:** {rec.pressure}")
                        st.markdown(f"**DN:** {rec.pipe_sizes}")
                        st.markdown(f"**Excitation:** {rec.excitation}")
                        st.markdown(f"**IP:** {rec.ip}")
                        st.markdown(f"**Protocols:** {rec.protocols}")
                        st.markdown(f"**Diagnostics:** {rec.diagnostics}")
                    else: st.warning("No compatible model")
            with st.expander("🌍 Other Vendors"):
                for ev in vendors.get('extra',[]): st.markdown(f"**{ev['vendor']}** — {ev['model']} ({ev['accuracy']}) — {ev['notes']}")
            for w in vendors['warnings']: st.warning(w)

        # LAYER 4
        with st.container(border=True):
            st.markdown("<div style='background-color:#F3E5F5;padding:15px;border-radius:10px;border-left:5px solid #6A1B9A'><h2 style='color:#6A1B9A;margin:0'>Layer 4 — TCO & Drift Prediction</h2></div>", unsafe_allow_html=True)
            st.write("")
            t1,t2,t3,t4 = st.columns(4)
            with t1: st.metric("CAPEX Score", f"{tco.capex_score}/30")
            with t2: st.metric("Calibration", f"Every {tco.calib_months} mo")
            with t3: st.metric("Liner Life", tco.liner_life)
            with t4: st.metric("Electrode Life", tco.electrode_life)
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

            # Excel Download — with history pre-filled
            st.markdown("---")
            tag = tag_number if tag_number else "NEW-INSTRUMENT"
            history_for_export = load_history(tag_filter=tag) if tag != "NEW-INSTRUMENT" else []
            excel_buf = generate_maintenance_excel(tag, fluid, cat, m, vendors['emerson'], vendors['eh'], vendors['krohne'], tco, history_for_export)
            st.download_button(label="📄 Download Maintenance Excel", data=excel_buf,
                file_name=f"MagFlow_Maintenance_{tag.replace('/','_')}_{datetime.now().strftime('%Y%m%d')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True)
            st.caption("💡 3 sheets: Guideline (drift risks + instrument info), Maintenance Checklist (daily→annual), Historical Data (pre-filled from shared database)")

            # Validation
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

# ===== PROJECT IMPORT FUNCTIONS =====

def get_excel_from_github():
    """Download current Data_Collection_v2.xlsx from GitHub."""
    try:
        gh = Github(st.secrets["GITHUB_TOKEN"])
        repo = gh.get_repo(GITHUB_REPO)
        contents = repo.get_contents(EXCEL_FILE_PATH)
        excel_bytes = base64.b64decode(contents.content)
        wb = openpyxl.load_workbook(io.BytesIO(excel_bytes))
        return wb, contents.sha, None
    except Exception as e:
        return None, None, str(e)

def push_excel_to_github(wb, sha, commit_msg):
    """Push updated Excel back to GitHub."""
    try:
        gh = Github(st.secrets["GITHUB_TOKEN"])
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
    """Extract existing electrodes, liners, fluids from the Excel."""
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
    """Compare extracted instruments with existing database."""
    new_findings = []
    for inst in instruments:
        electrode = inst.get('electrode', '').strip()
        liner = inst.get('liner', '').strip()
        fluid = inst.get('fluid', '').strip()
        tag = inst.get('tag', '')

        if electrode and electrode not in ('VTA', 'N/A', '') and electrode not in existing["electrodes"]:
            new_findings.append({"type": "electrode", "value": electrode, "tag": tag, "fluid": fluid, "context": f"Found in {tag} for {fluid}"})

        if liner and liner not in ('VTA', 'N/A', '') and liner not in existing["liners"]:
            new_findings.append({"type": "liner", "value": liner, "tag": tag, "fluid": fluid, "context": f"Found in {tag} for {fluid}"})

        if fluid and fluid not in ('VTA', 'N/A', '') and fluid not in existing["fluids"]:
            new_findings.append({"type": "fluid", "value": fluid, "tag": tag, "electrode": electrode, "liner": liner, "context": f"New fluid in {tag}"})
        elif fluid in existing["fluid_electrodes"] and electrode not in ('VTA', 'N/A', ''):
            existing_elec = existing["fluid_electrodes"][fluid]
            if electrode not in existing_elec:
                new_findings.append({"type": "fluid_electrode_variant", "value": electrode, "tag": tag, "fluid": fluid, "context": f"New electrode variant for existing fluid {fluid} (current: {existing_elec})"})

    return new_findings

def apply_updates_to_excel(wb, approved_updates):
    """Apply approved updates to the Excel workbook."""
    changes = []
    for update in approved_updates:
        utype = update["type"]
        value = update["value"]
        try:
            if utype == "electrode":
                ws = wb["Electrode Materials"]
                last_row = ws.max_row + 1
                ws.cell(row=last_row, column=1, value=value)
                ws.cell(row=last_row, column=2, value="Auto-imported")
                ws.cell(row=last_row, column=3, value=datetime.now().strftime('%Y-%m-%d'))
                changes.append(f"Added electrode: {value}")

            elif utype == "liner":
                ws = wb["Liner Materials"]
                last_row = ws.max_row + 1
                ws.cell(row=last_row, column=1, value=value)
                ws.cell(row=last_row, column=2, value="Auto-imported")
                ws.cell(row=last_row, column=3, value=datetime.now().strftime('%Y-%m-%d'))
                changes.append(f"Added liner: {value}")

            elif utype == "fluid":
                ws = wb["Fluid-Material Matrix"]
                last_row = ws.max_row + 1
                ws.cell(row=last_row, column=1, value=value)
                ws.cell(row=last_row, column=2, value=update.get("electrode", "VTA"))
                ws.cell(row=last_row, column=3, value=update.get("liner", "VTA"))
                changes.append(f"Added fluid: {value}")

            elif utype == "fluid_electrode_variant":
                ws = wb["Fluid-Material Matrix"]
                for row in ws.iter_rows(min_row=2):
                    if row[0].value and str(row[0].value).strip() == update["fluid"]:
                        current = str(row[1].value or "")
                        if value not in current:
                            row[1].value = current + " / " + value if current else value
                            changes.append(f"Added electrode variant {value} to fluid {update['fluid']}")
                        break
        except Exception as e:
            changes.append(f"Error updating {utype} {value}: {str(e)}")

    return wb, changes

def extract_datasheet_with_ai(pdf_bytes):
    try:
        client = anthropic.Anthropic(api_key=st.secrets["ANTHROPIC_API_KEY"])
        pdf_b64 = base64.standard_b64encode(pdf_bytes).decode("utf-8")
        prompt = """Extract ALL flowmeter instrument data from this JESA datasheet PDF.
For each instrument (each tag number), return JSON array:
[{"project":"...","tag":"...","service":"...","fluid":"...","dn":400,"flow_normal":450,"flow_max":495,"temp_design":80,"pressure_design":24.5,"conductivity":">20 µS/cm","electrode":"Platinum","liner":"PFA","tube":"316L SS","grounding":"Grounding straps","accuracy":"±0.2%","vendor":"VTA","model":"VTA"}]
Return ONLY valid JSON, no markdown, no explanation."""
        response = client.messages.create(
            model="claude-haiku-4-5-20251001", max_tokens=2000,
            messages=[{"role": "user", "content": [
                {"type": "document", "source": {"type": "base64", "media_type": "application/pdf", "data": pdf_b64}},
                {"type": "text", "text": prompt}
            ]}])
        text = response.content[0].text.strip().replace("```json","").replace("```","").strip()
        return json.loads(text), None
    except Exception as e:
        return [], str(e)

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
            ws.append_row([inst.get('project', project_name), inst.get('tag',''), inst.get('service',''),
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

# ===== TAB 4: PROJECT IMPORT =====
with mt4:
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

        # ===== REVIEW & APPROVE =====
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
                        st.markdown(f"**Project:** {inst.get('project', pname)}")
                        st.markdown(f"**Fluid:** {inst.get('fluid','')}")
                        st.markdown(f"**DN:** {inst.get('dn','')} mm")
                        st.markdown(f"**Flow:** {inst.get('flow_normal','')} → {inst.get('flow_max','')} m³/h")
                        st.markdown(f"**Temp design:** {inst.get('temp_design','')} °C")
                        st.markdown(f"**Pressure design:** {inst.get('pressure_design','')} bar-g")
                    with c2:
                        st.markdown(f"**Electrode:** {inst.get('electrode','')}")
                        st.markdown(f"**Liner:** {inst.get('liner','')}")
                        st.markdown(f"**Tube:** {inst.get('tube','')}")
                        st.markdown(f"**Grounding:** {inst.get('grounding','')}")
                        st.markdown(f"**Accuracy:** {inst.get('accuracy','')}")
                        st.markdown(f"**Vendor/Model:** {inst.get('vendor','')} / {inst.get('model','')}")

            # ===== NEW MATERIALS DETECTED =====
            if new_findings:
                st.divider()
                st.subheader("🆕 New Materials Detected")
                st.warning(f"**{len(new_findings)} new item(s)** not found in the current database. Select which ones to add:")

                approved = []
                for i, finding in enumerate(new_findings):
                    ftype = finding['type']
                    icons = {"electrode": "⚡", "liner": "🟡", "fluid": "💧", "fluid_electrode_variant": "🔀"}
                    labels = {"electrode": "New Electrode", "liner": "New Liner", "fluid": "New Fluid", "fluid_electrode_variant": "New Electrode Variant"}
                    icon = icons.get(ftype, "🆕")
                    label = labels.get(ftype, "New")
                    checked = st.checkbox(
                        f"{icon} **{label}:** `{finding['value']}` — {finding['context']}",
                        value=True, key=f"approve_{i}")
                    if checked:
                        approved.append(finding)

                st.session_state['import_approved'] = approved
            else:
                st.info("✅ No new materials detected — all extracted data matches the existing database.")
                st.session_state['import_approved'] = []

            st.divider()
            col_save1, col_save2 = st.columns(2)

            with col_save1:
                if st.button("💾 Save to Project Database", use_container_width=True):
                    with st.spinner("Saving to Google Sheets..."):
                        saved, err = save_project_to_gsheet(instruments, pname)
                    if err:
                        st.error(f"❌ {err}")
                    else:
                        st.success(f"✅ {saved} instrument(s) saved to Project Database")

            with col_save2:
                approved_updates = st.session_state.get('import_approved', [])
                if approved_updates:
                    if st.button(f"🔄 Apply {len(approved_updates)} update(s) to Database", type="primary", use_container_width=True):
                        with st.spinner("Updating Data_Collection_v2.xlsx on GitHub..."):
                            wb = st.session_state['import_wb']
                            sha = st.session_state['import_sha']
                            wb, changes = apply_updates_to_excel(wb, approved_updates)
                            commit_msg = f"Auto-update: {len(approved_updates)} new material(s) from {pname or 'datasheet import'}"
                            success, err = push_excel_to_github(wb, sha, commit_msg)
                        if success:
                            st.success(f"✅ Database updated on GitHub!")
                            for c in changes:
                                st.markdown(f"  • {c}")
                            st.info("♻️ The app will reload with the new data in ~1 minute.")
                            for key in ['import_instruments','import_wb','import_sha','import_new','import_approved']:
                                if key in st.session_state: del st.session_state[key]
                        else:
                            st.error(f"❌ GitHub update failed: {err}")
                else:
                    st.button("✅ No updates to apply", disabled=True, use_container_width=True)

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
        else:
            st.info("No projects imported yet. Upload a datasheet to get started.")
