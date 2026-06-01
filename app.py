"""MagFlow AI v5 — Voice Fix + Chat History + Dashboard"""
import streamlit as st
import anthropic
from datetime import datetime
from magflow_ai import MagFlowAI, ProcessInput

st.set_page_config(page_title="MagFlow AI — JESA", page_icon="🔧", layout="wide")

@st.cache_resource
def load_model():
    return MagFlowAI('Data_Collection_v2.xlsx')

ai = load_model()

def build_context():
    fluids = ", ".join(f['name'] for f in ai.data.fluids[:25]) + f"... ({len(ai.data.fluids)} total)"
    return f"""You are MagFlow AI Assistant for JESA/OCP electromagnetic flowmeters.
Database: {len(ai.data.fluids)} fluids, {len(ai.data.electrodes)} electrodes, {len(ai.data.liners)} liners, {len(ai.data.vendor_models)} models, {len(ai.data.drift_indicators)} drift indicators, {len(ai.data.project_data)} project entries.
Fluids: {fluids}
Rules: PFA default for acids, PTFE for large DN, Soft rubber for rock slurry (MIA/Daoui), Min 5 uS/cm, Max 10 m/s liquids / 3-6 m/s slurries, Grounding ring for FRP/PVC.
Diagnostics: Emerson SMV, E+H Heartbeat, Krohne OPTICHECK.
Answer in same language as question. Specify source: [JESA Database], [General Knowledge], or [Web Search].
For new tech/products: USE web_search, provide URLs."""

def search_local(query):
    q = query.lower().replace('-',' '); results = []
    words = [w for w in q.split() if len(w) > 2]
    combined = q.replace(' ','')
    for f in ai.data.fluids:
        fn = f['name'].lower(); fnc = fn.replace(' ','')
        if any(w in fn for w in words) or combined in fnc or fnc in combined:
            results.append(f"Fluid: {f['name']} | Type: {f['type']} | Electrode: {f['electrode']} | Liner: {f['liner']} | Grounding: {f['grounding']} | Remarks: {f['remarks']}")
    for e in ai.data.electrodes:
        en = e['name'].lower()
        if any(w in en for w in words) or en.replace(' ','') in combined:
            results.append(f"Electrode: {e['name']} | Emerson: {e['emerson']} | E+H: {e['eh']} | Krohne: {e['krohne']} | Cost: {e['cost']} | Abrasion: {e['abrasion']}")
    for m in ai.data.vendor_models:
        mn = (m['model'] + ' ' + m['vendor']).lower()
        if any(w in mn for w in words):
            results.append(f"Model: {m['vendor']} {m['model']} | Accuracy: {m['accuracy']} | Diagnostics: {m['diagnostics'][:50]}")
    for d in ai.data.drift_indicators:
        if any(w in d['indicator'].lower() for w in words):
            results.append(f"Drift: {d['indicator']} | {d['desc']}")
    for p in ai.data.project_data:
        pf = (p.get('fluid','') + ' ' + p.get('project','')).lower()
        if any(w in pf for w in words):
            results.append(f"Project: {p['project']} | {p['fluid']} | Electrode: {p.get('electrode','')} | Liner: {p.get('liner','')}")
    return results[:15]

def ask_claude(question, local_results, history=None):
    if history is None:
        history = []
    try:
        client = anthropic.Anthropic(api_key=st.secrets["ANTHROPIC_API_KEY"])
        ctx = ""
        if local_results:
            ctx = "\n\nJESA database results:\n" + "\n".join(f"- {r}" for r in local_results)
        conv_summary = ""
        if history:
            conv_summary = "\n\nPrevious conversation:\n"
            for msg in history[-6:]:
                role = "User" if msg["role"] == "user" else "Assistant"
                conv_summary += f"{role}: {msg['content'][:200]}\n"
        response = client.messages.create(
            model="claude-haiku-4-5-20251001",
            max_tokens=1500,
            system=build_context(),
            tools=[{"type": "web_search_20250305", "name": "web_search"}],
            messages=[{"role": "user", "content": question + ctx + conv_summary}]
        )
        parts = []
        for block in response.content:
            if hasattr(block, 'text'):
                parts.append(block.text)
        return "\n".join(parts) if parts else "No response."
    except Exception as e:
        return f"Error: {str(e)}"

def generate_pdf_html(fluid, cat, mat, ve, veh, vk, tco):
    vn = ve.model if ve else (veh.model if veh else (vk.model if vk else 'N/A'))
    vd = ve.diagnostics if ve else (veh.diagnostics if veh else (vk.diagnostics if vk else 'N/A'))
    today = datetime.now().strftime('%Y-%m-%d')
    mnt = tco.maintenance
    rh = ""
    for risk in tco.drift_risks:
        ic = "🔴" if risk.level=="High" else "🟡" if risk.level=="Medium" else "🟢"
        sh = "".join(f"<li>{s}</li>" for s in risk.steps)
        th = "".join(f"<li><b>{v}:</b> {t}</li>" for v,t in risk.tools.items())
        rh += f"<div style='border:1px solid #ddd;padding:10px;margin:5px 0;border-radius:5px'><b>{ic} {risk.indicator} — {risk.level}</b><p>{risk.description}</p><b>Steps:</b><ol>{sh}</ol><b>Tools:</b><ul>{th}</ul><b>Frequency:</b> {risk.frequency}</div>"
    def tl(items):
        return "".join(f"<tr><td style='padding:8px;border:1px solid #ddd'>☐</td><td style='padding:8px;border:1px solid #ddd'>{i}</td><td style='padding:8px;border:1px solid #ddd'>____/____/____</td><td style='padding:8px;border:1px solid #ddd'>____________</td></tr>" for i in items)
    sec = ""
    for title, items, color in [("Continuous",mnt.continuous,"#E8F5E9"),("Monthly",mnt.monthly,"#E3F2FD"),("Quarterly",mnt.quarterly,"#FFF3E0"),("Semi-Annual",mnt.semi_annual,"#FCE4D6"),("Annual",mnt.annual,"#F3E5F5"),("Multi-Year",mnt.multi_year,"#FFF9C4"),("Replacement",mnt.replacement,"#FFCDD2")]:
        if items:
            sec += f"<div style='margin:15px 0'><h3 style='background:{color};padding:10px;border-radius:5px'>{title}</h3><table style='width:100%;border-collapse:collapse'><tr style='background:#f5f5f5'><th style='padding:8px;border:1px solid #ddd;width:40px'>✓</th><th style='padding:8px;border:1px solid #ddd'>Task</th><th style='padding:8px;border:1px solid #ddd;width:120px'>Date</th><th style='padding:8px;border:1px solid #ddd;width:120px'>Technician</th></tr>{tl(items)}</table></div>"
    return f"""<!DOCTYPE html><html><head><meta charset='utf-8'><title>MagFlow Maintenance</title><style>body{{font-family:Arial;margin:20px;font-size:13px}}h1{{color:#1B2A4A;border-bottom:3px solid #1B2A4A;padding-bottom:10px}}.g{{display:grid;grid-template-columns:1fr 1fr;gap:10px;margin:15px 0}}.i{{background:#f8f9fa;padding:10px;border-radius:5px;border-left:4px solid #1B2A4A}}.i b{{color:#1B2A4A}}@media print{{body{{margin:10mm}}}}</style></head><body><h1>🔧 MagFlow AI — Maintenance Sheet</h1><p style='color:#666'>Generated: {today}</p><div class='g'><div class='i'><b>Fluid:</b> {fluid}</div><div class='i'><b>Category:</b> {cat}</div><div class='i'><b>Electrode:</b> {mat.electrode} ({mat.electrode_cost})</div><div class='i'><b>Liner:</b> {mat.liner} ({mat.liner_cost})</div><div class='i'><b>Grounding:</b> {mat.grounding}</div><div class='i'><b>O-Ring:</b> {mat.o_ring}</div><div class='i'><b>Model:</b> {vn}</div><div class='i'><b>Diagnostics:</b> {vd[:60]}</div><div class='i'><b>CAPEX:</b> {tco.capex_score}/30</div><div class='i'><b>Calibration:</b> Every {tco.calib_months} mo</div><div class='i'><b>Liner Life:</b> {tco.liner_life}</div><div class='i'><b>Electrode Life:</b> {tco.electrode_life}</div></div><h2>Drift Risks</h2>{rh}<h2>Maintenance Checklist</h2>{sec}<div style='margin-top:30px;border-top:2px solid #1B2A4A;padding-top:15px'><p><b>Approved by:</b> _________________ <b>Date:</b> ____/____/____</p><p style='color:#999;font-size:11px'>Generated by MagFlow AI v5 — JESA</p></div></body></html>"""

# ===== SIDEBAR =====
with st.sidebar:
    st.markdown("### 💬 MagFlow Assistant")
    voice_lang = st.selectbox("🌍 Voice language", ["French","English","Arabic","Spanish"], key="vlang", label_visibility="collapsed")
    lang_map = {"French":"fr-FR","English":"en-US","Arabic":"ar-SA","Spanish":"es-ES"}
    lang_code = lang_map.get(voice_lang, "en-US")
    voice_html = f"""<div>
    <button id="startBtn" onclick="startVoice()" style='background:#1B2A4A;color:white;border:none;padding:8px 16px;border-radius:20px;cursor:pointer;width:100%;font-size:14px'>🎤 Tap to speak</button>
    <button id="stopBtn" onclick="stopVoice()" style='background:#E53935;color:white;border:none;padding:8px 16px;border-radius:20px;cursor:pointer;width:100%;font-size:14px;display:none'>⏹ Send</button>
    <p id="vs" style='color:#888;font-size:11px;margin:4px 0'></p>
    <script>
    var recognition=null; var fullText='';
    function startVoice(){{
        if(!('webkitSpeechRecognition' in window)&&!('SpeechRecognition' in window)){{document.getElementById('vs').innerText='❌ Use Chrome';return;}}
        const S=window.SpeechRecognition||window.webkitSpeechRecognition;
        recognition=new S();
        recognition.lang='{lang_code}';
        recognition.continuous=true;
        recognition.interimResults=true;
        fullText='';
        document.getElementById('startBtn').style.display='none';
        document.getElementById('stopBtn').style.display='block';
        document.getElementById('vs').innerText='🔴 Listening... Click Send when done';
        recognition.onresult=function(e){{
            var t='';
            for(var i=0;i<e.results.length;i++){{t+=e.results[i][0].transcript+' ';}}
            fullText=t.trim();
            document.getElementById('vs').innerText='🔴 '+fullText;
        }};
        recognition.onerror=function(e){{document.getElementById('vs').innerText='❌ '+e.error;}};
        recognition.start();
    }}
    function stopVoice(){{
        if(recognition){{recognition.stop();}}
        document.getElementById('startBtn').style.display='block';
        document.getElementById('stopBtn').style.display='none';
        if(fullText){{
            document.getElementById('vs').innerText='✅ Sending: '+fullText;
            const u=new URL(window.parent.location);
            u.searchParams.set('voice',fullText);
            window.parent.history.replaceState({{}},'',u);
            window.parent.location.reload();
        }} else {{
            document.getElementById('vs').innerText='No speech detected';
        }}
    }}
    </script></div>"""
    st.components.v1.html(voice_html, height=80)
    voice_query = st.query_params.get("voice", None)
    if voice_query:
        st.query_params.clear()
        st.session_state.voice_input = voice_query
    st.caption("Or type your question:")
    if "messages" not in st.session_state:
        st.session_state.messages = [{"role": "assistant", "content": "👋 How can I help you today?"}]
    for msg in st.session_state.messages:
        with st.chat_message(msg["role"]):
            st.markdown(msg["content"])
    voice_text = st.session_state.pop("voice_input", None)
    prompt = st.chat_input("Ask a question...", key="chat_input")
    if voice_text:
        prompt = voice_text
    if prompt:
        st.session_state.messages.append({"role": "user", "content": prompt})
        with st.chat_message("user"):
            st.markdown(prompt)
        local_results = search_local(prompt)
        chat_history = [{"role": m["role"], "content": m["content"]} for m in st.session_state.messages[1:] if m["role"] in ["user", "assistant"]][:-1]
        with st.chat_message("assistant"):
            with st.spinner("Thinking..."):
                response = ask_claude(prompt, local_results, chat_history)
            st.markdown(response)
        st.session_state.messages.append({"role": "assistant", "content": response})
    if st.button("🗑️ Clear chat", use_container_width=True):
        st.session_state.messages = [{"role": "assistant", "content": "👋 Chat cleared."}]
        st.rerun()

# ===== MAIN =====
st.markdown("""<div style='text-align:center;padding:10px 0'><h1 style='color:#1B2A4A;margin:0'>🔧 MagFlow AI</h1><p style='font-size:16px;color:#555'>AI-Based Predictive Maintenance System for Electromagnetic Flowmeters</p><p style='font-size:12px;color:#999'>PFE ENSA — JESA (OCP × Worley) | Instrumentation & Control Department</p><p style='font-size:11px;color:#bbb'>Developed by Maroua Hakkak — 2026</p></div>""", unsafe_allow_html=True)
st.divider()

main_tab1, main_tab2 = st.tabs(["🔧 Recommendation Engine", "📊 Dashboard"])

with main_tab2:
    st.header("📊 JESA Project Analytics")
    st.subheader("Flowmeters per Project")
    pc = {}
    for p in ai.data.project_data:
        pn = p.get('project','')
        if pn and not any(k in pn for k in ['LEGEND','NOTE','Original','UPDATED']):
            pc[pn] = pc.get(pn, 0) + 1
    if pc:
        for pn, cnt in sorted(pc.items(), key=lambda x: -x[1]):
            st.markdown(f"**{pn}** — {cnt}")
            st.progress(cnt / max(pc.values()))
    d1, d2 = st.columns(2)
    with d1:
        st.subheader("Most Used Electrodes")
        ec = {}
        for p in ai.data.project_data:
            e = p.get('electrode','').strip()
            if e and e != 'N/A': ec[e] = ec.get(e, 0) + 1
        for en, cnt in sorted(ec.items(), key=lambda x: -x[1])[:10]:
            st.markdown(f"**{en}** — {cnt}")
            st.progress(cnt / max(ec.values()))
    with d2:
        st.subheader("Most Used Liners")
        lc = {}
        for p in ai.data.project_data:
            l = p.get('liner','').strip()
            if l and l != 'N/A': lc[l] = lc.get(l, 0) + 1
        for ln, cnt in sorted(lc.items(), key=lambda x: -x[1])[:10]:
            st.markdown(f"**{ln}** — {cnt}")
            st.progress(cnt / max(lc.values()))
    st.subheader("Database Summary")
    s1,s2,s3,s4,s5,s6 = st.columns(6)
    with s1: st.metric("Fluids", len(ai.data.fluids))
    with s2: st.metric("Electrodes", len(ai.data.electrodes))
    with s3: st.metric("Liners", len(ai.data.liners))
    with s4: st.metric("Models", len(ai.data.vendor_models))
    with s5: st.metric("Drift", len(ai.data.drift_indicators))
    with s6: st.metric("Projects", len(ai.data.project_data))

with main_tab1:
    tab1,tab2,tab3,tab4 = st.tabs(["🧪 Process Fluid","📐 Pipe Dimensions","🌡️ Operating Conditions","📝 Special Notes"])
    with tab1:
        c1,c2 = st.columns(2)
        with c1:
            fluid_list = ai.get_fluid_names()
            fluid = st.selectbox("Service Fluid", fluid_list)
        with c2:
            cat = ai.get_fluid_category(fluid)
            colors = {'Corrosive':'🔴','Abrasive':'🟠','Charged':'🟡','Clean':'🟢','Unknown':'⚪'}
            st.markdown(f"### Fluid Category: {colors.get(cat,'')} **{cat}**")
            conductivity = st.number_input("Conductivity (µS/cm)", value=5000.0, min_value=0.0, step=100.0)
    with tab2:
        c1,c2,c3 = st.columns(3)
        with c1: dn = st.number_input("DN (mm)", value=80, min_value=3, max_value=3000, step=10)
        with c2: pipe_mat = st.selectbox("Pipe Material", ["FRP","Stainless Steel","Carbon Steel","ZeCor","PVC","HDPE"])
        with c3: pipe_liner = st.selectbox("Pipe Liner (Internal)", ["Not applicable","HDPE","Rubber lined","Epoxy coated","Glass lined"])
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
        special = st.text_area("Special Conditions", placeholder="e.g., ATEX Zone 1, SIL 2...", height=80)
        notes = st.text_area("User Notes", placeholder="e.g., client prefers Emerson...", height=80)

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

        with st.container(border=True):
            st.markdown("<div style='background-color:#E8F5E9;padding:15px;border-radius:10px;border-left:5px solid #2E7D32'><h2 style='color:#2E7D32;margin:0'>Layer 1 — Input Validation</h2></div>", unsafe_allow_html=True)
            st.write("")
            if v.is_valid:
                st.success(f"✅ Magflow suitable — Velocity: {v.velocity} m/s — Category: {v.fluid_category}")
                if v.recommended_dn > 0: st.caption(f"💡 Sizing: DN {v.recommended_dn} mm recommended")
            else:
                st.error("❌ Magflow NOT suitable")
                for e in v.errors: st.error(e)
            for w in v.warnings: st.warning(w)
        if not v.is_valid: st.stop()

        with st.container(border=True):
            st.markdown("<div style='background-color:#E3F2FD;padding:15px;border-radius:10px;border-left:5px solid #1565C0'><h2 style='color:#1565C0;margin:0'>Layer 2 — Material Selection</h2></div>", unsafe_allow_html=True)
            st.write("")
            c1,c2,c3 = st.columns(3)
            with c1:
                ab = " ⚠️ Special Order" if 'special' in m.electrode_avail.lower() else ""
                st.markdown(f"<p style='font-size:18px;font-weight:bold;margin-bottom:2px'>Electrode</p><p style='font-size:15px;color:#333'>{m.electrode}{ab}</p><p style='font-size:12px;color:#2E7D32'>Cost: {m.electrode_cost}</p>", unsafe_allow_html=True)
                if m.electrode_alt: st.caption(f"Alternatives: {', '.join(m.electrode_alt)}")
            with c2:
                ab2 = " ⚠️ Special Order" if 'special' in m.liner_avail.lower() else ""
                st.markdown(f"<p style='font-size:18px;font-weight:bold;margin-bottom:2px'>Liner</p><p style='font-size:15px;color:#333'>{m.liner}{ab2}</p><p style='font-size:12px;color:#2E7D32'>Cost: {m.liner_cost}</p>", unsafe_allow_html=True)
                if m.liner_alt: st.caption(f"Alternatives: {', '.join(m.liner_alt)}")
            with c3: st.markdown(f"<p style='font-size:18px;font-weight:bold;margin-bottom:2px'>Grounding</p><p style='font-size:15px;color:#333'>{m.grounding}</p>", unsafe_allow_html=True)
            c4,c5,c6 = st.columns(3)
            with c4: st.markdown(f"<p style='font-size:18px;font-weight:bold;margin-bottom:2px'>Penetrant Ring</p><p style='font-size:15px;color:#333'>{m.penetrant or 'Not required'}</p>", unsafe_allow_html=True)
            with c5: st.markdown(f"<p style='font-size:18px;font-weight:bold;margin-bottom:2px'>O-Ring</p><p style='font-size:15px;color:#333'>{m.o_ring}</p>", unsafe_allow_html=True)
            with c6: st.markdown(f"<p style='font-size:18px;font-weight:bold;margin-bottom:2px'>Flange Coating</p><p style='font-size:15px;color:#333'>{m.flange_coat}</p>", unsafe_allow_html=True)
            if m.liner_dn_warning: st.warning(f"📐 {m.liner_dn_warning}")
            if m.tube_warning: st.info(f"🔧 {m.tube_warning}")
            for w in m.warnings: st.warning(w)
            if m.remarks: st.info(f"📝 {m.remarks}")

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
                        so = " ⚠️ Special Order" if rec.special_order else ""
                        st.markdown(f"<p style='font-size:16px;font-weight:bold;color:#333'>{rec.model}</p><p style='font-size:12px;color:#666'>{rec.model_type}{so}</p>", unsafe_allow_html=True)
                        st.markdown(f"**Accuracy:** {rec.accuracy}")
                        st.markdown(f"**Pressure:** {rec.pressure}")
                        st.markdown(f"**DN Range:** {rec.pipe_sizes}")
                        st.markdown(f"**Excitation:** {rec.excitation}")
                        st.markdown(f"**IP:** {rec.ip}")
                        st.markdown(f"**Protocols:** {rec.protocols}")
                        st.markdown(f"**Diagnostics:** {rec.diagnostics}")
                    else: st.warning("No compatible model")
            with st.expander("🌍 Other Vendors"):
                for ev in vendors.get('extra',[]): st.markdown(f"**{ev['vendor']}** — {ev['model']} ({ev['accuracy']}) — {ev['notes']}")
            for w in vendors['warnings']: st.warning(w)

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
            st.markdown("<p style='font-size:18px;font-weight:bold'>📋 Maintenance Checklist</p>", unsafe_allow_html=True)
            mnt = tco.maintenance
            for title,items,icon in [("Continuous",mnt.continuous,"🔄"),("Monthly",mnt.monthly,"📅"),("Quarterly",mnt.quarterly,"🔍"),("Semi-Annual",mnt.semi_annual,"⚙️"),("Annual",mnt.annual,"🛠️"),("Multi-Year",mnt.multi_year,"📊"),("Replacement",mnt.replacement,"♻️")]:
                if items:
                    with st.expander(f"{icon} {title} ({len(items)} tasks)"):
                        for item in items:
                                st.markdown(f"• {item}")
            st.markdown("---")
            pdf_html = generate_pdf_html(fluid, cat, m, vendors['emerson'], vendors['eh'], vendors['krohne'], tco)
            st.download_button(label="📄 Download Maintenance Sheet", data=pdf_html,
                file_name=f"MagFlow_{fluid.replace(' ','_')}_{datetime.now().strftime('%Y%m%d')}.html",
                mime="text/html", use_container_width=True)
            st.caption("💡 Open in browser → Cmd+P → Save as PDF")
            st.markdown("<p style='font-size:18px;font-weight:bold'>🔍 Validation vs JESA Projects</p>", unsafe_allow_html=True)
            if tco.validation:
                for match in tco.validation:
                    ic = "✅" if "MATCH" in match.match_type else ("⚠️" if "PARTIAL" in match.match_type else "❌")
                    with st.expander(f"{ic} {match.match_type} — {match.project} ({match.fluid}) — {match.confidence}%"):
                        for k,vd in match.details.items(): st.markdown(f"**{k}:** {vd}")
                        if match.explanation: st.markdown(f"**Why:** {match.explanation}")
                        if match.cost_impact: st.markdown(f"**Cost:** {match.cost_impact}")
                total = len(tco.validation); matches = sum(1 for mt in tco.validation if mt.confidence >= 80)
                st.caption(f"Summary: {matches}/{total} matches. {'Validated ✓' if matches > total/2 else 'Review needed ⚠'}")
            else: st.info("No matching JESA project found.")
            if tco.notes_response:
                st.markdown("<p style='font-size:18px;font-weight:bold'>💡 AI Response</p>", unsafe_allow_html=True)
                st.markdown(tco.notes_response)
        st.divider()
        st.caption("Source: JESA Internal DB, JESA Flow App 2024, Vendor datasheets, JESA Projects")
