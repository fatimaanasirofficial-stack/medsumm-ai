import streamlit as st
from groq import Groq
import fitz
from docx import Document
import os
import re
import time
import pathlib
from datetime import datetime
from dotenv import load_dotenv
from fpdf import FPDF
import plotly.graph_objects as go

load_dotenv()
client = Groq(api_key=os.getenv("GROQ_API_KEY"))

st.set_page_config(page_title="MedSumm AI", page_icon="🏥", layout="wide")

THEMES = {
    "Ocean":   {"primary": "#00d4ff", "high": "#ff4d6d", "mid": "#ffb84d", "ok": "#2ee6a8", "bg": "#0f1c2e", "bg2": "#1a2a3a"},
    "Emerald": {"primary": "#2ee6a8", "high": "#ff5c5c", "mid": "#ffd166", "ok": "#2ee6a8", "bg": "#0f1f1a", "bg2": "#1a2e25"},
    "Violet":  {"primary": "#b388ff", "high": "#ff6ec7", "mid": "#ffd166", "ok": "#2ee6a8", "bg": "#1a0f2e", "bg2": "#2a1a3e"},
}

NORMAL_RANGES = {
    "Heart Rate (bpm)":  (60,  100),
    "Systolic BP":       (90,  120),
    "Diastolic BP":      (60,  80),
    "Temperature (F)":   (97,  99),
    "SpO2 (%)":          (95,  100),
    "Glucose (mg/dL)":   (70,  100),
}

if "theme"   not in st.session_state: st.session_state.theme   = "Ocean"
if "history" not in st.session_state: st.session_state.history = []

with st.sidebar:
    st.markdown("### 🎨 Appearance")
    chosen = st.selectbox(
        "Color Theme", list(THEMES.keys()),
        index=list(THEMES.keys()).index(st.session_state.theme)
    )
    st.session_state.theme = chosen

    st.markdown("### ⚙️ Settings")
    detail_level = st.select_slider(
        "Summary Detail",
        options=["Brief", "Standard", "Detailed"],
        value="Standard"
    )

    st.markdown("---")
    st.markdown("### 📜 Session History")
    if st.session_state.history:
        for h in reversed(st.session_state.history[-5:]):
            st.caption(f"📄 {h['name']} — {h['time']}")
    else:
        st.caption("No reports analyzed yet.")

    st.markdown("---")
    st.markdown(
        "Built by **Fatima Nasir** · Biomedical Engineer\n\n"
        "[![LinkedIn](https://img.shields.io/badge/LinkedIn-Connect-blue?logo=linkedin)](https://www.linkedin.com/in/fatima-nasir-bme)"
    )
    st.caption("Powered by Streamlit + Groq (Llama 3.3)")

t = THEMES[st.session_state.theme]

st.markdown(f"""
<style>
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800&display=swap');
html, body, [class*="css"] {{ font-family: 'Inter', sans-serif; }}

.metric-card {{
    background: linear-gradient(135deg, {t['bg']}, {t['bg2']});
    padding: 18px; border-radius: 12px; text-align: center;
    border: 1px solid {t['primary']}44;
}}
.metric-card b    {{ font-size: 1.4rem; color: {t['primary']}; display: block; }}
.metric-card span {{ font-size: 0.75rem; color: #9ca3af;
                     text-transform: uppercase; letter-spacing: 0.05em; }}

.flag-high {{
    background: #3a1a1a; padding: 10px 14px; border-radius: 8px;
    border-left: 4px solid {t['high']}; margin-bottom: 6px;
    font-size: 0.9rem; color: #f9a8a8;
}}
.flag-mid {{
    background: #3a2e1a; padding: 10px 14px; border-radius: 8px;
    border-left: 4px solid {t['mid']}; margin-bottom: 6px;
    font-size: 0.9rem; color: #fde68a;
}}
.flag-ok {{
    background: #1a2e1a; padding: 10px 14px; border-radius: 8px;
    border-left: 4px solid {t['ok']}; margin-bottom: 6px;
    font-size: 0.9rem; color: #a7f3d0;
}}

mark {{ background-color: {t['primary']}55; color: white;
        padding: 1px 4px; border-radius: 3px; }}

.confidence-box {{
    background: {t['bg2']}; border-radius: 8px; padding: 12px 16px;
    border-left: 4px solid {t['primary']}; font-size: 0.82rem;
    color: #9ca3af; margin-top: 16px;
}}

.trim-warning {{
    background: #3a2e1a; border-radius: 8px; padding: 10px 14px;
    border-left: 4px solid {t['mid']}; font-size: 0.85rem;
    color: #ffd166; margin-bottom: 12px;
}}

.footer-box {{
    text-align: center; color: #6b7280; font-size: 0.82rem;
    padding: 24px 0 8px; border-top: 1px solid #1f2937; margin-top: 40px;
}}
.footer-box a {{
    color: {t['primary']}; text-decoration: none; font-weight: 600;
}}
.footer-box a:hover {{ text-decoration: underline; }}
.footer-name {{
    font-size: 1rem; font-weight: 700; color: white; margin-bottom: 4px;
}}

.stButton > button {{
    background: linear-gradient(90deg, {t['primary']}, {t['primary']}99) !important;
    border: none !important; font-weight: 600 !important;
    color: #000 !important; border-radius: 8px !important;
    padding: 0.5rem 1.5rem !important;
}}
.stButton > button:hover {{ opacity: 0.85 !important; }}
</style>
""", unsafe_allow_html=True)

# ── HEADER ──
st.markdown(
    f'<h1 style="font-size:2.4rem;font-weight:800;'
    f'background:linear-gradient(90deg,{t["primary"]},#ffffff);'
    f'-webkit-background-clip:text;-webkit-text-fill-color:transparent;'
    f'margin-bottom:0;">🏥 MedSumm AI</h1>',
    unsafe_allow_html=True
)
st.markdown(
    '<p style="color:#9ca3af;margin-top:0;font-size:1rem;">'
    'AI-powered medical report analysis — structured, fast, transparent.</p>',
    unsafe_allow_html=True
)

st.warning("⚠️ AI-generated content. Not a substitute for professional medical advice. Always consult a licensed provider.")

# ── HELPERS ──
def extract_text(uploaded_file):
    name = uploaded_file.name
    if name.endswith(".pdf"):
        doc = fitz.open(stream=uploaded_file.read(), filetype="pdf")
        return "\n".join(page.get_text() for page in doc), len(doc)
    elif name.endswith(".docx"):
        doc = Document(uploaded_file)
        return "\n".join(p.text for p in doc.paragraphs), None
    else:
        return uploaded_file.read().decode("utf-8"), None


def extract_vitals(text):
    patterns = {
        "Heart Rate (bpm)": r"(?:heart rate|hr|pulse)[:\s]+(\d{2,3})",
        "Systolic BP":       r"(?:bp|blood pressure)[:\s]+(\d{2,3})\s*/\s*\d{2,3}",
        "Diastolic BP":      r"(?:bp|blood pressure)[:\s]+\d{2,3}\s*/\s*(\d{2,3})",
        "Temperature (F)":   r"(?:temp|temperature)[:\s]+(\d{2,3}(?:\.\d)?)",
        "SpO2 (%)":          r"(?:spo2|oxygen saturation)[:\s]+(\d{2,3})",
        "Glucose (mg/dL)":   r"(?:glucose)[:\s]+(\d{2,3})",
    }
    found = {}
    for label, pat in patterns.items():
        m = re.search(pat, text, re.IGNORECASE)
        if m:
            found[label] = float(m.group(1))
    return found


def summarize_report(report_text, detail):
    detail_map = {
        "Brief":    "Keep each section to 1-2 sentences max.",
        "Standard": "Keep each section concise but complete.",
        "Detailed": "Provide thorough, clinically rich explanations in each section.",
    }
    trimmed = False
    if len(report_text) > 15000:
        report_text = report_text[:15000]
        trimmed = True

    prompt = f"""You are a clinical AI assistant with deep knowledge of biomedical sciences,
pathophysiology, and laboratory medicine. A patient or clinician has uploaded a medical report.
Analyze it with the precision of a biomedical engineer and the clarity of a patient advocate.

{detail_map[detail]}

Structure your response EXACTLY as follows:

## 🏥 Patient Overview
Brief demographics and reason for this report if mentioned.

## 🔬 Key Findings
List each finding. For every lab value include:
- The value found
- The normal reference range
- What it means physiologically (1 sentence, plain language)

## 🔴 Critical / Abnormal Values
Flag anything outside normal range. Use exactly:
🔴 HIGH: [value] — [what this means and why it matters]
🟡 BORDERLINE: [value] — [what to watch for]
✅ NORMAL: [value] — [brief reassurance]

## 🩺 Diagnosis / Clinical Impression
What the report suggests clinically. Always note this is AI interpretation only.

## 📋 Recommended Actions
Specific, actionable next steps. Always end with:
"→ Please discuss these findings with your doctor before taking any action."

## 💬 Plain Language Summary
Exactly 3 sentences. No jargon, no abbreviations. Imagine explaining to a worried patient.

## ⚠️ AI Limitations
One sentence about what this analysis cannot replace.

Medical Report:
{report_text}"""

    response = client.chat.completions.create(
        model="llama-3.3-70b-versatile",
        messages=[{"role": "user", "content": prompt}],
        temperature=0.3,
        max_tokens=2000,
    )
    return response.choices[0].message.content, trimmed


def generate_pdf(summary_text, filename):
    pdf = FPDF()
    pdf.add_page()
    pdf.set_font("Arial", "B", 16)
    pdf.cell(0, 10, "MedSumm AI - Medical Report Summary", ln=True)
    pdf.set_font("Arial", "", 9)
    pdf.cell(0, 8,
        f"Source: {filename}  |  "
        f"Generated: {datetime.now().strftime('%Y-%m-%d %H:%M')}  |  "
        f"Model: Llama 3.3 70B (Groq)", ln=True)
    pdf.set_font("Arial", "I", 8)
    pdf.cell(0, 6, "Created by Fatima Nasir | linkedin.com/in/fatima-nasir-bme", ln=True)
    pdf.ln(4)
    pdf.set_font("Arial", "", 11)
    clean = summary_text.encode("latin-1", "replace").decode("latin-1")
    pdf.multi_cell(0, 7, clean)
    pdf.ln(6)
    pdf.set_font("Arial", "I", 8)
    pdf.multi_cell(0, 5,
        "DISCLAIMER: This summary is AI-generated and is not a substitute "
        "for professional medical advice. Always consult a licensed healthcare provider.")
    return pdf.output(dest="S").encode("latin-1")


def render_summary(summary):
    for line in summary.split("\n"):
        upper = line.upper()
        if "🔴" in line or ("HIGH" in upper and "🔴" in line):
            st.markdown(f"<div class='flag-high'>{line}</div>", unsafe_allow_html=True)
        elif "🟡" in line or "BORDERLINE" in upper:
            st.markdown(f"<div class='flag-mid'>{line}</div>", unsafe_allow_html=True)
        elif "✅" in line:
            st.markdown(f"<div class='flag-ok'>{line}</div>", unsafe_allow_html=True)
        else:
            st.markdown(line)


def highlight_text(text, keyword):
    if not keyword:
        return text
    return re.sub(
        f"({re.escape(keyword)})", r"<mark>\1</mark>",
        text, flags=re.IGNORECASE
    )


def vitals_chart(vitals):
    labels = list(vitals.keys())
    values = list(vitals.values())
    colors = []
    for label, val in vitals.items():
        if label in NORMAL_RANGES:
            lo, hi = NORMAL_RANGES[label]
            colors.append("#ff4d6d" if (val < lo or val > hi) else t["primary"])
        else:
            colors.append(t["primary"])

    fig = go.Figure(go.Bar(
        x=labels, y=values,
        marker_color=colors,
        text=[str(v) for v in values],
        textposition="auto"
    ))

    for i, label in enumerate(labels):
        if label in NORMAL_RANGES:
            lo, hi = NORMAL_RANGES[label]
            fig.add_shape(type="line", x0=i-0.4, x1=i+0.4, y0=hi, y1=hi,
                          line=dict(color="yellow", width=2, dash="dot"))
            fig.add_shape(type="line", x0=i-0.4, x1=i+0.4, y0=lo, y1=lo,
                          line=dict(color="#2ee6a8", width=2, dash="dot"))

    fig.update_layout(
        height=320,
        margin=dict(t=40, b=20, l=20, r=20),
        paper_bgcolor="rgba(0,0,0,0)",
        plot_bgcolor="rgba(0,0,0,0)",
        font=dict(color="white"),
        showlegend=False,
        title=dict(
            text=(
                "Detected Vitals  |  "
                "<span style='color:yellow'>--- Upper normal</span>  "
                "<span style='color:#2ee6a8'>--- Lower normal</span>  "
                "<span style='color:#ff4d6d'>■ Abnormal</span>"
            ),
            font=dict(size=12), x=0
        )
    )
    return fig


# ── MAIN ──
uploaded_files = st.file_uploader(
    "Upload Medical Report(s)",
    type=["pdf", "docx", "txt"],
    accept_multiple_files=True
)

if uploaded_files:
    tabs = st.tabs([f.name for f in uploaded_files])

    for tab, uploaded_file in zip(tabs, uploaded_files):
        with tab:
            try:
                raw_text, page_count = extract_text(uploaded_file)
            except Exception as e:
                st.error(f"Could not read file: {e}")
                continue

            if not raw_text.strip():
                st.error("No readable text found. This may be a scanned or image-only document.")
                continue

            word_count = len(raw_text.split())
            vitals     = extract_vitals(raw_text)

            c1, c2, c3, c4 = st.columns(4)
            c1.markdown(f"<div class='metric-card'><b>{word_count:,}</b><span>Words</span></div>", unsafe_allow_html=True)
            c2.markdown(f"<div class='metric-card'><b>{page_count or '—'}</b><span>Pages</span></div>", unsafe_allow_html=True)
            c3.markdown(f"<div class='metric-card'><b>{len(vitals)}</b><span>Vitals Found</span></div>", unsafe_allow_html=True)
            c4.markdown(f"<div class='metric-card'><b>Llama 3.3</b><span>AI Model</span></div>", unsafe_allow_html=True)

            st.markdown("")

            if vitals:
                st.plotly_chart(vitals_chart(vitals), use_container_width=True)

            with st.expander("📄 View Extracted Text (with search)"):
                keyword = st.text_input("🔎 Highlight keyword", key=f"kw_{uploaded_file.name}")
                if keyword:
                    st.markdown(highlight_text(raw_text, keyword), unsafe_allow_html=True)
                else:
                    st.text_area("Raw Report Text", raw_text, height=200, key=f"raw_{uploaded_file.name}")

            if st.button("🔍 Analyze Report", type="primary", key=f"btn_{uploaded_file.name}"):
                with st.spinner("Llama 3.3 is reading your report..."):
                    start = time.time()
                    try:
                        summary, was_trimmed = summarize_report(raw_text, detail_level)
                        elapsed = round(time.time() - start, 2)

                        st.session_state.history.append({
                            "name": uploaded_file.name,
                            "time": datetime.now().strftime("%H:%M:%S")
                        })

                        if was_trimmed:
                            st.markdown(
                                "<div class='trim-warning'>⚠️ Report was trimmed to 15,000 characters. Core content preserved.</div>",
                                unsafe_allow_html=True
                            )

                        st.markdown("---")
                        head1, head2 = st.columns([5, 1])
                        head1.subheader("📋 AI Summary")
                        head2.caption(f"⏱️ {elapsed}s")

                        render_summary(summary)

                        st.markdown(
                            "<div class='confidence-box'>"
                            "🤖 <b>AI Confidence Note:</b> This summary is generated from text extraction only. "
                            "Handwritten reports, scanned images, or poor PDF quality may reduce accuracy. "
                            "Always cross-reference with the original document and consult your doctor."
                            "</div>",
                            unsafe_allow_html=True
                        )

                        st.markdown("")
                        dl1, dl2, dl3 = st.columns(3)
                        with dl1:
                            st.download_button(
                                "⬇️ Download as TXT",
                                data=summary,
                                file_name=f"summary_{uploaded_file.name}.txt",
                                mime="text/plain",
                                key=f"txt_{uploaded_file.name}"
                            )
                        with dl2:
                            try:
                                pdf_bytes = generate_pdf(summary, uploaded_file.name)
                                st.download_button(
                                    "⬇️ Download as PDF",
                                    data=pdf_bytes,
                                    file_name=f"summary_{uploaded_file.name}.pdf",
                                    mime="application/pdf",
                                    key=f"pdf_{uploaded_file.name}"
                                )
                            except Exception:
                                pass
                        with dl3:
                            st.download_button(
                                "⬇️ Download Raw Text",
                                data=raw_text,
                                file_name=f"raw_{uploaded_file.name}.txt",
                                mime="text/plain",
                                key=f"rawdl_{uploaded_file.name}"
                            )

                    except Exception as e:
                        st.error(f"Analysis failed: {e}")

else:
    st.info("👆 Upload one or more medical reports to get started.")

# ── FOOTER ──
st.markdown(f"""
<div class='footer-box'>
    <div class='footer-name'>Fatima Nasir</div>
    Biomedical Engineer &amp; AI Developer
    <br><br>
    <a href='https://www.linkedin.com/in/fatima-nasir-bme' target='_blank'>🔗 LinkedIn</a>
    &nbsp;·&nbsp;
    <a href='https://github.com/fatimaanasirofficial-stack/medsumm-ai' target='_blank'>💻 GitHub</a>
    <br><br>
    <span style='color:#374151;'>
        MedSumm AI &nbsp;·&nbsp; Powered by Groq (Llama 3.3 70B) &nbsp;·&nbsp; Not for clinical use
    </span>
</div>
""", unsafe_allow_html=True) 