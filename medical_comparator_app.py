import streamlit as st
import os
import io
import json
import pandas as pd
from anthropic import Anthropic
import openpyxl
from openpyxl.styles import Font, Alignment, PatternFill
from openpyxl.utils import get_column_letter
from reportlab.lib.pagesizes import letter, landscape
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib import colors

st.set_page_config(page_title="AI Medical Quotation & TOB Comparator", page_icon="🏥", layout="wide")

st.markdown("""
<style>
    .main-header { font-size: 24px; font-weight: 700; color: #0b2545; margin-bottom: 5px; }
    .sub-header { font-size: 14px; color: #6b7280; margin-bottom: 25px; }
    .stButton>button { width: 100%; background-color: #0d9488; color: white; font-weight: bold; border-radius: 8px; padding: 10px; }
    .stButton>button:hover { background-color: #0b7c72; color: white; }
</style>
""", unsafe_allow_html=True)

st.markdown('<div class="main-header">🏥 Secureme — Executive Medical Quotation &amp; TOB Comparator</div>', unsafe_allow_html=True)
st.markdown('<div class="sub-header">North Star Insurance Broker | Ultra-Fast AI Underwriting Engine</div>', unsafe_allow_html=True)

DEFAULT_KEY = os.environ.get("ANTHROPIC_AUTH_TOKEN", os.environ.get("ANTHROPIC_API_KEY", ""))
DEFAULT_URL = os.environ.get("ANTHROPIC_BASE_URL", "")
DEFAULT_MODEL = os.environ.get("ANTHROPIC_MODEL", "auto/claude-sonnet-5-5") # Use faster sonnet model by default for speed

def clean_secret(value: str) -> str:
    if not value: return ""
    v = value.strip().strip('"\'').strip()
    if v.lower().startswith("bearer "): v = v[7:].strip()
    return "".join(v.split())

def clean_url(value: str) -> str:
    if not value: return ""
    v = value.strip().strip('"\'').strip().rstrip("/")
    if not v.lower().startswith(("http://", "https://")):
        v = "https://" + v
    if "/v1" not in v.lower():
        v = v + "/v1"
    return v

with st.sidebar:
    st.header("🔑 API Configuration")
    api_key_input = st.text_input("API Key / Token", type="password", value=DEFAULT_KEY)
    base_url_input = st.text_input("Gateway Base URL", value=DEFAULT_URL)
    model_choice = st.text_input("Model Name (Fast Mode)", value="auto/claude-sonnet-5-5", help="Use Sonnet for fast 15-second comparison.")

    st.markdown("---")
    st.markdown("### ✍️ Broker Sign-Off Info")
    broker_rep = st.text_input("Broker Representative", value="Medical Underwriting Team")
    client_company = st.text_input("Client / Company Name", value="Valued Client LLC")

    st.markdown("---")
    if st.button("🔍 Test Connection"):
        key = clean_secret(api_key_input)
        url = clean_url(base_url_input)
        model = model_choice.strip() or "auto/claude-sonnet-5-5"
        if not key:
            st.error("API Key is empty.")
        else:
            try:
                client = Anthropic(api_key=key, base_url=url or None)
                client.messages.create(model=model, max_tokens=10, messages=[{"role": "user", "content": "Hi"}])
                st.success(f"✅ Connected successfully using model `{model}`!")
            except Exception as e:
                st.error(f"❌ Connection failed: {e}")

    st.markdown("---")
    st.markdown("### 📋 Instructions")
    st.markdown("1. Upload 2–6 Quotations &amp; TOB PDFs.\n2. Click **Run AI Comparison** (Optimized for ~15s speed).\n3. Download executive PDF &amp; Excel reports.")

uploaded_files = st.file_uploader(
    "Upload Insurer Quotation & Table of Benefits (PDFs)",
    type=["pdf", "docx", "txt"],
    accept_multiple_files=True
)

def extract_text(file_bytes, filename):
    name = filename.lower()
    if name.endswith(".pdf"):
        try:
            import pypdf
            reader = pypdf.PdfReader(io.BytesIO(file_bytes))
            # Extract first 4 pages of TOB/Quotes where summary & premiums live for lightning speed
            text = ""
            for i, page in enumerate(reader.pages[:4]):
                text += f"\n--- Page {i+1} ---\n" + (page.extract_text() or "")
            return text
        except Exception as e:
            return f"Error extracting PDF: {e}"
    if name.endswith(".docx"):
        try:
            import docx
            d = docx.Document(io.BytesIO(file_bytes))
            return "\n".join(p.text for p in d.paragraphs[:100])
        except Exception as e:
            return f"Error extracting DOCX: {e}"
    try:
        return file_bytes.decode("utf-8", errors="ignore")[:15000]
    except Exception as e:
        return f"Error reading text: {e}"

def generate_professional_pdf(comparison_df, members_df, num_options, rep_name, client_name):
    buffer = io.BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=landscape(letter),
                            rightMargin=15, leftMargin=15, topMargin=15, bottomMargin=15)
    elements = []
    styles = getSampleStyleSheet()

    title_style = ParagraphStyle(
        'MainTitle', parent=styles['Heading1'],
        fontSize=15, leading=18, alignment=1,
        textColor=colors.HexColor('#0b2545'), fontName='Helvetica-Bold', spaceAfter=2
    )
    broker_style = ParagraphStyle(
        'BrokerHeader', parent=styles['Normal'],
        fontSize=10, leading=13, alignment=0,
        textColor=colors.HexColor('#0d9488'), fontName='Helvetica-Bold'
    )
    client_style = ParagraphStyle(
        'ClientHeader', parent=styles['Normal'],
        fontSize=9, leading=12, alignment=2,
        textColor=colors.HexColor('#1f2937'), fontName='Helvetica-Bold'
    )
    section_title_style = ParagraphStyle(
        'SecTitle', parent=styles['Heading2'],
        fontSize=11, leading=14, textColor=colors.HexColor('#0b2545'),
        fontName='Helvetica-Bold', spaceBefore=10, spaceAfter=4
    )
    cell_style = ParagraphStyle('Cell', parent=styles['Normal'], fontSize=7.5, leading=9, textColor=colors.HexColor('#1f2937'))
    head_style = ParagraphStyle('Head', parent=styles['Normal'], fontSize=8, leading=10,
                                textColor=colors.white, fontName='Helvetica-Bold', alignment=1)
    sign_style = ParagraphStyle('Sign', parent=styles['Normal'], fontSize=8, leading=11, textColor=colors.HexColor('#374151'))

    broker_text = "<b>Secureme</b><br/><font size=8 color='#6b7280'>North Star Insurance Broker</font>"
    title_text = "<b>Executive Medical Quotation &amp; TOB Comparison</b>"
    client_text = f"<b>Prepared For:</b> {client_name}<br/><font size=8 color='#6b7280'>Date: {pd.Timestamp.today().strftime('%d %b %Y')}</font>"

    header_table = Table([[Paragraph(broker_text, broker_style), Paragraph(title_text, title_style), Paragraph(client_text, client_style)]], colWidths=[180, 380, 220])
    header_table.setStyle(TableStyle([('VALIGN', (0,0), (-1,-1), 'MIDDLE'), ('BOTTOMPADDING', (0,0), (-1,-1), 8)]))
    elements.append(header_table)
    elements.append(Spacer(1, 4))

    if members_df is not None and not members_df.empty:
        elements.append(Paragraph("<b>1. Insured Member Roster &amp; Census Details</b>", section_title_style))
        m_headers = list(members_df.columns)
        m_data = [[Paragraph(f"<b>{h}</b>", head_style) for h in m_headers]]
        for _, row in members_df.iterrows():
            m_data.append([Paragraph(str(v).replace("&", "&amp;").replace("<", "&lt;"), cell_style) for v in row])

        m_table = Table(m_data, colWidths=[150, 100, 100, 120, 100] if len(m_headers)==5 else None, repeatRows=1)
        m_table.setStyle(TableStyle([
            ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#13315c')),
            ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
            ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
            ('BOTTOMPADDING', (0, 0), (-1, -1), 3),
            ('TOPPADDING', (0, 0), (-1, -1), 3),
            ('GRID', (0, 0), (-1, -1), 0.5, colors.HexColor('#dbe3ec')),
            ('ROWBACKGROUNDS', (0, 1), (-1, -1), [colors.white, colors.HexColor('#f4f7fa')]),
        ]))
        elements.append(m_table)
        elements.append(Spacer(1, 8))

    elements.append(Paragraph("<b>2. Comparative Benefits &amp; Financial Summary (Best-Value Highlighted)</b>", section_title_style))
    headers = list(comparison_df.columns)
    headers[0] = "Details and benefits"

    comp_data = [[Paragraph(f"<b>{h}</b>", head_style) for h in headers]]

    for _, row in comparison_df.iterrows():
        row_cells = []
        param_name = str(row.iloc[0]).strip()
        is_grand_total = "GRAND TOTAL" in param_name.upper()

        for i, v in enumerate(row):
            val = str(v).replace("&", "&amp;").replace("<", "&lt;")
            if i == 0:
                row_cells.append(Paragraph(f"<b>{val}</b>", cell_style))
            else:
                if is_grand_total:
                    val = f"<font color='#047857'><b>★ {val}</b></font>"
                row_cells.append(Paragraph(val, cell_style))
        comp_data.append(row_cells)

    avail_width = 780
    col1_w = 180
    opt_w = max(90, (avail_width - col1_w) / max(1, num_options))
    col_widths = [col1_w] + [opt_w] * num_options

    c_table = Table(comp_data, colWidths=col_widths, repeatRows=1)
    c_table.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#0b2545')),
        ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
        ('VALIGN', (0, 0), (-1, -1), 'TOP'),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 4),
        ('TOPPADDING', (0, 0), (-1, -1), 4),
        ('LEFTPADDING', (0, 0), (-1, -1), 6),
        ('RIGHTPADDING', (0, 0), (-1, -1), 6),
        ('GRID', (0, 0), (-1, -1), 0.5, colors.HexColor('#dbe3ec')),
        ('ROWBACKGROUNDS', (0, 1), (-1, -1), [colors.white, colors.HexColor('#f4f7fa')]),
    ]))
    elements.append(c_table)
    elements.append(Spacer(1, 10))

    sign_block = [
        [
            Paragraph(f"<b>Prepared &amp; Reviewed By:</b><br/>{rep_name}<br/><b>Secureme (North Star Insurance Broker)</b>", sign_style),
            Paragraph("<b>Client Acknowledgment &amp; Approval:</b><br/>Name: ________________________<br/>Signature: ______________________", sign_style)
        ]
    ]
    sign_table = Table(sign_block, colWidths=[380, 400])
    sign_table.setStyle(TableStyle([
        ('BOX', (0,0), (-1,-1), 1, colors.HexColor('#0d9488')),
        ('BACKGROUND', (0,0), (-1,-1), colors.HexColor('#e6f5f3')),
        ('TOPPADDING', (0,0), (-1,-1), 8),
        ('BOTTOMPADDING', (0,0), (-1,-1), 8),
        ('LEFTPADDING', (0,0), (-1,-1), 10),
        ('RIGHTPADDING', (0,0), (-1,-1), 10),
    ]))
    elements.append(sign_table)

    doc.build(elements)
    buffer.seek(0)
    return buffer.getvalue()

def style_excel(writer, df_comp, df_members, num_options):
    ws = writer.sheets["Comparison Matrix"]
    ws.freeze_panes = "B2"
    for col_idx in range(1, len(df_comp.columns) + 1):
        cell = ws.cell(row=1, column=col_idx)
        cell.font = Font(bold=True, color="FFFFFF", size=11)
        cell.fill = PatternFill("solid", fgColor="0B2545")
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    for row in ws.iter_rows(min_row=2):
        for cell in row:
            cell.alignment = Alignment(vertical="top", wrap_text=True)
            cell.font = Font(size=10)
    ws.column_dimensions['A'].width = 32
    for col in range(2, len(df_comp.columns) + 1):
        ws.column_dimensions[get_column_letter(col)].width = 28

    if df_members is not None and not df_members.empty:
        ws_m = writer.sheets["Member Roster"]
        ws_m.freeze_panes = "A2"
        for col_idx in range(1, len(df_members.columns) + 1):
            cell = ws_m.cell(row=1, column=col_idx)
            cell.font = Font(bold=True, color="FFFFFF", size=11)
            cell.fill = PatternFill("solid", fgColor="13315c")
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        for row in ws_m.iter_rows(min_row=2):
            for cell in row:
                cell.alignment = Alignment(vertical="top", wrap_text=True)
                cell.font = Font(size=10)
        for col in range(1, len(df_members.columns) + 1):
            ws_m.column_dimensions[get_column_letter(col)].width = 25

if st.button("🚀 Run AI Comparison", disabled=not uploaded_files):
    key = clean_secret(api_key_input)
    url_clean = clean_url(base_url_input)
    model = model_choice.strip() or "auto/claude-sonnet-5-5"

    if not key:
        st.error("Please enter your API Key in the sidebar.")
    else:
        with st.spinner("⚡ Extracting & comparing documents (Optimized for ~15s speed)..."):
            try:
                texts = {}
                for file in uploaded_files:
                    texts[file.name] = extract_text(file.read(), file.name)

                document_contents = ""
                for fname, txt in texts.items():
                    document_contents += f"\n\n--- FILE: {fname} ---\n{txt}\n--- END FILE ---\n"

                num_opts = len(uploaded_files)

                prompt = f"""
You are an expert UAE medical insurance underwriter. Below is text from {num_opts} quotation / TOB files.

Perform a fast, side-by-side comparison. Return your response strictly as a JSON object with TWO keys: "member_roster" and "comparison_table".

1. "member_roster": An array of objects representing insured members:
   - "Member Name"
   - "Age / DOB"
   - "Gender"
   - "Category / Relation"
   - "Insurer / Option"

2. "comparison_table": An array of objects for side-by-side comparison, with keys "Parameter" and Option columns ("Option 1", "Option 2", ... matching upload order).
Include these exact parameter rows in this order:
- Insurer
- Plan Name
- Quotation Number
- Client Name
- Company / Employer Name
- Policy Period
- Medical Network
- Annual Limit
- Area of Coverage
- In-patient Services
- Out-patient Services
- Hospital Accommodation
- Consultation Deductible (Co-Pay)
- Diagnostics & Laboratory Services
- Medicine Limit & Copay
- Alternative Medicines
- Physiotherapy
- Psychiatric Treatment
- Mental Health
- Home Nursing
- Maternity Benefits
- New Born Baby Cover
- Routine Dental Benefits
- Routine Optical Benefits
- Pre-existing / Chronic Conditions
- Emergency Road Ambulance
- Vaccination for Children
- Repatriation
- Direct Access to Hospitals
- Claims Settlement Basis
- Actual Premium (excl. VAT)
- Basmah Fee
- ICP / Policy Fees
- VAT 5%
- GRAND TOTAL

Return ONLY valid JSON (no markdown fences, no prose) shaped like:
{{
  "member_roster": [
    {{"Member Name": "John Doe", "Age / DOB": "35 / 1991", "Gender": "M", "Category / Relation": "Principal", "Insurer / Option": "Option 1"}},
  ],
  "comparison_table": [
    {{"Parameter": "Insurer", "Option 1": "...", "Option 2": "..."}},
  ]
}}

DOCUMENTS:
{document_contents}
"""

                client = Anthropic(api_key=key, base_url=url_clean or None)
                response = client.messages.create(
                    model=model,
                    max_tokens=4000,
                    messages=[{"role": "user", "content": prompt}],
                )

                response_text = (response.content[0].text or "").strip()
                if "```json" in response_text:
                    response_text = response_text.split("```json", 1)[1].split("```", 1)[0].strip()
                elif "```json" in response_text:
                    response_text = response_text.split("```json", 1)[1].split("```", 1)[0].strip()
                elif "```" in response_text:
                    response_text = response_text.split("```", 1)[1].split("```", 1)[0].strip()

                start = response_text.find("{")
                end = response_text.rfind("}")
                parsed = json.loads(response_text[start:end + 1])

                rows_comp = parsed.get("comparison_table", [])
                rows_members = parsed.get("member_roster", [])

                if not rows_comp:
                    st.warning("AI returned no table. Raw response:")
                    st.text(response_text[:4000])
                else:
                    df_comp = pd.DataFrame(rows_comp)
                    df_members = pd.DataFrame(rows_members) if rows_members else None

                    st.success(f"✨ Comparison generated successfully in seconds via `{model}`!")

                    if df_members is not None and not df_members.empty:
                        st.subheader("👥 Insured Member Census & Roster")
                        st.dataframe(df_members, use_container_width=True, hide_index=True)

                    st.subheader("📊 Comparative Underwriting Matrix")
                    st.dataframe(df_comp, use_container_width=True, hide_index=True)

                    xbuf = io.BytesIO()
                    with pd.ExcelWriter(xbuf, engine="openpyxl") as writer:
                        df_comp_excel = df_comp.copy()
                        df_comp_excel = df_comp_excel.rename(columns={df_comp_excel.columns[0]: "Details and benefits"})
                        df_comp_excel.to_excel(writer, index=False, sheet_name="Comparison Matrix")

                        if df_members is not None and not df_members.empty:
                            df_members.to_excel(writer, index=False, sheet_name="Member Roster")

                        style_excel(writer, df_comp_excel, df_members, num_opts)

                        raw_rows = [(name, txt[:10000]) for name, txt in texts.items()]
                        pd.DataFrame(raw_rows, columns=["Source File", "Extracted Text"]).to_excel(
                            writer, index=False, sheet_name="Extracted Text")
                    xbuf.seek(0)

                    c1, c2 = st.columns(2)
                    with c1:
                        st.download_button(
                            "📥 Download Excel (.xlsx with Member Roster tab)",
                            data=xbuf.getvalue(),
                            file_name="Secureme_Medical_Quotation_Comparison.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        )
                    with c2:
                        st.download_button(
                            "📄 Download Executive PDF Report (.pdf)",
                            data=generate_professional_pdf(df_comp, df_members, num_opts, broker_rep, client_company),
                            file_name="Secureme_Medical_Quotation_Comparison.pdf",
                            mime="application/pdf",
                        )

            except Exception as e:
                st.error(f"Error during AI processing: {e}")
