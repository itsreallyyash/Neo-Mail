import streamlit as st
import pandas as pd
import re
import os
import tempfile
import base64
from datetime import datetime
from docx import Document
from docx.oxml.table import CT_Tbl
from docx.oxml.text.paragraph import CT_P
from docx.text.paragraph import Paragraph
from docx.table import Table
from jinja2 import Environment, BaseLoader
from weasyprint import HTML
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders
import io

# Set page config
st.set_page_config(
    page_title="Investment Email Generator",
    page_icon="📧",
    layout="wide",
    initial_sidebar_state="expanded"
)
def convert_template_to_html(template_content):
    """Convert plain text template with newlines to HTML with <br> tags"""
    return template_content.replace('\n', '<br>')
# Initialize session state
if 'email_template' not in st.session_state:
    st.session_state.email_template = """Dear {{client_names}},<br><br>

Please find below updated summary of payment details processed on {{processing_date}}.<br><br>

{{investment_tables}}<br><br>

Feel free to connect for any clarifications.<br><br>

Best regards,<br>
Investor Relations Team<br><br>

{{company_signature}}<br>
{{company_address}}<br>
E-mail: {{company_email}}"""

if 'pdf_template' not in st.session_state:
    st.session_state.pdf_template = """<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <title>Investment Update</title>
    <link href="https://fonts.googleapis.com/css2?family=Montserrat:wght@600;700&family=Inter:wght@400;600&display=swap" rel="stylesheet">
    <style>
        @page {
            size: A4;
            margin: 32px 0 32px 0;
            @bottom-center {
                content: "© 2025 Neo Wealth | Private & Confidential";
                color: #fff;
                background: #2d204c;
                font-family: 'Montserrat', Arial, sans-serif;
                font-size: 12px;
                padding-top: 6px;
                padding-bottom: 6px;
                width: 100%;
                text-align: center;
            }
        }
        body {
            font-family: 'Inter', Arial, sans-serif;
            background: #fff;
            margin: 0;
            padding: 0;
        }
        .main-container {
            width: 680px;
            margin: 0 auto;
            background: #fff;
            padding: 36px 0 20px 0;
            box-sizing: border-box;
        }
        .header {
            display: flex;
            align-items: flex-start;
            justify-content: space-between;
            margin-bottom: 24px;
        }
        .logo-block {
            display: flex;
            flex-direction: column;
            align-items: flex-start;
        }
        .logo {
            background: #232156;
            color: #fff;
            font-family: 'Montserrat', sans-serif;
            font-weight: 700;
            font-size: 26px;
            border-radius: 100px;
            width: 72px;
            height: 72px;
            display: flex;
            align-items: center;
            justify-content: center;
            margin-bottom: 8px;
        }
        .subtitle {
            font-size: 15px;
            color: #494949;
            margin-bottom: 4px;
        }
        .deal-title {
            font-size: 22px;
            font-family: 'Montserrat', sans-serif;
            font-weight: 700;
            color: #6a4c93;
            margin-bottom: 2px;
        }
        .company-name {
            font-size: 23px;
            font-family: 'Montserrat', sans-serif;
            font-weight: 700;
            color: #2d2d2d;
            margin-bottom: 0;
        }
        .decor-flower {
            position: absolute;
            top: 30px;
            right: 60px;
            opacity: 0.13;
            width: 160px;
        }
        .flex-columns {
            display: flex;
            flex-direction: row;
            gap: 24px;
            margin-top: 12px;
        }
        .left-col {
            flex: 2;
        }
        .right-col {
            flex: 1.4;
            margin-top: 7px;
        }
        .section-block {
            margin-bottom: 18px;
        }
        .section-title {
            background: #232156;
            color: #fff;
            font-family: 'Montserrat', sans-serif;
            font-weight: 700;
            font-size: 15px;
            padding: 6px 18px 6px 13px;
            border-radius: 5px 5px 0 0;
            display: inline-block;
            margin-bottom: 0;
        }
        .section-content {
            background: #f6f6ff;
            color: #2d2d2d;
            font-size: 14px;
            padding: 16px 18px 16px 14px;
            border-radius: 0 0 9px 9px;
            margin-bottom: 0;
            font-family: 'Inter', Arial, sans-serif;
        }
        .recent-updates-list {
            list-style: none;
            margin: 0;
            padding: 0;
        }
        .recent-updates-list li {
            position: relative;
            padding-left: 24px;
            margin-bottom: 9px;
            font-size: 14px;
        }
        .recent-updates-list li:before {
            content: "";
            position: absolute;
            left: 0;
            top: 7px;
            width: 7px;
            height: 7px;
            background: #6a4c93;
            border-radius: 100%;
        }
        .summary-box {
            background: #f5f1fb;
            border-radius: 11px;
            padding: 18px 22px 16px 22px;
            font-size: 16px;
            box-shadow: 0 2px 7px rgba(106,76,147,0.10);
        }
        .summary-label {
            font-size: 12.5px;
            color: #6a4c93;
            font-family: 'Montserrat', sans-serif;
            font-weight: 600;
            margin-top: 17px;
        }
        .summary-value {
            font-size: 16.5px;
            color: #2d2d2d;
            font-family: 'Montserrat', sans-serif;
            font-weight: 700;
            margin-bottom: 0;
        }
        .summary-divider {
            border: none;
            border-top: 1px dashed #cdc5e3;
            margin: 8px 0 8px 0;
        }
        /* Disclaimer Page Styling */
        .disclaimer-container {
            width: 680px;
            margin: 0 auto;
            padding: 40px 0 0 0;
            min-height: 1000px;
            position: relative;
            box-sizing: border-box;
        }
        .disclaimer-title {
            color: #6a4c93;
            font-size: 19px;
            font-family: Montserrat,Arial,sans-serif;
            font-weight: 600;
            border-left: 3.5px solid #6a4c93;
            padding-left: 9px;
            margin-bottom: 32px;
            margin-top: 70px;
        }
        .disclaimer-text {
            font-size: 15px;
            color: #555;
            margin-bottom: 32px;
            margin-top: 12px;
            line-height: 1.6;
        }
        .contact-title {
            font-size: 15px;
            color: #555;
            font-style: italic;
            margin-bottom: 42px;
            margin-top: 44px;
        }
        .contact-logo-row {
            display: flex;
            align-items: center;
            margin-bottom: 22px;
        }
        .contact-logo-img {
            width: 62px;
            margin-right: 24px;
        }
        .contact-logo-name {
            color: #232156;
            font-family: Montserrat,Arial,sans-serif;
            font-weight: 700;
            font-size: 18px;
        }
        .contact-address {
            font-size: 14.5px;
            color: #333;
            margin-bottom: 22px;
        }
        .disclaimer-flower {
            position: absolute;
            top: 64px;
            right: 80px;
            opacity: 0.11;
            width: 180px;
        }
    </style>
</head>
<body>
    <!-- PAGE 1 -->
    <div class="main-container">
        <div class="header">
            <div class="logo-block">
                <div class="logo">neo</div>
                <span style="color:#232156;font-size:13px;font-family:Montserrat,Arial,sans-serif;font-weight:600;margin-bottom:2px;">Do Good.</span>
            </div>
            {% if decor_flower %}
            <img src="{{ decor_flower }}" alt="" class="decor-flower"/>
            {% endif %}
        </div>

        <div class="subtitle">Investment Update</div>
        <div class="deal-title">Deal 1:</div>
        <div class="company-name">{{ company_name }}</div>

        <div class="flex-columns">
            <!-- Left Column -->
            <div class="left-col">
                <div class="section-block">
                    <div class="section-title">| Borrower Profile</div>
                    <div class="section-content">
                        {{ borrower_profile | safe }}
                    </div>
                </div>
                <div class="section-block">
                    <div class="section-title">| Recent Updates</div>
                    <div class="section-content">
                        <ul class="recent-updates-list">
                        {% for update in recent_updates %}
                            <li>{{ update }}</li>
                        {% endfor %}
                        </ul>
                    </div>
                </div>
            </div>

            <!-- Right Column: Investment Summary -->
            <div class="right-col">
                <div class="section-title" style="margin-bottom:0">| Investment Summary</div>
                <div class="summary-box">
                    <div class="summary-label">Instrument</div>
                    <div class="summary-value">{{ investment_summary["Instrument"] }}</div>
                    <hr class="summary-divider"/>

                    <div class="summary-label">IRR</div>
                    <div class="summary-value">{{ investment_summary["IRR"] }}</div>
                    <hr class="summary-divider"/>

                    <div class="summary-label">Date of Investment</div>
                    <div class="summary-value">{{ investment_summary["Date of Investment"] }}</div>
                    <hr class="summary-divider"/>

                    <div class="summary-label">Tenure</div>
                    <div class="summary-value">{{ investment_summary["Tenure"] }}</div>
                    <hr class="summary-divider"/>

                    <div class="summary-label">Collateral Description</div>
                    <div class="summary-value" style="font-size:13.2px;line-height:1.33;">
                        {{ investment_summary["Collateral Description"] }}
                    </div>
                    <hr class="summary-divider"/>

                    <div class="summary-label">Collateral Cover</div>
                    <div class="summary-value">{{ investment_summary["Collateral Cover"] }}</div>
                </div>
            </div>
        </div>
    </div>

    <!-- PAGE BREAK -->
    <div style="page-break-before: always"></div>

    <!-- PAGE 2: DISCLAIMER & CONTACT -->
    <div class="disclaimer-container">
        {% if decor_flower %}
        <img src="{{ decor_flower }}" alt="" class="disclaimer-flower"/>
        {% endif %}

        <div class="disclaimer-title">Disclaimer</div>
        <div class="disclaimer-text">
            This document is intended only for the personal use of the prospective investors/contributors (herein after referred to as the Clients) to whom it is addressed or delivered and must not be reproduced or redistributed in any form to any other person without any prior written consent of Neo Wealth Management (NWM). The document does not purport to be all-inclusive, nor does it contain exhaustive information which a prospective investor may desire for its decision making. The document is neither approved, certified nor its contents is verified by SEBI.<br><br>
            The contents of this document are provisional and may be subject to change at the discretion of NWM. NWM reserves the right (but is not required) to correct any errors or omissions on the document. In the preparation of the material contained in this website/email/document, NWM has used information, calculation, cashflows that is publicly available, certain research reports including information developed in-house applying valid assumptions. NWM warrants that the contents of this document are true to the best of its knowledge, however, assume no liability for the relevance, accuracy, or completeness of the contents herein.<br>
            For the entire terms and conditions, please refer to: https://www.neo-group.in/
        </div>

        <div class="contact-title">
            For any queries, please contact:
        </div>
        <div class="contact-logo-row">
            {% if contact_logo %}
            <img src="{{ contact_logo }}" class="contact-logo-img" />
            {% endif %}
            <span class="contact-logo-name">Do Good.</span>
        </div>
        <div class="contact-address">
            903, B-Wing, 9th Floor, Marathon FutureX, Mafatlal Mills Compound, N.M. Joshi Marg,<br>
            Lower Parel Mumbai 400013 IN
        </div>
    </div>
</body>
</html>
"""

if 'deals_data' not in st.session_state:
    st.session_state.deals_data = {}

if 'csv_data' not in st.session_state:
    st.session_state.csv_data = None

if 'header_html' not in st.session_state:
    st.session_state.header_html = ""

if 'footer_html' not in st.session_state:
    st.session_state.footer_html = ""

if 'company_info' not in st.session_state:
    st.session_state.company_info = {
        'company_signature': 'Neo Wealth',
        'company_email': 'ir@neoassetmanagement.com',
        'company_address': '123, Example Address, Mumbai, India'
    }

# ─── HELPER FUNCTIONS ────────────────────────────────────────────────────────

def iter_block_items(doc):
    """Yield each Paragraph or Table in the document in true Word order."""
    for child in doc.element.body:
        if isinstance(child, CT_P):
            yield Paragraph(child, doc)
        elif isinstance(child, CT_Tbl):
            yield Table(child, doc)

def parse_word_doc(uploaded_file) -> dict:
    """Parse Word document and extract deal information using exact logic from original code."""
    # Save uploaded file temporarily
    with tempfile.NamedTemporaryFile(delete=False, suffix='.docx') as tmp_file:
        tmp_file.write(uploaded_file.getvalue())
        tmp_path = tmp_file.name
    
    try:
        doc = Document(tmp_path)
        deals = {}
        current = None
        section = None
        expecting_table = False

        for block in iter_block_items(doc):
            # — Paragraphs —
            if isinstance(block, Paragraph):
                text = block.text.strip()
                if not text:
                    continue

                m = re.match(r"Deal\s*\d+\s*[:\-]?\s*(.+)", text, flags=re.I)
                if m:
                    cname = m.group(1).rstrip(":").strip()
                    current = {
                        "company_name": cname,
                        "borrower_profile": [],
                        "investment_summary": {},
                        "recent_updates": []
                    }
                    deals[cname] = current
                    section = None
                    expecting_table = False
                    continue

                low = text.lower()
                if "borrower profile" in low:
                    section = "borrower_profile"
                    continue
                if "investment summary" in low:
                    expecting_table = True
                    section = None
                    continue
                if "recent updates" in low:
                    section = "recent_updates"
                    continue

                if current:
                    if section == "borrower_profile":
                        current["borrower_profile"].append(text)
                    elif section == "recent_updates":
                        bullet = re.sub(r"^[•\-\*\s]+", "", text)
                        current["recent_updates"].append(bullet)

            # — Tables —
            elif isinstance(block, Table) and expecting_table and current:
                for row in block.rows:
                    if len(row.cells) >= 2:
                        k = row.cells[0].text.strip()
                        v = row.cells[1].text.strip()
                        if k:
                            current["investment_summary"][k] = v
                expecting_table = False

        # Convert borrower_profile lists to HTML formatted strings
        for d in deals.values():
            d["borrower_profile"] = "<br><br>".join(d["borrower_profile"])
        
        return deals
    finally:
        os.unlink(tmp_path)

def create_investment_table_html(row):
    """Create HTML table for investment data using exact logic from original code."""
    return f"""
    <div style="margin-bottom: 30px;">
      <table style="width:100%;border:1px solid #666;border-collapse:collapse;font-family:Arial,sans-serif">
        <tr style="background:#232156;color:#fff">
          <td colspan="2" style="padding:12px;font-weight:bold">{row['Security Name'].strip()}</td>
        </tr>
        <tr style="background:#f5f5f5">
          <td style="padding:8px;border:1px solid #666">No. of NCDs (nos.)</td>
          <td style="padding:8px;border:1px solid #666;text-align:right">{row['No. of NCDs (nos.)']}</td>
        </tr>
        <tr>
          <td style="padding:8px;border:1px solid #666">Face Value</td>
          <td style="padding:8px;border:1px solid #666;text-align:right">{row['Face Value']}</td>
        </tr>
        <tr style="background:#f5f5f5">
          <td style="padding:8px;border:1px solid #666">Opening Principal</td>
          <td style="padding:8px;border:1px solid #666;text-align:right">{row['Opening Principal Outstanding as on XX date']}</td>
        </tr>
        <tr>
          <td style="padding:8px;border:1px solid #666">Principal Repaid</td>
          <td style="padding:8px;border:1px solid #666;text-align:right">{row['Principal repaid on "Period"']}</td>
        </tr>
        <tr style="background:#f5f5f5">
          <td style="padding:8px;border:1px solid #666">Closing Principal</td>
          <td style="padding:8px;border:1px solid #666;text-align:right">{row['Closing Principal Outstanding as on XX date']}</td>
        </tr>
        <tr>
          <td style="padding:8px;border:1px solid #666">Net Interest Paid</td>
          <td style="padding:8px;border:1px solid #666;text-align:right">{row['Net Interest paid for the "Period"']}</td>
        </tr>
      </table>
    </div>
    """

def generate_investor_pdf(company_data, template_html):
    """Generate PDF using exact logic from original code."""
    try:
        env = Environment(loader=BaseLoader())
        template = env.from_string(template_html)
        html_content = template.render(**company_data)
        
        # Create PDF in memory
        pdf_buffer = io.BytesIO()
        HTML(string=html_content, base_url=".").write_pdf(pdf_buffer)
        pdf_buffer.seek(0)
        
        return base64.b64encode(pdf_buffer.read()).decode()
    except Exception as e:
        st.error(f"Error generating PDF: {str(e)}")
        return None

def test_email_connection(smtp_config):
    """Test SMTP connection using exact logic from original code."""
    try:
        server = smtplib.SMTP(smtp_config['server'], smtp_config['port'])
        server.starttls()
        server.login(smtp_config['username'], smtp_config['password'])
        server.quit()
        return True, "Connection successful!"
    except Exception as e:
        return False, f"Connection failed: {str(e)}"

def send_email_with_pdfs(to_email, subject, html_content, pdf_data, smtp_config):
    """Send email with PDF attachments using exact logic from original code."""
    try:
        msg = MIMEMultipart()
        msg['From'] = smtp_config['from_email']
        msg['To'] = to_email
        msg['Subject'] = subject
        
        msg.attach(MIMEText(html_content, 'html'))
        
        # Attach PDFs
        for pdf_name, pdf_b64_content in pdf_data.items():
            part = MIMEBase('application', 'pdf')
            part.set_payload(base64.b64decode(pdf_b64_content))
            encoders.encode_base64(part)
            part.add_header('Content-Disposition', f'attachment; filename="{pdf_name}"')
            msg.attach(part)
        
        server = smtplib.SMTP(smtp_config['server'], smtp_config['port'])
        server.starttls()
        server.login(smtp_config['username'], smtp_config['password'])
        server.sendmail(smtp_config['from_email'], to_email, msg.as_string())
        server.quit()
        
        return True, "Email sent successfully!"
    except Exception as e:
        return False, f"Failed to send email: {str(e)}"

# ─── STREAMLIT APP ────────────────────────────────────────────────────────────

def main():
    st.title("📧 Investment Email Generator")
    st.markdown("*Exact functionality from your original code with Streamlit interface*")
    st.markdown("---")
    
    # Sidebar navigation
    st.sidebar.title("Navigation")
    page = st.sidebar.selectbox(
        "Choose a section:",
        ["🔧 Configuration", "✏️ Template Editor", "🧪 Test Mode", "📤 Send Mode"]
    )
    
    if page == "🔧 Configuration":
        configuration_page()
    elif page == "✏️ Template Editor":
        template_editor_page()
    elif page == "🧪 Test Mode":
        test_mode_page()
    elif page == "📤 Send Mode":
        send_mode_page()

def configuration_page():
    st.header("🔧 Configuration")
    
    # File uploads section
    col1, col2 = st.columns(2)
    
    with col1:
        st.subheader("📄 Upload Word Document")
        word_file = st.file_uploader(
            "Upload Word document with deal information",
            type=['docx'],
            help="Upload the Word document containing deal information with sections: Deal X, Borrower Profile, Investment Summary, Recent Updates"
        )
        
        if word_file is not None:
            with st.spinner("Parsing Word document..."):
                try:
                    deals_data = parse_word_doc(word_file)
                    st.session_state.deals_data = deals_data
                    st.success(f"✅ Parsed {len(deals_data)} deals from document")
                    
                    # Show parsed deals
                    with st.expander("View Parsed Deals"):
                        for deal_name, deal_info in deals_data.items():
                            st.write(f"**{deal_name}**")
                            st.write(f"- Investment Summary: {len(deal_info['investment_summary'])} items")
                            st.write(f"- Recent Updates: {len(deal_info['recent_updates'])} items")
                            st.write(f"- Borrower Profile: {'Available' if deal_info['borrower_profile'] else 'Empty'}")
                            
                            # Show sample data
                            if deal_info['investment_summary']:
                                st.write("Sample Investment Summary:")
                                for k, v in list(deal_info['investment_summary'].items())[:3]:
                                    st.write(f"  • {k}: {v}")
                except Exception as e:
                    st.error(f"Error parsing Word document: {str(e)}")
                    st.error("Please ensure your document follows the format: Deal X: [Company Name], Borrower Profile, Investment Summary (table), Recent Updates")
    
    with col2:
        st.subheader("📊 Upload CSV Data")
        csv_file = st.file_uploader(
            "Upload CSV with investment data",
            type=['csv'],
            help="Upload the CSV file containing columns: I_email, Client Name/ Buyer Name, Security Name, No. of NCDs (nos.), Face Value, etc."
        )
        
        if csv_file is not None:
            try:
                df = pd.read_csv(csv_file)
                df.columns = df.columns.str.strip()  # Clean column names
                st.session_state.csv_data = df
                st.success(f"✅ Loaded CSV with {len(df)} records")
                
                # Show data preview
                with st.expander("View CSV Data Preview"):
                    st.dataframe(df.head())
                    st.write("**Columns found:**")
                    st.write(list(df.columns))
                    
                # Validate required columns
                required_cols = ['I_email', 'Security Name', 'Client Name/ Buyer Name']
                missing_cols = [col for col in required_cols if col not in df.columns]
                if missing_cols:
                    st.warning(f"⚠️ Missing required columns: {missing_cols}")
                else:
                    st.success("✅ All required columns found")
                    
                # Show recipient count
                if 'I_email' in df.columns:
                    unique_emails = df['I_email'].dropna().nunique()
                    st.info(f"📧 Found {unique_emails} unique recipient emails")
                    
                    # Show securities mapping
                    security_counts = df.groupby('Security Name').size().reset_index(name='count')
                    st.write("**Securities in CSV:**")
                    st.dataframe(security_counts)
                    
            except Exception as e:
                st.error(f"Error loading CSV: {str(e)}")
    
    # Header/Footer HTML section
    st.subheader("🎨 Email Template Components")
    
    col1, col2 = st.columns(2)
    
    with col1:
        st.write("**Header HTML**")
        header_file = st.file_uploader(
            "Upload header HTML file",
            type=['html'],
            key="header_upload",
            help="Upload the header HTML file (body content will be extracted)"
        )
        if header_file:
            header_content = header_file.read().decode('utf-8')
            # Extract body content using regex like original code
            header_match = re.search(r'<body[^>]*>(.*?)</body>', header_content, flags=re.S|re.I)
            st.session_state.header_html = header_match.group(1) if header_match else header_content
            st.success("✅ Header HTML loaded")
            
            with st.expander("Preview Header"):
                st.components.v1.html(f"<div style='border:1px solid #ddd; padding:10px;'>{st.session_state.header_html}</div>", height=150)
    
    with col2:
        st.write("**Footer HTML**")
        footer_file = st.file_uploader(
            "Upload footer HTML file",
            type=['html'],
            key="footer_upload",
            help="Upload the footer HTML file (body content will be extracted)"
        )
        if footer_file:
            footer_content = footer_file.read().decode('utf-8')
            # Extract body content using regex like original code
            footer_match = re.search(r'<body[^>]*>(.*?)</body>', footer_content, flags=re.S|re.I)
            st.session_state.footer_html = footer_match.group(1) if footer_match else footer_content
            st.success("✅ Footer HTML loaded")
            
            with st.expander("Preview Footer"):
                st.components.v1.html(f"<div style='border:1px solid #ddd; padding:10px;'>{st.session_state.footer_html}</div>", height=150)
    
    # Company information
    st.subheader("🏢 Company Information")
    
    col1, col2 = st.columns(2)
    with col1:
        company_signature = st.text_input("Company Name", value=st.session_state.company_info['company_signature'])
        company_email = st.text_input("Company Email", value=st.session_state.company_info['company_email'])
    
    with col2:
        company_address = st.text_area("Company Address", value=st.session_state.company_info['company_address'])
    
    # Save button
    if st.button("💾 Save Company Information"):
        st.session_state.company_info = {
            'company_signature': company_signature,
            'company_email': company_email,
            'company_address': company_address
        }
        st.success("✅ Company information saved!")

def template_editor_page():
    st.header("✏️ Template Editor")
    
    # Template type selector
    template_type = st.selectbox("Select Template to Edit:", ["Email Template", "PDF Template"])
    
    if template_type == "Email Template":
        st.subheader("📧 Email Template")
        st.info("Available variables: {client_names}, {processing_date}, {investment_tables}, {company_signature}, {company_address}, {company_email}")
        
        template_text = st.text_area(
            "Edit email template:",
            value=st.session_state.email_template,
            height=400,
            help="Use the variables above to customize your email template. {investment_tables} will be replaced with the formatted tables from CSV data."
        )
        
        col1, col2 = st.columns([1, 4])
        with col1:
            if st.button("💾 Save Email Template"):
                st.session_state.email_template = template_text
                st.success("✅ Email template saved!")
        
        with col2:
            if st.button("🔄 Reset to Default"):
                st.session_state.email_template = """Dear {client_names},

Please find below updated summary of payment details processed on {processing_date}.

{investment_tables}

Feel free to connect for any clarifications.

Best regards,
Investor Relations Team

{company_signature}
{company_address}
E-mail: {company_email}"""
                st.success("✅ Template reset to default!")
                st.experimental_rerun()
    
    else:  # PDF Template
        st.subheader("📄 PDF Template")
        st.info("Available variables: {company_name}, {investment_summary}, {borrower_profile}, {recent_updates}, {company_signature}, {company_address}, {company_email}, {processing_date}")
        
        template_text = st.text_area(
            "Edit PDF template (HTML):",
            value=st.session_state.pdf_template,
            height=600,
            help="This is an HTML template for generating PDFs. Use Jinja2 syntax for variables and loops."
        )
        
        col1, col2 = st.columns([1, 4])
        with col1:
            if st.button("💾 Save PDF Template"):
                st.session_state.pdf_template = template_text
                st.success("✅ PDF template saved!")
        
        with col2:
            if st.button("🔄 Reset to Default"):
                # Reset to default template (defined in initialization)
                st.session_state.pdf_template = """<!DOCTYPE html>
<html>
<head>
    <meta charset="utf-8">
    <style>
        body { 
            font-family: Arial, sans-serif; 
            margin: 20px; 
            line-height: 1.4; 
        }
        .header { 
            background: #232156; 
            color: white; 
            padding: 15px; 
            text-align: center;
            margin-bottom: 20px;
        }
        .section { 
            margin: 20px 0; 
            page-break-inside: avoid;
        }
        .summary-table { 
            width: 100%; 
            border-collapse: collapse; 
            margin: 10px 0;
        }
        .summary-table th, .summary-table td { 
            border: 1px solid #ddd; 
            padding: 8px; 
            text-align: left; 
        }
        .summary-table th { 
            background-color: #f2f2f2; 
            font-weight: bold;
        }
        .profile-section {
            background-color: #f9f9f9;
            padding: 15px;
            border-left: 4px solid #232156;
            margin: 15px 0;
        }
        .updates-list {
            padding-left: 20px;
        }
        .footer {
            margin-top: 30px;
            padding-top: 20px;
            border-top: 1px solid #ddd;
            font-size: 12px;
            color: #666;
        }
    </style>
</head>
<body>
    <div class="header">
        <h1>{{ company_name }}</h1>
        <p>Investment Report - {{ processing_date }}</p>
    </div>
    
    <div class="section">
        <h2>Investment Summary</h2>
        <table class="summary-table">
            {% for key, value in investment_summary.items() %}
            <tr>
                <th>{{ key }}</th>
                <td>{{ value }}</td>
            </tr>
            {% endfor %}
        </table>
    </div>
    
    <div class="section">
        <h2>Borrower Profile</h2>
        <div class="profile-section">
            {{ borrower_profile|safe }}
        </div>
    </div>
    
    {% if recent_updates %}
    <div class="section">
        <h2>Recent Updates</h2>
        <ul class="updates-list">
            {% for update in recent_updates %}
            <li>{{ update }}</li>
            {% endfor %}
        </ul>
    </div>
    {% endif %}
    
    <div class="footer">
        <p><strong>{{ company_signature }}</strong></p>
        <p>{{ company_address }}</p>
        <p>Email: {{ company_email }}</p>
        <p>Processing Date: {{ processing_date }}</p>
    </div>
</body>
</html>"""
                st.success("✅ PDF template reset to default!")
                st.experimental_rerun()

def test_mode_page():
    st.header("🧪 Test Mode")
    
    # Check if data is loaded
    if st.session_state.csv_data is None or not st.session_state.deals_data:
        st.warning("⚠️ Please upload Word document and CSV data in the Configuration section first.")
        return
    
    df = st.session_state.csv_data
    deals_data = st.session_state.deals_data
    
    # Show recipients list
    st.subheader("📋 Recipients Overview")
    if 'I_email' in df.columns:
        recipients_summary = df.groupby('I_email').agg({
            'Client Name/ Buyer Name': lambda x: ', '.join(x.unique()),
            'Security Name': lambda x: ', '.join(x.unique()),
            'Security Name': 'count'
        }).rename(columns={'Security Name': 'Number of Securities'}).reset_index()
        
        st.dataframe(recipients_summary, use_container_width=True)
        st.info(f"Total recipients: {len(recipients_summary)}")
    else:
        st.error("No 'I_email' column found in CSV data")
        return
    
    # Email preview section
    st.subheader("📧 Email Preview")
    
    # Select recipient for preview
    if len(recipients_summary) > 0:
        selected_email = st.selectbox(
            "Select recipient for preview:",
            recipients_summary['I_email'].tolist()
        )
        
        if selected_email:
            # Get data for selected recipient
            recipient_data = df[df['I_email'] == selected_email]
            
            # Generate investment tables HTML using exact original logic
            tables_html = ""
            for _, row in recipient_data.iterrows():
                tables_html += create_investment_table_html(row)
            
            # Prepare template variables
            client_names = ', '.join(recipient_data['Client Name/ Buyer Name'].unique())
            processing_date = datetime.now().strftime('%d %b %Y')
            
            template_vars = {
                'client_names': client_names,
                'processing_date': processing_date,
                'investment_tables': tables_html,
                **st.session_state.company_info
            }
            
            # Generate email content
            env = Environment(loader=BaseLoader())
            template = env.from_string(st.session_state.email_template)
            email_body = template.render(**template_vars)

            # Create full email HTML using exact original logic (600px width)
            full_email_html = f"""<!DOCTYPE html>
<html><head><meta charset="utf-8"></head><body style="margin:0;padding:0;background:#fff;">
<table width="600" align="center" cellpadding="0" cellspacing="0" border="0"
       style="margin:0 auto;background-color:#fff;">
  <tr><td style="padding:0;">{st.session_state.header_html}</td></tr>
  <tr><td style="padding:20px 0 20px 0;">{email_body}</td></tr>
  <tr><td style="padding:0;">{st.session_state.footer_html}</td></tr>
</table>
</body></html>"""
            
            # Display email preview
            st.components.v1.html(full_email_html, height=800, scrolling=True)
            
            # PDF Preview section
            st.subheader("📄 PDF Previews")
            
            # Get securities for this recipient
            securities = recipient_data['Security Name'].unique()
            
            for security in securities:
                security_clean = security.strip()
                if security_clean in deals_data:
                    st.write(f"**PDF Preview for: {security_clean}**")
                    
                    # Prepare PDF data using exact original logic
                    pdf_data = {
                        **deals_data[security_clean],
                        **st.session_state.company_info,
                        'processing_date': processing_date
                    }
                    
                    # Generate PDF preview
                    pdf_b64 = generate_investor_pdf(pdf_data, st.session_state.pdf_template)
                    
                    if pdf_b64:
                        # Create download button
                        client_name = recipient_data['Client Name/ Buyer Name'].iloc[0]
                        pdf_filename = f"{client_name}_{security_clean}".replace(" ", "_") + ".pdf"
                        
                        st.download_button(
                            label=f"📥 Download {pdf_filename}",
                            data=base64.b64decode(pdf_b64),
                            file_name=pdf_filename,
                            mime="application/pdf"
                        )
                        
                        # Show PDF preview
                        st.markdown(
                            f'<iframe src="data:application/pdf;base64,{pdf_b64}" width="100%" height="500"></iframe>',
                            unsafe_allow_html=True
                        )
                    else:
                        st.error(f"Failed to generate PDF for {security_clean}")
                else:
                    st.warning(f"⚠️ No deal data found for security: {security_clean}")
                    st.write("Available securities in Word document:")
                    st.write(list(deals_data.keys()))

def send_mode_page():
    st.header("📤 Send Mode")
    
    # Check if data is loaded
    if st.session_state.csv_data is None or not st.session_state.deals_data:
        st.warning("⚠️ Please upload Word document and CSV data in the Configuration section first.")
        return
    
    # SMTP Configuration
    st.subheader("📧 Email Configuration")
    
    col1, col2 = st.columns(2)
    
    with col1:
        email_provider = st.selectbox("Email Provider", ["Gmail", "Outlook"])
        email_address = st.text_input("Email Address")
    
    with col2:
        if email_provider == "Gmail":
            st.info("💡 Use App Password for Gmail (not your regular password)")
            smtp_config = {
                'server': 'smtp.gmail.com',
                'port': 587
            }
        else:
            smtp_config = {
                'server': 'smtp-mail.outlook.com',
                'port': 587
            }
        
        email_password = st.text_input("Password", type="password")
    
    smtp_config.update({
        'username': email_address,
        'password': email_password,
        'from_email': email_address
    })
    
    # Test connection
    col1, col2 = st.columns([1, 3])
    with col1:
        if st.button("🔍 Test Email Connection"):
            if email_address and email_password:
                with st.spinner("Testing connection..."):
                    success, message = test_email_connection(smtp_config)
                    if success:
                        st.success(f"✅ {message}")
                    else:
                        st.error(f"❌ {message}")
            else:
                st.error("Please enter email address and password")
    
    st.markdown("---")
    
    # Send Options
    st.subheader("🚀 Send Options")
    
    df = st.session_state.csv_data
    total_recipients = df['I_email'].dropna().nunique() if 'I_email' in df.columns else 0
    
    col1, col2 = st.columns(2)
    
    with col1:
        st.metric("Total Recipients", total_recipients)
        send_mode = st.radio(
            "Send Mode:",
            ["Send to All", "Send to Selected", "Send Test Email"]
        )
    
    with col2:
        if send_mode == "Send to Selected":
            if 'I_email' in df.columns:
                selected_recipients = st.multiselect(
                    "Select Recipients:",
                    df['I_email'].dropna().unique().tolist()
                )
            else:
                st.error("No email column found")
                selected_recipients = []
        elif send_mode == "Send Test Email":
            test_email = st.text_input("Test Email Address:")
            selected_recipients = [test_email] if test_email else []
        else:
            selected_recipients = df['I_email'].dropna().unique().tolist() if 'I_email' in df.columns else []
    
    # Show what will be sent
    if selected_recipients:
        st.subheader("📋 Email Summary")
        
        processing_date = datetime.now().strftime('%d %b %Y')
        subject = f"RE: Investment Update as on {processing_date}"
        
        st.write(f"**Subject:** {subject}")
        st.write(f"**Recipients:** {len(selected_recipients)}")
        
        # Show PDFs that will be generated
        total_pdfs = 0
        for recipient_email in selected_recipients:
            if send_mode == "Send Test Email":
                # For test emails, use first recipient's data from CSV
                recipient_data = df.head(1)
            else:
                recipient_data = df[df['I_email'] == recipient_email]
            
            securities = recipient_data['Security Name'].unique()
            available_securities = [s.strip() for s in securities if s.strip() in st.session_state.deals_data]
            total_pdfs += len(available_securities)
        
        st.write(f"**Total PDFs to be generated:** {total_pdfs}")
    
    # Final send button
    if st.button("📤 Send Emails", type="primary", disabled=not selected_recipients):
        if not email_address or not email_password:
            st.error("Please configure email settings first")
            return
        
        if not selected_recipients:
            st.error("No recipients selected")
            return
        
        # Progress tracking
        progress_bar = st.progress(0)
        status_text = st.empty()
        results_container = st.container()
        
        success_count = 0
        total_count = len(selected_recipients)
        
        processing_date = datetime.now().strftime('%d %b %Y')
        
        for i, recipient_email in enumerate(selected_recipients):
            status_text.text(f"Processing {recipient_email}... ({i+1}/{total_count})")
            
            try:
                # Get recipient data (special handling for test emails)
                if send_mode == "Send Test Email":
                    # Use first row of data for test email
                    recipient_data = df.head(1).copy()
                    recipient_data['I_email'] = recipient_email
                    recipient_data['Client Name/ Buyer Name'] = "Test Client"
                else:
                    recipient_data = df[df['I_email'] == recipient_email]
                
                if recipient_data.empty:
                    results_container.warning(f"⚠️ No data found for {recipient_email}")
                    continue
                
                # Generate investment tables HTML using exact original logic
                tables_html = ""
                for _, row in recipient_data.iterrows():
                    tables_html += create_investment_table_html(row)
                
                # Prepare template variables
                client_names = ', '.join(recipient_data['Client Name/ Buyer Name'].unique())
                template_vars = {
                    'client_names': client_names,
                    'processing_date': processing_date,
                    'investment_tables': tables_html,
                    **st.session_state.company_info
                }
                
                # Generate email content using exact original logic
                env = Environment(loader=BaseLoader())
                template = env.from_string(st.session_state.email_template)
                email_body = template.render(**template_vars)
                
                # Create full email HTML using exact original logic (600px width)
                full_email_html = f"""<!DOCTYPE html>
<html><head><meta charset="utf-8"></head><body style="margin:0;padding:0;background:#fff;">
<table width="600" align="center" cellpadding="0" cellspacing="0" border="0"
       style="margin:0 auto;background-color:#fff;">
  <tr><td style="padding:0;">{st.session_state.header_html}</td></tr>
  <tr><td style="padding:20px 0 20px 0;">{email_body}</td></tr>
  <tr><td style="padding:0;">{st.session_state.footer_html}</td></tr>
</table>
</body></html>"""
                
                # Generate PDFs for attachments using exact original logic
                pdf_data = {}
                securities = recipient_data['Security Name'].unique()
                
                for security in securities:
                    security_clean = security.strip()
                    if security_clean in st.session_state.deals_data:
                        # Prepare PDF data using exact original logic
                        pdf_vars = {
                            **st.session_state.deals_data[security_clean],
                            **st.session_state.company_info,
                            'processing_date': processing_date
                        }
                        
                        # Generate PDF
                        pdf_b64 = generate_investor_pdf(pdf_vars, st.session_state.pdf_template)
                        
                        if pdf_b64:
                            client_name = recipient_data['Client Name/ Buyer Name'].iloc[0]
                            pdf_filename = f"{client_name}_{security_clean}".replace(" ", "_") + ".pdf"
                            pdf_data[pdf_filename] = pdf_b64
                    else:
                        results_container.warning(f"⚠️ No deal data found for security: {security_clean}")
                
                # Send email using exact original logic
                subject = f"RE: Investment Update as on {processing_date}"
                success, message = send_email_with_pdfs(
                    recipient_email, subject, full_email_html, pdf_data, smtp_config
                )
                
                if success:
                    success_count += 1
                    results_container.success(f"✅ Sent to {recipient_email} ({len(pdf_data)} PDFs)")
                else:
                    results_container.error(f"❌ Failed to send to {recipient_email}: {message}")
                
            except Exception as e:
                results_container.error(f"❌ Error processing {recipient_email}: {str(e)}")
            
            # Update progress
            progress_bar.progress((i + 1) / total_count)
        
        # Final status
        status_text.text(f"✅ Completed! Successfully sent {success_count}/{total_count} emails")
        
        if success_count == total_count:
            st.success(f"🎉 All {total_count} emails sent successfully!")
        elif success_count > 0:
            st.warning(f"⚠️ Sent {success_count}/{total_count} emails. Check errors above.")
        else:
            st.error("❌ No emails were sent successfully. Please check your configuration and try again.")

if __name__ == "__main__":
    main()