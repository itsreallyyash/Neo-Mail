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

# Initialize session state
if 'email_template' not in st.session_state:
    st.session_state.email_template = """Dear {client_names},

Please find below updated summary of payment details processed on {processing_date}.

{investment_tables}

Feel free to connect for any clarifications.

Best regards,
Investor Relations Team

{company_signature}
{company_address}
E-mail: {company_email}"""

if 'deals_data' not in st.session_state:
    st.session_state.deals_data = {}

if 'csv_data' not in st.session_state:
    st.session_state.csv_data = None

if 'header_html' not in st.session_state:
    st.session_state.header_html = ""

if 'footer_html' not in st.session_state:
    st.session_state.footer_html = ""

# ─── HELPER FUNCTIONS ────────────────────────────────────────────────────────

def iter_block_items(doc):
    """Yield each Paragraph or Table in the document in true Word order."""
    for child in doc.element.body:
        if isinstance(child, CT_P):
            yield Paragraph(child, doc)
        elif isinstance(child, CT_Tbl):
            yield Table(child, doc)

def parse_word_doc(uploaded_file) -> dict:
    """Parse Word document and extract deal information."""
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

                if section == "borrower_profile":
                    current["borrower_profile"].append(text)
                elif section == "recent_updates":
                    bullet = re.sub(r"^[•\-\*\s]+", "", text)
                    current["recent_updates"].append(bullet)

            elif isinstance(block, Table) and expecting_table:
                for row in block.rows:
                    k = row.cells[0].text.strip()
                    v = row.cells[1].text.strip()
                    if k:
                        current["investment_summary"][k] = v
                expecting_table = False

        for d in deals.values():
            d["borrower_profile"] = "<br><br>".join(d["borrower_profile"])
        
        return deals
    finally:
        os.unlink(tmp_path)

def create_investment_table_html(row):
    """Create HTML table for investment data."""
    return f"""
    <div style="margin-bottom: 30px;">
      <table style="width:100%;border:1px solid #666;border-collapse:collapse;font-family:Arial,sans-serif">
        <tr style="background:#232156;color:#fff">
          <td colspan="2" style="padding:12px;font-weight:bold">{row['Security Name']}</td>
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

def generate_pdf_preview(company_data, template_html):
    """Generate PDF preview as base64 string."""
    try:
        env = Environment(loader=BaseLoader())
        template = env.from_string(template_html)
        html_content = template.render(**company_data)
        
        # Create PDF in memory
        pdf_buffer = io.BytesIO()
        HTML(string=html_content).write_pdf(pdf_buffer)
        pdf_buffer.seek(0)
        
        return base64.b64encode(pdf_buffer.read()).decode()
    except Exception as e:
        st.error(f"Error generating PDF preview: {str(e)}")
        return None

def test_smtp_connection(smtp_config):
    """Test SMTP connection."""
    try:
        server = smtplib.SMTP(smtp_config['server'], smtp_config['port'])
        server.starttls()
        server.login(smtp_config['username'], smtp_config['password'])
        server.quit()
        return True, "Connection successful!"
    except Exception as e:
        return False, f"Connection failed: {str(e)}"

def send_email_with_attachments(to_email, subject, html_content, pdf_data, smtp_config):
    """Send email with PDF attachments."""
    try:
        msg = MIMEMultipart()
        msg['From'] = smtp_config['from_email']
        msg['To'] = to_email
        msg['Subject'] = subject
        
        msg.attach(MIMEText(html_content, 'html'))
        
        # Attach PDFs
        for pdf_name, pdf_content in pdf_data.items():
            part = MIMEBase('application', 'pdf')
            part.set_payload(base64.b64decode(pdf_content))
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
    st.markdown("---")
    
    # Sidebar navigation
    st.sidebar.title("Navigation")
    page = st.sidebar.selectbox(
        "Choose a section:",
        ["🔧 Configuration", "🧪 Test Mode", "📤 Send Mode"]
    )
    
    if page == "🔧 Configuration":
        configuration_page()
    elif page == "🧪 Test Mode":
        test_mode_page()
    elif page == "📤 Send Mode":
        send_mode_page()

def configuration_page():
    st.header("🔧 Configuration")
    
    # File uploads
    col1, col2 = st.columns(2)
    
    with col1:
        st.subheader("📄 Upload Word Document")
        word_file = st.file_uploader(
            "Upload Word document with deal information",
            type=['docx'],
            help="Upload the Word document containing deal information"
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
                except Exception as e:
                    st.error(f"Error parsing Word document: {str(e)}")
    
    with col2:
        st.subheader("📊 Upload CSV Data")
        csv_file = st.file_uploader(
            "Upload CSV with investment data",
            type=['csv'],
            help="Upload the CSV file containing investment data"
        )
        
        if csv_file is not None:
            try:
                df = pd.read_csv(csv_file)
                df.columns = df.columns.str.strip()
                st.session_state.csv_data = df
                st.success(f"✅ Loaded CSV with {len(df)} records")
                
                # Show data preview
                with st.expander("View CSV Data Preview"):
                    st.dataframe(df.head())
                    
                # Show recipient count
                if 'I_email' in df.columns:
                    unique_emails = df['I_email'].dropna().nunique()
                    st.info(f"📧 Found {unique_emails} unique recipient emails")
                else:
                    st.warning("⚠️ No 'I_email' column found in CSV")
            except Exception as e:
                st.error(f"Error loading CSV: {str(e)}")
    
    # Header/Footer HTML
    st.subheader("🎨 Email Template Components")
    
    col1, col2 = st.columns(2)
    
    with col1:
        st.write("**Header HTML**")
        header_file = st.file_uploader(
            "Upload header HTML file",
            type=['html'],
            key="header_upload"
        )
        if header_file:
            header_content = header_file.read().decode('utf-8')
            # Extract body content
            header_match = re.search(r'<body[^>]*>(.*?)</body>', header_content, flags=re.S|re.I)
            st.session_state.header_html = header_match.group(1) if header_match else header_content
            st.success("✅ Header HTML loaded")
    
    with col2:
        st.write("**Footer HTML**")
        footer_file = st.file_uploader(
            "Upload footer HTML file",
            type=['html'],
            key="footer_upload"
        )
        if footer_file:
            footer_content = footer_file.read().decode('utf-8')
            # Extract body content
            footer_match = re.search(r'<body[^>]*>(.*?)</body>', footer_content, flags=re.S|re.I)
            st.session_state.footer_html = footer_match.group(1) if footer_match else footer_content
            st.success("✅ Footer HTML loaded")
    
    # Email template editor
    st.subheader("✏️ Email Template Editor")
    st.info("Available variables: {client_names}, {processing_date}, {investment_tables}, {company_signature}, {company_address}, {company_email}")
    
    template_text = st.text_area(
        "Edit email template:",
        value=st.session_state.email_template,
        height=300,
        help="Use the variables above to customize your email template"
    )
    
    col1, col2 = st.columns([1, 4])
    with col1:
        if st.button("💾 Save Template"):
            st.session_state.email_template = template_text
            st.success("✅ Template saved!")
    
    # Company information
    st.subheader("🏢 Company Information")
    
    col1, col2 = st.columns(2)
    with col1:
        company_signature = st.text_input("Company Name", value="Neo Wealth")
        company_email = st.text_input("Company Email", value="ir@neoassetmanagement.com")
    
    with col2:
        company_address = st.text_area("Company Address", value="123, Example Address, Mumbai, India")
    
    # Store company info in session state
    st.session_state.company_info = {
        'company_signature': company_signature,
        'company_email': company_email,
        'company_address': company_address
    }

def test_mode_page():
    st.header("🧪 Test Mode")
    
    # Check if data is loaded
    if not st.session_state.csv_data is not None or not st.session_state.deals_data:
        st.warning("⚠️ Please upload Word document and CSV data in the Configuration section first.")
        return
    
    df = st.session_state.csv_data
    deals_data = st.session_state.deals_data
    
    # Show recipients list
    st.subheader("📋 Recipients List")
    if 'I_email' in df.columns:
        recipients_df = df.groupby('I_email').agg({
            'Client Name/ Buyer Name': lambda x: ', '.join(x.unique()),
            'Security Name': lambda x: ', '.join(x.unique())
        }).reset_index()
        
        st.dataframe(recipients_df, use_container_width=True)
        st.info(f"Total recipients: {len(recipients_df)}")
    else:
        st.error("No 'I_email' column found in CSV data")
        return
    
    # Email preview
    st.subheader("📧 Email Preview")
    
    # Select recipient for preview
    if len(recipients_df) > 0:
        selected_email = st.selectbox(
            "Select recipient for preview:",
            recipients_df['I_email'].tolist()
        )
        
        if selected_email:
            # Get data for selected recipient
            recipient_data = df[df['I_email'] == selected_email]
            
            # Generate investment tables HTML
            tables_html = ""
            for _, row in recipient_data.iterrows():
                tables_html += create_investment_table_html(row)
            
            # Prepare template variables
            client_names = ', '.join(recipient_data['Client Name/ Buyer Name'].unique())
            template_vars = {
                'client_names': client_names,
                'processing_date': datetime.now().strftime('%d %b %Y'),
                'investment_tables': tables_html,
                **st.session_state.company_info
            }
            
            # Generate email content
            env = Environment(loader=BaseLoader())
            template = env.from_string(st.session_state.email_template)
            email_body = template.render(**template_vars)
            
            # Create full email HTML
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
            st.components.v1.html(full_email_html, height=600, scrolling=True)
            
            # PDF Preview
            st.subheader("📄 PDF Preview")
            
            # Get securities for this recipient
            securities = recipient_data['Security Name'].unique()
            
            for security in securities:
                if security in deals_data:
                    st.write(f"**PDF for: {security}**")
                    
                    # Prepare PDF template (simplified)
                    pdf_template = """
                    <html>
                    <head><style>
                        body { font-family: Arial, sans-serif; margin: 20px; }
                        .header { background: #232156; color: white; padding: 10px; }
                        .content { margin: 20px 0; }
                        .summary-table { width: 100%; border-collapse: collapse; }
                        .summary-table th, .summary-table td { 
                            border: 1px solid #ddd; padding: 8px; text-align: left; 
                        }
                        .summary-table th { background-color: #f2f2f2; }
                    </style></head>
                    <body>
                        <div class="header">
                            <h2>{{ company_name }}</h2>
                        </div>
                        <div class="content">
                            <h3>Investment Summary</h3>
                            <table class="summary-table">
                                {% for key, value in investment_summary.items() %}
                                <tr><td>{{ key }}</td><td>{{ value }}</td></tr>
                                {% endfor %}
                            </table>
                            
                            <h3>Borrower Profile</h3>
                            <p>{{ borrower_profile|safe }}</p>
                            
                            <h3>Recent Updates</h3>
                            <ul>
                                {% for update in recent_updates %}
                                <li>{{ update }}</li>
                                {% endfor %}
                            </ul>
                        </div>
                    </body>
                    </html>
                    """
                    
                    # Generate PDF preview
                    pdf_data = {
                        **deals_data[security],
                        **st.session_state.company_info,
                        'processing_date': datetime.now().strftime('%d %b %Y')
                    }
                    
                    pdf_b64 = generate_pdf_preview(pdf_data, pdf_template)
                    
                    if pdf_b64:
                        st.markdown(
                            f'<iframe src="data:application/pdf;base64,{pdf_b64}" width="100%" height="400"></iframe>',
                            unsafe_allow_html=True
                        )
                else:
                    st.warning(f"No deal data found for security: {security}")

def send_mode_page():
    st.header("📤 Send Mode")
    
    # Check if data is loaded
    if not st.session_state.csv_data is not None or not st.session_state.deals_data:
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
    if st.button("🔍 Test Email Connection"):
        if email_address and email_password:
            with st.spinner("Testing connection..."):
                success, message = test_smtp_connection(smtp_config)
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
            ["Send to All", "Send to Selected"]
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
        else:
            selected_recipients = df['I_email'].dropna().unique().tolist() if 'I_email' in df.columns else []
    
    # Final send button
    if st.button("📤 Send Emails", type="primary"):
        if not email_address or not email_password:
            st.error("Please configure email settings first")
            return
        
        if not selected_recipients:
            st.error("No recipients selected")
            return
        
        # Progress tracking
        progress_bar = st.progress(0)
        status_text = st.empty()
        
        success_count = 0
        total_count = len(selected_recipients)
        
        for i, recipient_email in enumerate(selected_recipients):
            status_text.text(f"Sending to {recipient_email}...")
            
            try:
                # Get recipient data
                recipient_data = df[df['I_email'] == recipient_email]
                
                # Generate email content
                tables_html = ""
                for _, row in recipient_data.iterrows():
                    tables_html += create_investment_table_html(row)
                
                client_names = ', '.join(recipient_data['Client Name/ Buyer Name'].unique())
                template_vars = {
                    'client_names': client_names,
                    'processing_date': datetime.now().strftime('%d %b %Y'),
                    'investment_tables': tables_html,
                    **st.session_state.company_info
                }
                
                env = Environment(loader=BaseLoader())
                template = env.from_string(st.session_state.email_template)
                email_body = template.render(**template_vars)
                
                full_email_html = f"""<!DOCTYPE html>
<html><head><meta charset="utf-8"></head><body style="margin:0;padding:0;background:#fff;">
<table width="600" align="center" cellpadding="0" cellspacing="0" border="0"
       style="margin:0 auto;background-color:#fff;">
  <tr><td style="padding:0;">{st.session_state.header_html}</td></tr>
  <tr><td style="padding:20px 0 20px 0;">{email_body}</td></tr>
  <tr><td style="padding:0;">{st.session_state.footer_html}</td></tr>
</table>
</body></html>"""
                
                # Generate PDFs for attachments
                pdf_data = {}
                securities = recipient_data['Security Name'].unique()
                
                for security in securities:
                    if security in st.session_state.deals_data:
                        # Create PDF content (simplified for demo)
                        pdf_template = """
                        <html><body>
                            <h2>{{ company_name }}</h2>
                            <h3>Investment Summary</h3>
                            {% for key, value in investment_summary.items() %}
                            <p><strong>{{ key }}:</strong> {{ value }}</p>
                            {% endfor %}
                        </body></html>
                        """
                        
                        pdf_vars = {
                            **st.session_state.deals_data[security],
                            **st.session_state.company_info
                        }
                        
                        pdf_b64 = generate_pdf_preview(pdf_vars, pdf_template)
                        if pdf_b64:
                            client_name = recipient_data['Client Name/ Buyer Name'].iloc[0]
                            pdf_filename = f"{client_name}_{security}.pdf".replace(" ", "_")
                            pdf_data[pdf_filename] = pdf_b64
                
                # Send email
                subject = f"RE: Investment Update as on {template_vars['processing_date']}"
                success, message = send_email_with_attachments(
                    recipient_email, subject, full_email_html, pdf_data, smtp_config
                )
                
                if success:
                    success_count += 1
                else:
                    st.error(f"Failed to send to {recipient_email}: {message}")
                
            except Exception as e:
                st.error(f"Error sending to {recipient_email}: {str(e)}")
            
            # Update progress
            progress_bar.progress((i + 1) / total_count)
        
        # Final status
        status_text.text(f"✅ Completed! Successfully sent {success_count}/{total_count} emails")
        
        if success_count == total_count:
            st.success(f"🎉 All {total_count} emails sent successfully!")
        else:
            st.warning(f"⚠️ Sent {success_count}/{total_count} emails. Check errors above.")

if __name__ == "__main__":
    main()