import pandas as pd
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders
import os
from datetime import datetime
from jinja2 import Environment, FileSystemLoader
from weasyprint import HTML

def generate_investor_pdf(company_data, template_path, output_path):
    env = Environment(loader=FileSystemLoader(os.path.dirname(template_path) or "."))
    template = env.get_template(os.path.basename(template_path))
    html_out = template.render(**company_data)
    HTML(string=html_out, base_url='.').write_pdf(output_path)
    return output_path

def send_email_with_pdfs(to_email, subject, body, pdf_paths, smtp_config):
    msg = MIMEMultipart()
    msg['From'] = smtp_config['from_email']
    msg['To'] = to_email
    msg['Subject'] = subject
    msg.attach(MIMEText(body, 'html'))

    for pdf_path in pdf_paths:
        with open(pdf_path, "rb") as f:
            part = MIMEBase('application', 'pdf')
            part.set_payload(f.read())
            encoders.encode_base64(part)
            part.add_header('Content-Disposition', f'attachment; filename="{os.path.basename(pdf_path)}"')
            msg.attach(part)

    server = smtplib.SMTP(smtp_config['server'], smtp_config['port'])
    server.starttls()
    server.login(smtp_config['username'], smtp_config['password'])
    server.sendmail(smtp_config['from_email'], to_email, msg.as_string())
    server.quit()
    print(f"✅ Email with PDFs sent to {to_email}")

class InvestmentEmailGenerator:
    def load_data(self, file_path=None, data=None):
        if file_path:
            self.df = pd.read_csv(file_path)
            self.df.columns = self.df.columns.str.strip()
        elif data is not None:
            self.df = data
        else:
            raise Exception("No CSV file provided and no fallback implemented.")
        return self.df

    def create_investment_table_html(self, row):
        return f"""
        <div style="margin-bottom: 30px;">
            <table style="border-collapse: collapse; width: 100%; font-family: Arial, sans-serif; border: 1px solid #666;">
                <tr style="background-color: #232156; color: white;">
                    <td colspan="2" style="padding: 12px 20px; font-weight: bold; font-size: 16px; border: 1px solid #666;">{row['Security Name']}</td>
                </tr>
                <tr style="background-color: #f5f5f5;">
                    <td style="padding: 12px 20px; border: 1px solid #666;">No. of NCDs (nos.)</td>
                    <td style="padding: 12px 20px; border: 1px solid #666; text-align: right;">{row['No. of NCDs (nos.)']}</td>
                </tr>
                <tr>
                    <td style="padding: 12px 20px; border: 1px solid #666;">Face Value</td>
                    <td style="padding: 12px 20px; border: 1px solid #666; text-align: right;">{row['Face Value']}</td>
                </tr>
                <tr style="background-color: #f5f5f5;">
                    <td style="padding: 12px 20px; border: 1px solid #666;">Opening Principal</td>
                    <td style="padding: 12px 20px; border: 1px solid #666; text-align: right;">{row['Opening Principal Outstanding as on XX date']}</td>
                </tr>
                <tr>
                    <td style="padding: 12px 20px; border: 1px solid #666;">Principal Repaid</td>
                    <td style="padding: 12px 20px; border: 1px solid #666; text-align: right;">{row['Principal repaid on "Period"']}</td>
                </tr>
                <tr style="background-color: #f5f5f5;">
                    <td style="padding: 12px 20px; border: 1px solid #666;">Closing Principal</td>
                    <td style="padding: 12px 20px; border: 1px solid #666; text-align: right;">{row['Closing Principal Outstanding as on XX date']}</td>
                </tr>
                <tr>
                    <td style="padding: 12px 20px; border: 1px solid #666;">Net Interest Paid</td>
                    <td style="padding: 12px 20px; border: 1px solid #666; text-align: right;">{row['Net Interest paid for the "Period"']}</td>
                </tr>
            </table>
        </div>
        """

def test_email_connection(config):
    try:
        server = smtplib.SMTP(config['server'], config['port'])
        server.starttls()
        server.login(config['username'], config['password'])
        server.quit()
        print("✅ Email configuration successful!")
        return True
    except Exception as e:
        print(f"❌ Email configuration failed: {str(e)}")
        return False

if __name__ == "__main__":
    print("🚀 Investment Email Generator")
    csv_file = input("\nEnter CSV file path: ").strip()

    print("\nChoose sending option:")
    print("1. Gmail Send")
    print("2. Outlook Send")
    choice = input("\nEnter choice (1-2): ").strip()

    if choice == "1":
        gmail_user = input("Enter your Gmail address: ").strip()
        gmail_app_password = input("Enter your Gmail App Password (16 chars): ").strip()
        smtp_config = {
            'server': 'smtp.gmail.com',
            'port': 587,
            'username': gmail_user,
            'password': gmail_app_password,
            'from_email': gmail_user
        }
    elif choice == "2":
        outlook_user = input("Enter your Outlook email: ").strip()
        outlook_password = input("Enter your Outlook password: ").strip()
        smtp_config = {
            'server': 'smtp-mail.outlook.com',
            'port': 587,
            'username': outlook_user,
            'password': outlook_password,
            'from_email': outlook_user
        }
    else:
        print("Invalid choice. Exiting.")
        exit()

    if not test_email_connection(smtp_config):
        print("❌ SMTP config not working. Exiting.")
        exit()

    generator = InvestmentEmailGenerator()
    generator.load_data(file_path=csv_file)

    company_config = {
        'processing_date': datetime.now().strftime('%d %b %Y'),
        'company_signature': 'Neo Wealth',
        'company_address': '123, Example Address, Mumbai, India',
        'company_email': 'ir@neoassetmanagement.com'
    }

    base_company_data = {
        "borrower_profile": "BharatPe is a leading Indian fintech platform focused on empowering small merchants...",
        "recent_updates": [
            "The subsidiary, Trillionloans NBFC has recently got a rating of BBB+ from Fitch",
            "Current Balance sheet includes Net Current assets of ~400cr",
            "Current total debt as of Feb 2025, is at ~500cr along with a monthly Consolidated EBITDA run rate of ~10cr.",
            "Active metrics have improved including a merchant base of 24L and a Total Payments value of ~15,000 cr."
        ],
        "investment_summary": {
            "instrument": "Non-Convertible Debentures",
            "make_whole": "12 Months",
            "tenure": "2 years",
            "collateral_description": (
                "Charge on all current and non-current assets of the company excl. investments in Trillion Loans NBFC and Unity SFB. "
                "Negative lien on pledge of stake in Unity SFB<br>Security Cover >=1.0x"
            )
        },
        "decor_flower": "flower.svg",
        "contact_logo": "neo-logo.png"
    }
    pdf_template_path = "template.html"

    # --- Group by i_email ---
    if 'I_email' not in generator.df.columns:
        raise Exception("CSV must contain 'I_email' column for recipient grouping.")

    for i_email, group in generator.df.groupby('I_email'):
        if pd.isna(i_email) or not i_email:
            continue

        investor_names = group['Client Name/ Buyer Name'].unique()
        all_tables = []
        pdf_paths = []

        for idx, row in group.iterrows():
            # Table for this security
            all_tables.append(generator.create_investment_table_html(row))
            # PDF for this security
            company_data = base_company_data.copy()
            company_data["company_name"] = row["Security Name"]
            os.makedirs("generated_pdfs", exist_ok=True)
            pdf_filename = f"{row['Client Name/ Buyer Name'].lower().replace(' ', '_')}_{row['Security Name'].lower().replace(' ', '_')}.pdf"
            pdf_path = os.path.join("generated_pdfs", pdf_filename)
            generate_investor_pdf(company_data, pdf_template_path, pdf_path)
            pdf_paths.append(pdf_path)

        # Join all tables in one email
        tables_html = "".join(all_tables)
        # Compose email body
        body = f"""
        Dear {', '.join(investor_names)},<br><br>
        Please find below updated summary of payment details processed on {company_config['processing_date']}. Attached are the repayment schedules for your securities.<br><br>
        {tables_html}
        <br>Feel free to connect in case you need any further information/clarifications.<br><br>
        Best regards,<br>
        Investors Relations Team<br><br>
        {company_config['company_signature']}<br>
        {company_config['company_address']}<br>
        E-mail: {company_config['company_email']}<br>
        """
        subject = f"RE: Investment NCD Interest and Principal Repayment for {company_config['processing_date']}"

        send_email_with_pdfs(i_email, subject, body, pdf_paths, smtp_config)

    print("\n🎉 All emails with PDFs sent!")
