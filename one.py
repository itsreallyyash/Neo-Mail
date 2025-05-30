# #!/usr/bin/env python3
# """
# Full Investment Email Generator with robust Word parsing
# """

# import re
# import os
# import pandas as pd
# import smtplib
# from datetime import datetime
# from docx import Document
# from jinja2 import Environment, FileSystemLoader
# from weasyprint import HTML
# from email.mime.text import MIMEText
# from email.mime.multipart import MIMEMultipart
# from email.mime.base import MIMEBase
# from email import encoders

# # ─── 1. PARSE YOUR WORD DOC ────────────────────────────────────────────────────────
# from docx import Document
# from docx.oxml.table import CT_Tbl
# from docx.oxml.text.paragraph import CT_P
# from docx.text.paragraph import Paragraph
# from docx.table import Table
# import re

# from docx import Document
# from docx.oxml.table import CT_Tbl
# from docx.oxml.text.paragraph import CT_P
# from docx.text.paragraph import Paragraph
# from docx.table import Table
# import re

# from docx import Document
# from docx.oxml.text.paragraph import CT_P
# from docx.oxml.table import CT_Tbl
# from docx.text.paragraph import Paragraph
# from docx.table import Table
# import re

# def iter_block_items(doc):
#     """
#     Yield each Paragraph or Table in the document in true Word order.
#     """
#     for child in doc.element.body:
#         if isinstance(child, CT_P):
#             yield Paragraph(child, doc)
#         elif isinstance(child, CT_Tbl):
#             yield Table(child, doc)

# def parse_word_doc(path: str) -> dict:
#     """
#     Parses your Deals doc.  Grabs:
#       - borrower_profile (paras)
#       - investment_summary (first Table after that heading)
#       - recent_updates (bullets)
#     """
#     doc            = Document(path)
#     deals          = {}
#     current        = None
#     section        = None
#     expecting_table = False

#     for block in iter_block_items(doc):
#         # — Paragraphs —
#         if isinstance(block, Paragraph):
#             text = block.text.strip()
#             if not text:
#                 continue

#             # New Deal header?
#             m = re.match(r"Deal\s*\d+\s*[:\-]?\s*(.+)", text, flags=re.I)
#             if m:
#                 cname = m.group(1).rstrip(":").strip()
#                 current = {
#                     "company_name": cname,
#                     "borrower_profile": [],
#                     "investment_summary": {},
#                     "recent_updates": []
#                 }
#                 deals[cname] = current
#                 section = None
#                 expecting_table = False
#                 continue

#             low = text.lower()
#             if "borrower profile" in low:
#                 section = "borrower_profile"
#                 continue
#             if "investment summary" in low:
#                 # next TABLE we encounter is the summary
#                 expecting_table = True
#                 section = None
#                 continue
#             if "recent updates" in low:
#                 section = "recent_updates"
#                 continue

#             if section == "borrower_profile":
#                 current["borrower_profile"].append(text)
#             elif section == "recent_updates":
#                 bullet = re.sub(r"^[•\-\*\s]+", "", text)
#                 current["recent_updates"].append(bullet)

#         # — Tables —
#         elif isinstance(block, Table) and expecting_table:
#             for row in block.rows:
#                 k = row.cells[0].text.strip()
#                 v = row.cells[1].text.strip()
#                 if k:
#                     current["investment_summary"][k] = v
#             expecting_table = False

#     # join borrower_profile paragraphs
#     for d in deals.values():
#         d["borrower_profile"] = "<br><br>".join(d["borrower_profile"])

#     return deals

# # ─── 2. HTML TABLE + PDF + EMAIL HELPERS ───────────────────────────────────────────
# class InvestmentEmailGenerator:
#     def load_data(self, file_path=None, data=None):
#         if file_path:
#             self.df = pd.read_csv(file_path)
#             self.df.columns = self.df.columns.str.strip()
#         elif data is not None:
#             self.df = data
#         else:
#             raise ValueError("No CSV provided and no fallback data.")
#         return self.df

#     def create_investment_table_html(self, row):
#         """Renders one security’s row as a stylized HTML table."""
#         return f"""
#         <div style="margin-bottom: 30px;">
#           <table style="width:100%;border:1px solid #666;border-collapse:collapse;font-family:Arial,sans-serif">
#             <tr style="background:#232156;color:#fff">
#               <td colspan="2" style="padding:12px;font-weight:bold">{row['Security Name'].strip()}</td>
#             </tr>
#             <tr style="background:#f5f5f5">
#               <td style="padding:8px;border:1px solid #666">No. of NCDs (nos.)</td>
#               <td style="padding:8px;border:1px solid #666;text-align:right">{row['No. of NCDs (nos.)']}</td>
#             </tr>
#             <tr>
#               <td style="padding:8px;border:1px solid #666">Face Value</td>
#               <td style="padding:8px;border:1px solid #666;text-align:right">{row['Face Value']}</td>
#             </tr>
#             <tr style="background:#f5f5f5">
#               <td style="padding:8px;border:1px solid #666">Opening Principal</td>
#               <td style="padding:8px;border:1px solid #666;text-align:right">{row['Opening Principal Outstanding as on XX date']}</td>
#             </tr>
#             <tr>
#               <td style="padding:8px;border:1px solid #666">Principal Repaid</td>
#               <td style="padding:8px;border:1px solid #666;text-align:right">{row['Principal repaid on "Period"']}</td>
#             </tr>
#             <tr style="background:#f5f5f5">
#               <td style="padding:8px;border:1px solid #666">Closing Principal</td>
#               <td style="padding:8px;border:1px solid #666;text-align:right">{row['Closing Principal Outstanding as on XX date']}</td>
#             </tr>
#             <tr>
#               <td style="padding:8px;border:1px solid #666">Net Interest Paid</td>
#               <td style="padding:8px;border:1px solid #666;text-align:right">{row['Net Interest paid for the "Period"']}</td>
#             </tr>
#           </table>
#         </div>
#         """


# def generate_investor_pdf(company_data: dict, template_path: str, output_path: str) -> str:
#     """Renders company_data via Jinja2 → HTML → PDF (WeasyPrint)."""
#     env = Environment(loader=FileSystemLoader(os.path.dirname(template_path) or "."))
#     tpl = env.get_template(os.path.basename(template_path))
#     html = tpl.render(**company_data)
#     HTML(string=html, base_url=".").write_pdf(output_path)
#     return output_path


# def send_email_with_pdfs(to_email, subject, body_html, pdf_paths, smtp_cfg):
#     msg = MIMEMultipart()
#     msg['From']    = smtp_cfg['from_email']
#     msg['To']      = to_email
#     msg['Subject'] = subject
#     msg.attach(MIMEText(body_html, 'html'))

#     for pdf in pdf_paths:
#         with open(pdf, "rb") as f:
#             part = MIMEBase('application', 'pdf')
#             part.set_payload(f.read())
#             encoders.encode_base64(part)
#             part.add_header('Content-Disposition',
#                             f'attachment; filename="{os.path.basename(pdf)}"')
#             msg.attach(part)

#     server = smtplib.SMTP(smtp_cfg['server'], smtp_cfg['port'])
#     server.starttls()
#     server.login(smtp_cfg['username'], smtp_cfg['password'])
#     server.sendmail(smtp_cfg['from_email'], to_email, msg.as_string())
#     server.quit()


# def test_email_connection(cfg) -> bool:
#     try:
#         s = smtplib.SMTP(cfg['server'], cfg['port'])
#         s.starttls()
#         s.login(cfg['username'], cfg['password'])
#         s.quit()
#         return True
#     except Exception as e:
#         print("SMTP connection failed:", e)
#         return False


# # ─── 3. MAIN PIPELINE ─────────────────────────────────────────────────────────────
# if __name__ == "__main__":
#     # 1) Inputs
#     tables_csv     = "/Users/yashshah/Downloads/AIEnterprise/DS_Data.csv"
#     word_doc_path  = "/Users/yashshah/Downloads/AIEnterprise/Downsell Communicaiton  - All Deals.docx"
#     template_path  = "/Users/yashshah/Downloads/AIEnterprise/template.html"
#     header_path = "/Users/yashshah/Downloads/AIEnterprise/header-img.html"
#     footer_path = "/Users/yashshah/Downloads/AIEnterprise/neowmfooter-signature.html"
#     with open(header_path, 'r', encoding='utf-8') as f:
#         header_html = f.read()
#     with open(footer_path, 'r', encoding='utf-8') as f:
#         footer_html = f.read()
#     # 2) SMTP config
#     print("\nChoose sending option:\n 1. Gmail   2. Outlook")
#     choice = input("Choice (1-2): ").strip()
#     if choice == "1":
#         user = input("Gmail address: ").strip()
#         pw   = input("Gmail app-password (16 chars): ").strip()
#         smtp_cfg = {
#             'server':     'smtp.gmail.com',
#             'port':       587,
#             'username':   user,
#             'password':   pw,
#             'from_email': user
#         }
#     else:
#         user = input("Outlook address: ").strip()
#         pw   = input("Outlook password:      ").strip()
#         smtp_cfg = {
#             'server':     'smtp-mail.outlook.com',
#             'port':       587,
#             'username':   user,
#             'password':   pw,
#             'from_email': user
#         }

#     if not test_email_connection(smtp_cfg):
#         print("❌ SMTP config not working. Exiting.")
#         exit(1)

#     # 3) Load CSV
#     gen       = InvestmentEmailGenerator()
#     tables_df = gen.load_data(file_path=tables_csv)

#     # 4) Parse Word doc
#     deals = parse_word_doc(word_doc_path)

#     # 5) Shared company info
#     base_config = {
#         'processing_date':   datetime.now().strftime('%d %b %Y'),
#         'company_signature': 'Neo Wealth',
#         'company_address':   '123, Example Address, Mumbai, India',
#         'company_email':     'ir@neoassetmanagement.com',
#         'decor_flower':      'flower.svg',
#         'contact_logo':      'neo-logo.png'
#     }

#     # 6) Group by investor e-mail & send
#     for investor_email, group in tables_df.groupby('I_email'):
#         if pd.isna(investor_email) or not investor_email.strip():
#             continue

#         # a) Build body HTML tables
#         tables_html = "".join(
#             gen.create_investment_table_html(row)
#             for _, row in group.iterrows()
#         )

#         # b) Generate one PDF per security
#         pdf_paths = []
#         for _, row in group.iterrows():
#             sec = row['Security Name'].strip()
#             if sec not in deals:
#                 print(f"⚠︎ Warning: no deal data for '{sec}', skipping PDF.")
#                 continue

#             data = {**deals[sec], **base_config}
#             os.makedirs("generated_pdfs", exist_ok=True)
#             fname = f"{row['Client Name/ Buyer Name']}_{sec}".replace(" ", "_") + ".pdf"
#             out   = os.path.join("generated_pdfs", fname)
#             generate_investor_pdf(data, template_path, out)
#             pdf_paths.append(out)

#         # c) Compose & send
#         names   = [n.strip() for n in group['Client Name/ Buyer Name'].unique()]
#         body    = f"""
#         Dear {', '.join(names)},<br><br>
#         Please find below updated summary of payment details processed on {base_config['processing_date']}.<br><br>
#         {tables_html}
#         <br>
#         Feel free to connect for any clarifications.<br><br>
#         Best regards,<br>
#         Investor Relations Team<br><br>
#         {base_config['company_signature']}<br>
#         {base_config['company_address']}<br>
#         E-mail: {base_config['company_email']}<br>
#         """
#         subject = f"RE: Investment Update as on {base_config['processing_date']}"

#         # combine header + your generated body + footer
#         full_email_html = header_html + body + footer_html

#         send_email_with_pdfs(investor_email, subject, full_email_html, pdf_paths, smtp_cfg)
#         print(f"✅ Sent to {investor_email} ({len(pdf_paths)} PDFs).")

#     print("\n🎉 All done!")
#!/usr/bin/env python3
"""
Full Investment Email Generator with robust Word parsing
and 600px-wide header/body/footer email layout.
"""

import re
import os
import pandas as pd
import smtplib
from datetime import datetime
from docx import Document
from docx.oxml.table import CT_Tbl
from docx.oxml.text.paragraph import CT_P
from docx.text.paragraph import Paragraph
from docx.table import Table
from jinja2 import Environment, FileSystemLoader
from weasyprint import HTML
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders

# ─── 1. PARSE YOUR WORD DOC ────────────────────────────────────────────────────────
def iter_block_items(doc):
    """Yield each Paragraph or Table in the document in true Word order."""
    for child in doc.element.body:
        if isinstance(child, CT_P):
            yield Paragraph(child, doc)
        elif isinstance(child, CT_Tbl):
            yield Table(child, doc)

def parse_word_doc(path: str) -> dict:
    """
    Parses your Deals doc.  Grabs:
      - borrower_profile (paras)
      - investment_summary (first Table after that heading)
      - recent_updates (bullets)
    """
    doc            = Document(path)
    deals          = {}
    current        = None
    section        = None
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

            if section == "borrower_profile":
                current["borrower_profile"].append(text)
            elif section == "recent_updates":
                bullet = re.sub(r"^[•\-\*\s]+", "", text)
                current["recent_updates"].append(bullet)

        # — Tables —
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

# ─── 2. TABLE‐HTML + PDF + EMAIL HELPERS ───────────────────────────────────────────
class InvestmentEmailGenerator:
    def load_data(self, file_path=None, data=None):
        if file_path:
            self.df = pd.read_csv(file_path)
            self.df.columns = self.df.columns.str.strip()
        elif data is not None:
            self.df = data
        else:
            raise ValueError("No CSV provided and no fallback data.")
        return self.df

    def create_investment_table_html(self, row):
        """Renders one security’s row as a stylized HTML table."""
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

def generate_investor_pdf(company_data, template_path, output_path):
    env = Environment(loader=FileSystemLoader(os.path.dirname(template_path) or "."))
    tpl = env.get_template(os.path.basename(template_path))
    html = tpl.render(**company_data)
    HTML(string=html, base_url=".").write_pdf(output_path)
    return output_path

def send_email_with_pdfs(to_email, subject, html_content, pdf_paths, smtp_cfg):
    msg = MIMEMultipart()
    msg['From']    = smtp_cfg['from_email']
    msg['To']      = to_email
    msg['Subject'] = subject
    msg.attach(MIMEText(html_content, 'html'))

    for pdf in pdf_paths:
        part = MIMEBase('application', 'pdf')
        with open(pdf, "rb") as f:
            part.set_payload(f.read())
        encoders.encode_base64(part)
        part.add_header('Content-Disposition', f'attachment; filename="{os.path.basename(pdf)}"')
        msg.attach(part)

    server = smtplib.SMTP(smtp_cfg['server'], smtp_cfg['port'])
    server.starttls()
    server.login(smtp_cfg['username'], smtp_cfg['password'])
    server.sendmail(smtp_cfg['from_email'], to_email, msg.as_string())
    server.quit()

def test_email_connection(cfg) -> bool:
    try:
        s = smtplib.SMTP(cfg['server'], cfg['port'])
        s.starttls()
        s.login(cfg['username'], cfg['password'])
        s.quit()
        return True
    except Exception as e:
        print("SMTP connection failed:", e)
        return False

# ─── 3. MAIN PIPELINE ─────────────────────────────────────────────────────────────
if __name__ == "__main__":
    # 1) Paths (you said “keep them the same”)
    tables_csv     = "/Users/yashshah/Downloads/AIEnterprise/DS_Data.csv"
    word_doc_path  = "/Users/yashshah/Downloads/AIEnterprise/Downsell Communicaiton  - All Deals.docx"
    template_path  = "/Users/yashshah/Downloads/AIEnterprise/template.html"
    header_path    = "/Users/yashshah/Downloads/AIEnterprise/header-img.html"
    footer_path    = "/Users/yashshah/Downloads/AIEnterprise/neowmfooter-signature.html"

    # Read & extract only the <body>…</body> contents
    header_raw = open(header_path, 'r', encoding='utf-8').read()
    footer_raw = open(footer_path, 'r', encoding='utf-8').read()
    header_snip = re.search(r'<body[^>]*>(.*?)</body>', header_raw, flags=re.S|re.I)
    footer_snip = re.search(r'<body[^>]*>(.*?)</body>', footer_raw, flags=re.S|re.I)
    header_html = header_snip.group(1) if header_snip else header_raw
    footer_html = footer_snip.group(1) if footer_snip else footer_raw

    # 2) SMTP config
    print("\nChoose sending option:\n 1. Gmail   2. Outlook")
    choice = input("Choice (1-2): ").strip()
    if choice == "1":
        user = input("Gmail address: ").strip()
        pw   = input("Gmail app-password (16 chars): ").strip()
        smtp_cfg = {
            'server':     'smtp.gmail.com',
            'port':       587,
            'username':   user,
            'password':   pw,
            'from_email': user
        }
    else:
        user = input("Outlook address: ").strip()
        pw   = input("Outlook password: ").strip()
        smtp_cfg = {
            'server':     'smtp-mail.outlook.com',
            'port':       587,
            'username':   user,
            'password':   pw,
            'from_email': user
        }
    if not test_email_connection(smtp_cfg):
        print("❌ SMTP config not working. Exiting.")
        exit(1)

    # 3) Load CSV & parse doc
    gen       = InvestmentEmailGenerator()
    tables_df = gen.load_data(file_path=tables_csv)
    deals     = parse_word_doc(word_doc_path)

    # 4) Shared metadata
    base_config = {
        'processing_date':   datetime.now().strftime('%d %b %Y'),
        'company_signature': 'Neo Wealth',
        'company_address':   '123, Example Address, Mumbai, India',
        'company_email':     'ir@neoassetmanagement.com',
        'decor_flower':      'flower.svg',
        'contact_logo':      'neo-logo.png'
    }

    # 5) Loop through each investor group
    for investor_email, group in tables_df.groupby('I_email'):
        if pd.isna(investor_email) or not investor_email.strip():
            continue

        # Build the HTML block of all your tables (100%-wide)
        tables_html = "".join(
            gen.create_investment_table_html(row)
            for _, row in group.iterrows()
        )

        # Generate PDFs per security
        pdf_paths = []
        for _, row in group.iterrows():
            sec = row['Security Name'].strip()
            if sec not in deals:
                print(f"⚠︎ Warning: no deal data for '{sec}', skipping PDF.")
                continue
            data = {**deals[sec], **base_config}
            os.makedirs("generated_pdfs", exist_ok=True)
            fname = f"{row['Client Name/ Buyer Name']}_{sec}".replace(" ", "_") + ".pdf"
            out   = os.path.join("generated_pdfs", fname)
            generate_investor_pdf(data, template_path, out)
            pdf_paths.append(out)

        # Compose the plain‐vanilla body
        names = [n.strip() for n in group['Client Name/ Buyer Name'].unique()]
        body  = f"""
        Dear {', '.join(names)},<br><br>
        Please find below updated summary of payment details processed on {base_config['processing_date']}.<br><br>
        {tables_html}
        <br>
        Feel free to connect for any clarifications.<br><br>
        Best regards,<br>
        Investor Relations Team<br><br>
        {base_config['company_signature']}<br>
        {base_config['company_address']}<br>
        E-mail: {base_config['company_email']}<br>
        """

        # Wrap everything in one 600px-wide email
        full_email_html = f"""<!DOCTYPE html>
<html><head><meta charset="utf-8"></head><body style="margin:0;padding:0;background:#fff;">
<table width="600" align="center" cellpadding="0" cellspacing="0" border="0"
       style="margin:0 auto;background-color:#fff;">
  <tr><td style="padding:0;">{header_html}</td></tr>
  <tr><td style="padding:20px 0 20px 0;">{body}</td></tr>
  <tr><td style="padding:0;">{footer_html}</td></tr>
</table>
</body></html>
"""

        # Send it
        subject = f"RE: Investment Update as on {base_config['processing_date']}"
        send_email_with_pdfs(investor_email, subject, full_email_html, pdf_paths, smtp_cfg)
        print(f"✅ Sent to {investor_email} ({len(pdf_paths)} PDFs).")

    print("\n🎉 All done!")
