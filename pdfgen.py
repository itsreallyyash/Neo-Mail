from jinja2 import Environment, FileSystemLoader
from weasyprint import HTML

# Your investment data (replace with dynamic load as needed)
company_data = {
    "company_name": "Resilient Innovations Private Limited",
    "borrower_profile": (
        "BharatPe is a leading Indian fintech platform focused on empowering small merchants and kirana stores "
        "by enabling UPI-based interoperable QR payments. QR codes are deployed all over India and Merchants are "
        "monitored and assessed by transaction volumes and size, and hence, these merchants are offered with POS and "
        "Speaker machines and merchant loans. <br><br>"
        "The company has presence of marquee investors like Peak XV, Insight Partners, Ribbit, Coatue, Tiger, "
        "Steadview, etc. and the company is run by a professional management with ex senior bankers on board including Nalin Negi, Rajnish Kumar, etc."
    ),
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
    # Optionally, use local images or URLs for your flower SVG/PNG and logo
    "decor_flower": "flower.svg",  # Path to your flower SVG/PNG
    "contact_logo": "neo-logo.png" # Path to your logo PNG
}

env = Environment(loader=FileSystemLoader('.'))
template = env.get_template("template.html")
html_out = template.render(**company_data)

HTML(string=html_out, base_url='.').write_pdf("investment_report.pdf")

print("PDF generated as investment_report.pdf")
