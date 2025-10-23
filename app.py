from flask import Flask, render_template, request, send_file
import pandas as pd
from fpdf import FPDF
import io

app = Flask(__name__)

# Home page
@app.route("/", methods=["GET", "POST"])
def index():
    if request.method == "POST":
        # Read uploaded Excel file
        excel_file = request.files["excel_file"]
        df = pd.read_excel(excel_file)

        # Get tour/transfer name
        tour_name = request.form["tour_name"]

        # Try to find matching row in Excel
        match = df[df["Particular"].str.contains(tour_name, case=False, na=False)]

        if match.empty:
            return "❌ Tour not found in the uploaded Excel file."

        # Extract details
        city = match.iloc[0]["City / Tour / Transfer"]
        particular = match.iloc[0]["Particular"]
        description = match.iloc[0]["Tour Description"]

        # --- Generate PDF voucher ---
        pdf = FPDF()
        pdf.add_page()
        pdf.set_font("Arial", size=12)

        pdf.cell(200, 10, txt="Travel LYKKE - Service Voucher", ln=True, align="C")
        pdf.ln(10)
        pdf.cell(200, 10, txt=f"City/Tour/Transfer: {city}", ln=True)
        pdf.cell(200, 10, txt=f"Particular: {particular}", ln=True)
        pdf.multi_cell(0, 10, txt=f"Description: {description}")

        # Output PDF as bytes
        pdf_bytes = pdf.output(dest="S").encode("latin1")

        return send_file(
            io.BytesIO(pdf_bytes),
            mimetype="application/pdf",
            as_attachment=True,
            download_name="voucher.pdf"
        )

    # GET request → show upload form
    return render_template("index.html")
