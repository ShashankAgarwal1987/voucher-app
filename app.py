from flask import Flask, render_template, request, send_file
import pandas as pd
from fpdf import FPDF
import io

app = Flask(__name__)

df = None  # Global Excel data

@app.route("/", methods=["GET", "POST"])
def index():
    global df
    if request.method == "POST":
        if "excel" in request.files:  # Upload Excel
            excel_file = request.files["excel"]
            df = pd.read_excel(excel_file)
            return render_template("index.html", message="Excel uploaded successfully!")

        if "query" in request.form:  # Generate Voucher
            query = request.form["query"].lower()
            if df is None:
                return render_template("index.html", error="Upload Excel first!")

            match = df[df["Particular"].str.lower().str.contains(query, na=False)]
            if match.empty:
                return render_template("index.html", error="No matching tour found!")

            output_text = match.iloc[0]["Formatted Output"]

            # Generate PDF in memory
            pdf = FPDF()
            pdf.add_page()
            pdf.set_font("Arial", size=12)
            pdf.multi_cell(0, 10, f"Service Voucher\n\n{output_text}")

            pdf_bytes = io.BytesIO()
            pdf.output(pdf_bytes)
            pdf_bytes.seek(0)

            return send_file(pdf_bytes, as_attachment=True, download_name="voucher.pdf")

    return render_template("index.html")

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=5000)
