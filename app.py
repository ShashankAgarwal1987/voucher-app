import os
import pandas as pd
import re
from datetime import datetime, timedelta
from flask import Flask, request, render_template, send_file
from fpdf import FPDF
from sentence_transformers import SentenceTransformer, util
import io

app = Flask(__name__)

# Load master Excel
MASTER_FILE = "Egypt_Itinerary_With_Formatted_Outputs.xlsx"
master_df = pd.read_excel(MASTER_FILE, sheet_name="Sheet1")
master_df.columns = master_df.columns.str.strip()

# Load embeddings model
print("Loading embeddings model…")
model = SentenceTransformer("all-MiniLM-L6-v2")

# Precompute embeddings for master data
particulars = master_df["Particular"].dropna().tolist()
particular_embeddings = model.encode(particulars, convert_to_tensor=True)


def semantic_match(activity, cutoff=0.6):
    """Find best semantic match for an activity using embeddings"""
    activity_embedding = model.encode(activity, convert_to_tensor=True)
    similarities = util.pytorch_cos_sim(activity_embedding, particular_embeddings)[0]
    best_idx = int(similarities.argmax())
    best_score = float(similarities[best_idx])

    if best_score >= cutoff:
        return particulars[best_idx], best_score
    else:
        return None, best_score


def transform_itinerary_semantic(start_date_str, itinerary_text):
    start_date = datetime.strptime(start_date_str, "%d-%b-%Y")
    output_rows = []

    lines = itinerary_text.strip().split("\n")

    for i, line in enumerate(lines):
        match = re.split(r"Day\s*\d+:\s*", line, flags=re.IGNORECASE)
        activities_str = match[1] if len(match) > 1 else line
        activities = [a.strip() for a in activities_str.split("+")]

        formatted_texts = []
        for activity in activities:
            # Try exact match first
            match_row = master_df[
                master_df["Particular"].str.contains(
                    re.escape(activity), case=False, na=False
                )
            ]
            if not match_row.empty:
                formatted_texts.append(str(match_row.iloc[0]["Formatted Output"]))
            else:
                # Semantic match
                best_match, score = semantic_match(activity)
                if best_match:
                    match_row = master_df[master_df["Particular"] == best_match]
                    formatted_texts.append(str(match_row.iloc[0]["Formatted Output"]))
                else:
                    formatted_texts.append(f"⚠️ No match found for: {activity}")

        combined_output = "\n".join(formatted_texts)
        date = (start_date + timedelta(days=i)).strftime("%d-%b-%Y")
        output_rows.append({"Date": date, "Formatted Output": combined_output})

    return pd.DataFrame(output_rows)


@app.route("/", methods=["GET", "POST"])
def index():
    if request.method == "POST":
        start_date = request.form["start_date"]
        itinerary_text = request.form["itinerary"]

        # Transform itinerary
        df = transform_itinerary_semantic(start_date, itinerary_text)

        # Generate PDF
        pdf = FPDF()
        pdf.add_page()
        pdf.set_font("Arial", size=12)

        for _, row in df.iterrows():
            pdf.multi_cell(0, 10, f"Date: {row['Date']}", align="L")
            pdf.multi_cell(0, 10, f"{row['Formatted Output']}\n", align="L")

        pdf_bytes = io.BytesIO()
        pdf.output(pdf_bytes, "F")
        pdf_bytes.seek(0)

        return send_file(
            pdf_bytes, download_name="Service_Voucher.pdf", as_attachment=True
        )

    return render_template("index.html")


if __name__ == "__main__":
    app.run(host="0.0.0.0", port=5000)
