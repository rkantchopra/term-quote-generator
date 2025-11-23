# app.py
# Full Flask app: file uploads + Google Sheets (Service Account) integration
import os
import json
import tempfile
import traceback
from flask import Flask, render_template, request, send_file, abort, redirect, url_for
import pandas as pd

from google.oauth2 import service_account
from googleapiclient.discovery import build

# Your quote generator (must exist in project)
from quote_generator import make_term_quote_from_excel

# ---------- Flask / config ----------
app = Flask(__name__)
app.secret_key = os.environ.get("FLASK_SECRET", "insecure-dev-secret")
UPLOAD_FOLDER = os.path.join(os.getcwd(), "uploads")
os.makedirs(UPLOAD_FOLDER, exist_ok=True)

# ---------- Google Sheets API helpers ----------
SCOPES = ["https://www.googleapis.com/auth/spreadsheets.readonly"]

def get_service_account_creds():
    """
    Create Credentials either from:
      - environment variable GOOGLE_APPLICATION_CREDENTIALS (path to file), OR
      - environment variable GOOGLE_SA_JSON containing raw JSON text.
    """
    # 1) If GOOGLE_APPLICATION_CREDENTIALS is set and file exists, use it
    gpath = os.environ.get("GOOGLE_APPLICATION_CREDENTIALS")
    if gpath and os.path.exists(gpath):
        creds = service_account.Credentials.from_service_account_file(gpath, scopes=SCOPES)
        return creds

    # 2) Otherwise try GOOGLE_SA_JSON (the JSON content stored as a secret)
    sa_json = os.environ.get("GOOGLE_SA_JSON")
    if sa_json:
        info = json.loads(sa_json)
        creds = service_account.Credentials.from_service_account_info(info, scopes=SCOPES)
        return creds

    raise RuntimeError("No Google service account credentials found. Set GOOGLE_APPLICATION_CREDENTIALS or GOOGLE_SA_JSON.")

def read_sheet_to_df(spreadsheet_id: str, sheet_name: str) -> pd.DataFrame:
    """Read a sheet by name into a pandas DataFrame using Sheets API."""
    creds = get_service_account_creds()
    service = build('sheets', 'v4', credentials=creds, cache_discovery=False)
    sheet = service.spreadsheets()
    try:
        resp = sheet.values().get(spreadsheetId=spreadsheet_id, range=sheet_name).execute()
    except Exception as e:
        # bubble up helpful message
        raise RuntimeError(f"Error reading sheet '{sheet_name}' from {spreadsheet_id}: {e}")
    values = resp.get('values', [])
    if not values:
        return pd.DataFrame()
    df = pd.DataFrame(values[1:], columns=values[0])
    return df

def download_and_process_spreadsheet(spreadsheet_id: str) -> str:
    """
    Reads 'Client Details' and 'Premiums' sheets, writes a combined XLSX, calls
    make_term_quote_from_excel(), and returns path to generated DOCX.
    """
    client_df = read_sheet_to_df(spreadsheet_id, "Client Details")
    prem_df   = read_sheet_to_df(spreadsheet_id, "Premiums")

    if client_df.empty and prem_df.empty:
        raise RuntimeError("Both 'Client Details' and 'Premiums' are empty or missing.")

    tmpdir = tempfile.gettempdir()
    combined_path = os.path.join(tmpdir, f"combined_{spreadsheet_id}.xlsx")
    with pd.ExcelWriter(combined_path, engine="openpyxl") as writer:
        if not client_df.empty:
            client_df.to_excel(writer, sheet_name="Client Details", index=False)
        else:
            pd.DataFrame().to_excel(writer, sheet_name="Client Details", index=False)
        if not prem_df.empty:
            prem_df.to_excel(writer, sheet_name="Premiums", index=False)
        else:
            pd.DataFrame().to_excel(writer, sheet_name="Premiums", index=False)

    output_path = make_term_quote_from_excel(combined_path)
    return output_path

# ---------- Upload / web UI routes ----------
@app.route("/", methods=["GET", "POST"])
def index():
    # GET -> show upload form
    # POST -> accept either:
    #  - a single combined uploaded Excel (file form name="file")
    #  - two uploaded files (client_file + premium_file)
    #  - or handled by other endpoints
    if request.method == "POST":
        # file upload handling
        single_file = request.files.get("file")
        client_file = request.files.get("client_file")
        premium_file = request.files.get("premium_file")

        # Single combined file
        if single_file and (not client_file and not premium_file):
            path = os.path.join(UPLOAD_FOLDER, "uploaded_combined.xlsx")
            single_file.save(path)
            try:
                out = make_term_quote_from_excel(path)
            except Exception as e:
                return f"Error generating quote: {e}\n\n{traceback.format_exc()}", 500
            return send_file(out, as_attachment=True)

        # Two separate files
        if client_file and premium_file:
            client_path = os.path.join(UPLOAD_FOLDER, "client.xlsx")
            premium_path = os.path.join(UPLOAD_FOLDER, "premium.xlsx")
            client_file.save(client_path)
            premium_file.save(premium_path)
            combined_path = os.path.join(UPLOAD_FOLDER, "combined_input.xlsx")
            try:
                with pd.ExcelWriter(combined_path, engine="openpyxl") as writer:
                    pd.read_excel(client_path).to_excel(writer, sheet_name="Client Details", index=False)
                    pd.read_excel(premium_path).to_excel(writer, sheet_name="Premiums", index=False)
                out = make_term_quote_from_excel(combined_path)
            except Exception as e:
                return f"Error combining files: {e}\n\n{traceback.format_exc()}", 500
            return send_file(out, as_attachment=True)

        return "Please upload either a combined Excel (file) or both client_file and premium_file.", 400

    # GET -> show template
    return render_template("index.html")

# ---------- Sheets API endpoint ----------
@app.route("/process_by_sheetid", methods=["POST"])
def process_by_sheetid():
    """
    Trigger processing by sending spreadsheet_id via POST JSON or form.
    Optional security: header 'X-SHEET-SECRET' or form field 'secret' must match env SHEET_SECRET if set.
    """
    # security check (optional)
    env_secret = os.environ.get("SHEET_SECRET")
    if env_secret:
        header_secret = request.headers.get("X-SHEET-SECRET")
        form_secret = request.form.get("secret")
        if not (header_secret == env_secret or form_secret == env_secret):
            return {"error": "Missing or invalid secret"}, 401

    data = request.get_json(silent=True) or request.form
    sheet_id = data.get("spreadsheet_id") or data.get("sheet_id") or data.get("sheetId")
    if not sheet_id:
        return {"error": "missing spreadsheet_id"}, 400

    try:
        out = download_and_process_spreadsheet(sheet_id)
    except Exception as e:
        return {"error": str(e), "trace": traceback.format_exc()}, 500

    return send_file(out, as_attachment=True)

# ---------- Utility route for local quick-test (uses uploaded sample) ----------
@app.route("/local_test", methods=["GET"])
def local_test():
    """
    Quick local test that uses the sample Excel you uploaded to the server.
    The developer provided sample path: /mnt/data/Term_Quote_Input_Template_Dual_Pay.xlsx
    This route exists only to help test locally — remove in production if you want.
    """
    sample_path = "/mnt/data/Term_Quote_Input_Template_Dual_Pay.xlsx"
    if not os.path.exists(sample_path):
        return {"error": f"Sample not found at {sample_path}"}, 404
    try:
        out = make_term_quote_from_excel(sample_path)
    except Exception as e:
        return {"error": str(e), "trace": traceback.format_exc()}, 500
    return send_file(out, as_attachment=True)

# ---------- Run ----------
if __name__ == "__main__":
    # debug on local dev only
    app.run(debug=True, host="0.0.0.0", port=5000)
