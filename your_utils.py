# your_utils.py
import re, requests

def download_sheet_as_xlsx(sheet_url: str, out_path: str):
    m = re.search(r"/d/([A-Za-z0-9_-]+)/", sheet_url)
    if not m:
        raise ValueError("Invalid Google Sheet URL.")
    fid = m.group(1)
    export = f"https://docs.google.com/spreadsheets/d/{fid}/export?format=xlsx"
    r = requests.get(export)
    r.raise_for_status()
    with open(out_path, "wb") as f:
        f.write(r.content)
    return out_path
