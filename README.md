# Term Quote Generator

## Run locally
1. python -m venv .venv
2. source .venv/bin/activate  (or .venv\\Scripts\\activate on Windows)
3. pip install -r requirements.txt
4. FLASK_APP=app.py flask run

Open http://localhost:5000

## Deploy to Render (Docker):
1. Push this repo to GitHub.
2. Go to https://render.com -> New -> Web Service.
3. Connect GitHub and select your repo + branch.
4. Choose Docker environment (auto-detect Dockerfile).
5. Create service. Render will build and deploy.

## Excel format expected
- Sheet: `Client Details` — columns like `Client Name`, `DOB`, `Age`, `City`, `Sum Assured`, `Policy Term`, `PPT`
- Sheet: `Premiums` — columns: `Insurance Company`, `Plan Name`, `Regular Premium`, `10 Pay Premium`, `Special Notes`
- Sheet: `Final Notes` (optional) — any advisory text

## Notes
- The DOCX is generated using python-docx.
- If you want the exact formatting from your health generator, paste that code and I will port exact helper functions & styles into this project.
s