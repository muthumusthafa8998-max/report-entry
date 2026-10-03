SERVICE ENTRY APP

Flask service-entry app for mobile entry into the master Excel workbook.

Local run:
  python -m venv .venv
  .venv\Scripts\activate
  pip install -r requirements.txt
  python app.py

Render deployment:
  Build: pip install -r requirements.txt
  Start: gunicorn app:app

The Render deployment uses DATA_DIR=/var/data and a persistent disk mounted at /var/data so entries saved to Service_Complaint_Tracker.xlsx survive restarts and deploys.

Features:
- Mobile portrait entry
- Complete / Pending status
- Automatic Sl. No.
- Save Entry
- Download latest Excel
- Clear all Excel data
