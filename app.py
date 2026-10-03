import os
from io import BytesIO
from datetime import datetime, date

import requests
from flask import Flask, render_template, request, jsonify, send_file
from openpyxl import Workbook

app = Flask(__name__)

SUPABASE_URL = os.environ.get("SUPABASE_URL", "").rstrip("/")
SUPABASE_KEY = os.environ.get("SUPABASE_KEY", "")
TABLE = "service_entries"

COLUMNS = [
    "Sl. No.",
    "Inward Date",
    "Outward Date",
    "Model No.",
    "Customer Details",
    "Complaints",
    "Status",
    "Remarks",
    "Item Replaced",
    "Final Remark",
    "Service Charge",
]

def check_config():
    if not SUPABASE_URL or not SUPABASE_KEY:
        raise RuntimeError("SUPABASE_URL and SUPABASE_KEY environment variables are required.")

def headers():
    return {
        "apikey": SUPABASE_KEY,
        "Authorization": f"Bearer {SUPABASE_KEY}",
        "Content-Type": "application/json",
    }

def supabase_request(method, path, **kwargs):
    check_config()
    url = f"{SUPABASE_URL}/rest/v1/{path}"
    h = headers()
    h.update(kwargs.pop("headers", {}))
    response = requests.request(method, url, headers=h, timeout=30, **kwargs)
    if not response.ok:
        raise RuntimeError(f"Supabase error {response.status_code}: {response.text[:500]}")
    return response

def get_entries():
    response = supabase_request(
        "GET",
        f"{TABLE}?select=*&order=id.asc"
    )
    return response.json()

@app.route("/")
def index():
    return render_template("index.html")

@app.route("/api/count")
def count():
    try:
        entries = get_entries()
        return jsonify({"count": len(entries)})
    except Exception as e:
        return jsonify({"error": str(e)}), 500

@app.route("/api/entries", methods=["GET"])
def entries():
    try:
        return jsonify(get_entries())
    except Exception as e:
        return jsonify({"error": str(e)}), 500

@app.route("/api/entries", methods=["POST"])
def add_entry():
    try:
        data = request.get_json(force=True) or {}

        required = ["inward_date", "model_no", "customer_details", "complaints", "status"]
        for field in required:
            if not str(data.get(field, "")).strip():
                return jsonify({"error": f"{field.replace('_', ' ').title()} is required."}), 400

        if data["status"] not in ("Complete", "Pending"):
            return jsonify({"error": "Status must be Complete or Pending."}), 400

        payload = {
            "inward_date": data.get("inward_date") or date.today().isoformat(),
            "outward_date": data.get("outward_date") or None,
            "model_no": str(data.get("model_no", "")).strip(),
            "customer_details": str(data.get("customer_details", "")).strip(),
            "complaints": str(data.get("complaints", "")).strip(),
            "status": data.get("status"),
            "remarks": str(data.get("remarks", "")).strip(),
            "item_replaced": str(data.get("item_replaced", "")).strip(),
            "final_remark": str(data.get("final_remark", "")).strip(),
            "service_charge": float(data.get("service_charge") or 0),
        }

        response = supabase_request(
            "POST",
            TABLE,
            json=payload,
            headers={"Prefer": "return=representation"}
        )
        return jsonify(response.json()[0]), 201

    except ValueError:
        return jsonify({"error": "Service Charge must be a number."}), 400
    except Exception as e:
        return jsonify({"error": str(e)}), 500

@app.route("/api/clear", methods=["POST"])
def clear():
    try:
        # Deletes every row. The database table/header structure remains.
        supabase_request("DELETE", f"{TABLE}?id=gte.0")
        return jsonify({"message": "All service data cleared."})
    except Exception as e:
        return jsonify({"error": str(e)}), 500

@app.route("/api/download")
def download():
    try:
        rows = get_entries()

        wb = Workbook()
        ws = wb.active
        ws.title = "Service Entries"
        ws.append(COLUMNS)

        for i, row in enumerate(rows, start=1):
            ws.append([
                i,
                row.get("inward_date") or "",
                row.get("outward_date") or "",
                row.get("model_no") or "",
                row.get("customer_details") or "",
                row.get("complaints") or "",
                row.get("status") or "",
                row.get("remarks") or "",
                row.get("item_replaced") or "",
                row.get("final_remark") or "",
                row.get("service_charge") or 0,
            ])

        # Basic formatting
        for cell in ws[1]:
            cell.font = cell.font.copy(bold=True)
        ws.freeze_panes = "A2"
        widths = [10, 14, 14, 16, 28, 35, 14, 28, 24, 28, 16]
        for idx, width in enumerate(widths, start=1):
            ws.column_dimensions[chr(64 + idx)].width = width

        output = BytesIO()
        wb.save(output)
        output.seek(0)

        filename = f"Service_Complaint_Tracker_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
        return send_file(
            output,
            as_attachment=True,
            download_name=filename,
            mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    except Exception as e:
        return jsonify({"error": str(e)}), 500

@app.get("/health")
def health():
    return jsonify({"status": "ok"})

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=int(os.environ.get("PORT", 3000)))
