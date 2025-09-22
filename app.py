from flask import Flask, redirect, url_for, session, request, render_template, flash, Response, stream_with_context
from googleapiclient.errors import HttpError
import re
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import Flow
from googleapiclient.discovery import build
from twilio.rest import Client
import os
import json
from google.auth.transport.requests import Request
import time
from datetime import date, datetime

app = Flask(__name__)
app.secret_key = "secret_key_here"

# -------------------- Twilio Config (keep in code) --------------------
account_sid = "ACd187ce44440d09f47d36ba63539f36d2"
auth_token = "09fc81d78c7c214eb32271c8a852db00"
twilio_number = "+12708121647"
twilio_client = Client(account_sid, auth_token)

# -------------------- Google OAuth Config --------------------
os.environ["OAUTHLIB_INSECURE_TRANSPORT"] = "1"     
CLIENT_SECRETS_FILE = "/etc/secrets/client_secret.json"    # <-- Render Secret File path
CREDENTIALS_FILE = "/etc/secrets/credentials.json"        # <-- optional if you pre-upload credentials
SCOPES = [
    "https://www.googleapis.com/auth/drive.readonly", 
    "https://www.googleapis.com/auth/spreadsheets.readonly"
]

REDIRECT_URI = os.getenv("REDIRECT_URI", "https://svs-sms-sending.onrender.com/oauth2callback")

def credentials_to_dict(credentials):
    return {
        'token': credentials.token,
        'refresh_token': credentials.refresh_token,
        'token_uri': credentials.token_uri,
        'client_id': credentials.client_id,
        'client_secret': credentials.client_secret,
        'scopes': credentials.scopes
    }

def save_credentials(creds):
    """Secret files are read-only on Render, so save to local file instead."""
    with open("credentials_local.json", "w") as f:
        json.dump(credentials_to_dict(creds), f)

def load_credentials():
    """Try local file first, fallback to secret file."""
    if os.path.exists("credentials_local.json"):
        with open("credentials_local.json") as f:
            return json.load(f)
    elif os.path.exists(CREDENTIALS_FILE):
        with open(CREDENTIALS_FILE) as f:
            return json.load(f)
    return None

# -------------------- Utility Functions --------------------
PHONE_RE = re.compile(r'^\+\d{10,15}$')
def is_valid_phone(phone: str) -> bool:
    return bool(PHONE_RE.match(phone))

def fetch_sheet_values(service, spreadsheet_id, range_name):
    max_retries = 5
    for attempt in range(max_retries):
        try:
            result = service.spreadsheets().values().get(
                spreadsheetId=spreadsheet_id,
                range=range_name
            ).execute()
            return result.get("values", [])
        except HttpError as e:
            if e.resp.status in [503, 500, 429]:
                wait_time = (2 ** attempt)
                print(f"Service unavailable. Retrying in {wait_time} seconds...")
                time.sleep(wait_time)
            else:
                raise
    raise Exception("Failed to fetch sheet after multiple retries")

# -------------------- Routes --------------------
@app.route("/")
def index():
    creds_data = load_credentials()
    if not creds_data:
        return redirect(url_for("authorize"))
    return redirect(url_for("list_sheets"))

@app.route("/authorize")
def authorize():
    session.pop("credentials", None)
    flow = Flow.from_client_secrets_file(
        CLIENT_SECRETS_FILE,
        scopes=SCOPES,
        redirect_uri=REDIRECT_URI
    )
    authorization_url, state = flow.authorization_url(
        access_type="offline",
        include_granted_scopes="true",
        prompt="consent"
    )
    session["state"] = state
    return redirect(authorization_url)

@app.route("/oauth2callback")
def oauth2callback():
    state = session["state"]
    flow = Flow.from_client_secrets_file(
        CLIENT_SECRETS_FILE,
        scopes=SCOPES,
        state=state,
        redirect_uri=REDIRECT_URI
    )
    flow.fetch_token(authorization_response=request.url)

    creds = flow.credentials
    if not creds.refresh_token:
        return "No refresh token. Remove app access in Google account and try again.", 400

    session["credentials"] = credentials_to_dict(creds)
    save_credentials(creds)
    return redirect(url_for("list_sheets"))

@app.route("/list_sheets")
def list_sheets():
    creds_data = load_credentials()
    if not creds_data:
        return redirect(url_for("authorize"))

    creds = Credentials(**creds_data)
    drive_service = build("drive", "v3", credentials=creds)

    results = drive_service.files().list(
        q="mimeType='application/vnd.google-apps.spreadsheet'",
        fields="files(id, name)"
    ).execute()
    files = results.get("files", [])

    return render_template("sheets.html", files=files)

# -------------------- Send SMS --------------------
@app.route("/send_sms", methods=["POST"])
def send_sms():
    sheet_id = request.form.get("sheet_id")
    if not sheet_id:
        return "Error: Sheet not selected", 400

    creds_data = load_credentials()
    if not creds_data:
        return redirect(url_for("authorize"))

    creds = Credentials(**creds_data)
    sheets_service = build("sheets", "v4", credentials=creds)
    result = sheets_service.spreadsheets().values().get(spreadsheetId=sheet_id, range="A:Z").execute()

    rows = result.get("values", [])
    if not rows:
        return "No data found in the sheet", 400

    header = rows[0]
    normalized_header = [h.strip().lower() for h in header]

    try:
        phone_index = normalized_header.index("phone")
        name_index = normalized_header.index("name")
        hallticket_index = normalized_header.index("hallticket")
        date_index = normalized_header.index("date")
        attendance_index = normalized_header.index("attendance")
    except ValueError as e:
        return f"Missing column in sheet: {e}", 400

    today_str = date.today().isoformat()
    count = 0

    for row in rows[1:]:
        row_date = row[date_index].strip() if len(row) > date_index else ""
        attendance = row[attendance_index].strip().lower() if len(row) > attendance_index else ""

        if row_date != today_str or attendance == "present":
            continue

        phone = row[phone_index].strip() if len(row) > phone_index else ""
        name = row[name_index].strip() if len(row) > name_index else ""
        hallticket = row[hallticket_index].strip() if len(row) > hallticket_index else ""

        if phone:
            if not phone.startswith("+"):
                phone = "+91" + phone

            message_body = f"Hi {name}, hallticket no {hallticket}, you were marked absent today."
            twilio_client.messages.create(body=message_body, from_=twilio_number, to=phone)
            count += 1

    flash(f"✅ SMS sent to {count} absentees!")
    return redirect(url_for("list_sheets"))


@app.route("/stream_send_sms")
def stream_send_sms():
    """
    SSE endpoint — streams status updates while sending SMS.
    Sends SMS only for rows where the 'date' column matches today.
    Expects ?sheet_id=<sheetId> as a query param.
    """
    sheet_id = request.args.get("sheet_id")
    if not sheet_id:
        return "Error: sheet_id required", 400

    creds_data = load_credentials()
    if not creds_data:
        return redirect(url_for("authorize"))

    creds = Credentials(
        token=creds_data["token"],
        refresh_token=creds_data.get("refresh_token"),
        token_uri=creds_data["token_uri"],
        client_id=creds_data["client_id"],
        client_secret=creds_data["client_secret"],
        scopes=creds_data["scopes"],
    )

    # Refresh token if expired
    try:
        if creds.expired and creds.refresh_token:
            creds.refresh(Request())
            save_credentials(creds)
    except Exception as e:
        def refresh_failed(e=e):
            yield f"data: {json.dumps({'status':'error','msg':'Failed to refresh credentials: '+str(e)})}\n\n"
        return Response(stream_with_context(refresh_failed()), mimetype="text/event-stream")

    sheets_service = build("sheets", "v4", credentials=creds)

    # Fetch values
    try:
        result = sheets_service.spreadsheets().values().get(
            spreadsheetId=sheet_id,
            range="A:Z"
        ).execute()
    except HttpError as e:
        def err(e=e):
            yield f"data: {json.dumps({'status':'error','msg':'Google Sheets error: '+str(e)})}\n\n"
        return Response(stream_with_context(err()), mimetype="text/event-stream")

    rows = result.get("values", [])
    if not rows:
        def no_data():
            yield f"data: {json.dumps({'status':'error','msg':'No data found in the sheet'})}\n\n"
        return Response(stream_with_context(no_data()), mimetype="text/event-stream")

    header = rows[0]
    normalized_header = [h.strip().lower() for h in header]

    try:
        phone_index = normalized_header.index("phone")
        name_index = normalized_header.index("name")
        hallticket_index = normalized_header.index("hallticket")
        date_index = normalized_header.index("date")
        attendance_index = normalized_header.index("attendance")  # 👈 NEW
    except ValueError as e:
        def missing_col(e=e):
            yield f"data: {json.dumps({'status':'error','msg':f'Missing column in sheet: {str(e)}'})}\n\n"
        return Response(stream_with_context(missing_col()), mimetype="text/event-stream")

    today_str = date.today().isoformat()

    @stream_with_context
    def event_stream():
        count = 0
        row_num = 1
        for row in rows[1:]:
            row_num += 1
            row_date = row[date_index].strip() if len(row) > date_index else ""
            attendance = row[attendance_index].strip().lower() if len(row) > attendance_index else ""

            # Skip if not today
            if row_date != today_str:
                payload = {'status': 'skipped', 'msg': 'Not today', 'row': row_num}
                yield f"data: {json.dumps(payload)}\n\n"
                continue

            # Skip if present
            if attendance == "present":
                payload = {'status': 'skipped', 'msg': 'Marked present', 'row': row_num}
                yield f"data: {json.dumps(payload)}\n\n"
                continue

            phone = row[phone_index].strip() if len(row) > phone_index else ""
            name = row[name_index].strip() if len(row) > name_index else ""
            hallticket = row[hallticket_index].strip() if len(row) > hallticket_index else ""

            if phone and not phone.startswith("+"):
                phone = "+91" + phone

            if not phone:
                payload = {'status': 'skipped', 'msg': 'Empty phone', 'row': row_num, 'name': name}
                yield f"data: {json.dumps(payload)}\n\n"
                continue

            if not is_valid_phone(phone):
                payload = {'status': 'invalid', 'msg': 'Invalid phone format', 'row': row_num, 'phone': phone}
                yield f"data: {json.dumps(payload)}\n\n"
                continue

            # Send SMS only if absent
            message_body = f"Hi {name}, hallticket no {hallticket}, you were marked absent today."
            try:
                twilio_client.messages.create(
                    body=message_body,
                    from_=twilio_number,
                    to=phone
                )
                count += 1
                payload = {'status': 'sent', 'msg': 'Absent SMS sent', 'row': row_num, 'phone': phone, 'name': name}
                yield f"data: {json.dumps(payload)}\n\n"
            except Exception as e:
                payload = {'status': 'failed', 'msg': f'Twilio error: {str(e)}', 'row': row_num}
                yield f"data: {json.dumps(payload)}\n\n"

            time.sleep(0.15)

        summary = {'status': 'done', 'msg': f'Completed — SMS sent to {count} absentees', 'sent_count': count}
        yield f"event: done\ndata: {json.dumps(summary)}\n\n"

    return Response(event_stream(), mimetype="text/event-stream")

@app.route("/preview_sheet")
def preview_sheet():
    sheet_id = request.args.get("sheet_id")
    if not sheet_id:
        return "No sheet selected", 400

    creds_data = load_credentials()
    if not creds_data:
        return redirect(url_for("authorize"))

    creds = Credentials(
        token=creds_data["token"],
        refresh_token=creds_data["refresh_token"],
        token_uri=creds_data["token_uri"],
        client_id=creds_data["client_id"],
        client_secret=creds_data["client_secret"],
        scopes=creds_data["scopes"]
    )

    if creds.expired and creds.refresh_token:
        creds.refresh(Request())
        save_credentials(creds)

    sheets_service = build("sheets", "v4", credentials=creds)
    result = sheets_service.spreadsheets().values().get(
        spreadsheetId=sheet_id,
        range="A:Z"
    ).execute()

    values = result.get("values", [])
    if not values:
        return "No data found in the sheet", 404

    # Normalize headers
    headers = [h.strip() for h in values[0]]             # For display
    headers_lower = [h.strip().lower() for h in values[0]]  # For dict keys

    # Convert rows to dicts with lowercase keys
    rows = values[1:] if len(values) > 1 else []
    data = [dict(zip(headers_lower, row)) for row in rows]

    # Get filters from query params
    filter_date = request.args.get("date")
    branch = request.args.get("branch")
    dept = request.args.get("department")
    name = request.args.get("name")
    hallticket = request.args.get("hallticket")
    status = request.args.get("status")

   # Apply filters only if user provided any
    if filter_date or branch or dept or name or hallticket or status:
        filtered = []
        for row in data:
            if filter_date and row.get("date", "").strip() != filter_date:
                continue
            if branch and row.get("branch", "").strip().lower() != branch.lower():
                continue
            if dept and row.get("department", "").strip().lower() != dept.lower():
                continue
            if name and row.get("name", "").strip().lower() != name.lower():
                continue
            if hallticket and row.get("hallticket", "").strip().lower() != hallticket.lower():
                continue
            if status and row.get("attendance", "").strip().lower() != status.lower():
                continue
            filtered.append(row)
    else:
        filtered = data  # show all if no filter selected

    # Default: show all rows if no filters applied
    if not (filter_date or branch or dept or name or hallticket or status):
        filtered = data

    return render_template(
        "preview.html",
        sheet_name=sheet_id,
        headers=headers,
        headers_lower=headers_lower,
        filtered_data=filtered,
        filters={
            "date": filter_date or "",
            "branch": branch or "",
            "department": dept or "",
            "name": name or "",
            "hallticket": hallticket or "",
            "status": status or ""
        }
    )

if __name__ == "__main__":
    app.run(debug=True)


