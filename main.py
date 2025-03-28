from datetime import datetime
import json
import os.path
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from googleapiclient.discovery import build
from sheet_helper import create_sheet_and_chart

SCOPES = [
    'https://www.googleapis.com/auth/presentations',
    'https://www.googleapis.com/auth/drive',
    'https://www.googleapis.com/auth/spreadsheets'
]

def get_credentials():
    creds = None
    if os.path.exists('token.json'):
        creds = Credentials.from_authorized_user_file('token.json', SCOPES)
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file('credentials.json', SCOPES)
            creds = flow.run_local_server(port=0)
        with open('token.json', 'w') as token:
            token.write(creds.to_json())
    return creds

def load_event_data():
    with open('event_data.json', 'r') as f:
        return json.load(f)['data']

def build_slide_content(event):
    return f"""Event Title: {event.get('event_title')}
Start Date: {event.get('start_date')}
End Date: {event.get('end_date')}
Repeat Type: {event.get('repeat_type')}
Duration (hrs): {event.get('duration')}
Inference Type(s): {', '.join(event.get('inference_type', []))}
Device(s) Used: {len(event.get('devices', []))}
Average Headcount: {event.get('analytics_summary', {}).get('average_count')}
Max Headcount: {event.get('analytics_summary', {}).get('max_count')}

Thank you for being part of this event!
"""

def create_presentation(slides_service, drive_service, content, title, sheet_id, chart_id):
    copied = drive_service.files().copy(
        fileId='1r5MIK3QZzhGK1g7USCJQ4bgb9ycX02eyKw22aBHshp8',
        body={'name': f'HawkEye Report - {title}'}
    ).execute()

    presentation_id = copied['id']
    print(f"✅ Created: https://docs.google.com/presentation/d/{presentation_id}/edit")

    slide_id = f"slide_{title[:10].replace(' ', '_')}"
    title_id = f"title_{title[:10].replace(' ', '_')}"
    body_id = f"body_{title[:10].replace(' ', '_')}"
    chart_id_obj = f"chart_{title[:10].replace(' ', '_')}"

    requests = [
        {"createSlide": {"objectId": slide_id, "insertionIndex": 0}},
        {"createShape": {
            "objectId": title_id,
            "shapeType": "TEXT_BOX",
            "elementProperties": {
                "pageObjectId": slide_id,
                "transform": {"scaleX": 1, "scaleY": 1, "translateX": 50, "translateY": 20, "unit": "PT"},
                "size": {"height": {"magnitude": 50, "unit": "PT"}, "width": {"magnitude": 400, "unit": "PT"}}
            }
        }},
        {"insertText": {"objectId": title_id, "insertionIndex": 0, "text": f"Event Report - {title}"}},
        {"updateTextStyle": {
            "objectId": title_id, "textRange": {"type": "ALL"},
            "style": {"fontSize": {"magnitude": 24, "unit": "PT"}, "bold": True},
            "fields": "fontSize,bold"
        }},
        {"createShape": {
            "objectId": body_id,
            "shapeType": "TEXT_BOX",
            "elementProperties": {
                "pageObjectId": slide_id,
                "transform": {"scaleX": 1, "scaleY": 1, "translateX": 50, "translateY": 90, "unit": "PT"},
                "size": {"height": {"magnitude": 300, "unit": "PT"}, "width": {"magnitude": 500, "unit": "PT"}}
            }
        }},
        {"insertText": {"objectId": body_id, "insertionIndex": 0, "text": content}},
        {"updateTextStyle": {
            "objectId": body_id, "textRange": {"type": "ALL"},
            "style": {"fontSize": {"magnitude": 14, "unit": "PT"}},
            "fields": "fontSize"
        }},
        {
            "createSheetsChart": {
                "objectId": chart_id_obj,
                "spreadsheetId": sheet_id,
                "chartId": chart_id,
                "linkingMode": "LINKED",
                "elementProperties": {
                    "pageObjectId": slide_id,
                    "transform": {
                        "scaleX": 1, "scaleY": 1,
                        "translateX": 50, "translateY": 400,
                        "unit": "PT"
                    },
                    "size": {
                        "height": {"magnitude": 300, "unit": "PT"},
                        "width": {"magnitude": 500, "unit": "PT"}
                    }
                }
            }
        }
    ]

    slides_service.presentations().batchUpdate(
        presentationId=presentation_id,
        body={"requests": requests}
    ).execute()

def main():
    creds = get_credentials()
    slides_service = build('slides', 'v1', credentials=creds)
    drive_service = build('drive', 'v3', credentials=creds)
    sheets_service = build('sheets', 'v4', credentials=creds)

    events = load_event_data()
    today = str(datetime.today().date())
    for event in events:
        if event.get('end_date') == today:
            content = build_slide_content(event)
            title = event.get("event_title", "Untitled")
            sheet_id, chart_id = create_sheet_and_chart(sheets_service, event)
            create_presentation(slides_service, drive_service, content, title, sheet_id, chart_id)

if __name__ == '__main__':
    main()
