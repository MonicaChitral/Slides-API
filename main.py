# Updated main.py to use single-slide portrait layout with table and 2 charts
from datetime import datetime
import json
import os.path
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from googleapiclient.discovery import build
from sheet_helper import create_sheet_and_charts

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

def create_presentation(slides_service, drive_service, event, sheet_id, chart_ids):
    title = event.get("event_title", "Untitled")
    logo_url = event.get("latest_image_url_id")  # Get the logo URL from event data
    copied = drive_service.files().copy(
        fileId='1BgMBoNIGRXCMzBJ26TLb8HUDbqcwGTHFfr5Bu1RyKeY',
        body={'name': f'HawkEye Report - {title}'}
    ).execute()
    presentation_id = copied['id']
    print(f"✅ Created: https://docs.google.com/presentation/d/{presentation_id}/edit")

    slide_id = f"slide_{title[:10].replace(' ', '_')}"
    title_id = f"title_{title[:10].replace(' ', '_')}"
    table_id = f"table_{title[:10].replace(' ', '_')}"
    kpi1_id = f"kpi1_{title[:10].replace(' ', '_')}"
    kpi2_id = f"kpi2_{title[:10].replace(' ', '_')}"

    # Get analytics data from event
    analytics = event.get('analytics', [])
    avg_count = event.get('analytics_summary', {}).get('average_count', 0)
    max_count = event.get('analytics_summary', {}).get('max_count', 0)

    # Get event details for the table - simplified
    event_details = [
        ["Information", "Details"],
        ["Name", event.get("event_title", "")],
        ["Date", event.get("start_date", "")],
        ["Location", "12 Melbourne Oxford"],
        ["Total Attendance", "5,000"]
    ]

    # Style table header
    {"updateTableCellProperties": {
        "objectId": table_id,
        "tableRange": {
            "location": {"rowIndex": 0, "columnIndex": 0},
            "rowSpan": 1,
            "columnSpan": 2
        },
        "tableCellProperties": {
            "tableCellBackgroundFill": {
                "solidFill": {
                    "color": {"rgbColor": {"red": 0.23, "green": 0.51, "blue": 0.79}}
                }
            }
        },
        "fields": "tableCellBackgroundFill"
    }},
    # Style header text
    {"updateTextStyle": {
        "objectId": table_id,
        "cellLocation": {"rowIndex": 0},
        "rowSpan": 1,
        "columnSpan": 2,
        "style": {
            "foregroundColor": {"opaqueColor": {"rgbColor": {"red": 1, "green": 1, "blue": 1}}},
            "bold": True
        },
        "fields": "foregroundColor,bold"
    }},
    # Style table content text
    {"updateTextStyle": {
        "objectId": table_id,
        "style": {
            "foregroundColor": {"opaqueColor": {"rgbColor": {"red": 1, "green": 1, "blue": 1}}}
        },
        "fields": "foregroundColor"
    }}

    # Define KPI cards with simpler styling
    kpi_cards = [
        {
            "id": kpi1_id,
            "title": "Total Count",
            "value": "30",  # Hardcoded for now, replace with actual data
            "change": "+30% vs last period",
            "x": 40,
            "y": 300,  # Adjusted position
            "width": 200,
            "height": 80,
            "background": {"red": 1, "green": 1, "blue": 1}  # White background
        },
        {
            "id": kpi2_id,
            "title": "Aggregated Count",
            "value": "40",  # Hardcoded for now, replace with actual data
            "change": "+40% vs last period",
            "x": 280,
            "y": 300,  # Adjusted position
            "width": 200,
            "height": 80,
            "background": {"red": 1, "green": 1, "blue": 1}  # White background
        }
    ]

    requests = [
        # Create slide
        {"createSlide": {"objectId": slide_id, "insertionIndex": 0}},

        # Add logo (only if URL is available)
        *([{
            "createImage": {
                "url": logo_url,
                "elementProperties": {
                    "pageObjectId": slide_id,
                    "size": {"width": {"magnitude": 100, "unit": "PT"}, "height": {"magnitude": 40, "unit": "PT"}},
                    "transform": {"scaleX": 1, "scaleY": 1, "translateX": 40, "translateY": 20, "unit": "PT"}
                }
            }
        }] if logo_url else []),

        # Title
        {"createShape": {
            "objectId": title_id,
            "shapeType": "TEXT_BOX",
            "elementProperties": {
                "pageObjectId": slide_id,
                "transform": {"scaleX": 1, "scaleY": 1, "translateX": 40, "translateY": 40, "unit": "PT"},
                "size": {"height": {"magnitude": 40, "unit": "PT"}, "width": {"magnitude": 400, "unit": "PT"}}
            }
        }},
        {"insertText": {"objectId": title_id, "text": "EVENT REPORT"}},
        {"updateTextStyle": {
            "objectId": title_id,
            "style": {
                "fontSize": {"magnitude": 24, "unit": "PT"},
                "foregroundColor": {"opaqueColor": {"rgbColor": {"red": 0.2, "green": 0.2, "blue": 0.2}}},
                "bold": True
            },
            "fields": "fontSize,foregroundColor,bold"
        }},

        # Create table
        {"createTable": {
            "objectId": table_id,
            "rows": len(event_details),
            "columns": 2,
            "elementProperties": {
                "pageObjectId": slide_id,
                "transform": {"scaleX": 1, "scaleY": 1, "translateX": 40, "translateY": 100, "unit": "PT"},
                "size": {"height": {"magnitude": 120, "unit": "PT"}, "width": {"magnitude": 500, "unit": "PT"}}
            }
        }}
    ]

    # Add table content
    for i, row in enumerate(event_details):
        for j, cell in enumerate(row):
            requests.append({
                "insertText": {
                    "objectId": table_id,
                    "cellLocation": {"rowIndex": i, "columnIndex": j},
                    "text": str(cell)
                }
            })

    # Style table header
    requests.extend([
        {"updateTableCellProperties": {
            "objectId": table_id,
            "tableRange": {
                "location": {"rowIndex": 0, "columnIndex": 0},
                "rowSpan": 1,
                "columnSpan": 2
            },
            "tableCellProperties": {
                "tableCellBackgroundFill": {
                    "solidFill": {
                        "color": {"rgbColor": {"red": 0.23, "green": 0.51, "blue": 0.79}}
                    }
                }
            },
            "fields": "tableCellBackgroundFill"
        }},
        {"updateTextStyle": {
            "objectId": table_id,
            "cellLocation": {"rowIndex": 0, "columnIndex": 0},
            "style": {
                "foregroundColor": {"opaqueColor": {"rgbColor": {"red": 1, "green": 1, "blue": 1}}},
                "bold": True
            },
            "fields": "foregroundColor,bold"
        }}
    ])

    # Add KPI cards
    for kpi in kpi_cards:
        requests.extend([
            # KPI card background
            {"createShape": {
                "objectId": kpi["id"],
                "shapeType": "RECTANGLE",
                "elementProperties": {
                    "pageObjectId": slide_id,
                    "transform": {"scaleX": 1, "scaleY": 1, "translateX": kpi["x"], "translateY": kpi["y"], "unit": "PT"},
                    "size": {"height": {"magnitude": kpi["height"], "unit": "PT"}, 
                           "width": {"magnitude": kpi["width"], "unit": "PT"}}
                }
            }},
            {"updateShapeProperties": {
                "objectId": kpi["id"],
                "shapeProperties": {
                    "shapeBackgroundFill": {
                        "solidFill": {"color": {"rgbColor": kpi["background"]}}
                    }
                },
                "fields": "shapeBackgroundFill"
            }},
            # KPI title
            {"createShape": {
                "objectId": f"{kpi['id']}_title",
                "shapeType": "TEXT_BOX",
                "elementProperties": {
                    "pageObjectId": slide_id,
                    "transform": {"scaleX": 1, "scaleY": 1, 
                                "translateX": kpi["x"] + 10, 
                                "translateY": kpi["y"] + 10, "unit": "PT"},
                    "size": {"height": {"magnitude": 20, "unit": "PT"}, 
                           "width": {"magnitude": kpi["width"] - 20, "unit": "PT"}}
                }
            }},
            {"insertText": {
                "objectId": f"{kpi['id']}_title",
                "text": kpi["title"]
            }},
            {"updateTextStyle": {
                "objectId": f"{kpi['id']}_title",
                "style": {
                    "foregroundColor": {"opaqueColor": {"rgbColor": {"red": 0.2, "green": 0.2, "blue": 0.2}}},
                    "fontSize": {"magnitude": 14, "unit": "PT"}
                },
                "fields": "foregroundColor,fontSize"
            }},
            # KPI value
            {"createShape": {
                "objectId": f"{kpi['id']}_value",
                "shapeType": "TEXT_BOX",
                "elementProperties": {
                    "pageObjectId": slide_id,
                    "transform": {"scaleX": 1, "scaleY": 1, 
                                "translateX": kpi["x"] + 10, 
                                "translateY": kpi["y"] + 35, "unit": "PT"},
                    "size": {"height": {"magnitude": 30, "unit": "PT"}, 
                           "width": {"magnitude": kpi["width"] - 20, "unit": "PT"}}
                }
            }},
            {"insertText": {
                "objectId": f"{kpi['id']}_value",
                "text": kpi["value"]
            }},
            {"updateTextStyle": {
                "objectId": f"{kpi['id']}_value",
                "style": {
                    "fontSize": {"magnitude": 24, "unit": "PT"},
                    "bold": True,
                    "foregroundColor": {"opaqueColor": {"rgbColor": {"red": 0.2, "green": 0.2, "blue": 0.2}}}
                },
                "fields": "fontSize,bold,foregroundColor"
            }},
            # KPI change indicator
            {"createShape": {
                "objectId": f"{kpi['id']}_change",
                "shapeType": "TEXT_BOX",
                "elementProperties": {
                    "pageObjectId": slide_id,
                    "transform": {"scaleX": 1, "scaleY": 1, 
                                "translateX": kpi["x"] + 10, 
                                "translateY": kpi["y"] + 70, "unit": "PT"},
                    "size": {"height": {"magnitude": 20, "unit": "PT"}, 
                           "width": {"magnitude": kpi["width"] - 20, "unit": "PT"}}
                }
            }},
            {"insertText": {
                "objectId": f"{kpi['id']}_change",
                "text": kpi["change"]
            }},
            {"updateTextStyle": {
                "objectId": f"{kpi['id']}_change",
                "style": {
                    "foregroundColor": {"opaqueColor": {"rgbColor": {"red": 0.2, "green": 0.2, "blue": 0.2}}},
                    "fontSize": {"magnitude": 12, "unit": "PT"}
                },
                "fields": "foregroundColor,fontSize"
            }}
        ])

    # Update KPI cards definition with real data
    kpi_cards = [
        {
            "id": kpi1_id,
            "title": "Total Count",
            "value": str(max_count),
            "change": "4.1% vs last month",
            "x": 40,
            "y": 200,  # Adjusted position
            "width": 200,
            "height": 80,
            "background": {"red": 0.4, "green": 0.8, "blue": 0.4}  # Light green
        },
        {
            "id": kpi2_id,
            "title": "Aggregated Count",
            "value": str(avg_count),
            "change": "1.1% vs last month",
            "x": 280,
            "y": 200,  # Adjusted position
            "width": 200,
            "height": 80,
            "background": {"red": 0.8, "green": 0.4, "blue": 0.4}  # Light red
        }
    ]

    # Add charts with adjusted positions
    chart_placement = [
        {"objectId": "chart1", "chartId": chart_ids[0], "translateY": 450},  # Adjusted position
        {"objectId": "chart2", "chartId": chart_ids[1], "translateY": 650}   # Adjusted position
    ]

    for chart in chart_placement:
        requests.append({
            "createSheetsChart": {
                "objectId": chart["objectId"],
                "spreadsheetId": sheet_id,
                "chartId": chart["chartId"],
                "linkingMode": "LINKED",
                "elementProperties": {
                    "pageObjectId": slide_id,
                    "transform": {
                        "scaleX": 1, "scaleY": 1,
                        "translateX": 40, "translateY": chart["translateY"],
                        "unit": "PT"
                    },
                    "size": {
                        "height": {"magnitude": 180, "unit": "PT"},  # Reduced height
                        "width": {"magnitude": 500, "unit": "PT"}
                    }
                }
            }
        })

    # Execute all requests
    slides_service.presentations().batchUpdate(
        presentationId=presentation_id,
        body={"requests": requests}
    ).execute()

    # Delete the second page if it exists
    presentation = slides_service.presentations().get(
        presentationId=presentation_id
    ).execute()
    
    if len(presentation.get('slides', [])) > 1:
        requests = [{
            'deleteObject': {
                'objectId': presentation['slides'][1]['objectId']
            }
        }]
        slides_service.presentations().batchUpdate(
            presentationId=presentation_id,
            body={'requests': requests}
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
            sheet_id, chart_ids = create_sheet_and_charts(sheets_service, event)
            create_presentation(slides_service, drive_service, event, sheet_id, chart_ids)

if __name__ == '__main__':
    main()
