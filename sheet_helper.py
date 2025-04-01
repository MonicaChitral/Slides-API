# Updated sheet_helper.py
import time
import json

def create_sheet_and_charts(sheets_service, event):
    title = event.get("event_title", "Untitled")[:30]
    spreadsheet = sheets_service.spreadsheets().create(body={
        'properties': {'title': f"{title} Sheet"},
        'sheets': [
            {'properties': {'title': 'TrendData'}},
            {'properties': {'title': 'SeatData'}}
        ]
    }).execute()

    sheet_id = spreadsheet['spreadsheetId']
    print(f"✅ Sheet created: https://docs.google.com/spreadsheets/d/{sheet_id}")

    # Prepare Trend Data from analytics
    trend_headers = [['Time', 'Count']]
    trend_data = []
    
    # Use real analytics data if available
    analytics = event.get('analytics', [])
    if analytics:
        for entry in analytics:
            time_str = entry.get('datetime', '').split('T')[1][:5]  # Get HH:MM from datetime
            trend_data.append([time_str, entry.get('headcount', 0)])
    else:
        # Fallback to dummy data
        trend_data = [
            ['8:00', 25],
            ['10:00', 40],
            ['12:00', 35],
            ['14:00', 30],
            ['16:00', 38],
            ['18:00', 35],
            ['20:00', 20]
        ]

    # Load seating data from seating.json
    try:
        with open('seating.json', 'r') as f:
            seating_data = json.load(f)
            seat_data = []
            for section in seating_data.get('sections', []):
                section_name = section.get('section_name', 'Unknown')
                weekly_data = section.get('weekly_data', {})
                last_week = weekly_data.get('last_week', {})
                seating_density = last_week.get('seating_density', 0)
                seat_data.append([section_name, seating_density])
    except (FileNotFoundError, json.JSONDecodeError):
        # Fallback to dummy data if seating.json is not available
        seat_data = [
            ['Section 1', -1.0],
            ['Section 2', -0.5],
            ['Section 3', 0.0],
            ['Section 4', 0.5],
            ['Section 5', 1.0]
        ]

    # Add headers for the seat data
    seat_headers = [['Section', 'Density']]
    seat_data = seat_headers + seat_data

    sheets_service.spreadsheets().values().update(
        spreadsheetId=sheet_id,
        range=f'TrendData!A1:B{len(trend_data) + 1}',
        valueInputOption='RAW',
        body={'values': trend_headers + trend_data}
    ).execute()

    sheets_service.spreadsheets().values().update(
        spreadsheetId=sheet_id,
        range=f'SeatData!A1:B{len(seat_data) + 1}',
        valueInputOption='RAW',
        body={'values': seat_data}
    ).execute()

    # Get sheet IDs
    metadata = sheets_service.spreadsheets().get(spreadsheetId=sheet_id).execute()
    sheet_id_map = {s['properties']['title']: s['properties']['sheetId'] for s in metadata['sheets']}

    trend_sheet_id = sheet_id_map['TrendData']
    seat_sheet_id = sheet_id_map['SeatData']

    requests = [
        {"addChart": {
            "chart": {
                "spec": {
                    "title": "Crowd Trend Analysis",
                    "basicChart": {
                        "chartType": "LINE",
                        "legendPosition": "NO_LEGEND",
                        "axis": [
                            {
                                "position": "BOTTOM_AXIS",
                                "title": ""
                            },
                            {
                                "position": "LEFT_AXIS",
                                "title": ""
                            }
                        ],
                        "lineSmoothing": True,
                        "series": [{
                            "series": {"sourceRange": {"sources": [{
                                "sheetId": trend_sheet_id,
                                "startRowIndex": 1,
                                "endRowIndex": len(trend_data) + 1,
                                "startColumnIndex": 1,
                                "endColumnIndex": 2
                            }]}},
                            "color": {"red": 0.3, "green": 0.5, "blue": 0.9},
                            "lineStyle": {"width": 2}
                        }],
                        "domains": [{
                            "domain": {
                                "sourceRange": {
                                    "sources": [{
                                        "sheetId": trend_sheet_id,
                                        "startRowIndex": 1,
                                        "endRowIndex": len(trend_data) + 1,
                                        "startColumnIndex": 0,
                                        "endColumnIndex": 1
                                    }]
                                }
                            }
                        }]
                    }
                },
                "position": {
                    "overlayPosition": {
                        "anchorCell": {
                            "sheetId": trend_sheet_id,
                            "rowIndex": 0,
                            "columnIndex": 3
                        }
                    }
                }
            }
        }},
        {"addChart": {
            "chart": {
                "spec": {
                    "title": "Seat Sections",
                    "basicChart": {
                        "chartType": "COLUMN",
                        "legendPosition": "NO_LEGEND",
                        "axis": [
                            {
                                "position": "BOTTOM_AXIS",
                                "title": "Section 1"
                            },
                            {
                                "position": "LEFT_AXIS",
                                "title": ""
                            }
                        ],
                        "series": [{
                            "series": {"sourceRange": {"sources": [{
                                "sheetId": seat_sheet_id,
                                "startRowIndex": 1,
                                "endRowIndex": len(seat_data) + 1,
                                "startColumnIndex": 1,
                                "endColumnIndex": 2
                            }]}},
                            "color": {"red": 0.1, "green": 0.3, "blue": 0.6},
                            "targetAxis": "LEFT_AXIS"
                        }],
                        "domains": [{
                            "domain": {
                                "sourceRange": {
                                    "sources": [{
                                        "sheetId": seat_sheet_id,
                                        "startRowIndex": 1,
                                        "endRowIndex": len(seat_data) + 1,
                                        "startColumnIndex": 0,
                                        "endColumnIndex": 1
                                    }]
                                }
                            }
                        }]
                    }
                },
                "position": {
                    "overlayPosition": {
                        "anchorCell": {
                            "sheetId": seat_sheet_id,
                            "rowIndex": 0,
                            "columnIndex": 3
                        }
                    }
                }
            }
        }}
    ]

    sheets_service.spreadsheets().batchUpdate(
        spreadsheetId=sheet_id,
        body={"requests": requests}
    ).execute()

    # Get chart IDs
    chart_metadata = sheets_service.spreadsheets().get(spreadsheetId=sheet_id, includeGridData=False).execute()
    chart_ids = []
    for sheet in chart_metadata.get("sheets", []):
        for chart in sheet.get("charts", []):
            chart_ids.append(chart["chartId"])

    if len(chart_ids) < 2:
        chart_ids = [1, 2]  # Fallback

    return sheet_id, chart_ids
