# sheet_helper.py
import time

def create_sheet_and_chart(sheets_service, event):
    title = event.get("event_title", "Untitled")[:30]
    spreadsheet = sheets_service.spreadsheets().create(body={
        'properties': {'title': f"{title} Chart Sheet"},
        'sheets': [{'properties': {'title': 'ChartData'}}]
    }).execute()

    sheet_id = spreadsheet['spreadsheetId']
    sheet_url = f"https://docs.google.com/spreadsheets/d/{sheet_id}"
    print(f"✅ Sheet created: {sheet_url}")

    headers = [['Timestamp', 'Avg Count', 'Max Count']]
    data = [[event.get('start_date', 'N/A'),
             event.get('analytics_summary', {}).get('average_count', 0),
             event.get('analytics_summary', {}).get('max_count', 0)]]

    sheets_service.spreadsheets().values().update(
        spreadsheetId=sheet_id,
        range='ChartData!A1:C2',
        valueInputOption='RAW',
        body={'values': headers + data}
    ).execute()

    metadata = sheets_service.spreadsheets().get(spreadsheetId=sheet_id).execute()
    sheet_id_map = {s['properties']['title']: s['properties']['sheetId'] for s in metadata['sheets']}
    chart_data_sheet_id = sheet_id_map.get('ChartData')

    if chart_data_sheet_id is None:
        raise Exception("Could not find 'ChartData' sheetId")

    requests_body = [
    {
        "addChart": {
            "chart": {
                "spec": {
                    "title": "Headcount Overview",
                    "basicChart": {
                        "chartType": "COLUMN",
                        "legendPosition": "BOTTOM_LEGEND",
                        "axis": [
                            {"position": "BOTTOM_AXIS", "title": "Date"},
                            {"position": "LEFT_AXIS", "title": "Count"}
                        ],
                        "domains": [{
                            "domain": {
                                "sourceRange": {
                                    "sources": [{
                                        "sheetId": chart_data_sheet_id,
                                        "startRowIndex": 1,
                                        "endRowIndex": 2,
                                        "startColumnIndex": 0,
                                        "endColumnIndex": 1
                                    }]
                                }
                            }
                        }],
                        "series": [
                            {
                                "series": {
                                    "sourceRange": {
                                        "sources": [{
                                            "sheetId": chart_data_sheet_id,
                                            "startRowIndex": 1,
                                            "endRowIndex": 2,
                                            "startColumnIndex": 1,
                                            "endColumnIndex": 2
                                        }]
                                    }
                                },
                                "colorStyle": {
                                    "rgbColor": {
                                        "red": 0.0,
                                        "green": 0.329,
                                        "blue": 0.667
                                    }
                                }
                            },
                            {
                                "series": {
                                    "sourceRange": {
                                        "sources": [{
                                            "sheetId": chart_data_sheet_id,
                                            "startRowIndex": 1,
                                            "endRowIndex": 2,
                                            "startColumnIndex": 2,
                                            "endColumnIndex": 3
                                        }]
                                    }
                                },
                                "colorStyle": {
                                    "rgbColor": {
                                        "red": 0.0,
                                        "green": 0.329,
                                        "blue": 0.667
                                    }
                                }
                            }
                        ]
                    }
                },
                "position": {"newSheet": True}
            }
        }
    }
]



    sheets_service.spreadsheets().batchUpdate(
        spreadsheetId=sheet_id,
        body={"requests": requests_body}
    ).execute()

    chart_metadata = sheets_service.spreadsheets().get(spreadsheetId=sheet_id, includeGridData=False).execute()
    charts = chart_metadata.get('sheets', [])
    chart_sheet_id = None
    for sheet in charts:
        title = sheet['properties']['title']
        if title.startswith("Chart"):
            chart_sheet_id = sheet['properties']['sheetId']
            break

    if chart_sheet_id is None:
        raise Exception("Chart sheet not created")

    chart_list = sheets_service.spreadsheets().get(spreadsheetId=sheet_id, includeGridData=False).execute()
    embedded_chart_id = None
    for sheet in chart_list.get("sheets", []):
        if "charts" in sheet:
            charts = sheet["charts"]
            if charts:
                embedded_chart_id = charts[0]["chartId"]
                break

    # fallback
    embedded_chart_id = 1 if embedded_chart_id is None else embedded_chart_id

    return sheet_id, embedded_chart_id
