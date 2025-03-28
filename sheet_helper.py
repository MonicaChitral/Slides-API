# sheet_helper.py

def create_sheet_and_chart(sheets_service, event):
    title = event.get("event_title", "Untitled")[:30]
    spreadsheet = sheets_service.spreadsheets().create(body={
        'properties': {'title': f"{title} Chart Sheet"},
        'sheets': [{'properties': {'title': 'ChartData'}}]
    }).execute()

    sheet_id = spreadsheet['spreadsheetId']
    print(f"✅ Sheet created: https://docs.google.com/spreadsheets/d/{sheet_id}")

    # Prepare data
    headers = [['Timestamp', 'Avg Count', 'Max Count']]
    data = [[event.get('start_date', 'N/A'),
             event.get('analytics_summary', {}).get('average_count', 0),
             event.get('analytics_summary', {}).get('max_count', 0)]]

    # Push data to sheet
    sheets_service.spreadsheets().values().update(
        spreadsheetId=sheet_id,
        range='ChartData!A1:C2',
        valueInputOption='RAW',
        body={'values': headers + data}
    ).execute()

    # Get ChartData sheet ID
    metadata = sheets_service.spreadsheets().get(spreadsheetId=sheet_id).execute()
    sheet_id_map = {s['properties']['title']: s['properties']['sheetId'] for s in metadata['sheets']}
    chart_data_sheet_id = sheet_id_map.get('ChartData')

    # Add embedded chart (not as image)
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
                            {"series": {"sourceRange": {
                                "sources": [{
                                    "sheetId": chart_data_sheet_id,
                                    "startRowIndex": 1,
                                    "endRowIndex": 2,
                                    "startColumnIndex": 1,
                                    "endColumnIndex": 2
                                }]}}},
                            {"series": {"sourceRange": {
                                "sources": [{
                                    "sheetId": chart_data_sheet_id,
                                    "startRowIndex": 1,
                                    "endRowIndex": 2,
                                    "startColumnIndex": 2,
                                    "endColumnIndex": 3
                                }]}}}
                        ]
                    }
                },
                "position": {
                    "overlayPosition": {
                        "anchorCell": {
                            "sheetId": chart_data_sheet_id,
                            "rowIndex": 5,
                            "columnIndex": 1
                        },
                        "offsetXPixels": 30,
                        "offsetYPixels": 30
                    }
                }
            }
        }
    }
]


    result = sheets_service.spreadsheets().batchUpdate(
        spreadsheetId=sheet_id,
        body={"requests": requests_body}
    ).execute()

    chart_id = result['replies'][0]['addChart']['chart']['chartId']
    return sheet_id, chart_id