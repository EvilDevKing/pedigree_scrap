import re
from constants import getGoogleService, getColumnLabelByIndex

def run(sheetId, sheetName):
    service = getGoogleService("sheets", "v4")
    worksheet = service.spreadsheets()
    sheetObj = worksheet.get(spreadsheetId=sheetId, fields='sheets(properties(sheetId,title))').execute()
    sheetInd = ""
    for sheet in sheetObj['sheets']:
        if sheet['properties']['title'] == sheetName:
            sheetInd = sheet['properties']['sheetId']
            break

    multichoices = worksheet.values().get(spreadsheetId=sheetId, range=f"mult choices!A2:B").execute().get('values')
    values = worksheet.values().get(spreadsheetId=sheetId, range=f"{sheetName}!A1:Z").execute().get('values')
    values.pop(0)

    for rid, row in enumerate(values):
        if len(row) != 0:
            for cid, col_val in enumerate(row):
                if re.match(r'\(.*?\)', col_val):
                    for choice in multichoices:
                        if choice[0] == col_val:
                            if len(choice) == 1 or choice[1].strip() == "":
                                batch_update_spreadsheet_request_body = {
                                    "requests": [
                                        {
                                            "repeatCell": {
                                                "range": {
                                                    "sheetId": sheetInd,
                                                    "startRowIndex": rid+1,
                                                    "endRowIndex": rid+2,
                                                    "startColumnIndex": cid,
                                                    "endColumnIndex": cid+1
                                                },
                                                "cell": {
                                                    "userEnteredFormat": {
                                                        "backgroundColor": {
                                                            "red": 50,
                                                            "green": 50,
                                                            "blue": 50
                                                        }
                                                    }
                                                },
                                                "fields": "userEnteredFormat.backgroundColor"
                                            }
                                        }
                                    ]
                                }
                                worksheet.batchUpdate(spreadsheetId=sheetId, body=batch_update_spreadsheet_request_body).execute()
                            else:
                                worksheet.values().update(
                                    spreadsheetId=sheetId,
                                    valueInputOption='RAW',
                                    range=f"{sheetName}!{getColumnLabelByIndex(cid)}{str(rid+2)}:Z{str(rid+2)}",
                                    body=dict(
                                        majorDimension='ROWS',
                                        values=[[choice[1]]])
                                ).execute()

if __name__ == "__main__":
    sheetId = input("Enter your worksheet id: ")
    if sheetId.strip() == "":
        print("The sheetid can't be empty.")
    else:
        sheetName = input("Enter your worksheet name: ")
        run(sheetId, sheetName)
    input("Press any key to exit...")