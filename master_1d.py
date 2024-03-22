from constants import getGoogleService
import time

def run(sheetId):
    service = getGoogleService("sheets", "v4")
    oned_result = service.spreadsheets().values().get(spreadsheetId=sheetId, range="1D crosses!A2:R").execute().get("values")
    overall_result = service.spreadsheets().values().get(spreadsheetId=sheetId, range="Overall Master crosses!A2:R").execute().get("values")
    print("Running...")
    for i, v in enumerate(oned_result):
        if len(v) == 0 or len(v) == 2: continue
        if (len(v) > 14 and "" in v[10:14]) or len(v) < 10:
            update_data = []
            for x in overall_result:
                if len(x) > 14 and v[2].lower() == x[2].lower() and "" not in x[10:14]:
                    update_data.append(x[10])
                    update_data.append(x[11])
                    update_data.append(x[12])
                    update_data.append(x[13])
                    break
            service.spreadsheets().values().update(
                spreadsheetId=sheetId,
                valueInputOption='RAW',
                range=f"1D crosses!K{i+2}:N{i+2}",
                body=dict(
                    majorDimension='ROWS',
                    values=[update_data]
                )
            ).execute()
            time.sleep(1)
        if (len(v) > 16 and "" in v[14:16]) or len(v) < 14:
            update_data = []
            for x in overall_result:
                if len(x) > 16 and v[3].lower() == x[3].lower() and "" not in x[14:16]:
                    update_data.append(x[14])
                    update_data.append(x[15])
                    break
            service.spreadsheets().values().update(
                spreadsheetId=sheetId,
                valueInputOption='RAW',
                range=f"1D crosses!O{i+2}:P{i+2}",
                body=dict(
                    majorDimension='ROWS',
                    values=[update_data]
                )
            ).execute()
            time.sleep(1)
        if (len(v) > 16 and v[16] == "") or len(v) < 16:
            update_data = []
            for x in overall_result:
                if len(x) > 16 and v[9].lower() == x[9].lower() and x[16] != "":
                    update_data.append(x[16])
                    break
            service.spreadsheets().values().update(
                spreadsheetId=sheetId,
                valueInputOption='RAW',
                range=f"1D crosses!Q{i+2}:Q{i+2}",
                body=dict(
                    majorDimension='ROWS',
                    values=[update_data]
                )
            ).execute()
            time.sleep(1)
    print("Done!")