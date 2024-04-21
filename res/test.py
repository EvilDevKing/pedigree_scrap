from constants import getGoogleService

service = getGoogleService("sheets", "v4")
sheet_metadata = service.spreadsheets().get(spreadsheetId="18wZ_UlyQKmhzygdb8nk8I6xAyIPvxJm3Ofh58d1NKZs").execute()
sheets = sheet_metadata.get('sheets', '')
flag = False
for sheet_info in sheets:
    if sheet_info["properties"]["title"] == "AQHA":
        flag = True
print(flag)