import re, os
import fitz
from constants import getGoogleService, OKC_DIR_NAME


def getCleanedName(org_name):
    if re.search(r'\(#\d+\)', org_name):
        return re.sub(r'\(#\d+\)', '', org_name)
    return org_name
def getPDFData():
    if not os.path.exists(OKC_DIR_NAME):
        return None

    data = []
    data.append(["Horse", "Rider", "Earning"])

    file_list = os.listdir(OKC_DIR_NAME)
    if len(file_list) != 0:
        for filename in file_list:
            with fitz.open(f'{OKC_DIR_NAME}/{filename}') as doc:
                data.append(["",filename,""])
                if "Slot Race" in doc[0].get_text():
                    for page in doc:
                        contents = page.get_text().split("\n")
                        indexes = [i for i, val in enumerate(contents) if re.match(r'\b\d+\.\d+(\+\d+)?\b', val) or val == "NT" or val == "SCR"]
                        temp_ind = 0
                        for ind in indexes:
                            if contents[temp_ind] == contents[ind]:
                                continue
                            else:
                                horse = re.sub(r'\d+\s', '', contents[ind-4])
                                rider = contents[ind-1]
                                if re.match(r'^\$\d+(\,\d+)?$', contents[ind+1]):
                                    earning = contents[ind+1]
                                else:
                                    earning = 0
                                data.append([getCleanedName(horse), rider, earning])
                                temp_ind = ind
                        
                elif "Contestant" not in doc[0].get_text():
                    for page in doc:
                        contents = page.get_text().split("\n")
                        start_index = contents.index("$$") + 1
                        sub_data = []
                        for i in range(start_index, len(contents)):
                            if re.match(r'^\d+$', contents[i]) and re.match(r'\b[A-Z](\s)?\w+\b', contents[i+1]):
                                if len(contents[i+1]) > 3:
                                    if len(sub_data) == 2:
                                        data.append([getCleanedName(sub_data[0]), sub_data[1], "0"])
                                        sub_data = []
                                    sub_data.append(contents[i+1])
                                    sub_data.append(contents[i+3])
                            if re.match(r'^\d+\s\w+', contents[i]):
                                if re.match(r'^\d+\s\w+', contents[i-1]): continue
                                if len(contents[i+2]) == 2 or len(contents[i+2]) == 1 and "$" in contents[i+2]: continue
                                if len(sub_data) == 2:
                                    data.append([getCleanedName(sub_data[0]), sub_data[1], "0"])
                                    sub_data = []
                                sub_data.append(re.sub(r'\d+\s', '', contents[i]))
                                if re.match(r'\b\d+\.\d+(\+\d+)?\b', contents[i+2]):
                                    sub_data.append(contents[i+1])
                                else:
                                    sub_data.append(contents[i+2])
                            if re.match(r'^\$\d+', contents[i]):
                                sub_data.append(contents[i])
                                if len(sub_data) == 1:
                                    sub_data = []
                            if len(sub_data) == 3:
                                data.append([getCleanedName(sub_data[0]), sub_data[1], sub_data[2]])
                                sub_data = []
                            if i == len(contents)-1 and len(sub_data) == 2:
                                data.append([getCleanedName(sub_data[0]), sub_data[1], "0"])
                data.append(["","",""])
            # os.remove(f'{OKC_DIR_NAME}/{filename}')
        return data
    else:
        return None
    
def insertDataToGS(sheetId, sheetName, data):
    service = getGoogleService("sheets", "v4")
    
    sheets = service.spreadsheets().get(spreadsheetId=sheetId, fields='sheets').execute()
    sheet_list = sheets.get('sheets', [])
    sheet_names = [sheet['properties']['title'] for sheet in sheet_list]
    if sheetName not in sheet_names:
        # Create a new spreadsheet
        request_body = {
            'requests': [
                {
                    'addSheet': {
                        'properties': {
                            'title': sheetName
                        }
                    }
                }
            ]
        }
        service.spreadsheets().batchUpdate(spreadsheetId=sheetId, body=request_body).execute()
    else:
        # Clear all data from selected sheet
        service.spreadsheets().values().clear(
            spreadsheetId=sheetId,
            body={},
            range=f'{sheetName}!A1:Z'
        ).execute()
    service.spreadsheets().values().update(
        spreadsheetId=sheetId,
        valueInputOption='RAW',
        range=f"{sheetName}!A1:C",
        body=dict(
            majorDimension='ROWS',
            values=data)
    ).execute()
    
def run(sheetId, sheetName):
    print("The script is working...")
    sheet_data = getPDFData()
    if sheet_data == None:
        print("Sorry. There is no <okc.result> folder or any pdf file.")
    else:
        insertDataToGS(sheetId, sheetName, sheet_data)
        print("Success!")