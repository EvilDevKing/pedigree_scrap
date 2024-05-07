import fitz
import os, re, time, sys, requests

from constants import *

search_cnt = 0
update_cnt = 0

class Unbuffered(object):
   def __init__(self, stream):
       self.stream = stream
   def write(self, data):
       self.stream.write(data)
       self.stream.flush()
   def writelines(self, datas):
       self.stream.writelines(datas)
       self.stream.flush()
   def __getattr__(self, attr):
       return getattr(self.stream, attr)

def getExtactName(org_name):
    if re.search(r'\s+\d+', org_name):
        return re.sub(r'\s+\d+', '', org_name).title()
    else:
        return org_name.title()

def extractPdf(file_path):
    NAME_INDEXES = [0, 1, 6, 3, 4, 14, 11, 12, 9, 10, 7, 8, 13, 2]
    NAME_INDEXES2 = [0, 1, 6, 3, 4, 14, 11, 12, 9, 10, 7, 8, 13, 2, 5]
    
    with fitz.open(file_path) as doc:
        rawText = doc[1].get_text()

    rawList = [item for item in rawText.split("\n") if item.strip() != ""]
    ind_start = -1
    try:
        ind_start = rawList.index('Page 1 out of 1')
    except:
        try:
            ind_start = rawList.index('Page 1 out of 2')
        except:
            ind_start = -1
    if ind_start == -1:
        print("Pdf parse error")
        return None
    ind_end = -1
    ind_other = -1
    hasCurrentOwner = True
    for i, text in enumerate(rawList):
        if "CURRENT OWNER" in text:
            ind_end = i
            break
    if ind_end == -1:
        hasCurrentOwner = False
        for i, text in enumerate(rawList):
            if "Horse Details" in text:
                ind_end = i
                ind_other = i
                break
        if ind_end == -1:
            print("Pdf parse error")
            return None
    else:
        for i, text in enumerate(rawList):
            if "Horse Details" in text:
                ind_other = i
                break
            
    data = rawList[ind_start+1:ind_end]
    if ind_end == -1:
        print("Pdf parse error")
        return None
    else:
        tmp_names = []
        tmp_name = ""
        tmp_loop = 0
        for val in data:
            if tmp_loop > 0:
                tmp_loop -= 1
                continue

            if re.search(r'^\d{2}/\d{2}/\d{4}', val):
                tmp_loop = 3 - len(val.strip().split(' '))
                tmp_names.append(tmp_name);
                tmp_name = ""
                continue

            tmp_name += val
            
        names = [None] * 15
        for index, name in enumerate(tmp_names):
            if hasCurrentOwner:
                names[NAME_INDEXES[index]] = getExtactName(name)
            else:
                names[NAME_INDEXES2[index]] = getExtactName(name)
        
        if hasCurrentOwner:
            if len(rawList[ind_other-1].split(" ")) == 1:
                names[5] = getExtactName(rawList[ind_other-3])
            else:
                names[5] = getExtactName(rawList[ind_other-2])
        return names
    
def searchFromAQHA(horse_name, ws, sheetId, sheetName, rind, cind, multichoices):
    global search_cnt
    r = requests.get(f"https://aqhaservices2.aqha.com/api/HorseRegistration/GetHorseRegistrationDetailByHorseName?horseName={horse_name.title()}")
    if r.status_code == 200:
        res = r.json()
        regist_num = res["RegistrationNumber"]
        body_data = {
            # "CustomerEmailAddress": "brittany.holy@gmail.com",
            "CustomerEmailAddress": "pascalmartin973@gmail.com",
            "CustomerID": 1,
            "HorseName": horse_name.title(),
            "RecordOutputTypeCode": "P",
            "RegistrationNumber": regist_num,
            "ReportCode": 21,
            "ReportId": 10008
        }
        r1 = requests.post("https://aqhaservices2.aqha.com/api/FreeRecords/SaveFreeRecord", json=body_data)
        if r1.status_code == 200:
            print("THREAD4: Found Horse (" + horse_name + ") on AQHA Server")
            search_cnt += 1
        else:
            print("THREAD1: Failed to send email from AQHA Server.")
    else:
        print("THREAD4: Not found Horse (" + horse_name + ") on AQHA Server")
        flag = False
        cell_range_str = f"{getColumnLabelByIndex(cind)}{rind}"
        for choice_val in multichoices:
            if choice_val[0] == f"({horse_name})":
                ws.values().update(
                    spreadsheetId=sheetId,
                    valueInputOption='RAW',
                    range=f"{sheetName}!{cell_range_str}:{cell_range_str}",
                    body=dict(
                        majorDimension='ROWS',
                        values=[[choice_val[1]]]
                    )
                ).execute()
                flag = True
        
        if not flag:
            ws.values().update(
                spreadsheetId=sheetId,
                valueInputOption='RAW',
                range=f"{sheetName}!{cell_range_str}:{cell_range_str}",
                body=dict(
                    majorDimension='ROWS',
                    values=[[""]]
                )
            ).execute()

def updateGSData(file_path, ws, sheetId, sheetName, indexOfHorseHeader, sheetData):
    global update_cnt
    ext_names = extractPdf(file_path)
    if ext_names is None: return
    os.remove(file_path)
    for id, row in enumerate(sheetData):
        if len(row) > 9:
            for ind, c_val in enumerate(row[9:]):
                if c_val == f"({ext_names[0].lower()})":
                    cell_range_str = f"{getColumnLabelByIndex(9+ind+indexOfHorseHeader)}{id+2}"
                    ws.values().update(
                        spreadsheetId=sheetId,
                        valueInputOption='RAW',
                        range=f"{sheetName}!{cell_range_str}:{cell_range_str}",
                        body=dict(
                            majorDimension='ROWS',
                            values=[[ext_names[1]]]
                        )
                    ).execute()
                    print(f"Updated: r-{id+2}, c-{ind}, <{ext_names[1]}>")
    update_cnt += 1
    time.sleep(2)
def start(sheetId, sheetName):
    sys.stdout = Unbuffered(sys.stdout)
    print("Forth process started")
    
    service = getGoogleService("sheets", "v4")
    worksheet = service.spreadsheets()
    multichoices = worksheet.values().get(spreadsheetId=sheetId, range=f"mult choices!A2:B").execute().get('values')
    values = worksheet.values().get(spreadsheetId=sheetId, range=f"{sheetName}!A1:Z").execute().get('values')
    header = values.pop(0)
    indexOfHorseHeader = header.index('Horse')
    for rind, row in enumerate(values):
        if len(row) > 9:
            for cind, c_val in enumerate(row):
                if re.match(r'\(.*?\)', c_val):
                    searchFromAQHA(c_val.lstrip("(").rstrip(")"), worksheet, sheetId, sheetName, rind+2, cind+indexOfHorseHeader, multichoices)
    while True:
        if update_cnt == search_cnt:
            createFileWith("res/t4.txt", "finish", "w")
            break
        files = getOrderFiles()
        if len(files) > 0:
            for file in files:
                updateGSData(ORDER_DIR_NAME + "/" + file, worksheet, sheetId, sheetName, indexOfHorseHeader, values)
        
    print("Forth process finished")