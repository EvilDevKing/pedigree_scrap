import fitz
import os, re, time, sys

from constants import *

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
    
def updateGSData(file_path, ws, sheetId, indexOfHorseHeader, sheetData):
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
                        range=f"horses!{cell_range_str}:{cell_range_str}",
                        body=dict(
                            majorDimension='ROWS',
                            values=[[ext_names[1]]]
                        )
                    ).execute()
    time.sleep(2)
def start(sheetId):
    sys.stdout = Unbuffered(sys.stdout)
    print("Forth process started")
    createOrderDirIfDoesNotExists()
    service = getGoogleService("sheets", "v4")
    worksheet = service.spreadsheets()
    values = worksheet.values().get(spreadsheetId=sheetId, range="horses!A1:Q").execute().get('values')
    header = values.pop(0)
    indexOfHorse = header.index('Horse')
    files = getOrderFiles()
    if len(files):
        for file in files:
            updateGSData(ORDER_BACKUP_DIR_NAME + "/" + file, worksheet, sheetId, indexOfHorse, values)
    print("Third process finished")