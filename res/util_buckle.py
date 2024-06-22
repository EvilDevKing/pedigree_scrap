import fitz
from PyPDF2 import *
import sys, os, re

from constants import *
SCOPES = ['https://www.googleapis.com/auth/spreadsheets']

def update_google_sheets(sheetId, data):
    service = getGoogleService('sheets', 'v4')
    service.spreadsheets().values().clear(
        spreadsheetId=sheetId,
        body={},
        range='horses!A1:Z'
    ).execute()
    service.spreadsheets().values().update(
            spreadsheetId=sheetId,
            valueInputOption='RAW',
            range="horses!A1:F",
            body=dict(
                majorDimension='ROWS',
                values=data)
        ).execute()

def getNames(filename):
    data = []
    with fitz.open(REPORT_DIR_NAME + "/" + filename) as doc:
        for page in doc:
            content = page.get_text()
            entries = content.split("\n")
            split_indicators = [i for i, value in enumerate(entries) if re.search(r'^\d+$', value.strip())]
            split_indicators.append(len(entries))
            start_index = 0
            for val in split_indicators:
                if start_index == 0:
                    start_index = val
                    continue
                sub_data = entries[start_index:val]
                start_index = val
                dataArr = []
                if len(sub_data) == 2 and sub_data[1] == '':
                    continue
                for i, dt in enumerate(sub_data):
                    if i == 1:
                        if "$" not in dt:
                            dt = re.sub(r'\s+', ' ', dt)
                            if "(" in dt:
                                r = dt.index("(")
                                dataArr.append(dt[:r].strip())
                                dataArr.append(sub_data[6].strip())
                            else:
                                dataArr.append(dt.strip())
                                dataArr.append(sub_data[7].strip())
                            dataArr.append("")
                    else:
                        match = re.search(r'(\d+\.\d+)', dt)
                        if match:
                            decimal = match.group(1)
                            dataArr[2] = decimal
                if len(dataArr) > 1:
                    data.append(dataArr)
    return data

def getPrices(filename):
    data = []
    pdffileObj = open(REPORT_DIR_NAME + "/" + filename, 'rb')
    pdfreader = PdfReader(pdffileObj)
    for page in pdfreader.pages:
        content = page.extract_text()
        entries = content.split("\n")
        start_index = [i for i, x in enumerate(entries) if "Owner Stallion Breeder" in x]
        for i in range(start_index[0]+1, len(entries), 8):
            numbers = []
            try:
                tmp = entries[i+7]
                if re.search(r'^\d+$', tmp): continue
                if "$" in tmp:
                    splt_tmp = tmp.split(" ")
                    for v in splt_tmp:
                        if v.replace(',', '').rstrip('$').isdigit():
                            numbers.append(int(v.replace(',', '').rstrip('$')))
                        if v.replace(',', '').lstrip('$').isdigit():
                            numbers.append(int(v.replace(',', '').lstrip('$')))
                else:
                    numbers = [""]
            except: xxx = 0
            if len(numbers) > 0:
                data.append(numbers)
    return data

def getRbData(filename):
    data = []
    names = getNames(filename)
    prices = getPrices(filename)
    for i, name in enumerate(names):
        price = prices[i]
        if len(price) > 1:
            data.append([name[0], name[1], name[2], price[0], price[1], price[2]])
        else:
            data.append([name[0], name[1], name[2], 0, 0, 0])
    data.append(["", "", "", "", "", ""])
    return data
    
def getPbData(filename):
    data = []
    with fitz.open(REPORT_DIR_NAME + "/" + filename) as doc:
        for page in doc:
            content = page.get_text()
            entries = content.split("\n")
            split_indicators = [i for i, value in enumerate(entries) if re.search(r'^\d+$', value.strip())]
            split_indicators.append(len(entries))
            start_index = 0
            for val in split_indicators:
                if start_index == 0:
                    start_index = val
                    continue
                sub_data = entries[start_index:val]
                start_index = val
                dataArr = []
                if len(sub_data) == 2 and sub_data[1] == '':
                    continue
                for i, dt in enumerate(sub_data):
                    if i == 1:
                        dt = re.sub(r'\s+', ' ', dt)
                        if "(" in dt:
                            r = dt.index("(")
                            dataArr.append(dt[:r].strip())
                            dataArr.append(sub_data[6].strip())
                        else:
                            dataArr.append(dt.strip())
                            dataArr.append(sub_data[7].strip())
                        dataArr.append("")
                    else:
                        if "$" in dt:
                            splt_dt = dt.split(" ")
                            if len(splt_dt) == 3:
                                a = splt_dt[0].index("$")
                                b = splt_dt[1].index("$")
                                c = splt_dt[2].index("$")
                                dataArr.append(int(splt_dt[0][a:].replace(',', '').lstrip('$')))
                                dataArr.append(int(splt_dt[1][b:].replace(',', '').lstrip('$')))
                                dataArr.append(int(splt_dt[2][c:].replace(',', '').lstrip('$')))
                            elif len(splt_dt) == 2:
                                if "$" in splt_dt[0]:
                                    d = splt_dt[0].index("$")
                                    dataArr.append(int(splt_dt[0][d:].replace(',', '').lstrip('$')))
                                if "$" in splt_dt[1]:
                                    e = splt_dt[1].index("$")
                                    dataArr.append(int(splt_dt[1][e:].replace(',', '').lstrip('$')))
                            else:
                                f = dt.index("$")
                                dataArr.append(int(dt[f:].replace(',', '').lstrip('$')))
                        else:
                            match = re.search(r'(\d+\.\d+)', dt)
                            if match:
                                decimal = match.group(1)
                                dataArr[2] = decimal
                if len(dataArr) == 3:
                    for i in range(3):
                        dataArr.append(0)
                data.append(dataArr)
    data.append(["", "", "", "", "", ""])
    return data

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

def run(sheetId):
    sys.stdout = Unbuffered(sys.stdout)
    print("Processing...")
    sheet_data = []
    sheet_data.append(["Horse", "Rider", "Time", "Owner Price", "Stallion Price", "Breeder Price"])
    if not os.path.exists(REPORT_DIR_NAME):
        os.makedirs(REPORT_DIR_NAME)
    
    files = os.listdir(REPORT_DIR_NAME)
    if len(files) == 0:
        print("Not found any report files in \"reports\" directory.")
    else:
        for filename in files:
            sheet_data.append(["", "", filename, "", "", ""])
            with fitz.open(REPORT_DIR_NAME + "/" + filename) as doc:
                content = doc[0].get_text()
                # with open('content.txt', 'w') as file:
                #     file.write(content)
                #     file.close()
                if re.search(r'Amateur', content):
                    sheet_data.extend(getPbData(filename))
                else:
                    sheet_data.extend(getRbData(filename))
        
        update_google_sheets(sheetId, sheet_data)
        print("### Successfully updated! ###")