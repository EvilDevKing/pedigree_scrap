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
            range="horses!A1:E",
            body=dict(
                majorDimension='ROWS',
                values=data)
        ).execute()

def getNames(filename):
    data = []
    with fitz.open("reports/" + filename) as doc:
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
                if len(dataArr) > 1:
                    data.append(dataArr)
    return data
    
def getPrices(filename):
    data = []
    pdffileObj = open("reports/" + filename, 'rb')
    pdfreader = PdfReader(pdffileObj)
    for page in pdfreader.pages:
        content = page.extract_text()
        entries = content.split("\n")
        for i in range(7, len(entries), 8):
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
    # print(len(prices))
    for i, name in enumerate(names):
        price = prices[i]
        if len(price) > 1:
            data.append([name[0], name[1], price[0], price[1], price[2]])
        else:
            data.append([name[0], name[1], 0, 0, 0])
    data.append(["", "", "", "", ""])
    return data
    
def getPbData(filename):
    data = []
    with fitz.open("reports/" + filename) as doc:
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
                if len(dataArr) == 2:
                    for i in range(3):
                        dataArr.append(0)
                data.append(dataArr)
    data.append(["", "", "", "", ""])
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
    sheet_data.append(["Horse", "Rider", "Owner Price", "Stallion Price", "Breeder Price"])
    if not os.path.exists("reports"):
        os.makedirs("reports")
    
    files = os.listdir("reports")
    if len(files) == 0:
        print("Not found any report files in \"reports\" directory.")
    else:
        for filename in files:
            sheet_data.append(["", "", filename, "", ""])
            with fitz.open("reports/" + filename) as doc:
                content = doc[0].get_text()
                with open('content.txt', 'w') as file:
                    file.write(content)
                    file.close()
                if re.search(r'Amateur', content):
                    sheet_data.extend(getPbData(filename))
                else:
                    sheet_data.extend(getRbData(filename))
        
        update_google_sheets(sheetId, sheet_data)
        print("### Successfully updated! ###")