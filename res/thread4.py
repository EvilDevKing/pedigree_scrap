import fitz
from selenium.webdriver.support.ui import WebDriverWait, Select
from selenium.webdriver.support import expected_conditions as ec
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
import os, re, time, sys

from constants import *

driver = None
search_cnt = 0

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
    
def searchFromAQHA(horse_name, ws, sheetId, sheetName, rind, cind):
    global driver, search_cnt
    if driver == None:
        url = "https://aqhaservices.aqha.com/members/record/freerecords"
        driver = getGoogleDriver()
        driver.get(url)
        WebDriverWait(driver, 100).until(lambda browser: browser.execute_script('return document.readyState') == 'complete')
        time.sleep(1)
    
    select_element = Select(WebDriverWait(driver, 30).until(ec.presence_of_element_located((By.CSS_SELECTOR, "select#ddlRecordName"))))
    select_element.select_by_value("10008")
    
    WebDriverWait(driver, 10).until(ec.element_to_be_clickable((By.CSS_SELECTOR, "input#chkHorseName"))).click()
    
    input_elem = WebDriverWait(driver, 10).until(ec.presence_of_element_located((By.CSS_SELECTOR, "input#txtEmail")))
    input_elem.click()
    input_elem.send_keys(Keys.CONTROL + "a")
    input_elem.send_keys(Keys.DELETE)
    # input_elem.send_keys("brittany.holy@gmail.com")
    input_elem.send_keys("pascalmartin973@gmail.com")
    time.sleep(0.5)
    
    input_elem = WebDriverWait(driver, 10).until(ec.presence_of_element_located((By.CSS_SELECTOR, "input#txtHorseName")))
    input_elem.click()
    input_elem.send_keys(Keys.CONTROL + "a")
    input_elem.send_keys(Keys.DELETE)
    input_elem.send_keys(horse_name)
    input_elem.send_keys(Keys.TAB)
    time.sleep(2)
    
    if driver.find_element(By.CSS_SELECTOR, "span.ng-binding").text != "":
        button_elem = WebDriverWait(driver, 20).until(ec.presence_of_element_located((By.CSS_SELECTOR, "div.text-right.m-m button")))
        button_elem.click()
        time.sleep(2)
        print("THREAD4: Found Horse (" + horse_name + ") on AQHA Server")
        search_cnt += 1
    else:
        print("THREAD4: Not found Horse (" + horse_name + ") on AQHA Server")
        cell_range_str = f"{getColumnLabelByIndex(cind)}{rind}"
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
    time.sleep(2)
def start(sheetId, sheetName):
    sys.stdout = Unbuffered(sys.stdout)
    print("Forth process started")
    
    service = getGoogleService("sheets", "v4")
    worksheet = service.spreadsheets()
    values = worksheet.values().get(spreadsheetId=sheetId, range=f"{sheetName}!A1:Z").execute().get('values')
    header = values.pop(0)
    indexOfHorseHeader = header.index('Horse')
    for rind, row in enumerate(values):
        if len(row) > 9:
            for cind, c_val in enumerate(row):
                if re.match(r'\(.*?\)', c_val):
                    searchFromAQHA(c_val.lstrip("(").rstrip(")"), worksheet, sheetId, sheetName, rind+2, cind+indexOfHorseHeader)
    file_cnt = 0
    while True:
        if file_cnt == search_cnt:
            createFileWith("res/t4.txt", "finish", "w")
            break
        files = getOrderFiles()
        if len(files) > 0:
            for file in files:
                updateGSData(ORDER_DIR_NAME + "/" + file, worksheet, sheetId, sheetName, indexOfHorseHeader, values)
                file_cnt += 1
        
    driver.quit()
    print("Forth process finished")

start("13b-fBnZpZFC_PTTuJ0Y9pYA-UYIgbsUDCCHjga5RBzs", "horses")