import fitz
from selenium.webdriver.support.ui import WebDriverWait, Select
from selenium.webdriver.support import expected_conditions as ec
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from bs4 import BeautifulSoup
import os, re, time, sys

from constants import *

browser = None
service = None

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
    return re.sub(r'\s+', ' ', org_name)

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
    
def findSireFromSite(cn):
    global browser
    if browser is None:
        url = "https://beta.allbreedpedigree.com/search?query_type=check&search_bar=horse&g=5&inbred=Standard"
        browser = getGoogleDriver()
        browser.get(url)
        WebDriverWait(browser, 10).until(lambda browser: browser.execute_script('return document.readyState') == 'complete')
        WebDriverWait(browser, 10).until(ec.element_to_be_clickable((By.XPATH, "//button[@class='btn-close']"))).click()
    WebDriverWait(browser, 10).until(ec.element_to_be_clickable((By.XPATH, "//div[@id='header-search-input-helper']"))).click()
    input_elem = WebDriverWait(browser, 10).until(ec.element_to_be_clickable((By.XPATH, "//input[@id='header-search-input']")))
    input_elem.send_keys(Keys.CONTROL + "a")
    input_elem.send_keys(Keys.DELETE)
    input_elem.send_keys(cn, Keys.ENTER)
    
    soup = BeautifulSoup(browser.page_source, 'html.parser')
    try:
        table = soup.find(class_="pedigree-table").find("tbody")
        return getSireNameFromTable(table)
    except:
        try:
            table = soup.find(class_="layout-table").find("tbody")
            tds = table.select("td:nth-child(1)")
            txt_vals = []
            links = []
            for td in tds:
                txt_vals.append(td.text.upper())
                links.append(td.find("a").get("href"))
            indexes = [i for i, x in enumerate(txt_vals) if x.lower() == cn.lower()]
            if len(indexes) == 1:
                browser.get(links[0])
                WebDriverWait(browser, 10).until(lambda browser: browser.execute_script('return document.readyState') == 'complete')
                soup = BeautifulSoup(browser.page_source, 'html.parser')
                table = soup.find(class_="pedigree-table").find("tbody")
                return getSireNameFromTable(table)
            else:
                try:
                    select = Select(browser.find_element(By.XPATH, "//select[@id='filter-match']"))
                    select.select_by_value("exact")
                    WebDriverWait(browser, 10).until(lambda browser: browser.execute_script('return document.readyState') == 'complete')
                    soup = BeautifulSoup(browser.page_source, 'html.parser')
                    table = soup.find(class_="pedigree-table").find("tbody")
                    return getSireNameFromTable(table)
                except:
                    return ""
        except:
            return ""

def updateGSData(file_path, sheetId):
    update_data = []
    ext_names = extractPdf(file_path)
    if ext_names is None: return
    update_data.append(ext_names[0].title())
    pre_index_list = [1, 5, 3, 5, 7, 11, 9, 13]
    for i in pre_index_list:
        update_data.append(ext_names[i].title())
    
    for i in range(7, 15):
        sire_name = findSireFromSite(ext_names[i])
        update_data.append(sire_name)
        
    os.remove(file_path)
    
    worksheet = service.spreadsheets()
    sheet_metadata = worksheet.get(spreadsheetId=sheetId).execute()
    for sheet in sheet_metadata['sheets']:
        sheet_name = sheet['properties']['title']
        result = worksheet.values().get(spreadsheetId=sheetId, range="%s!A1:Z" % sheet_name).execute()
        values = result.get('values')
        header = values[0]
        indexOfHorseHeader = header.index('Horse')
        for id, row in enumerate(values):
            if len(row) != 0:
                if update_data[0].lower() == row[indexOfHorseHeader].lower():
                    worksheet.values().update(
                        spreadsheetId=sheetId,
                        valueInputOption='RAW',
                        range="%s!A%s:Z%s" % (sheet_name, str(id+1), str(id+1)),
                        body=dict(
                            majorDimension='ROWS',
                            values=[update_data])
                    ).execute()
    time.sleep(2)
def start(sheetId):
    sys.stdout = Unbuffered(sys.stdout)
    print("Third process started")
    createOrderDirIfDoesNotExists()
    global service
    service = getGoogleService("sheets", "v4")
    updated_cnt = 0
    while True:
        if os.path.exists("t2.txt"):
            t2_result = None
            with open("t2.txt", "r") as file:
                t2_result = file.read()
                file.close()
            
            if updated_cnt == int(t2_result):
                os.remove("t2.txt")
                break
        files = getOrderFiles()
        if len(files):
            for file in files:
                updateGSData(ORDER_DIR_NAME + "/" + file, sheetId)
                updated_cnt += 1
    browser.quit()
    print("Third process finished")
    
start("13b-fBnZpZFC_PTTuJ0Y9pYA-UYIgbsUDCCHjga5RBzs")