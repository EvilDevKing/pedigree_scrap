import fitz
from selenium.webdriver.support.ui import WebDriverWait, Select
from selenium.webdriver.support import expected_conditions as ec
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
import os, re, time, sys

from constants import *

driver = None
worksheet = None
horse_names = []
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
    
def searchFromAQHA(horse_name):
    url = "https://aqhaservices.aqha.com/members/record/freerecords"
    
    driver.execute_script(f"window.open('{url}')")
    driver.switch_to.window(driver.window_handles[1])
    
    WebDriverWait(driver, 100).until(lambda driver: driver.execute_script('return document.readyState') == 'complete')
    time.sleep(2)
    
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
        driver.close()
        driver.switch_to.window(driver.window_handles[0])
        print("THREAD3: Found Horse (" + horse_name + ") on AQHA Server")
        createFileWith("res/t3.txt", str(search_cnt+1), "w")
    else:
        driver.close()
        driver.switch_to.window(driver.window_handles[0])
        print("THREAD3: Not found Horse (" + horse_name + ") on AQHA Server")

def findSireFromSite(cn):
    global driver
    if driver is None:
        url = "https://beta.allbreedpedigree.com/search?query_type=check&search_bar=horse&g=5&inbred=Standard"
        driver = getGoogleDriver()
        driver.get(url)
        WebDriverWait(driver, 10).until(lambda driver: driver.execute_script('return document.readyState') == 'complete')
        WebDriverWait(driver, 10).until(ec.element_to_be_clickable((By.CSS_SELECTOR, "button.btn-close"))).click()
    WebDriverWait(driver, 10).until(ec.element_to_be_clickable((By.CSS_SELECTOR, "div#header-search-input-helper"))).click()
    input_elem = WebDriverWait(driver, 10).until(ec.element_to_be_clickable((By.CSS_SELECTOR, "input#header-search-input")))
    input_elem.send_keys(Keys.CONTROL + "a")
    input_elem.send_keys(Keys.DELETE)
    input_elem.send_keys(cn, Keys.ENTER)
    
    try:
        table = driver.find_element(By.CSS_SELECTOR, "table.pedigree-table tbody")
        return getSireNameFromTable(table)
    except:
        try:
            tds = driver.find_elements(By.CSS_SELECTOR, "table.layout-table tbody td[class]:nth-child(1)")
            txt_vals = []
            links = []
            for td in tds:
                txt_vals.append(td.text)
                links.append(td.find_element(By.TAG_NAME, "a").get_attribute("href"))
            indexes = [x for x in txt_vals if x.lower() == cn.lower()]
            if len(indexes) == 1:
                driver.get(links[0])
                WebDriverWait(driver, 10).until(lambda driver: driver.execute_script('return document.readyState') == 'complete')
                table = driver.find_element(By.CSS_SELECTOR, "table.pedigree-table tbody")
                return getSireNameFromTable(table)
            else:
                try:
                    select = Select(driver.find_element(By.CSS_SELECTOR, "select#filter-match"))
                    select.select_by_value("exact")
                    WebDriverWait(driver, 10).until(lambda driver: driver.execute_script('return document.readyState') == 'complete')
                    table = driver.find_element(By.CSS_SELECTOR, "table.pedigree-table tbody")
                    return getSireNameFromTable(table)
                except:
                    return ""
        except:
            return ""

def updateGSData(file_name, sheetId, sheetName, indexOfHorse, sheetData):
    file_path = ORDER_DIR_NAME + "/" + file_name
    update_data = []
    ext_names = extractPdf(file_path)
    if ext_names is None: return
    pre_index_list = [1, 5, 3, 5, 7, 11, 9, 13]
    for i in pre_index_list:
        update_data.append(ext_names[i])
    
    os.replace(file_path, ORDER_BACKUP_DIR_NAME + "/" + file_name)
    
    for i in range(7, 15):
        sire_name = findSireFromSite(ext_names[i])
        if sire_name.strip() == "":
            searchFromAQHA(ext_names[i])
            update_data.append(f"({ext_names[i].lower()})")
        else:
            update_data.append(sire_name)
    
    for id, row in enumerate(sheetData):
        if len(row) != 0:
            if ext_names[0].lower() == row[indexOfHorse].lower():
                worksheet.values().update(
                    spreadsheetId=sheetId,
                    valueInputOption='RAW',
                    range=f"{sheetName}!{getColumnLabelByIndex(indexOfHorse+1)}{str(id+2)}:Z{str(id+2)}",
                    body=dict(
                        majorDimension='ROWS',
                        values=[update_data])
                ).execute()
    
    time.sleep(2)

def start(sheetId, sheetName):
    sys.stdout = Unbuffered(sys.stdout)
    print("Third process started")
    createOrderDirIfDoesNotExists()
    createOrderBackupDirIfDoesNotExists()
    global worksheet
    service = getGoogleService("sheets", "v4")
    worksheet = service.spreadsheets()
    values = worksheet.values().get(spreadsheetId=sheetId, range=f"{sheetName}!A1:Z").execute().get('values')
    header = values.pop(0)
    indexOfHorseHeader = header.index('Horse')
    updated_cnt = 0
    while True:
        if os.path.exists("res/t2.txt"):
            t2_result = None
            with open("res/t2.txt", "r") as file:
                t2_result = file.read()
                file.close()
            
            if updated_cnt == int(t2_result):
                os.remove("res/t2.txt")
                break
        files = getOrderFiles()
        if len(files):
            for file in files:
                updateGSData(file, sheetId, sheetName, indexOfHorseHeader, values)
                updated_cnt += 1
    driver.quit()
    print("Third process finished")