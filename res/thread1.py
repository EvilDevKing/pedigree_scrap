from selenium.webdriver.support.ui import WebDriverWait, Select
from selenium.webdriver.support import expected_conditions as ec
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
import time, sys

from constants import *

browser = None

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

def searchNameFromABP(service, sheetId, sheetName, indexOfHorseHeader, horse_name, index):
    sheet_data = []
    url = "https://beta.allbreedpedigree.com/search?query_type=check&search_bar=horse&g=5&inbred=Standard"
    
    browser.execute_script(f"window.open('{url}')")
    browser.switch_to.window(browser.window_handles[1])
    
    WebDriverWait(browser, 100).until(lambda browser: browser.execute_script('return document.readyState') == 'complete')
    WebDriverWait(browser, 10).until(ec.element_to_be_clickable((By.CSS_SELECTOR, "button.btn-close"))).click()
    time.sleep(2)

    WebDriverWait(browser, 10).until(ec.element_to_be_clickable((By.CSS_SELECTOR, "div#header-search-input-helper"))).click()
    input_elem = WebDriverWait(browser, 10).until(ec.element_to_be_clickable((By.CSS_SELECTOR, "input#header-search-input")))
    input_elem.send_keys(Keys.CONTROL + "a")
    input_elem.send_keys(Keys.DELETE)
    input_elem.send_keys(horse_name, Keys.ENTER)
    try:
        table = browser.find_element(By.CSS_SELECTOR, "table.pedigree-table tbody")
        sheet_data.append(getSheetDataFrom(table))
    except:
        try:
            tds = browser.find_elements(By.CSS_SELECTOR, "table.layout-table tbody td[class]:nth-child(1)")

            txt_vals = []
            links = []
            for td in tds:
                txt_vals.append(td.text)
                links.append(td.find_element(By.CSS_SELECTOR, "a").get_attribute("href"))
            indexes = [x for x in txt_vals if x.lower() == horse_name.lower()]
            if len(indexes) == 1:
                browser.get(links[0])
                WebDriverWait(browser, 10).until(lambda browser: browser.execute_script('return document.readyState') == 'complete')
                table = browser.find_element(By.CSS_SELECTOR, "table.pedigree-table tbody")
                sheet_data.append(getSheetDataFrom(table))
            else:
                try:
                    select = Select(browser.find_element(By.CSS_SELECTOR, "select#filter-match"))
                    select.select_by_value("exact")
                    WebDriverWait(browser, 10).until(lambda browser: browser.execute_script('return document.readyState') == 'complete')
                    table = browser.find_element(By.CSS_SELECTOR, "table.pedigree-table tbody")
                    sheet_data.append(getSheetDataFrom(table))
                except:
                    browser.close()
                    browser.switch_to.window(browser.window_handles[0])
                    print("THREAD1: Not found (" + horse_name + ") in https://beta.allbreedpedigree.com/")
        except:
            browser.close()
            browser.switch_to.window(browser.window_handles[0])
            print("THREAD1: Not found (" + horse_name + ") in https://beta.allbreedpedigree.com/")

    if len(sheet_data) != 0:
        service.spreadsheets().values().update(
            spreadsheetId=sheetId,
            valueInputOption='RAW',
            range=f"{sheetName}!{getColumnLabelByIndex(indexOfHorseHeader+1)}{str(index+2)}:Z{str(index+2)}",
            body=dict(
                majorDimension='ROWS',
                values=sheet_data)
        ).execute()

def fetchDataFromAQHA(sheetId, sheetName):
    search_cnt = 0
    global browser
    if browser == None:
        url = "https://aqhaservices.aqha.com/members/record/freerecords"
        browser = getGoogleDriver()
        browser.get(url)
        WebDriverWait(browser, 100).until(lambda browser: browser.execute_script('return document.readyState') == 'complete')
        time.sleep(2)
    service = getGoogleService("sheets", "v4")
    result = service.spreadsheets().values().get(spreadsheetId=sheetId, range=f"{sheetName}!A1:Z").execute().get('values')
    header = result.pop(0)
    indexOfHorseHeader = header.index("Horse")
    service.spreadsheets().values().clear(
        spreadsheetId=sheetId,
        body={},
        range=f"{sheetName}!{getColumnLabelByIndex(indexOfHorseHeader+1)}1:Z"
    ).execute()
    service.spreadsheets().values().update(
        spreadsheetId=sheetId,
        valueInputOption='RAW',
        range=f"{sheetName}!{getColumnLabelByIndex(indexOfHorseHeader+1)}1:Z1",
        body=dict(
            majorDimension='ROWS',
            values=[["Sire","Dams Sire","Grandsire Top","Grandsire Bottom","Grandsire Sire Top","Grandsire Sire Bottom","Granddams Sire Top","Granddams Sire bottom","Great-Grandsires Sire Top (1)","Great-Granddams Sire Top (2)","Great-Grandsires Sire Top (3)","Great-Granddams Sire Top (4)","Great-Grandsires Sire Bottom (5)","Great-Granddams Sire Bottom (6)","Great-Grandsires Sire Bottom (7)","Great-Granddams Sire Bottom (8)"]])
    ).execute()
    #### Fetch the all data from website ####
    cnt = 0
    for index, row_data in enumerate(result):
        if len(row_data) == 0 or row_data[indexOfHorseHeader] == "": continue
        name = row_data[0]
        select_element = Select(WebDriverWait(browser, 30).until(ec.presence_of_element_located((By.CSS_SELECTOR, "select#ddlRecordName"))))
        select_element.select_by_value("10008")
        
        WebDriverWait(browser, 10).until(ec.element_to_be_clickable((By.CSS_SELECTOR, "input#chkHorseName"))).click()
        
        input_elem = WebDriverWait(browser, 10).until(ec.presence_of_element_located((By.CSS_SELECTOR, "input#txtEmail")))
        input_elem.click()
        input_elem.send_keys(Keys.CONTROL + "a")
        input_elem.send_keys(Keys.DELETE)
        # input_elem.send_keys("brittany.holy@gmail.com")
        input_elem.send_keys("pascalmartin973@gmail.com")
        time.sleep(0.5)
        
        input_elem = WebDriverWait(browser, 10).until(ec.presence_of_element_located((By.CSS_SELECTOR, "input#txtHorseName")))
        input_elem.click()
        input_elem.send_keys(Keys.CONTROL + "a")
        input_elem.send_keys(Keys.DELETE)
        input_elem.send_keys(name)
        input_elem.send_keys(Keys.TAB)
        time.sleep(2)
        
        if browser.find_element(By.CSS_SELECTOR, "span.ng-binding").text != "":
            button_elem = WebDriverWait(browser, 20).until(ec.presence_of_element_located((By.CSS_SELECTOR, "div.text-right.m-m button")))
            button_elem.click()
            time.sleep(2)
            print("THREAD1: Found Horse (" + name + ") on AQHA Server")
            search_cnt += 1
        else:
            print("THREAD1: Not found Horse (" + name + ") on AQHA Server")
            searchNameFromABP(service, sheetId, sheetName, indexOfHorseHeader, name, index)
        
        time.sleep(5)
        cnt += 1
        print("Processed " + str(cnt))
    browser.quit()
    createFileWith("res/t1.txt", str(search_cnt), "w")

def start(sheetId, sheetName):
    sys.stdout = Unbuffered(sys.stdout)
    print("Main process started")
    fetchDataFromAQHA(sheetId, sheetName)
    print("---- Fetch Done ----")
    print("Main process finished")