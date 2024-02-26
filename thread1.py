from selenium.webdriver.support.ui import WebDriverWait, Select
from selenium.webdriver.support import expected_conditions as ec
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from bs4 import BeautifulSoup
import time, sys

from constants import *

browser = None
sheet_info = {}
success_research_names = 0

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

def getNamesFromGoogleSheet(sheetId):
    global service, sheet_info
    # try:
    service = getGoogleService("sheets", "v4")

    # Call the Sheets API
    worksheet = service.spreadsheets()
    sheet_metadata = worksheet.get(spreadsheetId=sheetId).execute()
    sheet_info['names'] = []
    sheet_info['properties'] = {}
    for sheet in sheet_metadata['sheets']:
        sheet_name = sheet['properties']['title']
        sheet_info['names'].append(sheet_name)
        result = worksheet.values().get(spreadsheetId=sheetId, range="%s!A1:A" % sheet_name).execute()
        values = result.get('values')
        header = values.pop(0)
        indexOfHorseHeader = header.index('Horse')
        sheet_info['properties'][sheet_name] = {}
        sheet_info['properties'][sheet_name]['header'] = header
        sheet_info['properties'][sheet_name]['rows'] = []
        sheet_info['properties'][sheet_name]['horse_names'] = []
        for row in values:
            if len(row) == 0:
                sheet_info['properties'][sheet_name]['horse_names'].append("")
                sheet_info['properties'][sheet_name]['rows'].append([])
            else:
                sheet_info['properties'][sheet_name]['horse_names'].append(row[indexOfHorseHeader])
                sheet_info['properties'][sheet_name]['rows'].append(row[0:indexOfHorseHeader+1])
        
        service.spreadsheets().values().clear(
            spreadsheetId=sheetId,
            body={},
            range='%s!%s1:Z' % (sheet_name, getColumnLabelByIndex(indexOfHorseHeader+1))
        ).execute()
        service.spreadsheets().values().update(
            spreadsheetId=sheetId,
            valueInputOption='RAW',
            range="%s!%s1:Z1" % (sheet_name, getColumnLabelByIndex(indexOfHorseHeader+1)),
            body=dict(
                majorDimension='ROWS',
                values=[["Sire","Dams Sire","Grandsire Top","Grandsire Bottom","Grandsire Sire Top","Grandsire Sire Bottom","Granddams Sire Top","Granddams Sire bottom","Great-Grandsires Sire Top (1)","Great-Granddams Sire Top (2)","Great-Grandsires Sire Top (3)","Great-Granddams Sire Top (4)","Great-Grandsires Sire Bottom (5)","Great-Granddams Sire Bottom (6)","Great-Grandsires Sire Bottom (7)","Great-Granddams Sire Bottom (8)"]])
        ).execute()

    # except: print("Error loading google service")

def searchName(horse_name):
    global success_research_names
    url = "https://aqhaservices.aqha.com/members/record/freerecords"
    
    browser.execute_script(f"window.open('{url}')")
    browser.switch_to.window(browser.window_handles[1])
    
    WebDriverWait(browser, 30).until(lambda browser: browser.execute_script('return document.readyState') == 'complete')
    time.sleep(2)
    
    select_element = Select(WebDriverWait(browser, 30).until(ec.presence_of_element_located((By.XPATH, "//select[@id='ddlRecordName']"))))
    select_element.select_by_value("10008")
    
    WebDriverWait(browser, 10).until(ec.element_to_be_clickable((By.XPATH, "//input[@id='chkHorseName']"))).click()
    
    input_elem = WebDriverWait(browser, 10).until(ec.presence_of_element_located((By.XPATH, "//input[@id='txtEmail']")))
    input_elem.click()
    input_elem.send_keys(Keys.CONTROL + "a")
    input_elem.send_keys(Keys.DELETE)
    # input_elem.send_keys("brittany.holy@gmail.com")
    input_elem.send_keys("pascalmartin973@gmail.com")
    time.sleep(0.5)
    
    input_elem = WebDriverWait(browser, 10).until(ec.presence_of_element_located((By.XPATH, "//input[@id='txtHorseName']")))
    input_elem.click()
    input_elem.send_keys(Keys.CONTROL + "a")
    input_elem.send_keys(Keys.DELETE)
    input_elem.send_keys(horse_name)
    input_elem.send_keys(Keys.TAB)
    time.sleep(3)
    
    if WebDriverWait(browser, 20).until(ec.presence_of_element_located((By.XPATH, "//span[@class='ng-binding']"))).text:
        button_elem = WebDriverWait(browser, 20).until(ec.presence_of_element_located((By.XPATH, "//div[@class='text-right m-m']/button")))
        button_elem.click()
        time.sleep(2)
        browser.close()
        browser.switch_to.window(browser.window_handles[0])
        success_research_names += 1
    else:
        browser.close()
        browser.switch_to.window(browser.window_handles[0])
        print("THREAD1: Not found Horse (" + horse_name + ") on AQHA Server")

def fetchDataFromWebsite(sheetId):
    global browser
    if browser == None:
        url = "https://beta.allbreedpedigree.com/search?query_type=check&search_bar=horse&g=5&inbred=Standard"
        browser = getGoogleDriver()
        browser.get(url)
        WebDriverWait(browser, 10).until(lambda browser: browser.execute_script('return document.readyState') == 'complete')
        WebDriverWait(browser, 10).until(ec.element_to_be_clickable((By.XPATH, "//button[@class='btn-close']"))).click()
        print("---- Started fetching the data from https://beta.allbreedpedigree.com ----")
    for sheet_name in sheet_info['names']:
        print("----- Selected %s -----" % sheet_name)
        header = sheet_info['properties'][sheet_name]['header']
        horse_names = sheet_info['properties'][sheet_name]['horse_names']
        rows = sheet_info['properties'][sheet_name]['rows']
        print("Total " + str(len(horse_names)) + " names")
        #### Fetch the all data from website ####
        cnt = 0
        for index, name in enumerate(horse_names):
            sheet_data = []
            if name != "":
                WebDriverWait(browser, 10).until(ec.element_to_be_clickable((By.XPATH, "//div[@id='header-search-input-helper']"))).click()
                input_elem = WebDriverWait(browser, 10).until(ec.element_to_be_clickable((By.XPATH, "//input[@id='header-search-input']")))
                input_elem.send_keys(Keys.CONTROL + "a")
                input_elem.send_keys(Keys.DELETE)
                input_elem.send_keys(name, Keys.ENTER)
                soup = BeautifulSoup(browser.page_source, 'html.parser')
                try:
                    table = soup.find(class_="pedigree-table").find("tbody")
                    tmp_data = getSheetDataFrom(table, rows[cnt])
                    tmp_data = [v for v in tmp_data if v.strip() == ""]
                    if len(tmp_data) > 0:
                        searchName(name)
                    else:
                        sheet_data.append(tmp_data)
                except:
                    try:
                        table = soup.find(class_="layout-table").find("tbody")
                        tds = table.select("td:nth-child(1)")
                        txt_vals = []
                        links = []
                        for td in tds:
                            txt_vals.append(td.text.upper())
                            links.append(td.find("a").get("href"))
                        indexes = [i for i, x in enumerate(txt_vals) if x.lower() == name.lower()]
                        if len(indexes) == 1:
                            browser.get(links[0])
                            WebDriverWait(browser, 10).until(lambda browser: browser.execute_script('return document.readyState') == 'complete')
                            soup = BeautifulSoup(browser.page_source, 'html.parser')
                            table = soup.find(class_="pedigree-table").find("tbody")
                            sheet_data.append(getSheetDataFrom(table, rows[cnt]))
                        else:
                            try:
                                select = Select(browser.find_element(By.XPATH, "//select[@id='filter-match']"))
                                select.select_by_value("exact")
                                WebDriverWait(browser, 10).until(lambda browser: browser.execute_script('return document.readyState') == 'complete')
                                soup = BeautifulSoup(browser.page_source, 'html.parser')
                                table = soup.find(class_="pedigree-table").find("tbody")
                                sheet_data.append(getSheetDataFrom(table, rows[cnt]))
                            except:
                                i = header.index('Horse')
                                while True:
                                    rows[cnt].append("")
                                    i += 1
                                    if i > len(header):
                                        break
                                sheet_data.append(rows[cnt])
                                print("THREAD1: Not found (" + name + ") in https://beta.allbreedpedigree.com/")
                                searchName(name)
                    except:
                        i = header.index('Horse')
                        while True:
                            rows[cnt].append("")
                            i += 1
                            if i > len(header):
                                break
                        sheet_data.append(rows[cnt])
                        print("THREAD1: Not found (" + name + ") in https://beta.allbreedpedigree.com/")
                        searchName(name)
                        
                service.spreadsheets().values().update(
                    spreadsheetId=sheetId,
                    valueInputOption='RAW',
                    range="%s!A%s:Z%s" % (sheet_name, str(index+2), str(index+2)),
                    body=dict(
                        majorDimension='ROWS',
                        values=sheet_data)
                ).execute()
                time.sleep(5)
            cnt += 1
            print("Processed " + str(cnt))
    browser.quit()

def start(sheetId):
    sys.stdout = Unbuffered(sys.stdout)
    print("Main process started")
    
    getNamesFromGoogleSheet(sheetId)
    fetchDataFromWebsite(sheetId)
    
    print("---- Fetch Done ----")
    print("Your sheet temporarily updated!")
    
    createFileWith("t1.txt", str(success_research_names), "w")
    print("Main process finished")