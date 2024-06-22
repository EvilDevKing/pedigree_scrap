import base64, os, time, sys, re
import fitz
from selenium.webdriver.support.ui import WebDriverWait, Select
from selenium.webdriver.support import expected_conditions as ec
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from constants import *

driver = None
worksheet = None

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
    org_name = org_name.replace("'", "")
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
        for i, val in enumerate(data):
            if i < len(data)-1:
                if re.search(r'^\d{2}/\d{2}/\d{4}', data[i+1]):
                    tmp_names.append(val)
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
    WebDriverWait(driver, 5).until(ec.element_to_be_clickable((By.CSS_SELECTOR, "div#header-search-input-helper"))).click()
    input_elem = WebDriverWait(driver, 5).until(ec.element_to_be_clickable((By.CSS_SELECTOR, "input#header-search-input")))
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

def updateGSData(file_path, sheetId, sheetName, indexOfHorse, sheetData, multichoices):
    update_data = []
    ext_names = extractPdf(file_path)
    if ext_names is None: return

    pre_index_list = [1, 5, 3, 5, 7, 11, 9, 13]
    for i in pre_index_list:
        update_data.append(ext_names[i])
    
    for i in range(7, 15):
        sire_name = findSireFromSite(ext_names[i])
        if sire_name.strip() == "":
            tmp_name = ""
            for choice_val in multichoices:
                if choice_val[0] == f"({ext_names[i].lower()})":
                    tmp_name = choice_val[1]
            if tmp_name != "":
                update_data.append(tmp_name)
            else:
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
    os.remove(file_path)

def checkMailAndDownloadOrderFile(sheetId, sheetName, multichoices):
    pdf_cnt = 0
    # Create Gmail API service
    service = getGoogleService('gmail', 'v1')
    while True:
        if os.path.exists("res/t1.txt"):
            with open("res/t1.txt", "r") as file:
                c = file.read()
                file.close()
                total_cnt = int(c)
                if pdf_cnt >= total_cnt:
                    os.remove("res/t1.txt")
                    break
        # Fetch messages from inbox
        results = service.users().messages().list(userId='me', labelIds=['INBOX'], maxResults=10, q="from:noreply@aqha.org").execute()
        messages = results.get('messages')
        if not messages:
            print("No messages found.")
        else:
            print("New Messages: " + str(len(messages)))
            for message in messages:
                msg_id = message['id']
                msg_body = service.users().messages().get(userId='me', id=msg_id).execute()
                try:
                    parts = msg_body['payload']['parts']
                    for part in parts:
                        if part['filename']:
                            if 'data' in part['body']:
                                data = part['body']['data']
                            else:
                                att_id = part['body']['attachmentId']
                                att = service.users().messages().attachments().get(userId='me', messageId=msg_id, id=att_id).execute()
                                data = att['data']
                            file_data = base64.urlsafe_b64decode(data.encode('UTF-8'))
                            filename = part['filename']
                            
                            createFileWith(ORDER_DIR_NAME + "/" + filename, file_data, 'wb')
                            print("Stored a pdf order file : " + filename)
                            time.sleep(0.5)
                            values = worksheet.values().get(spreadsheetId=sheetId, range=f"{sheetName}!A1:Z").execute().get('values')
                            header = values.pop(0)
                            indexOfHorseHeader = header.index('Horse')
                            updateGSData(ORDER_DIR_NAME + "/" + filename, sheetId, sheetName, indexOfHorseHeader, values, multichoices)
                            pdf_cnt += 1
                    service.users().messages().delete(userId='me', id=msg_id).execute()
                except: continue
        time.sleep(5)

def start(sheetId, sheetName):
    sys.stdout = Unbuffered(sys.stdout)
    print("Second process started")
    createOrderDirIfDoesNotExists()
    global worksheet, driver
    url = "https://beta.allbreedpedigree.com/search?query_type=check&search_bar=horse&g=5&inbred=Standard"
    driver = getGoogleDriver()
    driver.get(url)
    WebDriverWait(driver, 10).until(lambda driver: driver.execute_script('return document.readyState') == 'complete')
    WebDriverWait(driver, 10).until(ec.element_to_be_clickable((By.CSS_SELECTOR, "button.btn-close"))).click()
    service = getGoogleService("sheets", "v4")
    worksheet = service.spreadsheets()
    try:
        multichoices = worksheet.values().get(spreadsheetId=sheetId, range=f"mult choices!A2:B").execute().get('values')
        checkMailAndDownloadOrderFile(sheetId, sheetName, multichoices)
    except:
        print("There is no <mult choices> sheet!")
    driver.quit()
    print("Second process finished")