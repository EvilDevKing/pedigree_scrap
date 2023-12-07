import os, pickle
from selenium import webdriver

from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

REPORT_DIR_NAME = "reports"
ORDER_DIR_NAME = "orders"
UTILS_DIR_NAME = "utils"
scores = {}
def getGoogleService(service_name, version):
    # If modifying these scopes, delete the file token.json.
    SCOPES = ['https://www.googleapis.com/auth/spreadsheets', 'https://www.googleapis.com/auth/gmail.readonly', 'https://mail.google.com/']
    
    credential = None
    if os.path.exists(UTILS_DIR_NAME + '/token.json'):
        credential = Credentials.from_authorized_user_file(UTILS_DIR_NAME + '/token.json', SCOPES)
    if not credential or not credential.valid:
        if credential and credential.expired and credential.refresh_token:
            credential.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(UTILS_DIR_NAME + '/credentials.json', SCOPES)
            credential = flow.run_local_server(port=0)
        with open(UTILS_DIR_NAME + '/token.json', 'w') as token:
            token.write(credential.to_json())
            
    try:
        service = build(service_name, version, credentials=credential)
        return service
    except HttpError as err:
        print(err)
        return None
    
def getGoogleDriver():
    chrome_options = webdriver.ChromeOptions()
    # chrome_options.add_argument("--headless")
    chrome_options.add_argument("--log-level=3")
    driver = webdriver.Chrome(options=chrome_options)

    driver.execute_script(
        "Object.defineProperty(navigator, 'webdriver', {get: () => undefined})")
    driver.execute_cdp_cmd("Page.addScriptToEvaluateOnNewDocument", {
        "source":
            "const newProto = navigator.__proto__;"
            "delete newProto.webdriver;"
            "navigator.__proto__ = newProto;"
    })
    return driver

def getTextValue(list, index):
    try:
        return list[index].find("div", {"class":"block-name"}).get("title").title()
    except:
        return ""
    
def getSheetColumnLabels(start_index, n):
    column_labels = ["A","B","C","D","E","F","G","H","I","J","K","L","M","N","O","P","Q","R","S","T","U","V","W","X","Y","Z"]
    sheet_column_labels = []

    for i in range(start_index, n):
        if i < len(column_labels):
            sheet_column_labels.append(column_labels[i])
        else:
            # If you need more than 26 columns, you can extend the labels with combinations like AA, AB, etc.
            div, mod = divmod(i, len(column_labels))
            label = column_labels[mod]
            while div > 0:
                div, mod = divmod(div - 1, len(column_labels))
                label = column_labels[mod] + label
            sheet_column_labels.append(label)

    return sheet_column_labels
        
def getColumnLabelByIndex(ind):
    labels = getSheetColumnLabels(0, 50)
    return labels[ind]

def getSireNameFromTable(table):
    tds = table.select("td[id]")
    sire_elem = [element for element in tds if element.get("id").endswith("M")]
    sire_name = getTextValue(sire_elem, 0)
    return sire_name

def getSheetDataFrom(table, row):
    ## Extract the values will be input in spreadsheet ##
    tds = table.select("td[id]")
    level0_elem = [element for element in tds if len(element.get("id")) == 1 and element.get("id").endswith("M")] ## list of Group A - Expected count : 1
    level1_elem = [element for element in tds if len(element.get("id")) == 2 and element.get("id").endswith("M")] ## list of Group B - Expected count : 2
    level2_elem = [element for element in tds if len(element.get("id")) == 3 and element.get("id").endswith("M")] ## list of Group C - Expected count : 4
    level3_elem = [element for element in tds if len(element.get("id")) == 4 and element.get("id").endswith("M")] ## list of Group D - Expected count : 8

    row.append(getTextValue(level0_elem, 0))

    row.append(getTextValue(level1_elem, 1))
    row.append(getTextValue(level1_elem, 0))
    row.append(getTextValue(level1_elem, 1))

    row.append(getTextValue(level2_elem, 0))
    row.append(getTextValue(level2_elem, 2))
    row.append(getTextValue(level2_elem, 1))
    row.append(getTextValue(level2_elem, 3))

    i = 0
    while True:
        row.append(getTextValue(level3_elem, i))
        i += 1
        if i == len(level3_elem):
            break

    return row

def createOrderDirIfDoesNotExists():
    if not os.path.exists(ORDER_DIR_NAME):
        os.makedirs(ORDER_DIR_NAME)
        
def createFileWith(filename, filecontent, mode):
    with open(filename, mode) as f:
        f.write(filecontent)
    
def getOrderFiles():
    return os.listdir(ORDER_DIR_NAME)