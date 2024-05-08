from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as ec
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
import re

from constants import *

SCOPES = ['https://www.googleapis.com/auth/spreadsheets']

def run(sheetId, sheet_name):
    service = getGoogleService("sheets", "v4")
    worksheet = service.spreadsheets()
    exec = worksheet.values().get(spreadsheetId=sheetId, range="%s!D2:D" % sheet_name).execute()
    result = exec.get('values')
    browser = getGoogleDriver()
    browser.get("https://www.allbreedpedigree.com/")
    for id, row in enumerate(result):
        if len(row) != 0:
            horse_name = row[0]
            input_element = WebDriverWait(browser, 20).until(ec.element_to_be_clickable((By.XPATH, "//input[@name='h']")))
            input_element.click()
            input_element.send_keys(Keys.CONTROL + "a")
            input_element.send_keys(Keys.DELETE)
            input_element.send_keys(horse_name)
            input_element.send_keys(Keys.ENTER)
            try:
                wrapper = WebDriverWait(browser, 10).until(ec.presence_of_element_located((By.XPATH, "//div[@class='pedigree-wrapper']")))
                horse_desc = wrapper.find_element(By.XPATH, "//table[@id='PedigreeTable']").find_elements(By.TAG_NAME, "td")[0].text
                year = ''.join(re.findall(r'\b\d{4}\b', horse_desc.split("\n")[1]))
            except:
                year = ""
            
            try:
                worksheet.values().update(
                    spreadsheetId=sheetId,
                    valueInputOption='RAW',
                    range="%s!E%s:E%s" % (sheet_name, str(id+2), str(id+2)),
                    body=dict(values=[[year]])
                ).execute()
                print(f"Updated <{horse_name}> => <{year}>")
            except:
                print(f"There is no any sheet name like <{sheet_name}>")
    print("Successfully updated!")