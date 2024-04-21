from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as ec
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.action_chains import ActionChains

from constants import *
import time, re

def run(sheetId):
    data = []
    data.append(["Stallion", "Record Description", "Foals", "Price"])
    service = getGoogleService("sheets", "v4")
    result = service.spreadsheets().values().get(spreadsheetId=sheetId, range=f"Stallion master pedigree!A2:A").execute().get('values')
    
    driver = getGoogleDriver()
    driver.get("https://aqhaservices.aqha.com/members/records/aqharecords")
    WebDriverWait(driver, 30).until(lambda browser: browser.execute_script('return document.readyState') == 'complete')

    input_elem = driver.find_element(By.CSS_SELECTOR, "input#username")
    input_elem.click()
    input_elem.send_keys(Keys.CONTROL + "a")
    input_elem.send_keys(Keys.DELETE)
    input_elem.send_keys("brittany.holy")
    time.sleep(0.5)
    
    input_elem = driver.find_element(By.CSS_SELECTOR, "input#password")
    input_elem.click()
    input_elem.send_keys(Keys.CONTROL + "a")
    input_elem.send_keys(Keys.DELETE)
    input_elem.send_keys("GRWMs5Lw72L5MtS")
    time.sleep(0.5)

    driver.find_element(By.CSS_SELECTOR, "button.btn").click()
    time.sleep(1)
    WebDriverWait(driver, 30).until(lambda browser: browser.execute_script('return document.readyState') == 'complete')
    WebDriverWait(driver, 30).until(lambda browser: browser.execute_script('return document.querySelector("div#divProgress").innerHTML') == '')
    input_elem = driver.find_element(By.CSS_SELECTOR, "input#txtHorseName")
    for r in result:
        if r[0].strip() == "": continue
        print(f"<{r[0]}> selected.")
        ActionChains(driver).move_to_element(input_elem).click().perform()
        input_elem.send_keys(Keys.CONTROL + "a")
        input_elem.send_keys(Keys.DELETE)
        input_elem.send_keys(r[0])
        input_elem.send_keys(Keys.TAB)
        time.sleep(0.5)
        WebDriverWait(driver, 30).until(lambda browser: browser.execute_script('return document.querySelector("div#divProgress").innerHTML') == '')
        reg_elem = driver.find_element(By.CSS_SELECTOR, "input#txtRegNum")
        if reg_elem.get_attribute("value") != "":
            sire_accordian_btn = driver.find_element(By.CSS_SELECTOR, "a#sireAccordian")
            ActionChains(driver).move_to_element(sire_accordian_btn).click().perform()
            time.sleep(0.5)
            WebDriverWait(driver, 30).until(lambda browser: browser.execute_script('return document.querySelector("div#divProgress").innerHTML') == '')
            hasChildren = driver.execute_script("return document.querySelector('div#sireSection div.ui-grid-render-container-body div.ui-grid-viewport div.ui-grid-canvas').children.length") != 0
            if hasChildren:
                rows = driver.find_elements(By.CSS_SELECTOR, "div#sireSection div.ui-grid-render-container-body div.ui-grid-viewport div.ui-grid-canvas div.ui-grid-row")
                for row in rows:
                    cells = row.find_elements(By.CSS_SELECTOR, "div.ui-grid-cell")
                    if re.match(r'Get of Sire Summary for \d+ Crop Year', cells[1].text):
                        data.append([r[0], cells[1].text, cells[2].text, cells[3].text])

                page_navigation = driver.find_element(By.CSS_SELECTOR, "div#sireSection div.gridFooter div.pages")
                matched = re.search(r'\d+', page_navigation.text)
                page_num = int(matched.group())
                cur_page = page_navigation.find_element(By.CSS_SELECTOR, "input.currPage").get_attribute("value")
                cur_page = int(cur_page)
                if page_num != cur_page:
                    while True:
                        cur_page = page_navigation.find_element(By.CSS_SELECTOR, "input.currPage").get_attribute("value")
                        cur_page = int(cur_page)
                        if cur_page == page_num:
                            break
                        driver.find_elements(By.CSS_SELECTOR, "div#sireSection div.gridFooter div.pageBtn")[2].click()
                        time.sleep(0.5)
                        rows = driver.find_elements(By.CSS_SELECTOR, "div#sireSection div.ui-grid-render-container-body div.ui-grid-viewport div.ui-grid-canvas div.ui-grid-row")
                        for row in rows:
                            cells = row.find_elements(By.CSS_SELECTOR, "div.ui-grid-cell")
                            if re.match(r'Get of Sire Summary for \d+ Crop Year', cells[1].text):
                                data.append([r[0], cells[1].text, cells[2].text, cells[3].text])

    sheet_metadata = service.spreadsheets().get(spreadsheetId=sheetId).execute()
    sheets = sheet_metadata.get('sheets', '')
    flag = False
    for sheet_info in sheets:
        if sheet_info["properties"]["title"] == "AQHA":
            flag = True
    if not flag:
        body = {
            "requests":[{
                "addSheet":{
                    "properties":{
                        "title":"AQHA"
                    }
                }
            }]
        }

        service.spreadsheets().batchUpdate(spreadsheetId=sheetId, body=body).execute()
    else:
        service.spreadsheets().values().clear(
            spreadsheetId=sheetId,
            body={},
            range=f"AQHA!A1:E"
        ).execute()

    service.spreadsheets().values().update(
        spreadsheetId=sheetId,
        valueInputOption='RAW',
        range=f"AQHA!A1:E",
        body=dict(
            majorDimension='ROWS',
            values=data)
    ).execute()

    print("Done!")
    input("Press any key to exit...")

if __name__ == "__main__":
    sheetId = input("Enter your master sheet id:")
    if sheetId.strip() != "":
        run(sheetId)
    else:
        print("Invalid sheet id")