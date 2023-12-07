from selenium.webdriver.support.ui import WebDriverWait, Select
from selenium.webdriver.support import expected_conditions as ec
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from bs4 import BeautifulSoup

from constants import *
cn = "Doc's Jack Frost"
browser = getGoogleDriver()
browser.get("https://beta.allbreedpedigree.com/search?query_type=check&search_bar=horse&g=5&inbred=Standard")
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
    print(getSireNameFromTable(table))
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
        else:
            try:
                select = Select(browser.find_element(By.XPATH, "//select[@id='filter-match']"))
                select.select_by_value("exact")
            except:
                print("1")
        WebDriverWait(browser, 10).until(lambda browser: browser.execute_script('return document.readyState') == 'complete')
        soup = BeautifulSoup(browser.page_source, 'html.parser')
        table = soup.find(class_="pedigree-table").find("tbody")
        print(getSireNameFromTable(table))
    except:
        print("2")