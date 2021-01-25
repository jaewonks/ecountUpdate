import time
from bs4 import BeautifulSoup
from selenium import webdriver
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
import gspread
from oauth2client.service_account import ServiceAccountCredentials

rootPath = '..'
driver = webdriver.Chrome(
    executable_path='./chromedriver'.format(rootPath)
)

url = 'https://login.ecount.com/ECERP'
driver.get(url)
driver.find_element_by_id('com_code').send_keys('603781')
driver.find_element_by_id('id').send_keys('haju2')
driver.find_element_by_id('passwd').send_keys('Angel1014!')
driver.find_element_by_id('save').click()
WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.XPATH, "//*[@id='ma4']"))).click()
driver.find_element_by_id('ma35').click()
WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.XPATH, "//*[@id='ma212']"))).click()
WebDriverWait(driver, 20).until(EC.element_to_be_clickable((By.XPATH, "//*[@id='search']"))).click()

time.sleep(5)
html = driver.page_source
soup = BeautifulSoup(html, 'html.parser')
table = soup.find('table', id='grid-main')
trs = table.tbody.find_all('tr')
itemlist = []
for tr in trs:
    code = tr.find("td", {"data-label": "품목코드"}).get_text()
    if tr.find("td", {"data-label": "품목명[규격]"}) is None:
        break
    name = tr.find("td", {"data-label": "품목명[규격]"}).get_text()
    qty = tr.find("td", {"data-label": "재고수량"}).get_text()
    # print(code, name, qty)
    itemlist.append([code, name, int(qty)])

driver.quit()

# print(itemlist)

scope = ['https://spreadsheets.google.com/feeds']
json_file_name = './ecountproject-65a3a002b9e0.json'
credentials = ServiceAccountCredentials.from_json_keyfile_name(json_file_name, scope)
gc = gspread.authorize(credentials)
spreadsheet_url = 'https://docs.google.com/spreadsheets/d/1EjgZZ6Z1O2-0gPnORQXVFCjK5Cu9GmA96Y5Mcrmn0bE/edit#gid=1032236171'

# 문서 불러오기
doc = gc.open_by_url(spreadsheet_url)
# 시트 불러오기
worksheet = doc.worksheet('wholelist')
worksheet.update(str('A2:C'+ str(2 + len(itemlist)-1)), itemlist)


