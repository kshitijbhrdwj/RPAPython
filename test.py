
from selenium.webdriver import Chrome
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
import pandas as pd
import os,shutil
import time
import random
import lxml,html5lib

# open chrome
driver = Chrome("C://Users//kshitij//AppData//Local//Programs//Python//Python37//Scripts//chromedriver.exe")
driver.get('https://itdashboard.gov/')

# find 'DIVE IN' and click
dive_in_btn = driver.find_element_by_link_text('DIVE IN')
dive_in_btn.click()

# get agency name and total spending
agencies = []
spending = []
agencies_elements = driver.find_elements_by_xpath("//span[@class='h4 w200']")
spending_elements = driver.find_elements_by_xpath("//span[@class=' h1 w900']")

for element in agencies_elements:
    if element.text == "":
        break
    agencies.append(element.text)

count=0

for element in spending_elements:
    count+=1
    if count > len(spending_elements)/2:
        break
    else:
        spending.append(element.text)



# add above details to excel
my_list = [[agencies],[spending]]
df = pd.DataFrame.from_dict({'AGENCIES':agencies,'SPENDING':spending})




# select and click on random agency
# agency = random.choice(agencies)
agency = 'Department of Agriculture'
driver.find_element_by_xpath("//span[contains(@class,'h4 w200')]  [contains(text(),agency)]").click()



# get investment table of random selected agency

table_df = pd.DataFrame()
table_data = []
UII = []
table = WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.XPATH, '//table[@id="investments-table-object"]')))
x = pd.DataFrame()
x = pd.read_html(table.get_attribute('outerHTML'))
table_df = pd.DataFrame(x[0].values,columns=x[0].columns)
uii_data = driver.find_elements_by_xpath('//table[@id="investments-table-object"]/tbody/tr/td[1]/a[@href]')

download_path = os.path.join(os.getcwd(),agency)
os.mkdir(download_path)

for e in uii_data:
    e.click()
    dwn_buisness_pdf = driver.find_element_by_xpath('//*[@id="business-case-pdf"]/a')
    dwn_buisness_pdf.click()
    time.sleep(5)
    driver.execute_script("window.history.go(-2)")
    time.sleep(5)
    break               # breaking since downloading only one pdf


# move pdf to download_path
for file in os.listdir("C://Users//kshitij//Downloads"):
    if file.startswith('005-') and file.endswith('.pdf'):
        shutil.move(os.path.join("C://Users//kshitij//Downloads", file), download_path)






# writing data to excel
writer = pd.ExcelWriter('test.xlsx', engine='xlsxwriter')
frames = {'sheet1': df, agency: table_df}
for sheet, frame in frames.items():
    frame.to_excel(writer, sheet_name=sheet, index=False)

writer.save()



#close chrome
driver.close()
