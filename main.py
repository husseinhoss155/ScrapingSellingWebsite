from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import pandas as pd
import os
from UliPlot.XLSX import auto_adjust_xlsx_column_width
from pathlib import Path
import time
start_time = time.time()


driver = webdriver.Chrome('chromedriver.exe')

driver.get('https://taxdata.nashcountync.gov/search/commonsearch.aspx?mode=parid')

df = pd.read_excel('jobs.xlsx',converters={'Parcel':str})
parcels = df.to_numpy()

data = []

for i in range(30):
    row = []
    row.append(parcels[i][0])
    element = WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((By.ID, "inpParid"))
    )
    element.send_keys(str(parcels[i][0]))
    driver.find_element(By.ID,'btSearch').click()


    #Owner
    element = WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((By.CSS_SELECTOR, "#Owner\ Details > tbody > tr:nth-child(1) > td.DataletData"))
    )
    row.append(element.text)
    #Address
    row.append(driver.find_element(By.CSS_SELECTOR,'#Parcel > tbody > tr:nth-child(2) > td.DataletData').text)
    driver.find_element(By.CSS_SELECTOR,'#sidemenu > ul > li:nth-child(2) > a').click()

    #Sale date
    element = WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((By.CSS_SELECTOR, "#Sale\ Details > tbody > tr:nth-child(1) > td.DataletData"))
    )
    row.append(element.text)
    #price
    row.append(driver.find_element(By.CSS_SELECTOR, '#Sales > tbody > tr:nth-child(2) > td:nth-child(2)').text)
    driver.find_element(By.CSS_SELECTOR,'#sidemenu > ul > li:nth-child(1) > a').click()

    #acres
    element = WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((By.CSS_SELECTOR, "#Parcel > tbody > tr:nth-child(10) > td.DataletData"))
    )
    row.append(element.text)
    #PIN
    row.append(driver.find_element(By.CSS_SELECTOR,'#Parcel > tbody > tr:nth-child(1) > td.DataletData').text)
    #ph address
    row.append(driver.find_element(By.CSS_SELECTOR, '#Parcel > tbody > tr:nth-child(2) > td.DataletData').text)

    driver.find_element(By.CSS_SELECTOR,'#sidemenu > ul > li:nth-child(4) > a').click()

    #stories
    element = WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((By.CSS_SELECTOR, "#Residential > tbody > tr:nth-child(2) > td.DataletData"))
    )
    row.append(element.text)
    driver.find_element(By.CSS_SELECTOR,'#sidemenu > ul > li:nth-child(7) > a').click()

    #land value
    element = WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((By.CSS_SELECTOR, "#Values > tbody > tr:nth-child(1) > td.DataletData"))
    )
    row.append(element.text)
    #return home for new search
    driver.find_element(By.CSS_SELECTOR,'#topmenu > ul > li:nth-child(1) > a > span').click()

    element = WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((By.CSS_SELECTOR, "#content > div > table > tbody > tr > td > font:nth-child(4) > font > font > font > a > span"))
    )
    element.click()
    data.append(row)

df = pd.DataFrame(data=data,columns=['Parcel','Owner','Parcel address','sale date','price','acres','PIN','Physical address','Stories','Land value'])


#Exporting the dataframe in excel file with auto adjusting excel columns
with pd.ExcelWriter("Output.xlsx") as writer:
    df.to_excel(writer, sheet_name="MySheet")
    auto_adjust_xlsx_column_width(df, writer, sheet_name="MySheet", margin=0)

#Openeing the excel file
absolutePath = Path('Output.xlsx').resolve()
os.system(f'start Output.xlsx "{absolutePath}"')

print("--- %s seconds ---" % (time.time() - start_time))


