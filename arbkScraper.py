from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
import time
import pandas as pd
import random

options = Options()
options = webdriver.ChromeOptions() 
options.add_argument("--incognito")

companyData = []
ownerData = []
unitData = []
representativeData = []
activityData = []

driver = webdriver.Chrome('C:\\Users\\syste\\Desktop\\ARBKScraper\\chromedriver.exe', chrome_options=options)
driver.get('https://arbk.rks-gov.net/')

driver.maximize_window()
time.sleep(1)

selectActivity = driver.find_element(By.ID, "select2-ddlnace-container")
time.sleep(1)
selectActivity.click()
lists = driver.find_elements(By.XPATH, "/html[1]/body[1]/span[1]/span[1]/span[2]/ul[1]/li")
# lists = driver.find_elements(By.TAG_NAME, 'li') 
conuntActivity = len(lists)

for activty in range(0,700):
    if activty <= conuntActivity:
        lists[activty].click()

        selectedActivityPath =  driver.find_element(By.ID, 'select2-ddlnace-container')
        selectedActivity = selectedActivityPath.text

        btn = driver.find_element(By.XPATH, "/html[1]/body[1]/form[1]/header[1]/div[3]/div[1]/div[1]/div[1]/div[1]/div[1]/div[1]/div[1]/div[7]/div[1]/input[1]")
        btn.click() 
        time.sleep(4)

        links = driver.find_elements(By.XPATH, "//tr/td[2]/a[@href]")
        time.sleep(2)
        countLinks = len(links)

        # for total in range(0, countLinks):
        #     if total <= countLinks:
        for getLink in links:
            link = getLink.get_attribute('href')

            driver.execute_script("window.open('');")
            second = driver.switch_to.window(driver.window_handles[1])
            driver.get(link)
            time.sleep(1)

            id = random.randint(1,1000000000)
            company = pd.read_html(link, header=None)[0]
            getFirstColumn= company.loc[:, [1]]
            transposeColumn = getFirstColumn.transpose()
            transposeColumn['id'] = id
            transposeColumn['arbkLink'] = link
            companyData.append(transposeColumn)
            companyList = pd.concat(companyData)
            
            represnetative = pd.read_html(link, header=None)[1]
            represnetative['id'] = id
            representativeData.append(represnetative)
            represnetativeList = pd.concat(representativeData)

            owner = pd.read_html(link, header=None)[2]
            owner['id'] = id
            ownerData.append(owner)
            ownerList = pd.concat(ownerData)

            unit = pd.read_html(link, header=None)[3]
            unit['id'] = id
            unitData.append(unit)
            unitList = pd.concat(unitData)

            activity = pd.read_html(link, header=None)[4]
            activity['id'] = id
            activityData.append(activity)
            activityList = pd.concat(activityData)

            driver.close()
            driver.switch_to.window(driver.window_handles[0])
            time.sleep(1)

        excelFile = pd.ExcelWriter(str('ARBKData.xlsx'))
        companyList.to_excel(excelFile, sheet_name='Company')
        represnetativeList.to_excel(excelFile, sheet_name='Representative Data')
        ownerList.to_excel(excelFile, sheet_name='Owner Data')
        unitList.to_excel(excelFile, sheet_name='Unit Data')
        activityList.to_excel(excelFile, sheet_name='Activity Data')
        excelFile.save()

        selectActivity = driver.find_element(By.XPATH, "/html[1]/body[1]/form[1]/header[1]/div[3]/div[1]/div[1]/div[1]/div[1]/div[1]/div[1]/div[1]/div[5]/span[1]/span[1]/span[1]/span[1]")
        time.sleep(1)
        selectActivity.click()
        lists = driver.find_elements(By.XPATH, "/html[1]/body[1]/span[1]/span[1]/span[2]/ul[1]/li")
        activty+1
        time.sleep(2)

print("Successfully Scraper")