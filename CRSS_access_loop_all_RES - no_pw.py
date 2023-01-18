# This script automatically downloads daily MQ from CRSS
# Requirements:
#   1. Geckodriver should be in path
#   2. Pip install Zipfile, pyautogui, Selenium
#
# For improvements: 1. Automate connection to VPN
#                   2. Add Try-Except clauses for points of traceback
#                   3. Replace xpaths with a more accurate "find by" tag
#                   4. Replace time.sleep delays with better initiator for next line (i.e Page loaded)

from selenium import webdriver
from selenium.webdriver.common.keys import Keys
import time
from datetime import datetime, timedelta
import pyautogui
import os.path
import os
from zipfile import ZipFile
import shutil

# Function for checking if link is ready within 10 seconds
def pagewait(link, access, findby):
    for counter in range(1, 11):
        if counter < 10:
            try:
                if findby == "link": driver.get(link)
                elif findby == "xpath": driver.find_element_by_xpath(link).click()
                elif findby == "id": driver.find_element_by_id(link).click()
                break
            except:
                print("Retrying access to " + access + ": " + str(counter) + "/10")
                time.sleep(1)
        else:
            driver.quit()
            print("Failure access to " + access)
            exit()

# Set date formats
yesterday = datetime.now() - timedelta(days = 1)
dayformat1 = yesterday.strftime('%d')
date1 = yesterday.strftime('%Y%m%d')
if int(dayformat1) < 10:
    dayformat1 = dayformat1[1]
# print(dayformat1)

# Set data dump directory
strBCQfolderpath = "C:\\MQMasterFilePath"

# Set Firefox driver properties
profile = webdriver.FirefoxProfile()
profile.set_preference("browser.download.panel.shown", False)
profile.set_preference("browser.helperApps.neverAsk.openFile", 'application/zip')
profile.set_preference("browser.helperApps.neverAsk.saveToDisk", 'application/zip')
profile.set_preference("browser.download.folderList", 2)
profile.set_preference("browser.download.dir", strBCQfolderpath)

# Set participants list [Participant ID, username, password]
strAllParticipants = [
                        ["sampleID1", "userID1", "pw1"],
                        ["sampleID2", "userID2", "pw2"]
                      ]

# Initiate webdriver object instance and access CRSS website
driver = webdriver.Firefox(firefox_profile=profile)
strlogin = "https://crss-main.wesmsys.local/uaa/login?logout"
pagewait(strlogin, "login", "link")

# place here driver.setAcceptUntrustedCertificates(True)

# ----- START LOOP HERE -----

for strParticipantID in strAllParticipants:
    # Set credentials
    un = strParticipantID[1]
    pw = strParticipantID[2]

    # Set possible passwords
    pwMQ = (strParticipantID[0], strParticipantID[0] + "VIS", strParticipantID[0] + "SS")

    # Check if files are already present. delete if exists
    strFilePath = strBCQfolderpath + "\\" + strParticipantID[0] + "_METER_DATA_DAILY_null-null.zip"
    if os.path.exists(strFilePath):
        os.remove(strFilePath)

    # Clear and input username
    username = driver.find_element_by_id("username")
    username.clear()
    username.send_keys(un)

    # Clear and input password
    password = driver.find_element_by_id("password")
    password.clear()
    password.send_keys(pw)

    # Press Submit to login
    driver.find_element_by_id("login_submit_btn").click()
    time.sleep(10)

    # Access MQ/Settlement page
    strsettlementmq = "https://crss-main.wesmsys.local/#/tp-files"
    pagewait(strsettlementmq, "MQ/Settlement", "link")

    # Set initial loop break tag to False
    loopbreakout = False

    # Open the date picker
    strdatepickerxpath = "/html/body/div[1]/div/div[1]/div[2]/div/div[3]/div/div/div[2]/div[2]/div[3]/div/div/input"
    for counter in range(1, 11):
        if counter < 10:
            try:
                driver.find_element_by_xpath(strdatepickerxpath).click()
                break
            except:
                print("Retrying Date Picker: " + str(counter) + "/10")
                time.sleep(1)
        else:
            driver.quit()
            exit()

    # Set start row range:
    if int(dayformat1) < 20:
        rowrange = range(1, 7)
    else:
        rowrange = range(3, 7)

    # Loop search through the date picker
    for datepickerrows in rowrange:
        if loopbreakout == True:
            break
        for datepickercol in range(1, 8):
            strxpath = "/html/body/div[1]/div/div[1]/div[2]/div/div[3]/div/div/div[2]/div[2]/div[3]/div/div/div/div/table/tbody/tr[" + str(datepickerrows) + "]/td[" + str(datepickercol) + "]"
            #           /html/body/div[1]/div/div[1]/div[2]/div/div[3]/div/div/div[2]/div[2]/div[3]/div/div/div/div/table/tbody/tr[1]/td[5]
            # print("rows = " + str(datepickerrows))
            # print("col = " + str(datepickercol))
            try:
                driverdatepicker = driver.find_element_by_xpath(strxpath)
                if driverdatepicker.text == dayformat1:
                    time.sleep(1)
                    driverdatepicker.click()
                    loopbreakout = True
                    break
            except:
                pyautogui.alert("Date picker error")
                driver.quit()
                exit()

    # Click Search
    strxpathSearch = "/html/body/div[1]/div/div[1]/div[2]/div/div[3]/div/div/div[2]/div[2]/div[4]/button"
    for counter in range(1, 11):
        if counter < 10:
            try:
                driver.find_element_by_xpath(strxpathSearch).click()
                break
            except:
                print("Trying Search: " + str(counter) + "/10")
                time.sleep(1)
        else:
            pyautogui.alert("Search button not found")
            driver.quit()
            exit()

    # Click Metering
    strxpathMetering = "/html/body/div[1]/div/div[1]/div[2]/div/div[3]/div/div/div/div[2]/div[6]/div/div[1]/button"
    for counter in range(1, 11):
        if counter < 10:
            try:
                driver.find_element_by_xpath(strxpathMetering).click()
                break
            except:
                print("Trying Metering: " + str(counter) + "/10")
                time.sleep(1)
        else:
            print("Metering button not found")
            driver.quit()
            exit()

    # Click Download
    strxpathDL = "/html/body/div[1]/div/div[1]/div[2]/div/div[3]/div/div/div[2]/div[2]/div[6]/div/div[2]/div/div/span[4]/div/div/div/div[1]/div[2]/table/tbody/tr/td[3]/span/div/span/span/a[1]"
    for counter in range(1, 11):
        if counter < 10:
            try:
                driver.find_element_by_xpath(strxpathDL).click()
                break
            except:
                print("Trying download button: " + str(counter) + "/10")
                time.sleep(1)
        else:
            print("No Downloadable file!")
            driver.quit()
            exit()

    # Wait for download to finish
    while not os.path.exists(strFilePath):
        time.sleep(1)

    # Logout
    driver.get("https://crss-main.wesmsys.local/uaa/login?logout")

    # Extract .zip file Level 1
    for file in os.listdir(strBCQfolderpath):
        if file.endswith(".zip"):
            fileZip = strBCQfolderpath + "\\" + file
            with ZipFile(fileZip) as objZip:
                objZip.extractall(path=strBCQfolderpath)
            os.remove(fileZip)

    # Extract all .zip files Level 2 with PW
    for file in os.listdir(strBCQfolderpath):
        if file.endswith(".zip"):
            fileZip = strBCQfolderpath + "\\" + file
            with ZipFile(fileZip, "r") as objZip:
                for password in pwMQ:
                    try:
                        objZip.extractall(path=strBCQfolderpath, pwd=bytes("iemop" + password, "utf-8"))
                        break
                    except:
                        continue
            os.remove(fileZip)
            strBCQPathCreated = strBCQfolderpath + "\\" + password
            for file2 in os.listdir(strBCQPathCreated):
                source = strBCQPathCreated + "\\" + file2
                destination = strBCQfolderpath + "\\" + file2
                shutil.move(source, destination)
    os.rmdir(strBCQPathCreated)
    time.sleep(1)
# ----- END LOOP HERE -----

# Quit selenium window
time.sleep(1)
driver.quit()
