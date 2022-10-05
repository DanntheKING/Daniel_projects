import tkinter as tk
from tkinter import ttk
import selenium.common.exceptions
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.action_chains import ActionChains
import time
from openpyxl import load_workbook

# This allows for the program to interact with the webpage (The chromedriver.exe must be updated when a Google Chrome
#  browser update is released)
PATH = "C:\PycharmProjects\PythonProjects\Python\ArticleCounter\chromedriver.exe"
driver = webdriver.Chrome(PATH)
# Logs into Meltwater Account
try:
    driver.get("https://app.meltwater.com/explore/list?tab=searches")
    time.sleep(4)
    user = driver.find_element_by_id('input_0')
    user.send_keys("")
    time.sleep(4)
    user.send_keys(Keys.RETURN)
    time.sleep(4)
    password = driver.find_element_by_name('password')
    password.send_keys("")
    password.send_keys(Keys.RETURN)
except selenium.common.exceptions.NoSuchElementException:
    time.sleep(5)
    driver.refresh()
    driver.get("https://app.meltwater.com/explore/list?tab=searches")
    time.sleep(4)
    user = driver.find_element_by_id('input_0')
    user.send_keys("AFGlobalStrike@gmail.com")
    time.sleep(4)
    user.send_keys(Keys.RETURN)
    time.sleep(7)
    password = driver.find_element_by_name('password')
    password.send_keys("SACback2020!")
    password.send_keys(Keys.RETURN)
# Waits for Meltwater to load webpage
print("Wait time")
wait = WebDriverWait(driver, 60)
wait.until(EC.visibility_of_element_located((By.XPATH,
                                             '/html/body/ui-view/app-root/mi-app-chrome/mi-app-chrome-content/div[3]/div/md-content/ui-view/explore-list/div/list-header/div/div/ng-transclude/search-bar-slot/search-bar-wrapper/div/search-bar/form/md-chips/md-chips-wrap')))
print("Done waiting")

root = tk.Tk()
root.geometry("300x100")
root.title('Article Counter')
root.resizable(0, 0)

# configure the grid
root.columnconfigure(0, weight=1)
root.columnconfigure(1, weight=3)

# document
document_label = ttk.Label(root, text="Enter number of documents:")
document_label.grid(column=0, row=0, sticky=tk.W, padx=5, pady=5)

document = tk.Entry(root)
document.grid(column=1, row=0, sticky=tk.E, padx=5, pady=5)

excel_label = ttk.Label(root, text="Excel file name:")
excel_label.grid(column=0, row=1, sticky=tk.W, padx=5, pady=5)

excel_entry = ttk.Entry(root)
excel_entry.grid(column=1, row=1, sticky=tk.E, padx=5, pady=5)


def getinfo():
    articleCount = 0
    docCounter = int(document.get()) + 1
    Fname = excel_entry.get() + ".xlsx"
    # Will need the document count so the program will not look for extra documents that do not exist
    # Input the Excel file name you want the information to be exported to
    wb = load_workbook(Fname)
    wsheet = wb.active
    names = wsheet['A']
    # Counts the articles and exports that number to the Excel file
    try:
        for i in range(1, int(docCounter)):
            element = driver.find_element_by_xpath(
                "/html/body/ui-view/app-root/mi-app-chrome/mi-app-chrome-content/div[3]/div/md-content/ui-view/explore/div/md-content/div[1]/explore-host-container/ui-view/explore-overview/section/div[1]/content-stream-container/section/section/md-card/mw-content-stream/div/ng-transclude[3]/mw-content-stream-body/div[1]/mw-media-document[" + str(
                    i) + "]/div/div/div[1]/div[2]/mw-content-document-header/div/ng-transclude[1]/mw-content-document-source-avatar/mw-source-avatar/a/div[2]/mw-icon/md-icon").get_attribute(
                "md-svg-src")
            for x in range(0, len(names)):
                if names[x].value in element:
                    cell = 'B' + str(x + 1)
                    wsheet[cell].value += 1
                    wb.save(Fname)
                else:
                    print(element)

    # If the counter tries to count a document that doesn't exist then it will output the number of articles counted
    # so far
    except selenium.common.exceptions.NoSuchElementException:
        articleCount += 1
        print("Article count was not completed but stopped at document number %s" % i)

    finally:
        print("Article Counter is done")


submit = tk.Button(root, text="Submit", command=getinfo)
submit.grid(column=1, row=3, sticky=tk.E, padx=5, pady=5)

root.mainloop()
