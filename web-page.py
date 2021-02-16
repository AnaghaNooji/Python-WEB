#Script to filter last mail from the outlook

import pyperclip 
import win32com.client,sys
from datetime import datetime
from selenium import webdriver 
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver import ActionChains
from selenium.webdriver.common.keys import Keys
import pyautogui as py

outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")

inbox = outlook.GetDefaultFolder(6) # "6" refers to the index of a folder - in this case,
                                     #the inbox. You can change that number to reference
                                    # any other folder

messages = inbox.Items
#Get the emAIL/S
message = messages.GetLast()

#Get today's date
today = datetime.now()

body_content = message.Subject
print(today)
print(body_content)


f = open("test.txt",'w')
f.write(body_content)
#f.write(today)
f.close()

f= open("test.txt",'r')
for ln in f:
    ln = ln.rstrip()
    id = ln.split()
ticket_id = id[1]

pyperclip.copy(ticket_id)

#we are getting into the chrome and doing copy paste & stuffs
driver = webdriver.Chrome(ChromeDriverManager().install())
driver.maximize_window()
url="https://www.google.com"
driver.get(url)
an = 'aan'
#inplace of "q" you can give the name that is right in you search bar
user_input = driver.find_element_by_name('q')
user_input.click()
user_input.send_keys(ticket_id)

action=ActionChains(driver)

#to understand how select and copy works 
#ignore these three lines if not necessary
click =driver.find_element_by_name('q') 
action.double_click(click).perform()
action.key_down(Keys.CONTROL).send_keys("c").perform()

#press enter 
driver.find_element_by_name("q").send_keys(Keys.ENTER)
