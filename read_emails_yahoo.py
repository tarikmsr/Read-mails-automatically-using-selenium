#pip install pandas, selenium
#pip install openpyxl  //to read from excel

#change extesion proxy

import time
from selenium import webdriver
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.common.keys import Keys
from selenium.webdriver import ActionChains

import pandas as pd


list_data_file = pd.read_excel ('emails_list.xlsx')

listOfEmails= []
listOfPasswd=[]
for row in range(0,len(list_data_file)):
    listOfEmails.append(list_data_file.loc[row][0])
    listOfPasswd.append(list_data_file.loc[row][1])


nbr_messages_want_to_read=2
path="" 

driver = webdriver.Chrome(ChromeDriverManager().install())

driver = webdriver.Chrome(executable_path=path+"chromedriver.exe")


for i in range(0,len(listOfEmails)):  
    #open in other tab !
    Link='https://login.yahoo.com/?done=https%3A%2F%2Fwww.yahoo.com%2F&add={}'
    driver.get(Link.format(i))
    time.sleep(1)        
    username = driver.find_element_by_name('username')
    username.send_keys(listOfEmails[i] + Keys.ENTER)
    time.sleep(2)
    password = driver.find_element_by_id('login-passwd')
    password.send_keys(listOfPasswd[i]+ Keys.ENTER)

    #end login, start composing
    driver.get('https://mail.yahoo.com/d/search/referrer=unread')
    time.sleep(2)

    #get number of unread messages
    elements = driver.find_elements_by_class_name('u_Z13VSE6')
    print("Unread emails number : ",len(elements))

    link='https://mail.yahoo.com/d/search/referrer=unread/messages/{}'
    for j in range(1,nbr_messages_want_to_read+1) : #len(elements)+1

        driver.get(link.format(j))
        time.sleep(2)
    
    
    driver.find_element_by_id('ybarAccountMenuOpener').click()
    time.sleep(1)
    driver.find_element_by_id('profile-signout-link').click()
    


