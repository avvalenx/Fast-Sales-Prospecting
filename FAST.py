from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import pandas as pd
 
# TODO make interface and package for sales team
 
#create empty lists for each attribute being exported into excel
contact_names = []
contact_types = []
contact_days = []
contact_emails = []
contact_phones = []
 
#create dictionary to store buids and business names in
BUIDS = {}
excel_input = pd.read_excel('fast_input.xlsx', "Sheet1")
for num in range(len(excel_input['Business'])):
    BUIDS[str(excel_input['Business'][num])] = str(excel_input['Buid'][num])
 
#Open FAST and login
driver = webdriver.Edge()
driver.get('https://fansearch.web.att.com/fast/search?8&q')
driver.maximize_window()
username = driver.find_element(By.ID, 'GloATTUID')
password = driver.find_element(By.ID, 'GloPassword')
log_on_button = driver.find_element(By.ID, 'GloPasswordSubmit')
with open('creds.txt', 'r') as creds:
    username.send_keys(creds.readline().strip('\n'))
    password.send_keys(creds.readline())
log_on_button.click()
 
try:
    continue_button = WebDriverWait(driver, 10).until(EC.presence_of_element_located(By.ID, 'successButtonId'))
    continue_button.click()
except Exception:
    print('unable to login')
 
#main loop to grab contact data for each buid
for i, (business, buid) in enumerate(BUIDS.items()):
    print(i+1)
    driver.get('https://fansearch.web.att.com/fast/search?8&q')
    #find search bar and enter correct number
    try:
        search_bar = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, '/html/body/div[2]/div/div/div[2]/div/form/div/div/div[1]/div[2]/input[3]')))
        search_bar.send_keys(buid)
    except Exception:
        print('unable to locate search bar')
        df = pd.DataFrame({"Name": contact_names, "Type": contact_types, 'Days': contact_days, 'Email': contact_emails, 'Phone': contact_phones})
        #write dataframe created to excel sheet
        df.to_excel("fast_output.xlsx", sheet_name="Fast Contacts")
        print('data saved')
 
    #click search
    try:
        search_button = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, '/html/body/div[2]/div/div/div[2]/div/form/div/div/div[2]/button')))
        search_button.click()
    except Exception:
        print('unable to click search button')
 
    #click account
    try:
        account = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, '/html/body/div[2]/div/div/div[3]/div/form/div[2]/table/tbody/tr/td[3]/div/a')))
        account.click()
    except Exception:
        print('unable to click account link')
        df = pd.DataFrame({"Name": contact_names, "Type": contact_types, 'Days': contact_days, 'Email': contact_emails, 'Phone': contact_phones})
        #write dataframe created to excel sheet
        df.to_excel("fast_output.xlsx", sheet_name="Fast Contacts")
        print('data saved')
 
    #click customer contacts
    try:
        customer_contacts = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, '/html/body/div[2]/div/div/span/div[1]/ul/li[4]/a')))
        customer_contacts.click()
    except Exception:
        print('unable to click customer contacts tab')
 
    #add the business name to the columns for seperation
    contact_names.append(business)
    contact_types.append(buid)
    contact_days.append('')
    contact_emails.append('')
    contact_phones.append('')
 
    #loop to add name, type, email, and phone number into seperate list to be put in a data frame
    for name in WebDriverWait(driver, 10).until(EC.presence_of_all_elements_located((By.XPATH, '/html/body/div[2]/div/div/span/div[2]/form[1]/table/tbody/tr/td[1]/div'))):
        contact_names.append(name.text)
    for type in WebDriverWait(driver, 10).until(EC.presence_of_all_elements_located((By.XPATH, '/html/body/div[2]/div/div/span/div[2]/form[1]/table/tbody/tr/td[2]/div'))):
        contact_types.append(type.text)
    for day in WebDriverWait(driver, 10).until(EC.presence_of_all_elements_located((By.XPATH, '/html/body/div[2]/div/div/span/div[2]/form[1]/table/tbody/tr/td[6]/div'))):
        contact_days.append(day.text)
    for email in WebDriverWait(driver, 10).until(EC.presence_of_all_elements_located((By.XPATH, '/html/body/div[2]/div/div/span/div[2]/form[1]/table/tbody/tr/td[3]/div'))):
        contact_emails.append(email.text)
    for phone in WebDriverWait(driver, 10).until(EC.presence_of_all_elements_located((By.XPATH, '/html/body/div[2]/div/div/span/div[2]/form[1]/table/tbody/tr/td[10]/div'))):
        contact_phones.append(phone.text)
    try:
        home_button = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, '/html/body/div[1]/div/div/div[2]/ul[1]/li[1]/a')))
        home_button.click()
    except Exception:
        print('unable to click home button')
        df = pd.DataFrame({"Name": contact_names, "Type": contact_types, 'Days': contact_days, 'Email': contact_emails, 'Phone': contact_phones})
        #write dataframe created to excel sheet
        df.to_excel("fast_output.xlsx", sheet_name="Fast Contacts")
        print('data saved')
 
#create pandas dataframe of Name type email and phone
df = pd.DataFrame({"Name": contact_names, "Type": contact_types, 'Days': contact_days, 'Email': contact_emails, 'Phone': contact_phones})
#write dataframe created to excel sheet
df.to_excel("fast_output.xlsx", sheet_name="Fast Contacts")
print('done')