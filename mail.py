#from mail import sendmail
import time
import ast
import requests
import selenium
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import time as sl
import os, os.path, shutil, getpass, re, glob, socket, zipfile
import pandas as pd
import smtplib
import datetime
from datetime import datetime
import xlsxwriter

file = ''
filename = []
file_path = []
writer = ''

def DownloadExcelFile(driver,wait,fop,filename):
    #PATH = "D:\kwx1057920\Report_Automation\Python\chromedriver.exe"
    projects = ['NG MTN']
    global iles_path
    global yeah

    with open('C:\\Users\\'+getpass.getuser()+'\\Documents\\SAR_AMS_OWS_testbed_login.txt','r') as credentials:
        details = credentials.readlines()
        username,password,counter = details[0].strip(),details[1].strip(),int(details[2].strip())

    driver.maximize_window()

    driver.get("https://15fg-saapp.teleows.com/servicecreator/spl/HRWeeklyReportTrend/hrwkrep_managerslist_grid.spl?appId=11872&isVisualDesignerPreview=true&viewNewVersion=true")
    #driver.set_script_timeout("120")
    #email_list = pd.read_excel('C:/Users/user/Desktop/gfg.xlsx')

    #fill the login details username password and click the submit button
    username_tag = 'usernameInput'
    x_arg = f'//input[contains(@id,"{username_tag}")]'
    usernameInput = wait.until(EC.presence_of_element_located((By.XPATH, x_arg)))
    usernameInput.send_keys(username)

    password_tag = 'password'
    x_arg = f'//input[contains(@id,"{password_tag}")]'
    usernameInput = wait.until(EC.presence_of_element_located((By.XPATH, x_arg)))
    usernameInput.send_keys(password)

    login = 'btn_submit'
    x_arg = f'//div[contains(@id,"{login}")]'
    loginBtn = wait.until(EC.presence_of_element_located((By.XPATH, x_arg)))
    loginBtn.click()

    search_text = 'search_text'
    x_arg = f'//input[contains(@id,"{search_text}")]'
    projectInput = wait.until(EC.presence_of_element_located((By.XPATH, x_arg)))
    projectInput.send_keys("MTN")
    
    search = 'toolbarSearchButton'
    x_arg = f'//a[contains(@id,"{search}")]'
    searchBtn = wait.until(EC.presence_of_element_located((By.XPATH, x_arg)))
    searchBtn.click()

    #download the search result as excel and renames the file
    export = 'export'
    x_arg = f'//a[contains(@id,"{export}")]'
    exportBtn = wait.until(EC.presence_of_element_located((By.XPATH, x_arg)))
    exportBtn.click()
    driver.find_element_by_id('search_text').clear()

    downloadedfile = getDownLoadedFileName(50,driver)
    emailfile = downloadedfile
    #print(emailfile)
    time.sleep(5)

    new_name = []
    ath = "C:\\Users\\"+getpass.getuser()+"\\Downloads\\"
    iles_path = os.path.join(ath, emailfile)
    print(iles_path)
    driver.switch_to.window(driver.window_handles[0])

    sl.sleep(3)

    fop = []
    sitelist = ["https://15fg-saapp.teleows.com/servicecreator/spl/HRWeeklyReportTrend/hrwkrep_mainlist_grid.spl","https://15fg-saapp.teleows.com/servicecreator/spl/HRWeeklyReportTrend/hrwkrep_doa_intern_nysc_grid.spl?appId=11872&isVisualDesignerPreview=true&viewNewVersion=true"]
    for i in sitelist:
        fop = i
        print(fop)
        driver.switch_to.window(driver.window_handles[0])
        driver.get(i)

        #fill the search box for the projects and runs the search
        search_text = 'search_text'
        x_arg = f'//input[contains(@id,"{search_text}")]'
        projectInput = wait.until(EC.presence_of_element_located((By.XPATH, x_arg)))
        projectInput.send_keys("MTN")
        
        search = 'toolbarSearchButton'
        x_arg = f'//a[contains(@id,"{search}")]'
        searchBtn = wait.until(EC.presence_of_element_located((By.XPATH, x_arg)))
        searchBtn.click()

        #download the search result as excel and renames the file
        export = 'export'
        x_arg = f'//a[contains(@id,"{export}")]'
        exportBtn = wait.until(EC.presence_of_element_located((By.XPATH, x_arg)))
        exportBtn.click()
        driver.find_element_by_id('search_text').clear()

        a = 'ym-content'
        x_arg = f'//div[contains(@class,"{a}")]'
        b = len(x_arg)
        print(x_arg)
        print(b)
   
        
        latestDownloadedFileName = getDownLoadedFileName(50,driver) #waiting minutes to complete the download
        file = latestDownloadedFileName
        print(file)
        if file == None:
            pass
        else:
            time.sleep(5)
            #return file

            path = "C:\\Users\\"+getpass.getuser()+"\\Downloads\\"
            files_path = os.path.join(path, file)
            moment = datetime.now()
            #final_moment = moment.strftime("%Y-%m-%d %H-%M-%S")
            filename.append(files_path)
            print("Excel file has been Renamed")
            print(files_path)
            driver.switch_to.window(driver.window_handles[0])
            print(filename)

    print(filename)
    now = datetime.now()
    global nownow
    nownow = now.strftime("%Y-%m-%d")

    yeah = ('C:\\Users\\'+getpass.getuser()+'\\Documents\\HR WEEKLY REPORT\\MTN\\RNOC HR LIST_' + nownow + '.xlsx')
    writer = pd.ExcelWriter(yeah, engine='xlsxwriter')

    for excel_file in filename:
        #sheet = os.path.basename(excel_file)
        data = pd.read_excel(excel_file, sheet_name = None)
        key, value = list(data.items())[0]
        data = data.keys()
        print(key)
        df1 = pd.read_excel(excel_file)
        df1.fillna(value='', inplace=True)
        df1.to_excel(writer, sheet_name=key, index=True)
        print(df1)

        workbook = writer.book
        worksheet = writer.sheets[key]
        border_fmt = workbook.add_format({'bottom':1, 'top':1, 'left':1, 'right':1, 'text_wrap' : True})
        worksheet.conditional_format(xlsxwriter.utility.xl_range(0, 0, len(df1), len(df1.columns)), {'type': 'no_errors', 'format': border_fmt})
        header_format = workbook.add_format({
        'bold': True,
        'text_wrap': True,
        'valign': 'top',
        'fg_color': '#ADD8E6',
        'border': 1})

    # Write the column headers with the defined format.
        for col_num, value in enumerate(df1.columns.values):
            worksheet.write(0, col_num + 1, value, header_format)
        #for col_num, value in enumerate(df1.columns.values):
        #    worksheet.write()
    writer.save()
 
    if filename != []:
        sendMail(driver,writer)
    else:
        pass

def getDownLoadedFileName(waitTime,driver):
    driver.execute_script("window.open()")
        # switch to new tab
    driver.switch_to.window(driver.window_handles[1])
        # navigate to chrome downloads
    driver.get('chrome://downloads')

    # define the endTime
    endTime = time.time()+waitTime
    while True:
        try:
                # get downloaded percentage
            downloadPercentage = driver.execute_script(
                "return document.querySelector('downloads-manager').shadowRoot.querySelector('#downloadsList downloads-item').shadowRoot.querySelector('#progress').value")
                # check if downloadPercentage is 100 (otherwise the script will keep waiting)
            if downloadPercentage == 100:
                    # return the file name once the download is completed
                return driver.execute_script("return document.querySelector('downloads-manager').shadowRoot.querySelector('#downloadsList downloads-item').shadowRoot.querySelector('div#content  #file-link').text")
        except:
            pass
        time.sleep(1)
        if time.time() > endTime:
            break
    time.sleep(5)
    #driver.close()
 
def sendMail(driver,yeah):
    #download_path  = r'C:\Users\kWX1057920\Downloads'
    wait = WebDriverWait(driver, 600)
  
    driver.execute_script("window.open()")
    # switch to new tab
    driver.switch_to.window(driver.window_handles[1])

    with open('C:\\Users\\'+getpass.getuser()+'\\Documents\\SAR_AMS_OWS_testbed_login.txt','r') as credentials:
        details = credentials.readlines()
        username,password,counter = details[0].strip(),details[1].strip(),int(details[2].strip())

    driver.get("https://15fg-saapp.teleows.com/app/15fg/spl/Report_Email_Receiver/email_send_create_v2.spl")

    sl.sleep(5)
    driver.refresh()
    sl.sleep(3)

    #username_tag = 'usernameInput'
    #x_arg = f'//input[contains(@id,"{username_tag}")]'
    #usernameInput = wait.until(EC.presence_of_element_located((By.XPATH, x_arg)))
    #usernameInput.send_keys(username)

    print('kkk')

    #password_tag = 'password'
    #x_arg = f'//input[contains(@id,"{password_tag}")]'
    #usernameInput = wait.until(EC.presence_of_element_located((By.XPATH, x_arg)))
    #usernameInput.send_keys(password)

    #login = 'btn_submit'
    #x_arg = f'//div[contains(@id,"{login}")]'
    #loginBtn = wait.until(EC.presence_of_element_located((By.XPATH, x_arg)))
    #loginBtn.click()

    sl.sleep(3)
    email_list = pd.read_excel(iles_path)
  
# getting the names and the emails
    ile_path = email_list['Email']
    email = ';'.join(ile_path)

    #for i in range(len(emails)):
    #email = emails[i]
    
    # sending the email
    print(yeah)
    email_to = 'email_to'
    x_arg = f'//input[contains(@id,"{email_to}")]'
    emailToInput = wait.until(EC.presence_of_element_located((By.XPATH, x_arg)))
    emailToInput.send_keys(email)

    email_cc = 'email_cc'
    x_arg = f'//input[contains(@id,"{email_cc}")]'
    emailCCInput = wait.until(EC.presence_of_element_located((By.XPATH, x_arg)))
    emailCCInput.send_keys('liuyan159@huawei.com;joy.chinenye.ozoudeh@huawei.com;john.uzoma.onuoha@huawei.com')

    email_bcc = 'email_bcc'
    x_arg = f'//input[contains(@id,"{email_bcc}")]'
    emailBCCInput = wait.until(EC.presence_of_element_located((By.XPATH, x_arg)))
    emailBCCInput.send_keys('karunwi.oluwasegun.john@huawei.com;shodipe.ifeoluwa.oluwatobi@huawei.com')

    subject = 'title'
    x_arg = f'//input[contains(@id,"{subject}")]'
    usernameInput = wait.until(EC.presence_of_element_located((By.XPATH, x_arg)))
    usernameInput.send_keys("RNOC HR Project List (NG MTN)")

    content_tag = 'content'
    x_arg = f'//textarea[contains(@id,"{content_tag}")]'
    contentInput = wait.until(EC.presence_of_element_located((By.XPATH, x_arg)))
    contentInput.send_keys("""Dear Managers, \n\nKindly find attached updated HR List for your perusal. \n\nPlease check and share feedback. \n\n\nTHANKS, \nShodipe Ifeoluwa Oluwatobi\n08189678247\nshodipe.ifeoluwa.oluwatobi@huawei.com """)

    attachment = '_uploadFile'
    x_arg = f'//form/input[contains(@name,"{attachment}")]'
    at = r'C:\\Users\\'+getpass.getuser()+'\\Documents\\HR WEEKLY REPORT\\MTN\\RNOC HR LIST_' + nownow + '.xlsx'
    attachmentInput = wait.until(EC.presence_of_element_located((By.XPATH, x_arg)))
    attachmentInput.send_keys(at)
    
    print(yeah)
    sl.sleep(50) #file attachment delay

    submit = 'ServiceButton1'
    x_arg = f'//div[contains(@id,"{submit}")]'
    submitBtn = wait.until(EC.presence_of_element_located((By.XPATH, x_arg)))
    submitBtn.click()

    print("Mail Sending Completed...")
    driver.quit()
    print("Mail Sending Completed...")

def main():
    chrome_options = webdriver.ChromeOptions()
    chrome_options.add_argument('--ignore-certificate-errors')
    driver = webdriver.Chrome(options=chrome_options,executable_path=r"C:\Users\kWX1057920\Downloads\chromedriver.exe")
    wait = WebDriverWait(driver, 600)
    projects = ['NG MTN', 'NG Airtel']
    fop = []
    DownloadExcelFile(driver,wait,fop,filename)  
    sendMail(driver,writer)
    driver.quit()

main()