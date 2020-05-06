from bs4 import BeautifulSoup
from datetime import date,datetime
import os,sys,subprocess,time,requests,keyboard,glob,getpass
from io import StringIO
from pdfminer.high_level import extract_text,extract_text_to_fp
from pdfminer.layout import LAParams
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.common.exceptions import NoSuchElementException,ElementNotInteractableException,TimeoutException

# Declaring static paths 
userdetailstext = r"C:\Users\\"+getpass.getuser()+r"\Documents\userDetails.txt"
powershellscript = r"C:\Users\\"+getpass.getuser()+r"\Documents\schedule.ps1"
convertedhtml = r"C:\Users\\"+getpass.getuser()+r"\Documents\converted.html"
applocation = r"C:\Users\\"+getpass.getuser()+r"\Documents\WebExAutomation.exe"
calendarpath = r"C:\Users\\"+getpass.getuser()+r"\Documents\Calendar*.pdf"

# defining a function to get abspath for absolute path
def resource_path(relative_path):
    try:
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")

    return os.path.join(base_path, relative_path)

print("\nA WebEx Automated Script by SoCalledEngineer\n")

#powershell script function
def power_shell(script_details):
    ps = open(powershellscript,"w+")
    ps.write(script_details)
    ps.close()
    try:
        p=subprocess.run(['powershell',"-ExecutionPolicy", "Unrestricted", powershellscript], stdout=subprocess.PIPE, stderr=subprocess.STDOUT, shell=True)
    except:
        print("\nSeems application not able to run powershell scripts which means I can't set scheduled tasks. Sorry ;-)")
        time.sleep(10)
        sys.exit()

# For creating a userdetails file
filename = open(userdetailstext,"a")
filename.close()
# For Checking whether userdetails file is empty or not
filename=open(userdetailstext,"r")
# If empty then copy details into text file
if len(filename.read().split(","))<7:
    filename = open(userdetailstext,"w+")
    print("The folowing details are asked for joining Webex Meeting\n")
    print("Details will be asked only during first run of this application. Enter details carefully.\n")
    print("Enter Your Name")
    filename.write(input()+",")
    print("Enter your Email")
    filename.write(input()+",")
    print("Enter your outlook user email. The one with ur emp id in it. like 833223@cognizant.com")
    filename.write(input()+",")
    print("Enter your outlook password. Make sure this is correct or else script will fail.")
    filename.write(input()+",")
    print("Enter your gmail username")
    filename.write(input()+",")
    print("Enter your gmail password.")
    filename.write(input()+",")
    print("While logging into outlook u receive otp via\n#if mobile enter 0\n#if gmail enter 1")
    filename.write(input())
    filename.close()
    f = open(userdetailstext,"r")
    userDetails = f.read().split(",")
    f.close()
    username = '-'.join(userDetails[0].split(" "))
    useremail = userDetails[1]
    outlookuser = userDetails[2]
    outlookpass = userDetails[3]
    gmailuser = userDetails[4]
    gmailpass = userDetails[5]
    timings = ["8:00AM","12:00PM","6:00PM"]
    for set_time in timings:
        power_shell("$A = New-ScheduledTaskAction -Execute "+'"'+applocation+'"'+"\n$T = New-ScheduledTaskTrigger -Daily -At "+set_time+"\n$S = New-ScheduledTaskSettingsSet -WakeToRun -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries -Priority 3 -ExecutionTimeLimit (New-TimeSpan -Hours 1)\n$D = New-ScheduledTask -Action $A -Trigger $T -Settings $S\nRegister-ScheduledTask MyTasks\WebExAutomation"+set_time.replace(":","-")+" -InputObject $D")
        
#if not empty then continue using that already populated data
else:
    f = open(userdetailstext,"r")
    userDetails = f.read().split(",")
    f.close()
    username = '-'.join(userDetails[0].split(" "))
    useremail = userDetails[1]
    outlookuser = userDetails[2]
    outlookpass = userDetails[3]
    gmailuser = userDetails[4]
    gmailpass = userDetails[5]
    
#initiating global driver variable
driver = ""

# defining driver function
def driver_func():
    global driver
    try:
        option = Options()
        option.add_extension(resource_path('extension_1_5_0_0.crx'))
        driver = webdriver.Chrome(options=option,executable_path=os.path.abspath(resource_path("chromedriver.exe")))
    except Exception:
        try:
            driver = webdriver.Firefox(executable_path=os.path.abspath(resource_path("geckodriver.exe")))
            extension_path = resource_path('cisco_webex_extension-1.5.0-fx.xpi')
            driver.install_addon(extension_path, temporary=True)
        except Exception:
            print("\nThis requires either chrome or firefox to be installed. Please install either of those and rerun this script ;-)")
    driver.maximize_window()
    driver.set_page_load_timeout(500)

#defining otp extraction function
def otp_extraction():
    global count6,otp,otpnum
    otpmsg = WebDriverWait(driver, 60).until(EC.element_to_be_clickable(otp)).text
    if "Use this code for Microsoft verification" in otpmsg:
        otpnum = otpmsg.split("Use this code for Microsoft verification")[0].splitlines()[1][:6]
    elif "Use this code for Microsoft verification" not in otpmsg and count6==1:
        count6 = count6+1
        driver.refresh()
        otp_extraction()
    elif "Use this code for Microsoft verification" not in otpmsg and count6>1 and count6<3:
        count6=count6+1
        otp = (By.XPATH,"(//span[@class='y2'])[2]")
        gtime = (By.XPATH,"(//td[@class='xW xY ']/span)[2]")
        driver.refresh()
        otp_extraction()
    elif count6==3:
        print("\nSeems u have not received otp. Please download ur calender pdf manually. If already downloaded the program will automatically schedule the meeting accoding to pdf")
        driver.quit()
        pdf_extract()
        time.sleep(10)
        sys.exit()
    return otpnum

# defining meeting details function
def detail_extraction():
    global links,names,times
    for i in driver.find_elements_by_xpath("//div[@class='_3J0DlEF3-tXqehZJzmfbpa']"):
        links.append(i.get_attribute("title"))
    for j in driver.find_elements_by_xpath("//div[@class='_3J0DlEF3-tXqehZJzmfbpa _2cWS8pNAe8oDryhpkwqLhL']"):
        names.append(j.get_attribute("title"))
    for k in driver.find_elements_by_xpath("//div[@class='_2TlB2Y6FLDwMsOcsJiLqeA HscWDJMhQ_BFsXeYXDCXo']"):
        times.append(k.get_attribute("title"))

#defining meetings grab functon
def meetings_grab_func():
    global count5,otp,otpnumber,meetingname,meetinglink
    try:
        driver.get("https://outlook.office365.com/calendar/view/month")
        if "login.microsoftonline.com" in driver.current_url:
            email = (By.ID, "i0116")
            password = (By.ID, "i0118")
            nextbtn = (By.ID, "idSIButton9")
            gmail= (By.ID,"identifierId")
            gpass = (By.NAME,"password")
            nextgmail = (By.ID,"identifierNext")
            passnext = (By.ID,"passwordNext")
            otp = (By.XPATH,"(//span[@class='y2'])[1]")
            otpenter = (By.ID,"idTxtBx_SAOTCC_OTC")
            verify = (By.ID,"idSubmit_SAOTCC_Continue")
            WebDriverWait(driver, 60).until(EC.element_to_be_clickable(email)).send_keys(outlookuser)
            WebDriverWait(driver, 60).until(EC.element_to_be_clickable(nextbtn)).click()
            WebDriverWait(driver, 60).until(EC.element_to_be_clickable(password)).send_keys(outlookpass)
            WebDriverWait(driver, 60).until(EC.element_to_be_clickable(nextbtn)).click()
            time.sleep(10)
            driver.execute_script("window.open('https://accounts.google.com/AccountChooser?service=mail&continue=https://mail.google.com/mail/');")
            driver.switch_to.window(driver.window_handles[1])
            WebDriverWait(driver, 60).until(EC.element_to_be_clickable(gmail)).send_keys(gmailuser)
            WebDriverWait(driver, 60).until(EC.element_to_be_clickable(nextgmail)).click()
            WebDriverWait(driver, 60).until(EC.element_to_be_clickable(gpass)).send_keys(gmailpass)
            WebDriverWait(driver, 60).until(EC.element_to_be_clickable(passnext)).click()
            otpnumber = otp_extraction()
            driver.close()
            driver.switch_to.window(driver.window_handles[0])
            WebDriverWait(driver, 60).until(EC.element_to_be_clickable(otpenter)).send_keys(otpnumber)
            WebDriverWait(driver, 60).until(EC.element_to_be_clickable(verify)).click()
            try:
                WebDriverWait(driver, 2).until(EC.presence_of_element_located((By.ID,"idSpan_SAOTCC_Error_OTC")))
                driver.quit()
                print("\nSeems u have not received otp. Please download ur calender pdf manually. If already downloaded the program will automatically schedule the meeting accoding to pdf")
                pdf_extract()
                sys.exit()
            except (TimeoutException, NoSuchElementException, ElementNotInteractableException):
                WebDriverWait(driver, 60).until(EC.element_to_be_clickable(nextbtn)).click()
                WebDriverWait(driver, 60).until(EC.presence_of_element_located((By.XPATH,"(//div[@class='_2TlB2Y6FLDwMsOcsJiLqeA HscWDJMhQ_BFsXeYXDCXo'])")))
                detail_extraction()
        else:
            detail_extraction() 
    except (TimeoutException, NoSuchElementException, ElementNotInteractableException, Exception):
        if(count5==0):
            time.sleep(10)
            driver.quit()
            print("\nSeems Connection Speed is low...Will retry in the background and notify u once succeeded....")
            driver_func()
            count5=count5+1
            meetings_grab_func()
        else:
            time.sleep(10)
            driver.quit()
            driver_func()
            meetings_grab_func()
            
# defining url validation function
def url_ok(url2):
    r = requests.get(url2)
    return r.status_code == 200

# defining url checking function
def url_check(url,test,urlcount):
    while test:
        try:
            if url_ok(url):
                test = False
            else:
                url = url[:(len(url)-1)]
        except:
            if urlcount==0:
                print("\nSeems there is a problem with connection...will retry in the background")
                urlcount = urlcount+1
                time.sleep(5)
                url_check(url,test,urlcount)
            else:
                time.sleep(5)
                url_check(url,test,urlcount)
    return url
                        
# Data Extraction from pdf
def pdf_extract():
    output_string = StringIO()
    try:
        with open(glob.glob(calendarpath)[len(glob.glob(calendarpath))-1], "rb") as fin:
            extract_text_to_fp(fin, output_string, laparams=LAParams(), output_type='html', codec=None)
        f = open(convertedhtml,"w+",encoding='utf-8')
        f.write(output_string.getvalue().strip())
        f.close()
    except Exception:
        print("\nI Cannot find ur outlook calendar file. Please ensure that file exits in the same directory as this application ;-)")
        time.sleep(10)
        sys.exit()
    # Extracting data from converted html
    soup = BeautifulSoup(open(convertedhtml,encoding='utf-8'),'html.parser')
    global text1,text2,text3,text
    text1,text2,text3="","",""
    text = []
    total = soup.find_all('div')
    for child in total:
        span1 = child.find_all('span',{"style":"font-family: SegoeUI-Semibold; font-size:11px"})
        span2 = child.find_all('span',{"style":"font-family: SegoeUI; font-size:9px"})
        span3 = child.find_all('span',{"style":"font-family: SegoeUI; font-size:9px"})
        for child1 in span1:
            text1 = child1.get_text()
        for child2 in span2:
            if "2020" in child2.get_text():
                text2 = child2.get_text()
        for child3 in span3:
            if "cognizanttraining.webex.com" in child3.get_text() or "cognizantcorp.webex.com" in child3.get_text() or "cognizant.webex.com" in child3.get_text() or "cognizant.kpoint.com" in child3.get_text():
                text3 = child3.get_text()
                text.append(text1 + "~" + text2.strip() + "~" + text3)
    for t in text:
        meetingDetails = t.split("~")
        startTime = meetingDetails[1].split(" ")
        meetingStartTime = startTime[2]+startTime[3]
        full_time = datetime.strftime(datetime.strptime(startTime[2]+" "+startTime[3], "%I:%M %p"), "%H:%M").split(":")
        presentDate = date.today().strftime("%#m/%#d/%Y")
        presentTime = datetime.now()
        test = True
        url = ''.join(meetingDetails[2].splitlines())
        urlcount = 0
        if startTime[1] == presentDate:
            if presentTime < presentTime.replace(hour=int(full_time[0]), minute=int(full_time[1]), second=0, microsecond=0) and "https://" in url or "www." in url:
                power_shell("$A = New-ScheduledTaskAction -Execute "+'"'+applocation+'"'+" -Argument "+'"'+url_check(url,test,urlcount)+'"'+"\n$T = New-ScheduledTaskTrigger -Once -At "+meetingStartTime+"\n$S = New-ScheduledTaskSettingsSet -WakeToRun -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries -Priority 3 -ExecutionTimeLimit (New-TimeSpan -Hours 1)"+"\n$D = New-ScheduledTask -Action $A -Trigger $T -Settings $S"+"\nRegister-ScheduledTask MyTasks\Task"+startTime[1].replace("/","-")+startTime[2].replace(":","-")+startTime[3]+" -InputObject $D")
            elif presentTime > presentTime.replace(hour=int(full_time[0]), minute=int(full_time[1]), second=0, microsecond=0) and "https://" in url or "www." in url:
                power_shell(r"Unregister-ScheduledTask -TaskPath '\MyTasks\' -TaskName Task"+startTime[1].replace("/","-")+startTime[2].replace(":","-")+startTime[3]+" -Confirm:$false")
    print("\n***WebEx Meetings Scheduled for the day and removed previous tasks if existed***")
            
# A simple logic to handle application if no args are given
try:
    if len(sys.argv[1])>3:
        args_exists = '1'
except:
    args_exists = '0'
        
#if no args given
if args_exists == '0':
    count5,count6 = 0,1
    links,names,times = [],[],[]
    otp,otpnum,otpnumber = "","",""
    #calling calender func()
    if userDetails[6]=='1':
        startingTime=datetime.now()
        driver_func()
        meetings_grab_func()
        driver.quit()
        endingTime = datetime.now()
        if (endingTime-startingTime).total_seconds()<20:
            pdf_extract()
        else:
            for z in range(len(times)):
                full_time = datetime.strftime(datetime.strptime(times[z], "%I:%M %p"), "%H:%M").split(":")
                presentTime = datetime.now()
                presentDate = date.today().strftime("%#m-%#d-%Y")
                test = True
                urlcount = 0
                if presentTime < presentTime.replace(hour=int(full_time[0]), minute=int(full_time[1]), second=0, microsecond=0) and "https://" in links[z] or "www." in links[z]:
                    power_shell("$A = New-ScheduledTaskAction -Execute "+'"'+applocation+'"'+" -Argument "+'"'+url_check(links[z],test,urlcount)+'"'+"\n$T = New-ScheduledTaskTrigger -Once -At "+times[z].replace(" ","")+"\n$S = New-ScheduledTaskSettingsSet -WakeToRun -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries -Priority 3 -ExecutionTimeLimit (New-TimeSpan -Hours 1)"+"\n$D = New-ScheduledTask -Action $A -Trigger $T -Settings $S"+"\nRegister-ScheduledTask MyTasks\Task"+presentDate+(times[z].replace(" ","-")).replace(":","-")+" -InputObject $D")
                elif presentTime > presentTime.replace(hour=int(full_time[0]), minute=int(full_time[1]), second=0, microsecond=0) and "https://" in links[z] or "www." in links[z]:
                    power_shell(r"Unregister-ScheduledTask -TaskPath '\MyTasks\' -TaskName Task"+presentDate+(times[z].replace(" ","-")).replace(":","-")+" -Confirm:$false")
            print("\n***WebEx Meetings Scheduled for the day and removed previous tasks if existed***")
    elif userDetails[6]=='0':
        pdf_extract() 
elif args_exists == '1':
    count1,count2,count3,count4=0,0,0,0

    # defining a function to tap on join as participant
    def join_as_participant():
        iframeParent = driver.find_element_by_name('mainFrame')
        driver.switch_to.frame(iframeParent)
        iframeChild = driver.find_element_by_name('main')
        driver.switch_to.frame(iframeChild)
        button = driver.find_element_by_link_text('join as a participant')
        button.click()
        
    # defining a function to perform website opening operation and stuff
    def main_func():
        global count1
        try:
            driver.get(sys.argv[1])
            join_as_participant()
        except (TimeoutException, NoSuchElementException, ElementNotInteractableException, Exception):
            if(count1==0):
                time.sleep(10)
                print("\nSeems Connection Speed is low...Will retry in the background and notify u once succeeded....\n")
                count1=count1+1
                main_func()
            else:
                time.sleep(10)
                main_func()
                
    # defining secondary function for webex meet links
    def second_main_func():
        global count2
        try:
            driver.get(sys.argv[1])
        except (TimeoutException, NoSuchElementException, ElementNotInteractableException, Exception):
            if(count2==0):
                time.sleep(10)
                print("\nSeems Connection Speed is low...Will retry in the background and notify u once succeeded....")
                count2=count2+1
                second_main_func()
            else:
                time.sleep(10)
                second_main_func()
                
    #defining kpoint function
    def third_main_func():
        global count4
        try:
            driver.get(sys.argv[1])
            email = (By.ID, "i0116")
            password = (By.ID, "i0118")
            nextbtn = (By.ID, "idSIButton9")
            WebDriverWait(driver, 60).until(EC.element_to_be_clickable(email)).send_keys(outlookuser)
            WebDriverWait(driver, 60).until(EC.element_to_be_clickable(nextbtn)).click()
            WebDriverWait(driver, 60).until(EC.element_to_be_clickable(password)).send_keys(outlookpass)
            WebDriverWait(driver, 60).until(EC.element_to_be_clickable(nextbtn)).click()
            WebDriverWait(driver, 60).until(EC.element_to_be_clickable(nextbtn)).click()
        except (TimeoutException, NoSuchElementException, ElementNotInteractableException, Exception):
            if(count4==0):
                time.sleep(10)
                print("\nSeems Connection Speed is low...Will retry in the background and notify u once succeeded....")
                count4=count4+1
                third_main_func()
            else:
                time.sleep(10)
                third_main_func()
                
    # defining a recursive function to retry join if meeting not started at regular interval
    def retry_to_join():
        global count3
        try:
            user = driver.find_element_by_id('join.label.userName')
            user.send_keys(username)
            email = driver.find_element_by_id('join.label.emailAddress')
            email.send_keys(useremail)
            join = driver.find_element_by_class_name('join').click()
        except (TimeoutException, NoSuchElementException, ElementNotInteractableException, Exception):
            if(count3==0):
                print("\nSeems Meeting is Not Yet started. Will retry to join in a regular interval and notify you once joined....")
                time.sleep(120)
                count3=count3+1
                main_func()
                retry_to_join()
            else:
                time.sleep(120)
                main_func()
                retry_to_join() 
        return "\n**** You Have Successfully Joined In The Meeting ;-) ****"
    
    # If it is WebEx Training then following login will be followed 
    if "cognizanttraining" in sys.argv[1] or "cognizant.webex.com" in sys.argv[1]:
        driver_func()
        main_func()
        print(retry_to_join())
    # If it is WebEx personal Room then following logic will be followed
    elif "meet" in sys.argv[1] or "join" in sys.argv[1]:
        driver_func()
        second_main_func()
        time.sleep(100)
        keyboard.press_and_release('enter')
        print("\n**** Seems you have successfully joined ;-) ****")
    elif "kpoint" in sys.argv[1]:
        driver_func()
        third_main_func()
    #If the User Chooses Wrong Choice
    else:
        print("\nThis link is not Compatible yet")
time.sleep(10)
