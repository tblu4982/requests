# READ THIS!
# You will need to provide two csv files, one named "accounts.csv" and "vnums.csv"
# "accounts.csv" needs your V-number first, and your PIN on a second line!
# "vnums.csv" needs a list of V-numbers, each on their own line!
# This program will log into banner your credentials provided by "accounts.csv" and
# generate a list of V-numbers from "vnums.csv" to crawl each student's banner
# and generate a new csv file called "credit_hours.csv"
# "credit_hours.csv" contains a list of information about each student:
#       V-number, Student Name, Earned Credits, Year, Enrolled Credits

import subprocess
import sys

try:
    from selenium import webdriver
except ModuleNotFoundError:
    subprocess.check_call([sys.executable, "-m", "pip", "install", "selenium"])
    from selenium import webdriver
    
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import Select
from selenium.common.exceptions import NoSuchElementException
from datetime import datetime
import re #used for regular expressions
import sys #used to stop execution under certain circumstances
import csv

#method to return appropriate value for term dropdown
def get_timecode():
    #get current month and year
    month = datetime.now().month
    year = str(datetime.now().year)
    #change month to string
    if month >= 1:
        #set month to "01" if month is between 1 and 4
        if month >= 5:
            #set month to "05" if month is between 5 and 7
            if month >= 8:
                #set month to "08" if month is between 8 and 11
                if month == 12:
                    #change month to string if month is 12
                    month = "12"
                else:
                    month = "08"
            else:
                month = "05"
        else:
            month = "01"
    #return year and month to use as value for dropdown selector
    return year + month

def shift_timecode(date):
    year = int(date[:4])
    month = date[4:]
    
    if month == "12":
        year += 1
        
    year = str(year)
    
    match month:
        case "01":
            month = "05"
        case "05":
            month = "08"
        case "08":
            month = "12"
        case "12":
            month = "01"
            
    return year + month

def build_files():
    driver.find_element(By.XPATH, "//table[contains(@class, 'dataentrytable')]").submit()

    #scrape transcript for courses and semesters
    data = driver.find_elements(By.XPATH, "//table[@class = 'datadisplaytable']")

    #used to determine if we reached current/future semesters
    at_target = False
    at_bottom = False
    #used to hold course as we build it from data
    credit_hours = 0.0
    earned_credits = 0.0

    for i in data:
        output = i.text.split('\n')

    i = 0
    for line in output:
        #Find each semester as we iterate through scraped data
        #signifies the start of courses in progress
        if re.search("COURSES IN PROGRESS", line):
            at_target = True
        if re.search("Overall:", line):
            at_bottom = True
        
        if at_target:
            try:
                target = float(line)
                credit_hours += target
            except Exception:
                continue
                
        if at_bottom:
                i += 1
                if i == 3:
                    earned_credits = line
    
    return [earned_credits, credit_hours]

username = ''
password = ''

with open ("accounts.csv", 'r') as myfile:
    reader = csv.reader(myfile)
    arr = list(reader)
    for i in range(len(arr)):
        arr[i] = arr[i]
    username, password = arr[0][0], arr[1][0]

#--------ADD A CHECK TO SEE IF WEBDRIVER IS CURRENT VERSION--------
try:
    driver = webdriver.Edge()
    driver.get('https://ssb-prod.ec.vsu.edu/BNPROD/twbkwbis.P_WWWLogin')
except Exception as e:
    print(e)
    print("Error getting webdriver. It may be out of date!")
    sys.exit()
#--------ADD A CHECK TO SEE IF WEBDRIVER IS CURRENT VERSION--------

#Track if user is at landing page
home_url = "https://ssb-prod.ec.vsu.edu/BNPROD/twbkwbis.P_GenMenu?name=bmenu.P_MainMnu"

uid = driver.find_element(By.ID, "UserID")
pwd = driver.find_element(By.ID, "PIN")

uid.send_keys(username)
pwd.send_keys(password)

driver.find_element(By.XPATH, "//form").submit()

vnums = None

#create array of vnums from file
with open ("vnums.csv", 'r') as vnumfile:
    reader = csv.reader(vnumfile)
    vnums = list(reader)
    
for i in range(len(vnums)):
    vnums[i] = vnums[i][0]

credit_hours = []
error_vnums = []
fullnames = []

itr = 0
#crawl through banner until to get to transcript
driver.find_element(By.LINK_TEXT, "Faculty and Advisors").click()
index = 0
first = True
while index < len(vnums):
    driver.find_element(By.LINK_TEXT, "Student Information Menu").click()
    driver.find_element(By.LINK_TEXT, "ID Selection").click()
    #The following dropdown menu lacks an id
    #so to select for current semester, I noticed html code stored
    #dropdown values as dates, I used a method to grab current
    #month and year, and store it to a variable to use as value
    date = get_timecode()
    #We visit the 'Select Term' page only once, during the first time we crawl
    if first:
        term = Select(driver.find_element(By.XPATH, "//select[@name='term']"))
        term.select_by_value(date)
        driver.find_element(By.XPATH, "//td[@class='dedefault']").submit()
        first = False
    #requests student id and uses it to get to transcript
    sid = driver.find_element(By.ID, "Stu_ID")
    sid.send_keys(vnums[index])
    driver.find_element(By.XPATH, "//input[@type='submit' and @value='Submit']").submit()
    try:
        if driver.find_element(By.CSS_SELECTOR, "span.errortext"):
            driver.find_element(By.XPATH, "//input[@type='submit' and @value='Submit']").submit()
            driver.find_element(By.LINK_TEXT, "Faculty Services").click()
            driver.find_element(By.LINK_TEXT, "Term Selection").click()
            tdate = shift_timecode(date)
            print(f"timecode shifted from {date} to {tdate}, attempting again")
            term = Select(driver.find_element(By.XPATH, "//select[@name='term']"))
            term.select_by_value(tdate)
            driver.find_element(By.XPATH, "//td[@class='dedefault']").submit()
            driver.find_element(By.LINK_TEXT, "Student Information Menu").click()
            driver.find_element(By.LINK_TEXT, "ID Selection").click()
            sid = driver.find_element(By.ID, "Stu_ID")
            sid.send_keys(vnums[index])
            driver.find_element(By.XPATH, "//input[@type='submit' and @value='Submit']").submit()
    except NoSuchElementException:
        pass
    #Additional crawling logic if we've reached the last vnum in list
    if index == len(vnums):
        driver.find_element(By.XPATH, "//input[@type='submit' and @value='Submit']").submit()
        #gets student name
        try:
            driver.find_element(By.XPATH, "//input[@type='submit' and @value='Submit']").submit()
            #driver.find_element(By.XPATH, "//input[@type='submit' and @value='Submit']").submit()
            name = driver.find_element(By.TAG_NAME, 'b')
            fullnames.append(name.text.replace(" ", "").replace(".", ""))
        #if this error occurs, then v-number is not associated with a student
        except NoSuchElementException:
            print("Error occurred at " + vnums[index])
            error_vnums.append(vnums.pop(index))
            driver.find_element(By.XPATH, "//input[@type='submit' and @value='Submit']").submit()
            driver.find_element(By.LINK_TEXT, "Faculty Services").click()
            index += 1
            continue
    #For all other vnums that are not last in list
    else:
        try:
            #driver.find_element(By.XPATH, "//input[@type='submit' and @value='Submit']").submit()
            name = driver.find_element(By.TAG_NAME, 'b')
            fullnames.append(name.text.replace(" ", "").replace(".", ""))
        except NoSuchElementException:
            print("Error occurred at " + vnums[index])
            error_vnums.append(vnums.pop(index))
            driver.find_element(By.XPATH, "//input[@type='submit' and @value='Submit']").submit()
            driver.find_element(By.LINK_TEXT, "Faculty Services").click()
            index += 1
            continue
        

    driver.find_element(By.XPATH, "//input[@type='submit' and @value='Submit']").submit()
    #crawl webpage to academic transcript
    driver.find_element(By.LINK_TEXT, "Academic Transcript").click()

    levl_id = Select(driver.find_element(By.ID, "levl_id"))
    type_id = Select(driver.find_element(By.ID, "tprt_id"))

    credit_hours.append(build_files())

    driver.find_element(By.LINK_TEXT, "Faculty Services").click()
    itr += 1
    index += 1

#Web driver is no longer needed
driver.quit()
csv = ""
for i in range(len(vnums)):
    csv += f"{vnums[i]}, {credit_hours[i]},"

with open("credit hours.csv", 'w') as outfile:
    outfile.write("V-Number, Name, Earned Credits, Year, Credit Hours")
    for i in range(len(vnums)):
        earned_credits = float(credit_hours[i][0])
        class_year = ""
        if earned_credits >= 90.0:
            class_year = "Senior"
        elif earned_credits >= 60.0:
            class_year = "Junior"
        elif earned_credits >= 30.0:
            class_year = "Sophomore"
        else:
            class_year = "Freshman"
        if len(vnums[i]) > 9:
            outfile.write(f"\n{vnums[i][:-1]}, {fullnames[i]}, {earned_credits}, {class_year}, {credit_hours[i][1]}")
        else:
            outfile.write(f"\n{vnums[i]}, {fullnames[i]}, {earned_credits}, {class_year}, {credit_hours[i][1]}")

print('Program complete! Check files for advisory report(s).')
if len(error_vnums) > 0:
    print('Could not generate transcripts for the following V-Numbers:')
    for vnum in error_vnums:
        print(vnum)
