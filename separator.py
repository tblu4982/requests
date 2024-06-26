#This script requires pandas, if not already installed
#To use this script, you need a campus v2 report stored as a csv file
#Select that file and this script will sort the students out by classification and students who have no advisor

import re
import pandas as pd
import csv
from tkinter.filedialog import askopenfilename

#prompt user for source file path
file_path = askopenfilename(title = 'Select CSV File', filetypes = [('CSV Files', '*.csv')])
#search query for finding an advisor within student information
advisor_query = "Mohammed, Ahmed Farouk|Shelton, Joseph|Davis, Brittany|Wang, Ju|REYNOLDS, MICHAEL|Chen, Wei-Bang|Haris Rais, Muhammad"
#list to hold column titles from csv
columns = []
columns_found = False
#list to hold students who have no advisor
no_advisor = []
#lists to hold students by advisor
freshman = []
davis = []
wang = []
reynolds = []
chen = []
rais = []

#open csv and convert it into a 3d array
with open(file_path, 'r') as myfile:
    csvreader = csv.reader(myfile)
    #find the student's classification and store it in the appropriate list
    for row in csvreader:
        has_advisor = False

        for index in row:
            if not columns_found:
                if re.search("Student Alternate ID", index):
                    columns = row
                    columns_found = True
                    continue
            #checks to see if student has a CS advisor
            if re.search(advisor_query, index):
                if re.search("Mohammed, Ahmed Farouk|Shelton, Joseph", index):
                    freshman.append(row)
                if re.search("Davis, Brittany", index):
                    davis.append(row)
                if re.search("Wang, Ju", index):
                    wang.append(row)
                if re.search("REYNOLDS, MICHAEL", index):
                    reynolds.append(row)
                if re.search("Chen, Wei-Bang", index):
                    chen.append(row)
                if re.search("Haris Rais, Muhammad", index):
                    rais.append(row)
                has_advisor = True
        #if student doesn't have CS advisor, put them in the no advisors list
        if not has_advisor and columns_found:
            no_advisor.append(row)

#prepare lists to a dataframe
df1 = pd.DataFrame(freshman, columns = columns)
df2 = pd.DataFrame(davis, columns = columns)
df3 = pd.DataFrame(wang, columns = columns)
df4 = pd.DataFrame(reynolds, columns = columns)
df5 = pd.DataFrame(chen, columns = columns)
df6 = pd.DataFrame(rais, columns = columns)
df7 = pd.DataFrame(no_advisor, columns = columns)

#write dataframes to excel workbook
with pd.ExcelWriter('student_list.xlsx') as outfile:
    df1.to_excel(outfile, sheet_name = 'Freshmen')
    df2.to_excel(outfile, sheet_name = 'Davis')
    df3.to_excel(outfile, sheet_name = 'Wang')
    df4.to_excel(outfile, sheet_name = 'Reynolds')
    df5.to_excel(outfile, sheet_name = 'Chen')
    df6.to_excel(outfile, sheet_name = 'Haris Rais')
    df7.to_excel(outfile, sheet_name = "No Advisor")
        