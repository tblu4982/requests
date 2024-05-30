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
#lists to hold students by classification
fr = []
so = []
jr = []
sr = []

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
            elif re.search("Freshman|Sophomore|Junior|Senior", index):
                if re.search("Freshman", index):
                    fr.append(row)
                elif re.search("Sophomore", index):
                    so.append(row)
                elif re.search("Junior", index):
                    jr.append(row)
                elif re.search("Senior", index):
                    sr.append(row)
                else:
                    continue
            #checks to see if student has a CS advisor
            if re.search(advisor_query, index):
                has_advisor = True
        #if student doesn't have CS advisor, put them in the no advisors list
        if not has_advisor and columns_found:
            no_advisor.append(row)

#prepare lists to a dataframe
for c in no_advisor:
    print(c)
df1 = pd.DataFrame(fr, columns = columns)
df2 = pd.DataFrame(so, columns = columns)
df3 = pd.DataFrame(jr, columns = columns)
df4 = pd.DataFrame(sr, columns = columns)
df5 = pd.DataFrame(no_advisor, columns = columns)

#write dataframes to excel workbook
with pd.ExcelWriter('student_list.xlsx') as outfile:
    df1.to_excel(outfile, sheet_name = 'Freshmen')
    df2.to_excel(outfile, sheet_name = 'Sophomore')
    df3.to_excel(outfile, sheet_name = 'Junior')
    df4.to_excel(outfile, sheet_name = 'Senior')
    df5.to_excel(outfile, sheet_name = "No Advisor")
        