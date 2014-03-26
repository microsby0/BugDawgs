import os
import gspread
import time
from datetime import datetime

def setUpHeaders(student_sheet):
    student_sheet.update_acell('A1',"Events")
    student_sheet.update_acell('B1',"Name\n****Denotes pick up/drop off****")
    student_sheet.update_acell('C1',"Phone Number")
    student_sheet.update_acell('D1',"810")
    student_sheet.update_acell('E1',"Special Notes")
    student_sheet.update_acell('F1',"Call if you need me 706-614-3328 (cell) 706-542-1238 (office)\nYou can also send a text, which I can access more easily in meetings and at home. Thanks and Good Luck! \nMarianne")


def dataCount(rows):
    print "Number of Rows: " + str(len(rows)) #number of rows
    i=1
    for row in rows:
        print "Columns in Row: " + str(i) + ": " + str(len(row)) #number of columns
        i+=1

def clearCells(range):
    cell_list = student_sheet.range(range)
    for cell in cell_list:
        student_sheet.update_cell(int(cell.row), int(cell.col), "")

account = gspread.login(os.environ["GSPREAD_USERNAME"], os.environ["GSPREAD_PASSWORD"])

src_sheet  = account.open("Test").sheet1
student_sheet = account.open("Test Dest").sheet1
#final_sheet = account.open("Final Dest").sheet1

setUpHeaders(student_sheet)

src_rows = src_sheet.get_all_values()

content         = ""
groupName       = ""
visitDate       = ""
visitTime       = ""
visitLength     = ""
contactName     = ""
contactPhone    = ""
contactEmail    = ""
location        = ""
address         = ""
groupSize       = ""
groupAge        = ""

#message and headers are in row 1
currentRow = 2

#sformats dates in src sheet
for row in src_rows:
    date = datetime.strptime(row[2], "%m/%d/%Y")
    print date
    row[2] = {'formattedDate': date, 'rawDate': row[2]}

# sorts src_sheet based on date of event
src_rows = sorted(src_rows, key=lambda row: row[2]['formattedDate'])

print "\nSorted"

for row in src_rows:
    print row[2]['formattedDate']

#puts info from src_sheet into variables
for row in src_rows:
    groupName = row[1]

    #Visit Date
    d = datetime.strptime(row[2]['rawDate'], "%m/%d/%Y")
    #d = row[2]['formattedDate']
    visitDate = d.strftime("%A") + ", " + row[2]['rawDate']

    #Visit Time
    #formats time to AM/PM 12 hour
    time_values = time.strptime(row[3], "%H:%M:%S")
    visitTime = time.strftime("%I:%M %p", time_values)

    #Swap out for a field for visit end?
    visitLength     = row[4]

    contactName     = row[5]
    contactPhone    = row[6]
    contactEmail    = row[7]
    location        = row[8]
    address         = row[9]
    groupSize       = row[10]
    groupAge        = row[11]
    comments        = row[12]

    content = (
        groupName + ":\n" +
        visitDate + "\n\n" +
        visitTime + "\n\n" +
        "Contact: " + contactName + "\n" +
        contactPhone + "\n" +
        contactEmail + "\n\n" +
        "Age Group: " + groupAge + "\n" +
        "Number in Group: " + groupSize + "\n\n" +
        location + "\n" +
        address
    )

    student_sheet.update_cell(currentRow, 1, content)
    student_sheet.update_cell(currentRow,5, comments)
    currentRow += 1
    #print(content + "\n----------------------------------------")


student_rows = student_sheet.get_all_values()

for src_row in src_rows:
    groupName = src_row[1]
    contactName = src_row[5]
    print "Group " + groupName
    print "Contact " + contactName
    for student_row in student_rows:
        #student_row[0] = full content
        #student_row[1] = Names
        #student_row[3] = Phone number
        #student_row[4] = 810s
        #student_row[5] = special comments
        if groupName in student_row[0] and contactName in student_row[0]:
            print "\nMatch--------------------------"
            print "Src: " + groupName + " ," + contactName
            print "Student: " + student_row[0]












