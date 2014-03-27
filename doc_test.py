import os
import gspread
import time
from datetime import datetime
import re

def setUpHeaders(student_sheet):
    student_sheet.update_acell('A1',"Events")
    student_sheet.update_acell('B1',"Name\n****Denotes pick up/drop off****")
    student_sheet.update_acell('C1',"Phone Number")
    student_sheet.update_acell('D1',"810")
    student_sheet.update_acell('E1',"Special Notes")
    student_sheet.update_acell('F1',"Call if you need me 706-614-3328 (cell) 706-542-1238 (office)\nYou can also send a text, which I can access more easily in meetings and at home. Thanks and Good Luck! \nMarianne")

def clearCells(range):
    #A2:E number of rows
    cell_list = student_sheet.range(range)
    for cell in cell_list:
        student_sheet.update_cell(int(cell.row), int(cell.col), "")

def stringToDatetime(content):
    try:
        result = re.search("\d\d/\d\d/\d\d\d\d",content).group(0)
        result = datetime.strptime(result,"%m/%d/%Y")
        return result
    except AttributeError:
         print "Error---------------\n" + content



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

#formats dates in src sheet
for row in src_rows:
    date = datetime.strptime(row[2], "%m/%d/%Y")
    row[2] = date

# sorts src_sheet based on date of event

#puts info from src_sheet into variables
src_sheet_row_list = []
for row in src_rows:
    groupName = row[1]

    #Visit Date
    visitDate = "Date: " + row[2].strftime("%A, %m/%d/%Y")

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
    src_sheet_row_list.append([content,"","","",""])
    # student_sheet.update_cell(currentRow, 1, content)
    # student_sheet.update_cell(currentRow,5, comments)

    #print(content + "\n----------------------------------------")


student_rows = student_sheet.get_all_values()

mergeList = student_rows + src_sheet_row_list
mergeList.pop(0) #gets rid of header row

sortedMergedList = sorted(mergeList, key=lambda item: stringToDatetime(item[0]))



for row in sortedMergedList:
   print stringToDatetime(row[0]).strftime("%m/%d/%Y")


##################UNTESTED#########################################
# for row in sortedMergedList:
#     if(datetime.now() > stringToDatetime(row[0])):
#         #write to final_sheet
#         #remove from sortedMergedList
#         print "Its in the past"

# rangeToClear = "A2:B" + str(len(student_rows))
# clearCells(rangeToClear)

# currentRow = 2
# for row in sortedMergedList:
#     student_sheet.update_cell(currentRow,1,row[0]) #content
#     student_sheet.update_cell(currentRow,2,row[1]) #names
#     student_sheet.update_cell(currentRow,3,row[2]) #phone numbers
#     student_sheet.update_cell(currentRow,4,row[3]) #810s
#     student_sheet.update_cell(currentRow,5,row[4]) #special notes
#     currentRow += 1
# rangeToClear="A2:N" + str(len(src_rows))
# clearCells(rangeToClear)













