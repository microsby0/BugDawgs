import os
import gspread
import time
from datetime import datetime
import re

def setUpHeaders():
    student_sheet.update_acell('A1',"Events")
    student_sheet.update_acell('B1',"Name\n****Denotes pick up/drop off****")
    student_sheet.update_acell('C1',"Phone Number")
    student_sheet.update_acell('D1',"810")
    student_sheet.update_acell('E1',"Special Notes")
    student_sheet.update_acell('F1',"Call if you need me 706-614-3328 (cell) 706-542-1238 (office)\nYou can also send a text, which I can access more easily in meetings and at home. Thanks and Good Luck! \nMarianne")
    final_sheet.update_acell('A1',"Events")
    final_sheet.update_acell('B1',"Name\n****Denotes pick up/drop off****")
    final_sheet.update_acell('C1',"Phone Number")
    final_sheet.update_acell('D1',"810")
    final_sheet.update_acell('E1',"Special Notes")
    final_sheet.update_acell('F1',"Call if you need me 706-614-3328 (cell) 706-542-1238 (office)\nYou can also send a text, which I can access more easily in meetings and at home. Thanks and Good Luck! \nMarianne")

def clearCells(range,sheet):
    #A2:E number of rows
    cell_list = student_sheet.range(range)
    for cell in cell_list:
        sheet.update_cell(int(cell.row), int(cell.col), "")

def stringToDatetime(content):
    try:
        result = re.search("\d\d/\d\d/\d\d\d\d",content).group(0)
        result = datetime.strptime(result,"%m/%d/%Y")
        return result
    except AttributeError:
         print "Error---------------\n" + content

def populate():
    real_sheet = account.open("Insect Zoo Request Form").sheet1
    real_rows = real_sheet.get_all_values()
    curRow=1
    curCol=1
    for row in real_rows:
        for cell in row:
            src_sheet.update_cell(curRow,curCol,cell)
            curCol+=1
        curRow+=1
        curCol=1
        if curRow> 50:
            break

def formatDates(rows):
    for row in rows:
        try:
            date = datetime.strptime(row[2], "%m/%d/%Y")
            row[2] = date
        except ValueError:
            #switch this to alert someone later
            row[2] = datetime.now()

def createSrcList(rows):
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
    tempList=[]

    for row in src_rows:
        groupName = row[1]

        #Visit Date
        visitDate = "Date: " + row[2].strftime("%A, %m/%d/%Y")

        #Visit Time
        #formats time to AM/PM 12 hour
        try:
            time_values = time.strptime(row[3], "%H:%M:%S")
            visitTime = time.strftime("%I:%M %p", time_values)
        except ValueError:
            #switch this to alert someone later
            time_values = time.strptime("12:30:34", "%H:%M:%S")
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
        tempList.append([content,"","","",comments])
    return tempList

def removeOutdatedEvents(masterList):
    startRow = len(final_sheet.get_all_values()) + 1
    startCol = 1
    counter=0
    for row in masterList:
        if(datetime.now() > stringToDatetime(row[0])):
            for cell in row:
                final_sheet.update_cell(startRow,startCol,cell)
                startCol+=1
            startRow+=1
            startCol=1
            sortedMergedList.pop(counter)
        counter+=1


account = gspread.login(os.environ["GSPREAD_USERNAME"], os.environ["GSPREAD_PASSWORD"])

src_sheet  = account.open("Test").sheet1
student_sheet = account.open("Test Dest").sheet1
final_sheet = account.open("Final Dest").sheet1

#populate()

setUpHeaders()

#get data from src sheet
src_rows = src_sheet.get_all_values()
src_rows.pop(0)

#formats dates in src sheet
formatDates(src_rows)

#creates list of rows in src sheet
src_sheet_row_list = createSrcList(src_rows)

#gets list of rows from student sheet
student_rows = student_sheet.get_all_values()

#combines lists
mergeList = student_rows + src_sheet_row_list
mergeList.pop(0) #gets rid of header row

#sorts the master list
sortedMergedList = sorted(mergeList, key=lambda item: stringToDatetime(item[0]))

#prints out size and dates in the master list before its sorted
#print "Before: " + str(len(sortedMergedList))
# for row in sortedMergedList:
#    print stringToDatetime(row[0]).strftime("%m/%d/%Y")

removeOutdatedEvents(sortedMergedList)

#prints size and dates in master list after its sorted
#print "After: " + str(len(sortedMergedList))
# for row in sortedMergedList:
#    print stringToDatetime(row[0]).strftime("%m/%d/%Y")

#clears student sheet before it is updated
rangeToClear = "A2:E" + str(len(student_rows))
clearCells(rangeToClear,student_sheet)

currentRow = 2
for row in sortedMergedList:
    student_sheet.update_cell(currentRow,1,row[0]) #content
    student_sheet.update_cell(currentRow,2,row[1]) #names
    student_sheet.update_cell(currentRow,3,row[2]) #phone numbers
    student_sheet.update_cell(currentRow,4,row[3]) #810s
    student_sheet.update_cell(currentRow,5,row[4]) #special notes
    currentRow += 1

#clears all of src sheet
rangeToClear="A2:N" + str(len(src_rows))
clearCells(rangeToClear,src_sheet)













