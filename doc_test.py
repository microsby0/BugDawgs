import os
import gspread
import time
from datetime import datetime

def setUpHeaders(dest_sheet):
    dest_sheet.update_acell('A1',"Events")
    dest_sheet.update_acell('B1',"Name\n****Denotes pick up/drop off****")
    dest_sheet.update_acell('C1',"Phone Number")
    dest_sheet.update_acell('D1',"810")
    dest_sheet.update_acell('E1',"Special Notes")
    dest_sheet.update_acell('F1',"Call if you need me 706-614-3328 (cell) 706-542-1238 (office)\nYou can also send a text, which I can access more easily in meetings and at home. Thanks and Good Luck! \nMarianne")


def dataCount(rows):
    print "Number of Rows: " + str(len(rows)) #number of rows
    i=1
    for row in rows:
        print "Columns in Row: " + str(i) + ": " + str(len(row)) #number of columns
        i+=1

def clearCells(range):
    cell_list = dest_sheet.range(range)
    for cell in cell_list:
        dest_sheet.update_cell(int(cell.row), int(cell.col), "")

# FIXME: should probably make these into environment vars
account = gspread.login(os.environ["GSPREAD_USERNAME"], os.environ["GSPREAD_PASSWORD"])

#need a spreadsheet and its first worksheet
src_sheet  = account.open("Test").sheet1
dest_sheet = account.open("Test Dest").sheet1


# clearC    ells("A2:E20")

setUpHeaders(dest_sheet)

rows = src_sheet.get_all_values()
#dataCount(rows)

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

for row in rows:
    date = datetime.strptime(row[2], "%m/%d/%Y")
    print date
    row[2] = {'formattedDate': date, 'rawDate': row[2]}

# sort based on date of event
rows = sorted(rows, key=lambda row: row[2]['formattedDate'])

print "\n"

for row in rows:
    print row[2]['formattedDate']

for row in rows:
    groupName = row[1]

    #Visit Date
    d = datetime.strptime(row[2]['rawDate'], "%m/%d/%Y")
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

    dest_sheet.update_cell(currentRow, 1, content)
    dest_sheet.update_cell(currentRow,5, comments)
    currentRow += 1
    #print(content + "\n----------------------------------------")
