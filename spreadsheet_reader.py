# import modules

import xlrd, xlwt
#from xlrd import open_workbook
from xlutils.copy import copy

# define font types

bold = xlwt.easyxf('font: bold on')
underline = xlwt.easyxf('font: underline on')
title = xlwt.easyxf('font: bold on, underline on')

#functions

def titlesheet(ws): #titles the top of the spreadsheet with column headers
    ws.write(0,0,"Name",title)
    ws.write(0,1,"Estimate",title)
    ws.write(0,2,"Real",title)
    ws.write(0,3,"Suggested",title)
    return

def checkiftitled(ws): #checks a spreadsheet if it has column titles
    if ws.cell(0, 0).value == xlrd.empty_cell.value:
        return False
    else:
        return True

def listactivties(wb,code):
    if code == "quick":
        print("Please select an activity group to add to.")
    if code == "manage":
        print("Please select an activity group to manage.")

def menutext():
    print("Please select an option by entering the number of the desired action.")
    print("1: Quick add")
    print("2: Manage activities")
    print("3: Exit")
    return

def menucontrol():
    menutext()
    menuselect = input()
    if menuselect == "1":
        quickadd()
    else:
        if menuselect == "2":
            manageactivities()
        else:
            if menuselect == "3":
                exit()
            else:
                print("Sorry, the previous input was invalid. Please enter only a number.")
                menucontrol()
    return

# test if timesheet.xls exists

try:
    wb = copy(xlrd.open_workbook('timesheet.xls'))
except:
    wb = xlwt.Workbook()
    generalws = wb.add_sheet("General")
    titlesheet(generalws)
    wb.save('timesheet.xls')
    wb = copy(xlrd.open_workbook('timesheet.xls'))

menucontrol()



    

#w = copy(xlrd.open_workbook('example3.xls'))
#w.get_sheet(0).write(0,1,"foo")
#w.save('example3.xls')


#wb.save('example3.xls')

# worksheet = workbook.sheet_by_index(0)
#sheet2 = workbook.add_worksheet("hello")