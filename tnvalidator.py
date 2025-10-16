#!/usr/bin/env python3

#Changelog
#Release 1 - October 16, 2025
# - Made windows resizable
# - Added larger font option
# - Licensed under GPL v3
# *** Known Issues ***
# - Does not load/parse .xls files (please request this if it would be helpful to you)

#Alpha 7 - October 7, 2025
# - Corrected valid activity checking for converted/newly created units. Warns if Unit is in Valid Units tab but
#       not Clan tab, i.e., did not exist at beginning of turn (may or may not be error). Error if in neither.
# - Added color coding to report:
#   - Category title corresponds to Err level. Black = Clean, Orange = Warn, Red = Error.
#   - If ErrorType is clean, title is black and "No errors" message is in blue for easy scanning.
#   - If ErrorType is hit, title is red for errors/orange for warnings. Specific errors black for easy reading.
# - Changed to single-window interface
# - Added some additional error handling

#Alpha 6 - September 30, 2025
# - Corrected Skill and Research checks to only consider Tribes valid units
# - Added checking for excess (>3) skill attempts by a tribe
# - Added checking for duplicate attempts of the same skill by a tribe
# - Added checking for duplicate skill priorities by a tribe

#Alpha 5 - September 29, 2025
# - Reorganized report window
# - Fixed error when no report is selected
# - Fixed column parsing stopping on blank lines

import tkinter as tk
from tkinter import filedialog as fd
#from tkinter import simpledialog as sd
from tkinter.messagebox import showerror
from tkinter import ttk
import openpyxl
#import xlrd
import pathlib

## Data parsing functions. These are ripe for a re-write as the initial tests performed really just compared columns to
## columns. Multi-column series parsers added but this really needs to be refactored as a generaly purpose function.
#Parse a 3-column series. Only used for a couple of things (activity item/distinction validation, some skill checks)
def parseTriple(activeSheet, startcol):
    sheetData=[]
    for i in range(2, activeSheet.max_row + 1):
        if activeSheet.cell(i,startcol).value is None:
            pass
        else:
            triple=[str(activeSheet.cell(i,startcol,).value).upper(), str(activeSheet.cell(i,startcol+1,).value).upper(), str(activeSheet.cell(i,startcol+2,).value).upper()]
            sheetData.append(triple)
    return sheetData

#Parse a single column, omitting header and blank lines
def parseCol(activeSheet, col):
    sheetData=[]
    for i in range(2, activeSheet.max_row + 1):
        if activeSheet.cell(i,col).value is None:
            pass
        else:
            sheetData.append(activeSheet.cell(i,col).value)
    return sheetData

## Core data analysis function. Compares two lists/columns, case insensitive
#compare two lists, case insensitive. This is the primary check for valid units
def checkValidList(testList, validList):
    listErrors = []
    for i in range(len(testList)):
        match = None
        for vi in validList:
            if testList[i].upper() == vi.upper():
                match = testList[i]
               
        if match is None:
            #Add 2 to correct human-readable row number for zero indexing and omitting header row 
            errorData = (i+2,testList[i])
            listErrors.append(errorData)

    return listErrors

## File access function. Written to only pull data that is being used for functions, as I was concerned about performance
## pulling all sheet data in. With fix to sheet size reducing from 2.5mb to 200k, should probably be refactored to pull in
## entire workbook. XLS parsing function will be added here when I get around to it.
#read xlsx file and return reportData with all sheets we want read into memory, close file
def processOrdersXLSX(path, reportData):
    
    try:
        orders = openpyxl.load_workbook(path)
    except:
        showerror("File Input Error", "Could not open file, may be open or in use.")
        return None

    #pull all sheets we are interested in
    try:
        clanSheet = orders["Clan"]
        actSheet = orders["Tribes_Activities"]
        movSheet = orders["Tribe_Movement"]
        scoSheet = orders["Scout_Movement"]
        sklSheet = orders["Skill_Attempts"]
        resSheet = orders["Research_Attempts"]
        xfrSheet = orders["Transfers"]
        vgSheet = orders["Valid Goods"]
        vuSheet = orders["Valid Units"]
        vaSheet = orders["Valid Activity"]
    except:
        showerror("File Input Error", "Selected file is not a valid TN3 Order sheet.")
        return None

    #load relevant columns/triples into data structure
    reportData["clanUnits"] = parseCol(clanSheet,1)
    reportData["actUnits"] = parseCol(actSheet,1)
    reportData["movUnits"] = parseCol(movSheet,1)
    reportData["sctUnits"] = parseCol(scoSheet,1)
    reportData["sklUnits"] = parseCol(sklSheet,1)
    reportData["resUnits"] = parseCol(resSheet,1)
    reportData["xfrGoods"] = parseCol(xfrSheet,3)
    reportData["xfrFromUnits"] = parseCol(xfrSheet,1)
    reportData["xfrToUnits"] = parseCol(xfrSheet,2)
    reportData["validGoods"] = parseCol(vgSheet,1)
    reportData["validUnits"] = parseCol(vuSheet,1)
    reportData["validActs"] = parseTriple(vaSheet,1)
    reportData["clanActs"] = parseTriple(actSheet,2)
    reportData["skillAttempts"] = parseTriple(sklSheet,1)
    
    orders.close()
    
    return reportData


#main loop called when an order sheet is selected
def select_file():

    filetypes = [('TN3 XLSX Turnsheet', '*.xlsx')]
    
    #For when XLS support is added
    #filetypes = (
    #   ('TN3 XLSX Turnsheet', '*.xlsx'),
    #    ('TN3 XLS Turnsheet', '*.xls')
    #)

    ordersPath = fd.askopenfilename(
        title='Select TN3 Order File',
        filetypes=filetypes)
    
    #If no file is selected, bug out
    if ordersPath == "":
        return None
    
    path = pathlib.Path(ordersPath)

    #destroy any existing frame if this is a second run
    for widget in root.winfo_children():
        if isinstance(widget, tk.Frame):
            widget.destroy()
            root.geometry('600x200')
            root.update()
    
    #create reportData as empty dictionary so we can use named indexes (feels hacky but its a mystery)
    reportData = {
        "clanUnits":[],
        "actUnits":[],
        "clanActs":[],
        "movUnits":[],
        "sctUnits":[],
        "sklUnits":[],
        "resUnits":[],
        "xfrFromUnits":[],
        "xfrToUnits":[],
        "xfrGoods":[],
        "validGoods":[],
        "validUnits":[],
        "validActs":[],
        "skillAttempts":[]
    }
  
    #process file and load data into reportData
    if path.suffix == ".xlsx":
        reportData = processOrdersXLSX(path,reportData)

    #processOrdersXLSX will return None if the file cannot be opened or was invalid. Bail out.
    if reportData == None:
        return None

    #build results window

    results = tk.Frame(root)
    results.pack(fill = "both", expand = True)

    style = ttk.Style()
    style.configure("Treeview", font=("Arial", fontsize), rowheight=int(fontsize*2.2))
    
    treeview = ttk.Treeview(results, show="tree")
    treeview.column("#0", width=550, stretch=True)
    treescroll=ttk.Scrollbar(results, orient="vertical", command=treeview.yview)
    treeview.configure(yscrollcommand=treescroll.set)
    treeview.tag_configure("error", foreground="red")
    treeview.tag_configure("warning", foreground="orange")
    treeview.tag_configure("pass", foreground="blue")
    treeview.tag_configure("old", font=("Arial", 12))
    root.geometry('600x750')
    titleMessage= "Validating File: " + path.name
    tk.Label(results, text=titleMessage, font=("Arial", 12, "bold")).pack()

    #cull GM units from valid Units to get a Clan Unit List based on Valid Units

    clanNumber = reportData["clanUnits"][0][1:4]
    validClanUnits = []
    validGMUnits = []
    for i in range(len(reportData["validUnits"])):
        if reportData["validUnits"][i][1:4] == str(clanNumber):
            validClanUnits.append(reportData["validUnits"][i])
        else:
            validGMUnits.append(reportData["validUnits"][i])
        
    tk.Label(results, text="Orders for Clan " + clanNumber, font=("Arial", 12, "bold")).pack()

    treeview.pack(side="left", fill="both", expand=True)
    treescroll.pack(side="left",fill="y")

    #cull subunits to have only valid Clan Tribes
    validClanTribes = []
    for i in range(len(validClanUnits)):
        if len(validClanUnits[i]) == 4:
            validClanTribes.append(validClanUnits[i])
    
    #Display Valid Units
    vuRoot = treeview.insert("",0,text="Valid Units")
    cuRoot = treeview.insert(vuRoot,tk.END,text="Valid Clan Units")
    for i in range(len(validClanUnits)):
        treeview.insert(cuRoot,tk.END,text=str(validClanUnits[i]))
    vtRoot = treeview.insert(vuRoot,tk.END,text="Valid Clan Tribes" )
    for i in range(len(validClanTribes)):
        treeview.insert(vtRoot,tk.END,text=str(validClanTribes[i]))
    guRoot = treeview.insert(vuRoot,tk.END,text="Valid GM Units")
    for i in range(len(validGMUnits)):
        treeview.insert(guRoot,tk.END,text=str(validGMUnits[i]))

    #create organizational root
    actRoot = treeview.insert("",tk.END,text="Activity Orders")
    xfrRoot = treeview.insert("",tk.END,text="Transfer Orders")
    movRoot = treeview.insert("",tk.END,text="Movement and Scouting Orders")
    lrnRoot = treeview.insert("",tk.END,text="Skill and Research Orders")
    #Check for Invalid Units Assigned Movement Orders 
    errRoot = treeview.insert(movRoot,tk.END,text="Movement Unit Errors", open=True)
    vErrors = checkValidList(reportData["movUnits"], validClanUnits)
    
    if len(vErrors) == 0:
        treeview.insert(errRoot,tk.END,text="No Invalid Units Assigned Movement Orders", tags="pass")
    else:
        treeview.item(movRoot, tags="error")
        treeview.item(errRoot, tags="error")
        for i in range(len(vErrors)):
            errorText = "Invalid Unit " + str(vErrors[i][1]) + " assigned movement order on Row " + str(vErrors[i][0])
            treeview.insert(errRoot,tk.END,text=errorText)

    #Check for Invalid Units Assigned Scouting Orders
    errRoot = treeview.insert(movRoot,tk.END,text="Scouting Unit Errors", open=True)
    vErrors = checkValidList(reportData["sctUnits"], validClanUnits)
    if len(vErrors) == 0:
        treeview.insert(errRoot, tk.END, text="No Invalid Units Assigned Scouting Orders", tags="pass")
    else:
        treeview.item(movRoot, tags="error")
        treeview.item(errRoot, tags="error")
        for i in range(len(vErrors)):
            errorText = "Invalid Unit " + str(vErrors[i][1]) + " assigned scouting order on Row " + str(vErrors[i][0])
            treeview.insert(errRoot, tk.END, text=errorText)

    #Check for Invalid Units Assigned Skill Attempts
    errRoot = treeview.insert(lrnRoot,tk.END,text="Skill Attempt Unit Errors", open=True)
    vErrors = checkValidList(reportData["sklUnits"], validClanTribes)
    
    if len(vErrors) == 0:
        treeview.insert(errRoot, tk.END, text="No Invalid Units Assigned Skill Attempts", tags="pass")
    else:
        treeview.item(lrnRoot, tags="error")
        treeview.item(errRoot, tags="error")
        for i in range(len(vErrors)):
            errorText = "Invalid Unit " + str(vErrors[i][1]) + " assigned skill attempt on Row " + str(vErrors[i][0])
            treeview.insert(errRoot, tk.END, text=errorText)

    #check for Tribes assigned more than three skill attempts
    errRoot = treeview.insert(lrnRoot,tk.END,text="Tribes Assigned Excess Skill Attempts Errors", open=True)
    skillAttempts = {}
    skillErrors = {}
    for i in reportData["sklUnits"]:
        if i in skillAttempts:
            skillAttempts[i] += 1
        else:
            skillAttempts[i] = 1

    for key, value in skillAttempts.items():
        if value > 3:
            skillErrors[key] = value

    if len(skillErrors) == 0:
        treeview.insert(errRoot, tk.END, text="No Tribe Assigned More Than Three Skill Attempts", tags="pass")
    else:
        treeview.item(lrnRoot, tags="error")
        treeview.item(errRoot, tags="error")
        for key, value in skillErrors.items():
            errorText = "Tribe " + str(key) + " assigned " + str(value) + " skill attempts"
            treeview.insert(errRoot, tk.END, text=errorText)
    
    #check for duplicate Tribe/Skill attempts
    errRoot = treeview.insert(lrnRoot,tk.END,text="Duplicate Tribe/Skill Attempts Errors", open=True)
    skillAttempts = []
    skillErrors = []
    for i in range(len(reportData["skillAttempts"])):
        checkAttempt = [reportData["skillAttempts"][i][0], reportData["skillAttempts"][i][2]]
        if checkAttempt in skillAttempts:
            skillErrors.append(checkAttempt)
        else:
            skillAttempts.append(checkAttempt)
    if len(skillErrors) == 0:
        treeview.insert(errRoot, tk.END, text="No Tribe Attempting Duplicate Skills", tags="pass")
    else:
        treeview.item(lrnRoot, tags="error")
        treeview.item(errRoot, tags="error")
        for i in range(len(skillErrors)):
            errorText = "Tribe " + str(skillErrors[i][0]) + " duplicate attempts for skill " + str(skillErrors[i][1]) 
            treeview.insert(errRoot, tk.END, text=errorText)

    errRoot = treeview.insert(lrnRoot,tk.END,text="Duplicate Skill Attempt Priority Errors", open=True)
    skillAttempts = []
    skillErrors = []
    for i in range(len(reportData["skillAttempts"])):
        checkAttempt = [reportData["skillAttempts"][i][0], reportData["skillAttempts"][i][1]]
        if checkAttempt in skillAttempts:
            skillErrors.append(checkAttempt)
        else:
            skillAttempts.append(checkAttempt)
    if len(skillErrors) == 0:
        treeview.insert(errRoot, tk.END, text="No Tribe Attempting Skills At Same Priority", tags="pass")
    else:
        treeview.item(lrnRoot, tags="error")
        treeview.item(errRoot, tags="error")
        for i in range(len(skillErrors)):
            errorText = "Tribe " + str(skillErrors[i][0]) + " attempting multiple skills at priority " + str(skillErrors[i][1]) 
            treeview.insert(errRoot, tk.END, text=errorText)
    
    
    #Check for Invalid Units Assigned Research Attempts
    errRoot = treeview.insert(lrnRoot,tk.END,text="Research Attempt Unit Errors", open=True)
    vErrors = checkValidList(reportData["resUnits"], validClanTribes)

    if len(vErrors) == 0:
        treeview.insert(errRoot, tk.END, text="No Invalid Units Assigned Research Attempts", tags="pass")
    else:
        treeview.item(lrnRoot, tags="error")
        treeview.item(errRoot, tags="error")
        for i in range(len(vErrors)):
            errorText = "Invalid Unit " + str(vErrors[i][1]) + " assigned research attempt on Row " + str(vErrors[i][0])
            treeview.insert(errRoot, tk.END, text=errorText)

    #Check for Invalid Units Assigned Activities
    actUnitRoot = treeview.insert(actRoot,tk.END,text="Activity Orders Unit Issue", open=True)
    #Check clan tab first because new units ordinariliy should not perform activities. If unit is on valid list, give warning (converted unit, scouting orders). If not on valid list, give error.
    errRoot = treeview.insert(actUnitRoot,tk.END,text="Invalid Unit Assigned Activity [Error]", open=True)
    warnRoot = treeview.insert(actUnitRoot,tk.END,text="New Unit Assigned Activity [Warning/Informational]", open=True)
    vErrors = checkValidList(reportData["actUnits"], reportData["clanUnits"])
    errCount = 0
    warnCount = 0

    if len(vErrors) != 0:
        for i in range(len(vErrors)):
            if vErrors[i][1] in validClanUnits:
                warnCount += 1
                errorText = "New Unit " + str(vErrors[i][1]) + " assigned activity order on Row " + str(vErrors[i][0])
                treeview.insert(warnRoot, tk.END, text=errorText)
            else:
                errCount += 1
                errorText = "Invalid Unit " + str(vErrors[i][1]) + " assigned activity order on Row " + str(vErrors[i][0])
                treeview.insert(errRoot, tk.END, text=errorText)
    if errCount == 0:
        treeview.insert(errRoot, tk.END, text="No Invalid Units Assigned Activity Orders", tags="pass")
    else:
        treeview.item(actRoot, tags="error")
        treeview.item(errRoot, tags="error")
    if warnCount == 0:
        treeview.insert(warnRoot, tk.END, text="No New Units Assigned Activity Orders", tags="pass")
    else:
        if treeview.item(actRoot, option="tags") != ("error",):
            treeview.item(actRoot, tags="warning")
        treeview.item(warnRoot, tags="warning")
   
    #Check for invalid Activities
    errRoot = treeview.insert(actRoot,tk.END, text="Activity Orders Item/Distinction Errors", open=True)
    vErrors = []

    for i in range(len(reportData["clanActs"])):
        if reportData["clanActs"][i] not in reportData["validActs"]:
            vErrors.append(reportData["clanActs"][i])

    if len(vErrors) == 0:
        treeview.insert(errRoot, tk.END, text="No Activity Item/Distinction Errors Found", tags="pass")
    else:
        treeview.item(actRoot, tags="error")
        treeview.item(errRoot, tags="error")
        for i in range(len(vErrors)):
            errorText = "Invalid Item/Distinction for Activity " + vErrors[i][0] + ": " + vErrors[i][1] + " / " + vErrors[i][2]
            treeview.insert(errRoot, tk.END, text=errorText)

    #check for Activity Discontinuity
    errRoot = treeview.insert(actRoot,tk.END,text="Activity Order Discontinuity Errors", open=True)

    actAssignedUnits = []
    actErrors = []
    prevActUnit = None

    for i in range(len(reportData["actUnits"])):
        if reportData["actUnits"][i] not in actAssignedUnits:
            prevActUnit=prevActUnit=reportData["actUnits"][i]
            actAssignedUnits.append(reportData["actUnits"][i])
        
        else:
            if reportData["actUnits"][i] == prevActUnit:
                actAssignedUnits.append(reportData["actUnits"][i])
            else:
                prevActUnit=prevActUnit=reportData["actUnits"][i]
                actAssignedUnits.append(reportData["actUnits"][i])
                errorData = (i+2,reportData["actUnits"][i])
                actErrors.append(errorData)
    
    if len(actErrors) == 0:
        treeview.insert(errRoot, tk.END, text="No Activity Order Discontinuity Detected", tags="pass")
    else:
        treeview.item(actRoot, tags="error")
        treeview.item(errRoot, tags="error")
        for i in range(len(actErrors)):
            errorText = "Unit " + str(actErrors[i][1]) + " assigned non-contiguous activity order on Row " + str(actErrors[i][0])
            treeview.insert(errRoot, tk.END, text=errorText)

    #check for at least one Clan unit in each transfer
    errRoot = treeview.insert(xfrRoot,tk.END,text="Invalid Transfer Unit Errors", open=True)
    vErrors = []
    for i in range (len(reportData["xfrFromUnits"])):
        match = None
        for vi in validClanUnits:
            if reportData["xfrFromUnits"][i].upper() == vi.upper():
                match = reportData["xfrFromUnits"][i]
            if reportData["xfrToUnits"][i].upper() == vi.upper():
                match = reportData["xfrToUnits"][i]
        
        if match is None:
            #add 2 for omitted title row and zero index conversion
            vErrors.append(i+2)
             
    if len(vErrors) == 0:
        treeview.insert(errRoot, tk.END, text="No Transfer Orders Without Clan Unit", tags="pass")
    else:
        treeview.item(xfrRoot, tags="error")
        treeview.item(errRoot, tags="error")
        for i in range(len(vErrors)):
            errorText = "Transfer order on Row " + str(vErrors[i]) + " has no valid Clan unit"
            treeview.insert(errRoot, tk.END, text=errorText)

    #check for transfers from non-Clan Units (not a valid transfer order, error)
    errRoot = treeview.insert(xfrRoot,tk.END,text="Transfers From Non-Clan/GM Units [Error]", open=True)
    vErrors = checkValidList(reportData["xfrFromUnits"], reportData["validUnits"])
    if len(vErrors) == 0:
        treeview.insert(errRoot, tk.END, text="No Transfers From Non-Clan/GM Units", tags="pass")
    else:
        treeview.item(xfrRoot, tags="error")
        treeview.item(errRoot, tags="error")
        for i in range(len(vErrors)):
            tRow = vErrors[i][0]        
            errorText = "Transfer From Non-Clan/GM Unit " + str(vErrors[i][1]) + " to Unit " + reportData["xfrToUnits"][tRow-2] + " on Row " + str(tRow)
            treeview.insert(errRoot, tk.END, text=errorText)

    #check for transfers to non-Clan Units (valid but worth reviewing for mistakes, warning)
    errRoot = treeview.insert(xfrRoot,tk.END,text="Transfers to Non-Clan/GM Units [Warning/Informational]", open=True)
    vErrors = checkValidList(reportData["xfrToUnits"], reportData["validUnits"])
    if len(vErrors) == 0:
        treeview.insert(errRoot, tk.END, text="No Transfers To Non-Clan/GM Units", tags="pass")
    else:
        if treeview.item(xfrRoot, option="tags") != ("error",):
            treeview.item(xfrRoot, tags="warning")
        treeview.item(errRoot, tags="warning")
        for i in range(len(vErrors)):
            tRow = vErrors[i][0]        
            errorText = "Transfer To Non-Clan/GM Unit " + str(vErrors[i][1]) + " from Unit " + reportData["xfrFromUnits"][tRow-2] + " on Row " + str(tRow)
            treeview.insert(errRoot, tk.END, text=errorText)

    #check for invalid goods in transfers
    errRoot = treeview.insert(xfrRoot,tk.END,text="Invalid Transfer Goods Errors", open=True)
    vErrors = checkValidList(reportData["xfrGoods"], reportData["validGoods"])

    if len(vErrors) == 0:
        treeview.insert(errRoot, tk.END, text="No Invalid Goods in Transfer Orders", tags="pass")
    else:
        treeview.item(xfrRoot, tags="error")
        treeview.item(errRoot, tags="error")
        for i in range(len(vErrors)):
            errorText = "Invalid Good " + str(vErrors[i][1]) + " on Row " + str(vErrors[i][0])
            treeview.insert(errRoot, tk.END, text=errorText)

fontsize = 10

#fontchange
def font_resize():
    global fontsize
    fontsize = checkVar.get()

#create main window

root = tk.Tk()
root.title('TNValidator')
root.resizable(True, True)
root.geometry('600x200')

tk.Label(text="TN Order Validator by Clan 293", font=("Arial", 12, "bold")).pack()

#Font doodling
checkVar = tk.IntVar()
fontcheck = tk.Checkbutton(root, text="Check for Larger Font", variable = checkVar, onvalue = 12, offvalue = 10, command=font_resize)

tk.Label(text="Please select your TN Auto Order Sheet to Validate").pack()

# open button
open_button = tk.Button(
    root,
    text='Select File',
    command=select_file
)

open_button.pack()

tk.Label(text="Add New Units to Column A of Valid Units Sheet to Reduce False Positives").pack()
fontcheck.pack()
fontcheck.deselect()

root.mainloop()