TNValidator - An auto order spreadsheet validation tool for TribeNet TN3 by Clan 293

TNValidator analyzes your XLSX auto order spreadsheet to identify common errors prior to submission. It is available as a Python script, or as a packaged Windows .exe (created using "pyinstaller --onefile --noconsole" on Windows 11). The latest version of TNValidator can be downloaded at https://github.com/nmarigoni/tnvalidator

What does TNValidator actually do?

TNValidator will read a selected auto orders spreadsheet and generate a report identifying common data entry errors in various aspect of turn sheet. TNValidator is read only. It does not make any network connections and does not write any files to disk. The TN Orders file cannot be open in excel (or any other application) to validate.

![screenshot showing example report from application, displaying an expandable tree structure with categories of errors, color-coded by error type](https://github.com/nmarigoni/tnvalidator/blob/main/example-report.png?raw=true)

Errors are highlighted in red: these are entries that, to my knowledge, are never correct and will result in errors, failure of orders to process, and possibly emails from the GM.

Warnings are highlighted in orange: these are entries that may be correct but may also indicate a mistake. Warnings are minimal and are explained below.

Installation:
If using the packaged Windows executable, no installation is needed, just download the .exe and run.

If using the Python script, the only dependency you should need to install is openpyxl. All other libraries should be included in a typical Python install (at least on windows). TNValidator has been tested with Python 3.13 on Windows 11.

Usage Tips:
- Units created or converted through GM actions on the turn being evaluated can be added to the Valid Units worksheet to reduce false positive on errors. Adding units to Valid Units does not, to the best of my knowledge, impact the processing of orders in any way, and is recommended in other player aids to permit data validation when entering orders.

- There is a checkbox to have the report generated with a larger font (size 12 instead of size 10). Check the box before selecting the file to validate. The window can also be resized.

Contact via Discord with any questions.

# Functionality #

TNValidator uses the following data sets to perform checks:
- Valid Units = All units listed in the Valid Units worksheet, includes GM units used for special transfers. Note that newly created units can be added to the Valid Units 
- Valid Clan Units - All Valid Units that belong to your Clan
- GM Units - All Valid Units that are GM (clan 263) units
- Turn Start Units - All units listed on the Clan worksheets, i.e., units that existed before the beginning of the turn (used principally for determining valid units to perform activities).

TNValidator performs the following tests:

### Activity Orders ###

- **Invalid Unit Assigned Activity** - 
Checks the TRIBE column in the Tribes_Activities worksheet to determine if all activities are assigned to a Valid Unit. A failure is an error.

- **New Unit Assigned Activity** - 
Checks the TRIBE column in the Tribes_Activities worksheet to determine if a unit is assigned an activity that is a Valid Unit but not a Turn Start Unit. This typically identifies newly created units the user has added to the Valid Units worksheet. A failure is a warning because this is valid for some activities types, e.g., Scouting and other Warrior activities, but should be reviewed.

- **Activity Orders Item/Distinction Errors** - 
Checks the ACTIVITY/ITEM/DISTINCTION columns of the Tribes_Activities worksheet to determine if all assigned activities match a combination on the Valid Activity worksheet. A failure is an error.

- **Activity Order Discontinuity Errors** - 
Checks the TRIBE column in the Tribe_Activities worksheet to identify any circumstances where there is a discontinuous assignment of activities to the same unit, i.e., that multiple "groups" of activities are assigned to the same unit, with activities assigned to other units in between. This will cause processing errors because later groupings will overwrite existing groupings when processed. A failure is an error.

### Transfer Orders ###
- **Invalid Transfer Unit Errors** - 
Checks the From and To columns of the Transfers worksheet to ensure a Valid Clan Unit is on at least one side of each transfer, i.e., that each transfer actually involves your Clain. A failure is an error.

- **Transfers From Non-Clan/GM Units** - 
Checks the From column of the Transfers worksheet for transfers ordered from a unit that is not a Valid Clan Unit or GM Unit, as a player cannot order a transfer from another Clan. A failure is an error.

- **Transfers To Non-Clan/GM Units** - 
Checks the To column of the Transfers worksheet for transfers ordered to a unit that is not a Valid Clan Unit or GM Unit. While external transfers to another Clan are valid, these should be checked to ensure an apparent external transfer is not a mistyped intra-Clan transfer. A failure is a warning.

- **Invalid Transfer Goods** - 
Checks the Item column of the Transfers worksheet to determine if all transfer orders are for an item/good type that appears in the Goods column of the Valid Goods worksheet. To catch errors such as unnecessary/required pluralization. A failure is an error.

### Movement and Scouting ###
- **Movement Unit Errors** - 
Checks the TRIBE column of the Tribe_Movement worksheet to determine if all units assigned scouting orders are Valid Units. A failure is an error.

- **Scouting Unit Errors** - 
Checks the TRIBE column of the Scout_Movement worksheet to determine if all units assigned scouting orders are Valid Units. A failure is an error.

### Skill and Research Orders ###
- **Skill Attempt Unit Errors** - 
Checks the TRIBE column of the Skill_Attempts worksheet to determine if all units assigned skill attempt orders are both Valid Units and are Tribe-type Units rather than Subunits. A failure is an error.

- **Tribes Assigned Excess Skill Attempts Errors** - 
Checks the TRIBE column of the Skill_Attempts worksheet to determine if any unit is assigned more than three skill attempts. A failure is an error.

- **Duplicate Tribe/Skill Attempts Errors** - 
Checks the TRIBE and TOPIC columns of the Skill_Attempts worksheet to determine if any unit has been assigned the same skill attempt topic more than once. A failure is an error.

- **Duplicate Skill Attempt Priority Errors** - 
Checks the TRIBE and ORDER columns of the Skill_Attempts worksheet to determine if any unit has been assigned more than one skill attempt with the same order number/priority. A failure is an error.

- **Research Attempt Unit Errors** - 
Checks the TRIBE column of the Research_Attempts worksheet to determine if all units assigned research attempt orders are both Valid Units and are Tribe-type Units rather than Subunits. A failure is an error.
