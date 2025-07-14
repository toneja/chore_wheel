# Chore Wheel
Randomly assign chores to randomized groups of people.

**WINDOWS ONLY**

## Download
https://github.com/toneja/chore_wheel/releases/download/v1.2/HCRU.Chore.Assignment.Wheel.zip

## Installation
- Unzip the above file anywhere on your PC.
- Creates a new subfolder with the application (Rainbow wheel icon) and an “\_internal” folder containing the application’s dependencies. Do not delete the “\_internal” folder.

## Usage
- All employee and chore assignment information will be contained in one workbook.
- The workbook will have 2 sheets: Assignments, and Chore History. The Assignments sheet contains the list of employee names and their assigned chores. The Chore History sheet lists all past chores and must be present in order for the fair assignment of duties.
- Use the “Open File” button to open the chore assignment workbook.
- Click the “Assign Chores” button to spin the wheel and assign the next set of chores.
	- One week of chores will be assigned per button press. Monthly chores are assigned bi-monthly and will take the place of that employee’s weekly chore.
	- The file will automatically be saved once chores are assigned.

## Unavailable Employees
- To prevent an employee from being assigned a chore during a spin, add a column to the Assignments sheet titled “Out” and add “True” to that person’s row. This will skip their assignments for that spin of the wheel. “Out” and “True” **must** be capitalized.

## Editing Chore Data
- Individual cells may be edited in the application by double-clicking them. New rows may be added and rows may be deleted using the application buttons. These changes are not automatically saved, use the “Save” or “Save As” buttons to save these changes.
- The workbook may also still be edited in Excel if that is more practical.

## Possible Issues
- Cannot assign chores: make sure that the chore workbook is not still open in Excel.
- Unfair chore assignment: make sure that the chore history sheet is in the workbook.
- If there are issues with the assignments and you wish to redo them, open the workbook in Excel and delete the problem Weekly/Monthly assignment columns, then re-open it in the application and re-run the “Assign Chores” option.
