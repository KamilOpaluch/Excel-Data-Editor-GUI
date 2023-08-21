# Excel-Data-Editor-GUI

A simple GUI-based Excel editor using Python's tkinter and pandas. 
This editor provides a user-friendly interface to view and modify Excel files. 
It comes with features such as filtering by specific values, adding new rows, committing changes, and restarting to undo modifications.

![image](https://github.com/KamilOpaluch/Excel-Data-Editor-GUI/assets/142261174/b1a93d5c-f0ca-4b72-8540-7a06512b3c97)


## Features


Dropdown Filter: Select a specific value from the dropdown to filter the displayed rows in the table.

Edit Cells: Click on a cell to edit its content.

Commit Changes: Save your changes to the Excel file.

Add Line: Insert a new row to the table.

Restart Changes: Revert any changes made during the session and reload the Excel file.


## Dependencies


tkinter: For GUI development.

pandas: For handling Excel data.

ttkthemes: For theming the GUI.

openpyxl: As an engine to read/write Excel files.


## Usage


Place the Excel file (e.g., 'Excel_example.xlsx') in the same directory as the script.


## Classes and Methods


DateEntry: A custom tk.Entry widget to input dates in the format dd-mm-yyyy.

_on_key_press: Handles key press events to ensure correct date format.

_on_return_press: Callback when the return key is pressed.

ExcelGUI: The main application class.

__init__: Initializes the GUI layout and loads the Excel file.

close_window: Closes the main window.

filter_table: Filters the table based on the dropdown selection.

add_line: Adds a new row to the table and the dataframe.

commit_changes: Saves changes to the Excel file.

restart_changes: Reverts any changes made during the session.

on_cell_click: Handles cell click events to allow editing.

finish_editing: Updates the table and dataframe after editing a cell.


## Disclaimer
All the code provided in this project is for educational or informational purposes only, and is not intended to be a substitute for professional advice. 
The code is provided "AS IS" without warranty of any kind. 
Use of the code is solely at your own risk.


