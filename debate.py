"""
Program to help in Aarohan Debate
This program has been made to make the Aarohan Debate go smoother.

The program checks whether a participant has registered for the event,
checks their status, logs the time the participant leaves for the venue 
and what room they go to. It also allows for changes to be 
done in the case of any mistakes in user entry. 

The program integrates with Excel to save data as it happens. It logs
the timings of arrivals and departures from the mini audi.

Idea by : Venkataram S
Author : Sam Alex Koshy
Starting date : 2022/08/22 

"""

# Importing libraries.
import sys
import os
from datetime import datetime
from time import sleep

import re
# import json
# import openpyxl
from openpyxl import load_workbook
from openpyxl.worksheet.worksheet import Worksheet
from openpyxl.workbook.workbook import Workbook

def main():
    """The main functioning of the code. Keeps the code running."""
    # Load workbook.
    workbook = load_workbook("debate.xlsx")
    worksheet = workbook.active
    cmds = "Commands are: check <no>, add <no>, send [no], "
    cmds += "name <name>, list, change, save, cmds"
    print(cmds)
    # Keep code running until cmd quit entered.
    while(True):
        cmd = input("\nEnter command : ")
        if cmd.lower() == "quit":
            _save_workbook(workbook, True)
        elif cmd.lower().startswith("check"):
            _check_reg_no(worksheet, cmd)
        elif cmd.lower().startswith("add"):
            _add_participant(workbook, worksheet, cmd)
        elif cmd.lower().startswith("send"):
            _send_participant(workbook, worksheet, cmd)
        elif cmd.lower().startswith("name"):
            _check_name(worksheet, cmd.lower())
        elif cmd.startswith("change"):
            _change_participant_details(workbook, worksheet)
        elif cmd.startswith("cmd"):
            print(cmds)
        elif cmd.startswith("save"):
            _save_workbook(workbook, False)

def _check_reg_no(worksheet:Worksheet, word:str):
    """Check the status of a participant on typing their register number.

    If argument given is not a valid number, the function does not execute.

    This function prints the details of the participant if they are registered,
    and their status for the debate. If they are not a participant, prompt
    is given to check their name in the list.

    Parameters
    ---------
    - worksheet `Worksheet`:
        The worksheet with which to work with.
    - word `str`:
        The string that contains the register number of the participant, 
        along with the command word.
    """
    try:
       reg_no = int(re.split("check ", word)[1])
    except:
        return
    colB = worksheet['B']

    # Iterate through the column to find the register number
    for cell in colB:
        if cell.value == reg_no:
            # If found, print details.
            row:int = cell.row
            msg = ""
            for value in worksheet[row]:
                msg += f"{value.value}, "
            print(msg)

            # Print status of participant.
            if worksheet.cell(row=row, column=4).value!=None:
                print("Already sent the participant")
            elif worksheet.cell(row=row, column=3).value!=None:
                print("Participant is waiting in mini audi.")
            else:
                print("Participant has not arrived yet.")
            return

    # If register number not found, prompt for name.
    print(reg_no, "not found. Check if their name is there with name <name>.")
    name = input("Enter name : ")
    name = "name " + name
    _check_name(worksheet, name)

def _check_name(worksheet:Worksheet, word:str):
    """Check if the name exists in the list. If it exists, display their 
    details including row number.

    This function prints the details of the participant if their name is found.

    Parameters
    ----------
    - worksheet `Worksheet`:
        The worksheet with which to work with.
    - word `str`:
        The string that contains the name of the participant, 
        along with the command word.
    """
    try:
        name = re.split("name ", word)[1]
    except:
        return
    colA = worksheet['A']
    found = False
    deleted = False

    # Iterate through the column to find the name.
    for cell in colA:
        if cell.value is None:
            worksheet.delete_rows(cell.row, 1)
            deleted = True
        elif (re.search(cell.value.lower(), name) 
                or re.search(name, cell.value.lower())):
            # Print the details of the user if found.
            row = cell.row
            msg = f"{row}. "
            for value in worksheet[row]:
                msg += f"{value.value}, "
            print(msg)
            found = True
    if deleted:
        print("Rows were deleted, try again.")
    elif not found:
        print(name, "not found.")

def _add_participant(workbook:Workbook, worksheet:Worksheet, word:str):
    """Logs the time when the participant enters. 

    This function prints the details of the user, including whether they
    have already been added or if they have been sent.

    Parameters
    ----------
    - workbook `Workbook`:
        The workbook which has to be saved.
    - worksheet `Worksheet`:
        The worksheet with which to work with.
    - word `str`:
        The string that contains the register number of the participant, 
        along with the command word. 
    """
    try:
        reg_no = int(re.split("add", word)[1])
    except:
        return
    found = 0

    # Iterate through the column to find the register number
    for cell in ['B']:
        if cell.value == reg_no:
            row = cell.row
            if worksheet.cell(row=row, column=4).value!=None:
                print("Already sent the participant")
            elif worksheet.cell(row=row, column=3).value!=None:
                print("Already logged in")
            else:
                worksheet.cell(row=row, column=3, value=datetime.now())
                for value in worksheet[cell.row]:
                    print(value.value)
            found = 1
    # If participant is not found, send message.
    if found == 0:
        print(reg_no, "not found.")
    _auto_save_workbook(workbook)

def _send_participant(workbook:Workbook, worksheet:Worksheet, word:str):
    """Log the time the participant is sent to their panel room, along with
    the room they are sent to.

    This function prints the details of the participant after logging their
    details.

    Parameters
    ----------
    - workbook `Workbook`:
        The workbook which has to be saved.
    - worksheet `Worksheet`:
        The worksheet with which to work with.
    - word `str`:
        The string that contains the register number of the participant, 
        along with the command word. 
    """
    try:
        values = re.findall("\d+", word)
        reg_no = int(values[0])
        panel_room = int(values[1])
    except:
        return

    # Iterate through the column to find the register number
    for cell in worksheet['B']:
        if cell.value == reg_no:
            worksheet.cell(row=cell.row, column=4, value=datetime.now())
            worksheet.cell(row=cell.row, column=5, value=panel_room)
            for value in worksheet[cell.row]:
                print(value.value)
            print("Sent", reg_no)
    _auto_save_workbook(workbook)

def _change_participant_details(workbook:Workbook, worksheet:Worksheet):
    """Change the details of any participant.
    
    The name, register number or class of a participant can be changed.

    Parameters
    ----------
    - workbook `Workbook`:
        The workbook which has to be saved.
    - worksheet `Worksheet`:
        The worksheet with which to work with.
    """
    # Accept location of the change.
    row = int(input("Enter row address : "))
    col = int(input("Enter column address : "))
    # Validate the location, then make the change
    if col not in [1, 2, 5]:
        print("Invalid column number. Aborting change")
        return
    elif col == 1:
        val = input("Enter name : ")
    elif col == 2:
        val = int(input("Enter reg no : "))
    elif col == 5:
        val = input("Enter class : ")
    worksheet.cell(row=row, column=col, value=val)
    msg = f"{row}. "
    for value in worksheet[row]:
        msg += f"{value.value}, "
    print(msg)
    _auto_save_workbook(workbook)

def _save_workbook(workbook:Workbook, close:bool):
    """Save the workbook on demand.

    This function closes all excel windows to ensure that the workbook
    object gets saved to a file.

    Parameters
    ----------
    - workbook `Workbook`:
        The workbook which has to be saved.
    - close `bool`:
        Whether the program needs to be closed after saving.
    """
    # Close all files to ensure a save.
    try:
        os.system("taskkill/im EXCEL.EXE ")
    except:
        print("Already closed")
    # Give some time to ensure all sheets are closed.
    if close:
        sleep(10)
    # Try to save file.
    try:
        workbook.save("debate.xlsx")
    except PermissionError:
        print("Failed to save")
    else:
        print("Excel sheet saved")
        os.system("start EXCEL.EXE debate.xlsx")
        # Close the program.
        if close:
            sys.exit()

def _auto_save_workbook(workbook:Workbook):
    """Save the workbook when needed.

    Parameters
    ----------
    - workbook `Workbook`:
        The workbook which has to be saved.
    """
    try:
        workbook.save("debate.xlsx")
    except PermissionError:
        print("Failed to save")


if __name__ == "__main__":
    main()