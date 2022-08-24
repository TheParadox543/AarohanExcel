"""
Program to help in Aarohan Debate
This program has been made to make the Aarohan Debate go smoother.

The program checks whether a participant has registered for the event,
checks their status, creates a list of participants to be sent in the order 
by which they have arrived to mini audi. It also allows for changes to be 
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
import json
import openpyxl
from openpyxl import load_workbook
from openpyxl.worksheet.worksheet import Worksheet
from openpyxl.workbook.workbook import Workbook

def main():
    """The main functioning of the code. Keeps the code running."""
    # Load workbook.
    wb = load_workbook("debate.xlsx")
    ws = wb.active
    cmds = "Commands are: check <no>, add <no>, send [no], "
    cmds += "name <name>, list, change, cmds"
    try:
        with open("list.json", 'r') as file:
            # Load list of members to be sent to panel room.
            participants = json.load(file)
    except:
        participants = []
    os.system("start EXCEL.EXE debate.xlsx")
    print(cmds)
    # Keep code running until cmd quit entered.
    while(True):
        cmd = input("\nEnter command : ")
        if cmd.lower() == "quit":
            _save_wb(wb, participants, True)
        elif cmd.lower().startswith("check"):
            _check_reg(ws, cmd)
        elif cmd.lower().startswith("add"):
            participants = _add(ws, participants, cmd)
        elif cmd.lower().startswith("send"):
            participants = _send(ws, participants, cmd)
        elif cmd.startswith("list"):
            print(participants)
        elif cmd.lower().startswith("name"):
            _check_name(ws, cmd.lower())
        elif cmd.startswith("change"):
            _change(ws)
        elif cmd.startswith("cmd"):
            print(cmds)
        elif cmd.startswith("save"):
            _save_wb(wb, participants, False)

def _check_reg(ws:Worksheet, word):
    """Check the status of a participant on typing their register number.

    This function prints the details of the participant if they are registered,
    and their status for the debate. If they are not a participant, prompt
    is given to check their name in the list.
    """
    try:
       reg_no = int(re.split("check ", word)[1])
    except:
        return
    colB = ws['B']
    for cell in colB:
        if cell.value == reg_no:
            row:int = cell.row
            msg = ""
            for value in ws[row]:
                msg += str(value.value) + " "
            print(msg)
            if ws.cell(row=row, column=4).value!=None:
                print("Already sent the participant")
            elif ws.cell(row=row, column=3).value!=None:
                print("Participant is waiting in mini audi.")
            else:
                print("Participant has not arrived yet.")
            return
    print(reg_no, "not found. Check if their name is there with name <name>.")
    name = input("Enter name : ")
    name = "name " + name
    _check_name(ws, name)

def _check_name(ws:Worksheet, word:str):
    """Check if the name exists in the list. If it exists, display their 
    details including row number."""
    try:
        name = re.split("name ", word)[1]
    except:
        return
    colA = ws['A']
    found = 0
    for cell in colA:
        if (re.search(cell.value.lower(), name) 
                or re.search(name, cell.value.lower())):
            row = cell.row
            msg = f"{row} "
            for value in ws[row]:
                msg += str(value.value) + " "
            print(msg)
            found = 1
    if found == 0:
        print(name, "not found.")

def _add(ws:Worksheet, participants:list[int], word):
    """Add participants to the waiting list when they enter the venue.
    Also log the time when they enter."""
    try:
        reg_no = int(re.split("add", word)[1])
    except:
        return participants
    for cell in ws['B']:
        if cell.value == reg_no:
            row = cell.row
            if ws.cell(row=row, column=4).value!=None:
                print("Already sent the participant")
            elif ws.cell(row=row, column=3).value!=None:
                print("Already logged in")
            else:
                ws.cell(row=row, column=3, value=datetime.now())
                for value in ws[cell.row]:
                    print(value.value)
                participants.append(reg_no)
    try:
        with open("list.json", "w") as file:
            json.dump(participants, file)
    except:
        print("Couldn't save list")
        return participants
    return participants

def _send(ws:Worksheet, participants:list[int], word:str):
    try:
        reg_no = word[4:]
    except IndexError:
        return participants
    if reg_no == "" or reg_no == " ":
        try:
            participant = participants[0]
        except:
            print("No one to send now")
        else:
            print(participant)
    elif re.findall("\d", reg_no):
        reg_no = re.findall('\d+', reg_no)[0]
        reg_no = int(reg_no)
        try:
            participant = participants[0]
        except:
            print("No one to send now")
            return participants
        if reg_no == participant:
            for cell in ws['B']:
                if cell.value == reg_no:
                    ws.cell(row=cell.row, column=4, value=datetime.now())
                    for value in ws[cell.row]:
                        print(value.value)
            print("Sent", participant)
            participants.pop(0)
        else: 
            print("Wrong participant")
    try:
        with open("list.json", "w") as file:
            json.dump(participants, file)
    except:
        print("Couldn't save list")
        return participants
    return participants

def _change(ws:Worksheet):
    row = int(input("Enter row address : "))
    col = int(input("Enter column address : "))
    if col == 1:
        val = input("Enter name : ")
    elif col == 2:
        val = int(input("Enter reg no : "))
    elif col == 5:
        val = input("Enter class : ")
    ws.cell(row=row, column=col, value=val)
    for value in ws[row]:
        print(value.value)

def _save_wb(wb:Workbook, participants:list[int], close:bool):
    try:
        os.system("taskkill/im EXCEL.EXE ")
    except:
        print("Already closed")
    try:
        with open("list.json", "w") as file:
            json.dump(participants, file)
    except:
        print("Couldn't save list")
        return
    else:
        print("List saved")
    if close:
        sleep(10)
    # else: 
    #     sleep(1)
    try:
        wb.save("debate.xlsx")
    except PermissionError:
        print("Failed to save")
    else:
        print("Excel sheet saved")
        os.system("start EXCEL.EXE debate.xlsx")
        if close:
            sys.exit()

if __name__ == "__main__":
    main()