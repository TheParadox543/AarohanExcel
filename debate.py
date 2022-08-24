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
    wb = load_workbook("debate.xlsx")
    ws = wb.active
    cmds = "Commands are: check <no>, add <no>, send [no], "
    cmds += "name <name>, list, change, cmds"
    try:
        with open("list.json", 'r') as file:
            participants = json.load(file)
    except:
        participants = []
    os.system("start EXCEL.EXE debate.xlsx")
    print(cmds)
    while(True):
        cmd = input("\nEnter command:")
        if cmd.lower() == "quit":
            close_window(wb, participants)
        elif cmd.lower().startswith("check"):
            check_num(ws, cmd)
        elif cmd.lower().startswith("add"):
            participants = add(ws, participants, cmd)
        elif cmd.lower().startswith("send"):
            participants = send(ws, participants, cmd)
        elif cmd.startswith("list"):
            print(participants)
        elif cmd.lower().startswith("name"):
            check_name(ws, cmd.lower())
        elif cmd.startswith("change"):
            change(ws)
        elif cmd.startswith("cmd"):
            print(cmds)

def check_num(ws:Worksheet, word):
    try:
       reg_no = int(re.split("check ", word)[1])
    except:
        return
    colB = ws['B']
    for cell in colB:
        if cell.value == reg_no:
            row:int = cell.row
            for value in ws[row]:
                print(value.value)
            if ws.cell(row=row, column=4).value!=None:
                print("Already sent the participant")
            elif ws.cell(row=row, column=3).value!=None:
                print("Already logged in")
            else:
                print(datetime.now())
            return
    print(reg_no, "not found.")
    # name = input("Enter name:")
    # name += "name "
    # check_name(ws, name)

def check_name(ws:Worksheet, word:str):
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
            print("\n", row)
            msg = ""
            for value in ws[row]:
                msg += str(value.value) + " "
            print(msg)
            found = 1
    if found == 0:
        print(name, "not found.")

def add(ws:Worksheet, participants:list[int], word):
    try:
        reg_no = int(re.split("add", word)[1])
    except:
        return
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
        return
    return participants

def send(ws:Worksheet, participants:list[int], word:str):
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
        return
    return participants

def change(ws:Worksheet):
    row = int(input("Enter row address: "))
    col = int(input("Enter column address: "))
    if col == 1:
        val = input("Enter name: ")
    elif col == 2:
        val = int(input("Enter reg no: "))
    ws.cell(row=row, column=col, value=val)
    for value in ws[row]:
        print(value.value)

def close_window(wb:Workbook, participants:list[int]):
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
    sleep(5)
    try:
        wb.save("debate.xlsx")
    except PermissionError:
        print("Failed to save")
    else:
        print("Excel sheet saved")
        os.system("start EXCEL.EXE debate.xlsx")
        sys.exit()

if __name__ == "__main__":
    main()