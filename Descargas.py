# -*- coding: utf-8 -*-
"""
Created on Fri Jan 15 14:40:08 2021

@author: harim
"""


from os import path, listdir, startfile
from time import ctime, sleep
from re import search
from sys import exit
from docx import Document
from shutil import move


source = r"C:\X\X\Downloads"

Master_file = r"C:\X\X\X\X\Master.docx"

otros = [r"C:\X\X\X\X\X\Comprobantes de Pago"]


###############################################################################
    

def ParseRows (filename, table_number):
    """Parses the .docx file to return a dictionary of lists with the values of the rows."""
    
    doc = Document(filename)
    table = doc.tables[table_number]
    
    result = {}
    for i, row in enumerate(table.rows):
        if i == 0:
            continue
        for j, cell in enumerate(row.cells):
            if j == 0:
                result[cell.text] = result.get(cell.text, [])
                header = cell.text
            else:    
                result[header].append(cell.text.strip())
        
    return result   


print()########################################################################


ls = listdir(source)

now_date, now_hour = ctime()[:10], ctime()[11:13]

staged_files = []

for item in ls:
    file_dir = path.join(source, item)
    last_mod = ctime(path.getctime(file_dir))
    item_date, item_hour = last_mod[:10], last_mod[11:13]
    if item_date == now_date:
        if item_hour == now_hour or (int(item_hour)-1) == (int(now_hour)-1):
            staged_files.append(path.join(source, item))


staged_files = [ x for x in enumerate(reversed(staged_files)) ]

while True:
    
    if staged_files == []:
        
        print("\n\tTHERE ARE NO STAGED FILES...\n")
        print("\tPRESS 'Enter' TO EXIT.\n")
        input(">> ")
        sleep(1)
        exit(0)
    
    else:
        
        print("\n\tTHESE ARE THE STAGED FILES:")
        print("\n**********************************************************\n")
        
        for index, file in staged_files:
            print("\t{} --> {}".format(index, path.basename(str(file))))
            
        print("\n**********************************************************\n")
        print("\tREMOVE A FILE:       type [index]")
        print("\tCOMMIT THE FILE(S):  press 'Enter'\n")
        
        command = input(">> ")
    
    
    result = search(r"([\d]+)", command)
    files_index = [ x[0] for x in staged_files ]
    
    
    if command == "":
        break
        
    elif result == None:
        print("\n\tINVALID INDEX, TRY AGAIN...")
        
    elif int(result[1]) in files_index:
        
        for index, file in staged_files:
            if index == int(result[1]):
                staged_files.pop(staged_files.index((index, file)))
    
    else:
         print("\n\tINVALID INDEX, TRY AGAIN...")


univ_dir = ParseRows(Master_file, 2)
univ_dir = univ_dir["BASE"][0]

ls = listdir(univ_dir)

destinos = []

for item in ls:
    destinos.append(path.join(univ_dir, item))

for item in otros:
    destinos.append(item)
    
destinos = [ x for x in enumerate(destinos) ]


while True:
    
    print("\n\tWHERE DO YOU WANT THEM TO GO?")
    print("\n**********************************************************\n")
    
    for index, directory in destinos:
        print("\t{} --> {}".format(index, path.basename(str(directory))))
    
    print("\n**********************************************************\n")
    print("\tSET A DESTINY:     type [index]")
    print("\tCANCEL EVERYTING:  type 'exit'\n")
    
    command = input(">> ")
        
    result = search(r"([\d]+)", command)
    destinos_index = [ x[0] for x in destinos ]
    
    if command.lower() == "exit":
        sleep(1)
        exit(0)
        
    elif result == None:
        print("\n\tINVALID INDEX, TRY AGAIN...")
        
    elif int(result[1]) in destinos_index:
        
        target = ""
        for index, file in destinos:
            if index == int(result[1]):
                target = destinos[destinos.index((index, file))][1]
        
        for i, file in staged_files:
            try:
                move(file, target)
            except:
                print("\n\tERROR MOVING FILE [{}] '{}'...\n".format(i, file))
                sleep(1)
        
        print("\n\tSUCCESS: all files where moved correctly!\n")
        sleep(1)
        startfile(target)
        exit(0)
                
    else:
         print("\n\tINVALID INDEX, TRY AGAIN...")
    
        

