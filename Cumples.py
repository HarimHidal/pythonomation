# -*- coding: utf-8 -*-
"""
Created on Fri Jan 15 17:18:55 2021

@author: harim
"""

import pyautogui as gui, webbrowser as wb, sys, os, time, re, pyperclip


file = r"C:\X\X\X\Cumpleaños.txt"
mensajeHBD = "Feliz cumpleaños! Espero que te la pases muy bien y que cumplas muchos más, un abrazo!"
link = "https://docs.google.com/document/d/XXXXXXXXXX/export?format=txt"

###############################################################################


def ReadTXT (path):
    """Opens the file in the path and returns an ordered list with the lines."""
    
    with open(path, "r", encoding="UTF-8") as file:
        result, final = [], []
        lines = file.readlines()
    
    for line in lines:
        result.append(line.strip().strip("\ufeff"))
    
    result[:] = [x for x in result if x]
    
    i = -1
    for item in result:
        if item[0].isnumeric():
            final[i].append(item)
        else:
            final.append([item])
            i += 1
    
    return final
    

def CheckForBD (lista, month, day):
    result = []
    for mes in lista:
        if mes[0] == month:
            mes.pop(0)
            for person in mes:
                buscar = re.search(r"([\d]+) ([\w ]+)", person)
                if buscar[1] == day:
                    result.append(buscar[2])
    return result


print()########################################################################


if len(sys.argv) > 1:
    
    if sys.argv[1] in "update":
     
        print("CHECK WHETHER THERE IS A Cumpleaños.txt OR NOT...\n") 
        if os.path.isfile(file):     
            
            print("DELETE THE Cumpleaños.txt FILE...")
            os.remove(file)
        
        print("DOWNLOADING Cumpleaños DOC AS .txt FILE...\n")
        wb.open(link)

    print("ALL SET!\n")
    sys.exit()

else:
        
    print("PARSE THE DATE LOOKING FOR DAY AND MONTH...\n")
    result = re.search("^[a-zA-Z]+ ([a-zA-Z]+)  ?([0-9]+) [\d]+:[\d]+:[\d]+ [0-9]+$", time.ctime())
    
    eng_esp = {"Jan":"ENERO", "Feb":"FEBRERO", "Mar":"MARZO", "Apr":"ABRIL", "May":"MAYO", "Jun":"JUNIO", 
               "Jul":"JULIO", "Aug":"AGOSTO", "Sep":"SEPTIEMBRE", "Oct":"OCTUBRE", "Nov":"NOVIMEBRE", "Dec":"DICIEMBRE"}
    
    month = eng_esp[result[1]]
    
    print("CHECK WHETHER THERE IS A Cumpleaños.txt OR NOT...\n")                    
    if os.path.isfile(file):
        
        print("OPEN Cumpleaños.txt TO GET THE BIRTHDAYS...\n")
        lista = ReadTXT(file)
    
        cumpleañeros = CheckForBD(lista, month, result[2])
           
        if cumpleañeros == []:
            gui.alert("Hoy no hay cumpleaños.")
        
        else:
           
            print("OPEN WHATSAPP ON DESKTOP...\n")
            gui.hotkey("win", "3")
        
            print("ASK THE USER TO CONTINUE WHEN READY...\n")
            gui.alert("Hoy es cumpleaños de " + str(cumpleañeros) + "\nIs WhatsApp already open to continue?")
            gui.click(111,113, duration=0.2)
            gui.write(cumpleañeros[0][:4], interval=0.1)
            
            print("COPY THE BIRTHDAY MESSAGE IN THE CLIPBOARD...\n")
            pyperclip.copy(mensajeHBD)
        
        print("ALL SET!\n")
        print("NOTE: use [ update ] command-line argument to update the file.\n")
        sys.exit()

    else:
        
        print("DOWNLOADING Cumpleaños DOC AS .txt FILE...\n")
        wb.open(link)
    
        print("ALL SET!\n")
        print("NOTE: use [ update ] command-line argument to update the main file.\n")
        sys.exit()
        
        