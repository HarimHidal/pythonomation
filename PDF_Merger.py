# -*- coding: utf-8 -*-
"""
Created on Tue Feb  9 16:21:33 2021

@author: harim
"""


from PyPDF2 import PdfFileMerger
from os import listdir, path
from sys import exit
from time import sleep

base = r"C:\X\X\Desktop"


print()########################################################################


while True:
    
    print("\n\tBE SURE THE FOLLOWING PRE-REQUISITES ARE MET:\n")
    print("\t--> Los archivos están en el ESCRITORIO.")
    print("\t--> El nombre de los archivos va en orden alfabético de unión.\n")
    print("\tTO CONTINUE PRESS 'Enter'.\n")
    
    command = input(">> ")
    
    if command == "":
        break
    
    else:
        print("\n\tINVALID INPUT, TRY AGAIN...")
        
        
ls = listdir(base)
merging_list = []

for item in ls:
    if item.endswith(".pdf"):
        merging_list.append(path.join(base, item))


while True:
    
    if merging_list == []:
        print("\n\tTHERE ARE NO PDF FILES TO MERGE...\n")
        print("\tPRESS 'Enter' TO EXIT.\n")
        input(">> ")
        sleep(1)
        exit(0)
    
    elif len(merging_list) == 1:
        print("\n\tCANNOT MERGE JUST ONE PDF FILE...\n")
        print("\tPRESS 'Enter' TO EXIT.\n")
        input(">> ")
        sleep(1)
        exit(0)
        
    else:
        print("\n\tTHESE ARE THE PDF FILES TO MERGE:")
        print("\n**********************************************************\n")
        
        for index, file in enumerate(merging_list):
            print("\t{} --> {}".format(index+1, path.basename(str(file))))
            
        print("\n**********************************************************\n")
        print("\tTO CONTINUE:   press 'Enter'.")
        print("\tTO ABORT:      type 'exit'.\n")
        
        command = input(">> ")  
    
        if command == "":
            break

        elif command.lower() in "exit":
            sleep(1)
            exit(0)

        else:
            print("\n\tINVALID OPTION, TRY AGAIN...")

try:
    # OPENS MERGER    
    merger = PdfFileMerger()   
    
    for pdf_file in merging_list:
        merger.append(pdf_file)
    
    merger.write(r"C:\Users\harim\Desktop\POLAR.pdf")
    
    # CLOSES MERGER
    merger.close()    
except:
    print("\n\tUPS x_X LOOKS LIKE SOMETHING WENT WRONG...\n")

print("\n\tTHE PDF FILES WERE MERGED SUCCESSFULLY!\n")
sleep(1.5)
exit(0)