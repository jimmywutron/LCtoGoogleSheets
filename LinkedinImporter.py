#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Created on Tue Jan  7 13:22:04 2025

@author: jimmywu
"""

import requests, json, ezsheets, re, requests, pprint
from datetime import datetime, timezone
from tkinter import Tk
from tkinter.filedialog import askopenfilename

#%% Intialize sheet and variables for import

# Initialize sheet
spreadsheetID = '1qB0typdQr1e3KK38lp9e8zQi1kMf2Zn2Ud30Gakya58'
ss = ezsheets.Spreadsheet(spreadsheetID)
sheet = ss[0]

# Initialize variables
headings = sheet.getRow(1)
idCol = headings.index('LinkedIn') + 1
startIndex = headings.index('last updated (LI)') + 1
maxCol = len(headings) + 1
maxRow = len(sheet.getColumn(1)) + 1

# Get keys from headers
liKeys = headings[startIndex:]
KEYS_CLEAN = []
# Clean key names
for key in liKeys:
    key = key[:-4]
    KEYS_CLEAN.append(key)

#%% Import Data


Tk().withdraw() # we don't want a full GUI, so keep the root window from appearing

# Open file dialog
print("Select the JSON file containing the LinkedIn Database")
filename = askopenfilename(title="Select the JSON file containing the LinkedIn Database", filetypes=[("JSON files", "*.json")])

if filename:
    print(f"File selected: {filename}. Loading data...")
    try:
        with open(filename) as jsonFile:
           linkedinProfiles  = json.load(jsonFile)
        jsonFile.close()
        print("Data loaded")
    except:
        print("Unable to load file. Exiting.")
        quit()
else:
    print("No file selected.")
    quit()



invalidLinksKey = input("Would you like to import an existing list of invalid URLs? Type '1' for yes, '2' for no:  ")

if invalidLinksKey == '1':
    Tk().withdraw() # we don't want a full GUI, so keep the root window from appearing
   
    # Open file dialog
    filename = askopenfilename(title="Select a file", filetypes=[("JSON files", "*.json")])

    if filename:
        print(f"File selected: {filename}. Loading data...")
        try:
            with open(filename) as jsonFile:
               invalidUrls  = json.load(jsonFile)
            jsonFile.close()
            print("Data loaded")
        except:
            print("Unable to load file. Exiting.")
            quit()
    else:
        print("No file selected.")
    


#%%

# Iterate sheet and update data
ss.refresh()
for row in range(2,maxRow):
    key = sheet[idCol,row]
    if key not in linkedinProfiles:
        print('No Key Found')
    else:
        print('Updating Key: %s' %(key))
        
        for col in range(startIndex,maxCol):
            colHeader = sheet[col,1][:-5]
            print("Attempting to update %s" %(colHeader))
            if colHeader in linkedinProfiles[key]: 
                print("Updating %s" %(colHeader))
                # If updating occupation, check to see if there are any changes
                if colHeader == 'occupation (LI)':
                    currentRole = sheet[col,row]
                    newRole = linkedinProfiles[key][colHeader]
                    if currentRole == newRole:
                        linkedinProfiles[key]['role change'] = 'no'
                    else:
                        linkedinProfiles[key]['role change'] = 'yes'
                sheet[col,row] = linkedinProfiles[key][colHeader]
            
print('Finished Updating')

        

