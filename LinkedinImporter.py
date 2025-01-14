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

#%% Load Google Sheet and collect list of LinkedIn URLs

# Initialize sheet
spreadsheetID = '1qB0typdQr1e3KK38lp9e8zQi1kMf2Zn2Ud30Gakya58'
ss = ezsheets.Spreadsheet(spreadsheetID)
sheet = ss[0]

# Get LinkedIn URLs
headings = sheet.getRow(1)
linkdedinCol = headings.index('LinkedIn') + 1
linkedinUrls = sheet.getColumn(linkdedinCol)[1:]


#%% Create dictionary of profile data from URLs

api_key = 'd4-5cedTMKMQsGYQOvVaVg'
headers = {'Authorization': 'Bearer ' + api_key}
api_endpoint = 'https://nubela.co/proxycurl/api/v2/linkedin'

#%% Initialize dictionaries, lists, and counters

linkedinProfiles = {}
PROFILES_ADDED = 0
CREDITS_SPENT = 0
CREDIT_COST = 0.0264
invalidUrls = []

#%% Optional import existing JSON

databaseKey = input("Would you like to import an existing LinkedIn Database? Type '1' for yes, '2' for no:  ")

if databaseKey == '1':
    Tk().withdraw() # we don't want a full GUI, so keep the root window from appearing
    # filename = askopenfilename() # show an "Open" dialog box and return the path to the selected file

    # Open file dialog
    filename = askopenfilename(title="Select a file", filetypes=[("JSON files", "*.json")])

    if filename:
        print(f"File selected: {filename}. Loading data...")
        try:
            with open(filename) as jsonFile:
               linkedinProfiles  = json.load(jsonFile)
            jsonFile.close()
        except:
            print("Unable to load file. Exiting.")
            quit()
    else:
        print("No file selected.")
    


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
        except:
            print("Unable to load file. Exiting.")
            quit()
    else:
        print("No file selected.")
    
#%% Function to check if linkedIn URL is in a valid format

def is_valid_linkedin_url(url):
    # Regex to match LinkedIn profile URLs
    linkedin_regex = r"^https://(www\.)?linkedin\.com/in/[\w-]+/?$"
    
    # Check if the URL matches the LinkedIn profile pattern
    if not re.match(linkedin_regex, url):
        return False
    else:
        return True
    
#%% Add LinkedIn profiles to the linkedinProfiles dictionary

print("Scraping LinkedIn...")

for url in linkedinUrls:
    
    if len(url) < 1:
        pass
    # Handle duplicates
    elif url in linkedinProfiles:
        print("Duplicate entry")
    # Check to see if the URL is in a valid format
    elif is_valid_linkedin_url(url):
        CREDITS_SPENT += 1
        params = {

            'linkedin_profile_url': url,
            'use_cache': 'if-present',
            'fallback_to_cache': 'on-error',
            }
        try:
            response = requests.get(api_endpoint,
                                params=params,
                                headers=headers)
            response.raise_for_status()
            profile = json.loads(response.text)
            linkedinProfiles[url] = profile
            print('Added %s' %(profile['full_name']))
            
            # Update counter
            PROFILES_ADDED += 1
    
        except requests.exceptions.HTTPError as err:
            invalidUrls.append(url)
            if response.status_code == '404':
                print('%s is an invalid URL' %(url))
            else:
                print(f"HTTP Error occured: {err}")
            
        except requests.exceptions.RequestException as err:
            print(f"Error occurred: {err}")
#%%

print('Added %s profiles \n%s credits spent' %(PROFILES_ADDED,CREDITS_SPENT))
COST = CREDITS_SPENT*CREDIT_COST
print('Total Cost =  $%s USD' %(COST))

#%% Show remaining credits
headers = {'Authorization': 'Bearer ' + api_key}
api_endpoint = 'https://nubela.co/proxycurl/api/credit-balance'
response = requests.get(api_endpoint,
                        headers=headers)
print(response.text)


#%% Add date updated, current position, organization, and start date fields

for profile in linkedinProfiles:
    linkedinProfiles[profile]['last updated'] = str(datetime.now(timezone.utc))
    try:
        linkedinProfiles[profile]['current position'] = linkedinProfiles[profile]['experiences'][0]['title']
        linkedinProfiles[profile]['current company'] = linkedinProfiles[profile]['experiences'][0]['company']
    except:
        print("Unable to add current position and company data.")
    # Get start date as a string
    try:
        dateString = str(linkedinProfiles[profile]['experiences'][0]['starts_at']['year']) + "/" + str(linkedinProfiles[profile]['experiences'][0]['starts_at']['month']) + "/" + str(linkedinProfiles[profile]['experiences'][0]['starts_at']['day'])
        linkedinProfiles[profile]['start date'] = dateString
    except:
        print("Unable to add start date")
        
#%% Create Date String for File Naming

# Get date
d = datetime.now(timezone.utc)
MONTH = d.strftime('%m')
YEAR = d.strftime('%Y')
DAY = d.strftime('%d')
DATE_STRING = YEAR + MONTH + DAY


#%% Save data as Python

# =============================================================================
# print('Writing results......')
# fileName = DATE_STRING + "_AlumLinkedInData.py"
# resultFile = open(filename, 'w')
# resultFile.write('allData = ' + pprint.pformat(linkedinProfiles))
# resultFile.close()
# 
# resultFile = open('20240106_AlumLinkedInData_InvalidUrls.py', 'w')
# resultFile.write('allData = ' + pprint.pformat(invalidUrls))
# resultFile.close()
# 
# print('Done.')
# =============================================================================

#%% Save Data as JSON

print('Writing results......')
fileName = DATE_STRING + "_AlumLinkedInData.json"
with open(fileName, "w") as outfile: 
    json.dump(linkedinProfiles, outfile)
outfile.close()
print('Saved %s' %(filename))

fileName = DATE_STRING + "_InvalidUrls.json"
with open(fileName, "w") as outfile: 
    json.dump(invalidUrls, outfile)
print('Saved %s' %(filename))
outfile.close()
print('Done.')



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

# =============================================================================
# # Get keys from headers
# liKeys = headings[startIndex:]
# KEYS_CLEAN = []
# # Clean key names
# for key in liKeys:
#     key = key[:-4]
#     KEYS_CLEAN.append(key)
# =============================================================================
#%% Import Data
IMPORT_JSON_KEY = 0

if IMPORT_JSON_KEY == 1:
    
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
            if len(colHeader) > 1:
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

        
#%%

print('Summary:')
print('__________')
print('Added %s profiles \n%s credits spent' %(PROFILES_ADDED,CREDITS_SPENT))
COST = CREDITS_SPENT*CREDIT_COST
print('Total Cost =  $%s USD' %(COST))
headers = {'Authorization': 'Bearer ' + api_key}
api_endpoint = 'https://nubela.co/proxycurl/api/credit-balance'
response = requests.get(api_endpoint,
                        headers=headers)
print(response.text)
print('Total Profiles Updated: %s' %(len(linkedinProfiles)))
