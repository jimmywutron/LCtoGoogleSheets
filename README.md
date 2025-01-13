

# **Leadership Connect and LinkedIn to Google Sheets Importer**

Automatically updates the [Tech Labs Master Alumni List](https://docs.google.com/spreadsheets/d/1qB0typdQr1e3KK38lp9e8zQi1kMf2Zn2Ud30Gakya58/edit?gid=1154729072#gid=1154729072) using data from Leadership Connect lists and Linkedin URLS.

# Before you start

Please make sure to have the following packages installed in python: openpyxl, datetime, pandas, ezsheets. 

Use the package manager pip to install these packages.

```bash
pip install openpyxl pandas datetime pathlib ezsheets
```
Before starting, make sure to delete all .pickle files from your home directory. These files could be previous authorizations from a different Google Account. 

Following the instructions linked, set up the [Google Drive](https://developers.google.com/drive/api/quickstart/python) and [Google Sheets API](https://developers.google.com/sheets/api/quickstart/python). 

During setup, list yourself as a test user.

<img width="526" alt="Screenshot 2025-01-13 at 1 22 22 PM" src="https://github.com/user-attachments/assets/e67a6f4e-22c3-45a1-bfc6-2b7ef5d21852" />

Because we are using ezsheets, we will need to save the credentials file as both credentials.json and credentials-sheets.json in the home folder.

<img width="252" alt="Screenshot 2025-01-13 at 12 42 57 PM" src="https://github.com/user-attachments/assets/151cdc18-8302-4873-94a9-bed2f50a5cab" />

Now you should be ready to automatically update Google Sheets with Python.

# Updating from Leadership Connect

### Download the LeadershipConnect data

Open the Leadership Connect lists for [Congressional](https://app.leadershipconnect.io/contacts/list/69818380?view=table&sort=notes,desc) and [Executive](https://app.leadershipconnect.io/contacts/list/69818224?view=table&sort=notes,desc) alumni. Download the people data by clicking "Export All" in the top right corner.

<img width="1240" alt="Screenshot 2025-01-13 at 12 47 47 PM" src="https://github.com/user-attachments/assets/305f8695-a115-416c-955f-8f84e36423d2" />

Make sure all fields are selected, then click "Submit Export". 

<img width="1064" alt="Screenshot 2025-01-13 at 1 19 27 PM" src="https://github.com/user-attachments/assets/ca651a3d-e455-40f9-93e3-d802092d516a" />


Within a few minutes, the CSV file should be available for download in your email. Download this file to your local device. Repeat as needed.

<img width="400" alt="Screenshot 2025-01-13 at 1 24 52 PM" src="https://github.com/user-attachments/assets/97e3b2e1-02d4-4660-a349-158039bb80a0" />



### Running the program

Now you are ready to run the importer. To run the program, open the terminal and type

```bash
python LCImporter.py
```

or if you're in Windows

```bash
py LCImporter.py
```

Follow the instructions in the terminal to import the files and update the Google sheets, and the program will run on its own in the background.

# Updating from Linkedin

### Scraping Linkedin Data

The LinkedIn scraper uses ProxyCurl to get data from Linkedin profiles. ProxyCurl credits cost $0.0264 per profile updated. 

If you have a list of profiles that you don't want ProxyCurl to re-scrape and update, there is an option to import an existing .json list before scraping. Otherwise, to run the program open the terminal and type

```bash
python ProxyCurlScraper.py
```

or if you're in Windows

```bash
py ProxyCurlScraper.py
```

This program will save the scraped Linkedin data as a JSON file to your root folder. It will also save a list of invalid URLs.

<img width="237" alt="Screenshot 2025-01-13 at 2 04 35 PM" src="https://github.com/user-attachments/assets/a3cd935f-a15e-4bea-b1a5-0c1b5949f0d6" />
<img width="218" alt="Screenshot 2025-01-13 at 2 05 10 PM" src="https://github.com/user-attachments/assets/05be29ec-7961-44f0-966c-ab6fbb2520a8" />

You will use these JSON files to update the master spreadsheet.

### Importing the data

Now you are ready to import. Open the terminal and type

```bash
python LinkedinImporter.py
```

or if you're in Windows

```bash
py LinkedinImporter.py
```

Follow the directions to import the JSON data from the scraper and let the program run in the background to update the [Tech Labs Master Alumni List](https://docs.google.com/spreadsheets/d/1qB0typdQr1e3KK38lp9e8zQi1kMf2Zn2Ud30Gakya58/edit?gid=1154729072#gid=1154729072).
