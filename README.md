

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

Because we are using ezsheets, we will need to save the credentials file as both credentials.json and credentials-sheets.json in the home folder.

<img width="252" alt="Screenshot 2025-01-13 at 12 42 57 PM" src="https://github.com/user-attachments/assets/151cdc18-8302-4873-94a9-bed2f50a5cab" />

Now you should be ready to automatically update Google Sheets with Python.

# Updating from Leadership Connect

1. Download the LeadershipConnect data

Open the Leadership Connect lists for [Congressional](https://app.leadershipconnect.io/contacts/list/69818380?view=table&sort=notes,desc) and [Executive](https://app.leadershipconnect.io/contacts/list/69818224?view=table&sort=notes,desc) alumni. Download the people data by clicking "Export" in the top right corner.

<img width="1240" alt="Screenshot 2025-01-13 at 12 47 47 PM" src="https://github.com/user-attachments/assets/305f8695-a115-416c-955f-8f84e36423d2" />




# Running the program

To run the program, open the terminal and type

```bash
python LCImporter.py
```

or

```bash
py LCImporter.py
```

