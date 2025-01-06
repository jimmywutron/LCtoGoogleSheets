

# **Leadership Connect to Google Sheets Importer**

Automatically updates the [Tech Labs Master Alumni List](https://docs.google.com/spreadsheets/d/1qB0typdQr1e3KK38lp9e8zQi1kMf2Zn2Ud30Gakya58/edit?gid=1154729072#gid=1154729072) using data from Leadership Connect lists.

# Before you start

Please make sure to have the following packages installed in python: openpyxl, datetime, pandas, ezsheets. 

Use the package manager pip to install these packages.

```bash
pip install openpyxl pandas datetime pathlib ezsheets
```
Before starting, make sure to delete all .pickle files from your home directory, as these may be authorizations from a different account. 

Following the instructions linked, set up the [Google Drive](https://developers.google.com/drive/api/quickstart/python and [Google Sheets API](https://developers.google.com/sheets/api/quickstart/python). Because we are using ezsheets, we will need to save the credentials file as both credentials.json and credentials-sheets.json in the home folder.


