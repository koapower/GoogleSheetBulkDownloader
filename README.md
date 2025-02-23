# GoogleSheetBulkDownloader
A C# console tool that downloads multiple Google Sheets tabs as CSV files based on a configuration file. This tool uses the [GoogleSheetsWrapper](https://github.com/SteveWinward/GoogleSheetsWrapper/tree/main) library to interact with Google Sheets.

## How to use
> [!IMPORTANT]
> It is required to use a Service Account to access your Google Sheets spreadsheets. Please setup a Google API Service Account before using the tool. see [here](#SetupGoogleServiceAccountandgetcredentialjson)
1. Download the latest version at [Release Page](https://github.com/koapower/GoogleSheetBulkDownloader/releases).
2. Unzip the files in the same folder.
3. Place your credential json file in the same folder, just like the ````project-1234567890123-abcdefgh1319.json```` file.
4. Open DownloaderSetting.json with a text editor. You *must* change ````GoogleServiceAccountName```` and ````JsonCredentialFilePath```` to your own account and credential json file.

| Field Name | Description |
| ---------------- | ------------------------------- |
| ````OutputFolder```` | ````The output folder path. Csv will be downloaded into this folder.```` |
| ````TableConfigCsvFilePath```` | ````The tableconfig csv file. This is used to list all the tableconfig google sheets (not the actual sheets you want to download).```` |
| ````EndColumn```` | ````The max column the downloader read, default is 26, which means the downloader will read from A to Z column.```` |
| ````GoogleServiceAccountName```` | ````Your google service account name.```` |
| ````JsonCredentialFilePath```` | ````Your credential json file.```` |
5. Open any of your spreadsheet and create a sheet, name it ````TableConfig````. You will be putting all of the sheets you want to download here. It should follow the structure like this.
> [!NOTE]
> Spreadsheet ID can be found in your sheet link, ie https://docs.google.com/spreadsheets/d/spreadsheetId/edit?gid=0#gid=0

| outputname | spreadsheetId | tabName |
| ---------------- | ------------------------------- | ---------------- |
| Sheet1OutputName | Sheet1SpreadsheetId | Sheet1TabName |
| Sheet2OutputName | Sheet2SpreadsheetId | Sheet2TabName |
| MapData | 1a2b3c4d5e6f7g8h9i0j1k2l3m4n5o6p | Map |
6. Make sure *all* your google sheets (including the actual sheets you want to download) can be accessed by your service account. It can be done by clicking "File" -> "Share" and set your service account as a "Viewer".
7. Open TableConfig.csv and copy your ````TableConfig```` spreadsheet ID like this. Downloader supports multiple ````TableConfig```` spreadsheets.
````csv
name(optional),spreadsheetid,tabName
TableConfig1,TableConfig1_Spreadsheet_ID,TableConfig
TableConfig2,TableConfig2_Spreadsheet_ID,TableConfig
````
8. Run ````GoogleSheetBulkDownloader.exe````, the csv files will be downloaded to the output folder.

## Setup Google Service Account and get credential json
1. Create a Google Cloud Project.
2. Enable "Google Sheets API" and "Google Drive API".
3. Go to "APIs & Services > Credentials" and choose "CREATE CREDENTIALS > Service account".
4. Fill out the form and click "Create" and "Done".
5. Go to "IAM & Admin > Service Accounts > Keys" and click "ADD KEY > Create new key".
6. Select JSON key type and press "Create". (The file will be downloaded automatically.)

