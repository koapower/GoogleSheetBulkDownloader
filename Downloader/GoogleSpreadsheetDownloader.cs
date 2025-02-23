using CsvHelper;
using GoogleSheetBulkDownloader.Downloader.Item;
using GoogleSheetsWrapper;
using System.Globalization;

namespace GoogleSheetBulkDownloader.Downloader
{
    public class GoogleSpreadsheetDownloader
    {
        readonly DownloaderConfig settings;
        string jsonCredential;
        int endColumn;

        public GoogleSpreadsheetDownloader(DownloaderConfig settings)
        {
            this.settings = settings;
            jsonCredential = File.ReadAllText(settings.JsonCredentialFilePath);
            if (!int.TryParse(settings.EndColumn, CultureInfo.InvariantCulture, out endColumn))
            {
                endColumn = 26; //預設讀取到Z欄
            }
        }

        public void Download()
        {
            var tableConfigs = GetTableConfigs();
            var toDlList = GetToDLSheetInfos(tableConfigs);
            DownloadAllSheets(toDlList);
        }

        private List<TableConfigRow> GetTableConfigs()
        {
            Console.WriteLine("Reading TableConfig csv file...");
            var tableConfigs = new List<TableConfigRow>();
            using (var reader = new StreamReader(settings.TableConfigCsvFilePath))
            {
                using (var csv = new CsvReader(reader, CultureInfo.InvariantCulture))
                {
                    var records = csv.GetRecords<TableConfigRow>();
                    tableConfigs.AddRange(records);
                }
            }

            return tableConfigs;
        }

        private List<TableConfigRecord> GetToDLSheetInfos(List<TableConfigRow> tableConfigs)
        {
            Console.WriteLine("Reading TableConfig records from Google Sheets...");
            var repoConfig = GetBaseRepositoryConfig();
            var toDlSheets = new List<TableConfigRecord>();
            foreach (var row in tableConfigs)
            {
                var sheetHelper = CreateSheetHelper<TableConfigRecord>(row.SpreadsheetId, row.TabName);
                var repo = new TableConfigRepository(sheetHelper, repoConfig);
                var records = repo.GetAllRecords();
                toDlSheets.AddRange(records);
            }
            toDlSheets.Sort((x, y) => x.SpreadsheetId.CompareTo(y.SpreadsheetId));

            return toDlSheets;
        }

        private void DownloadAllSheets(List<TableConfigRecord> toDlList)
        {
            Console.WriteLine($"To Download Count: {toDlList.Count}, start downloading all sheets...");
            SheetHelper sheetHelper = null;
            foreach (var record in toDlList)
            {
                if (sheetHelper == null || sheetHelper.SpreadsheetID != record.SpreadsheetId)
                {
                    sheetHelper = CreateSheetHelper(record.SpreadsheetId, record.TabName);
                }
                if (sheetHelper.TabName != record.TabName)
                {
                    sheetHelper.UpdateTabName(record.TabName);
                }

                DownloadSheetAsCsv(sheetHelper, record.OutputFileName);
            }
        }

        private void DownloadSheetAsCsv(SheetHelper sheetHelper, string outputName)
        {
            var outputFolder = settings.OutputFolder;
            if(!Directory.Exists(outputFolder))
                Directory.CreateDirectory(outputFolder);

            var exporter = new SheetExporter(sheetHelper);
            var filepath = Path.Combine(outputFolder, $"{outputName}.csv");
            // Default to CultureInfo.InvariantCulture and "," as the delimiter
            using (var stream = new FileStream(filepath, FileMode.Create))
            {
                // Query the range 1,1,8 => A1:G (ie 1st column, 1st row, 8th column and last row in the sheet)
                var range = new SheetRange(sheetHelper.TabName, 1, 1, endColumn);
                Console.WriteLine($"Downloading tab {sheetHelper.TabName} to {filepath}...");
                exporter.ExportAsCsv(range, stream);
            }
        }

        private SheetHelper CreateSheetHelper(string sheetId, string tabName)
        {
            // Create a SheetHelper class
            var sheetHelper = new SheetHelper(
            // https://docs.google.com/spreadsheets/d/<SPREADSHEET_ID_IS_HERE>/edit#gid=0
                sheetId,
            // The email for the service account you created
                settings.GoogleServiceAccountName,
            // the name of the tab you want to access, leave blank if you want the default first tab
                tabName);

            sheetHelper.Init(jsonCredential);

            return sheetHelper;
        }

        private SheetHelper<T> CreateSheetHelper<T>(string sheetId, string tabName) where T : BaseRecord
        {
            // Create a SheetHelper class
            var sheetHelper = new SheetHelper<T>(
            // https://docs.google.com/spreadsheets/d/<SPREADSHEET_ID_IS_HERE>/edit#gid=0
                sheetId,
            // The email for the service account you created
                settings.GoogleServiceAccountName,
            // the name of the tab you want to access, leave blank if you want the default first tab
                tabName);

            sheetHelper.Init(jsonCredential);

            return sheetHelper;
        }

        private BaseRepositoryConfiguration GetBaseRepositoryConfig()
        {
            return new BaseRepositoryConfiguration()
            {
                // Does the table have a header row?
                HasHeaderRow = true,
                // Are there any blank rows before the header row starts?
                HeaderRowOffset = 0,
                // Are there any blank rows before the first row in the data table starts?                
                DataTableRowOffset = 0,
            };
        }
    }
}
