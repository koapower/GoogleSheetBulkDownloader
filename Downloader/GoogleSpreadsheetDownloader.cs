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
        int maxParallelDownloads;

        public GoogleSpreadsheetDownloader(DownloaderConfig settings)
        {
            this.settings = settings;
            jsonCredential = File.ReadAllText(settings.JsonCredentialFilePath);
            if (!int.TryParse(settings.EndColumn, CultureInfo.InvariantCulture, out endColumn))
            {
                endColumn = 26; //預設讀取到Z欄
            }
            if (!int.TryParse(settings.MaxParallelDownloads, CultureInfo.InvariantCulture, out maxParallelDownloads) || maxParallelDownloads < 1)
            {
                maxParallelDownloads = 5; //注意Google Sheets API的read quota是每user每分鐘60個request
            }
        }

        /// <param name="sheetNames">TableConfig csv裡name欄的過濾名單，null或空代表全部下載</param>
        public void Download(IReadOnlyList<string> sheetNames = null)
        {
            var tableConfigs = GetTableConfigs();
            if (sheetNames != null && sheetNames.Count > 0)
            {
                tableConfigs = FilterTableConfigs(tableConfigs, sheetNames);
            }
            var toDlList = GetToDLSheetInfos(tableConfigs);
            DownloadAllSheets(toDlList);
        }

        private List<TableConfigRow> FilterTableConfigs(List<TableConfigRow> tableConfigs, IReadOnlyList<string> sheetNames)
        {
            var filtered = new List<TableConfigRow>();
            foreach (var name in sheetNames)
            {
                var match = tableConfigs.FirstOrDefault(c => c.Name == name);
                if (match == null)
                {
                    Console.WriteLine($"Warning: '{name}' does not match any name in the TableConfig csv, skipped.");
                }
                else
                {
                    filtered.Add(match);
                }
            }
            return filtered;
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
                var records = ExecuteWithRetry(() =>
                {
                    var sheetHelper = CreateSheetHelper<TableConfigRecord>(row.SpreadsheetId, row.TabName);
                    var repo = new TableConfigRepository(sheetHelper, repoConfig);
                    return repo.GetAllRecords();
                });
                toDlSheets.AddRange(records);
            }
            toDlSheets.Sort((x, y) => x.SpreadsheetId.CompareTo(y.SpreadsheetId));

            return toDlSheets;
        }

        private void DownloadAllSheets(List<TableConfigRecord> toDlList)
        {
            Console.WriteLine($"To Download Count: {toDlList.Count}, start downloading all sheets...");
            //同一個spreadsheet的tab共用一個SheetHelper依序下載，不同spreadsheet之間平行下載
            var groups = toDlList.GroupBy(x => x.SpreadsheetId).ToList();
            var options = new ParallelOptions { MaxDegreeOfParallelism = maxParallelDownloads };
            Parallel.ForEach(groups, options, group =>
            {
                SheetHelper sheetHelper = null;
                foreach (var record in group)
                {
                    //SheetHelper只用來提供Service連線，不需要對齊目前的tab
                    //(呼叫UpdateTabName會多發一個metadata request，而且tab不存在時會直接在sheet上建立新tab)
                    sheetHelper ??= ExecuteWithRetry(() => CreateSheetHelper(record.SpreadsheetId, record.TabName));
                    DownloadSheetAsCsv(sheetHelper, record.TabName, record.OutputFileName);
                }
            });
        }

        private void DownloadSheetAsCsv(SheetHelper sheetHelper, string tabName, string outputName)
        {
            var outputFolder = settings.OutputFolder;
            var filepath = Path.Combine(outputFolder, $"{outputName}.csv");
            //outputName可以含子資料夾(例如Localization/ItemName)，這裡連同子資料夾一起建立
            var fileFolder = Path.GetDirectoryName(filepath);
            if (!string.IsNullOrEmpty(fileFolder) && !Directory.Exists(fileFolder))
                Directory.CreateDirectory(fileFolder);

            var exporter = new SheetExporter(sheetHelper);
            // Default to CultureInfo.InvariantCulture and "," as the delimiter
            ExecuteWithRetry(() =>
            {
                using (var stream = new FileStream(filepath, FileMode.Create))
                {
                    // Query the range 1,1,8 => A1:G (ie 1st column, 1st row, 8th column and last row in the sheet)
                    //range自帶tab名稱，所以values請求不依賴sheetHelper目前指向哪個tab
                    var range = new SheetRange(tabName, 1, 1, endColumn);
                    Console.WriteLine($"Downloading tab {tabName} to {filepath}...");
                    exporter.ExportAsCsv(range, stream);
                }
                return true;
            });
        }

        /// <summary>
        /// Google Sheets API的read quota是per minute計算的，被429擋下時等待一段時間再重試
        /// </summary>
        private T ExecuteWithRetry<T>(Func<T> action)
        {
            const int maxRetries = 3;
            const int baseDelaySeconds = 35;
            int attempt = 0;
            while (true)
            {
                try
                {
                    return action();
                }
                catch (Google.GoogleApiException e) when (e.HttpStatusCode == System.Net.HttpStatusCode.TooManyRequests && attempt < maxRetries)
                {
                    attempt++;
                    int delaySeconds = baseDelaySeconds * attempt;
                    Console.WriteLine($"Rate limited by Google API, waiting {delaySeconds}s before retry ({attempt}/{maxRetries})...");
                    Thread.Sleep(TimeSpan.FromSeconds(delaySeconds));
                }
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
