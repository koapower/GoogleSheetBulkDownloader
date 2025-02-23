using CsvHelper.Configuration.Attributes;

namespace GoogleSheetBulkDownloader.Downloader.Item
{
    /// <summary>
    /// 這是給local的csv檔案使用的，用來記載GoogleSheet上TableConfig的資訊(TableConfig的格式需和<see cref="TableConfigRecord"/>一致) 
    /// </summary>
    internal class TableConfigRow
    {
        [Name("name(optional)")]
        [Optional]
        public string Name { get; set; }
        [Name("spreadsheetid")]
        public string SpreadsheetId { get; set; }
        [Name("tabName")]
        public string TabName { get; set; }
    }
}
