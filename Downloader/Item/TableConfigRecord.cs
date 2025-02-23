using GoogleSheetsWrapper;

namespace GoogleSheetBulkDownloader.Downloader.Item
{
    internal class TableConfigRecord : BaseRecord
    {
        [SheetField(
                DisplayName = "outputname",
                ColumnID = 1,
                FieldType = SheetFieldType.String)]
        public string OutputFileName { get; set; }

        [SheetField(
                DisplayName = "spreadsheetId",
                ColumnID = 2,
                FieldType = SheetFieldType.String)]
        public string SpreadsheetId { get; set; }

        [SheetField(
                DisplayName = "tabName",
                ColumnID = 3,
                FieldType = SheetFieldType.String)]
        public string TabName { get; set; }

        public TableConfigRecord() { }
        public TableConfigRecord(IList<object> row, int rowId, int minColumnId = 1)
            : base(row, rowId, minColumnId) { }
    }
}
