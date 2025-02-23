using GoogleSheetsWrapper;

namespace GoogleSheetBulkDownloader.Downloader.Item
{
    internal class TableConfigRepository : BaseRepository<TableConfigRecord>
    {
        public TableConfigRepository() { }

        public TableConfigRepository(SheetHelper<TableConfigRecord> sheetsHelper, BaseRepositoryConfiguration config)
            : base(sheetsHelper, config) { }
    }
}
