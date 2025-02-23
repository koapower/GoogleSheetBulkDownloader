using GoogleSheetBulkDownloader.Downloader;
using Microsoft.Extensions.Configuration;

namespace GoogleSheetBulkDownloader
{
    internal class Program
    {
        static readonly string version = "1.1";
        static IConfigurationRoot settings;
        static GoogleSpreadsheetDownloader downloader;

        static void Main(string[] args)
        {
            Console.WriteLine($"GoogleSheetDownloader v{version}");
            Initialize();
            Download();
        }

        private static void Initialize()
        {
            settings = BuildConfig();
            var downloaderSetting = CreateDownloaderConfig(settings);
            downloader = new GoogleSpreadsheetDownloader(downloaderSetting);
        }

        private static IConfigurationRoot BuildConfig()
        {
            var builder = new ConfigurationBuilder();
            builder.AddJsonFile("DownloaderSettings.json", optional: false, reloadOnChange: true);

            return builder.Build();
        }

        private static DownloaderConfig CreateDownloaderConfig(IConfigurationRoot config)
        {
            return new DownloaderConfig
            {
                OutputFolder = config["OutputFolder"],
                EndColumn = config["EndColumn"],
                GoogleServiceAccountName = config["GoogleServiceAccountName"],
                TableConfigCsvFilePath = config["TableConfigCsvFilePath"],
                JsonCredentialFilePath = config["JsonCredentialFilePath"]
            };
        }

        private static void Download()
        {
            downloader.Download();
        }
    }
}
