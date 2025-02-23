using GoogleSheetBulkDownloader.Downloader;
using Microsoft.Extensions.Configuration;

namespace GoogleSheetBulkDownloader
{
    internal class Program
    {
        static readonly string version = "1.0";
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
            downloader = new GoogleSpreadsheetDownloader(settings);
        }

        private static IConfigurationRoot BuildConfig()
        {
            var builder = new ConfigurationBuilder();
            builder.AddJsonFile("DownloaderSettings.json", optional: false, reloadOnChange: true);

            return builder.Build();
        }

        private static void Download()
        {
            downloader.Download();
        }
    }
}
