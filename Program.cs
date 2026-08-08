using GoogleSheetBulkDownloader.Downloader;
using Microsoft.Extensions.Configuration;

namespace GoogleSheetBulkDownloader
{
    internal class Program
    {
        static readonly string version = "1.2";
        static IConfigurationRoot settings;
        static GoogleSpreadsheetDownloader downloader;

        static void Main(string[] args)
        {
            Console.WriteLine($"GoogleSheetDownloader v{version}");
            if (args.Contains("--help") || args.Contains("-h"))
            {
                PrintUsage();
                return;
            }

            var sheetNames = ParseSheetNames(args);
            Initialize();
            Download(sheetNames);
        }

        private static void PrintUsage()
        {
            Console.WriteLine("Usage: GoogleSheetBulkDownloader.exe [--sheets <name> <name> ...]");
            Console.WriteLine("  (no args)                Download every TableConfig listed in the local TableConfig csv.");
            Console.WriteLine("  --sheets <name> ...      Only download the TableConfigs whose 'name' column matches. Space separated, duplicates ignored.");
        }

        /// <summary>
        /// Collects the values following --sheets (until the next --option or end of args), de-duplicated.
        /// Returns an empty list when --sheets is absent, which means "download everything".
        /// </summary>
        private static List<string> ParseSheetNames(string[] args)
        {
            var names = new List<string>();
            bool collectingSheets = false;
            foreach (var arg in args)
            {
                if (arg.StartsWith("-"))
                {
                    collectingSheets = arg.Equals("--sheets", StringComparison.OrdinalIgnoreCase);
                    if (!collectingSheets)
                        Console.WriteLine($"Unknown option '{arg}', ignored. Use --help to see usage.");
                    continue;
                }

                if (collectingSheets && !names.Contains(arg))
                    names.Add(arg);
            }

            return names;
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
                JsonCredentialFilePath = config["JsonCredentialFilePath"],
                MaxParallelDownloads = config["MaxParallelDownloads"]
            };
        }

        private static void Download(List<string> sheetNames)
        {
            downloader.Download(sheetNames);
        }
    }
}
