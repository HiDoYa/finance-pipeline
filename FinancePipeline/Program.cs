using System;
using System.IO;
using System.CommandLine;
using System.CommandLine.Invocation;
using System.Collections.Generic;

namespace Financepipeline
{
    class Program
    {
        static int Main(string[] args)
        {
            var cmd = new RootCommand
            {
                new Argument<string>("username", "Mint username"),
                new Argument<string>("password", "Mint password"),
                new Option<string>("--google-cred-path", "Path to your google credentials (service account)"),
                new Option<string>("--spreadsheet-id", "Path to Google sheet id"),
                new Option<string>("--download-path", "Path to download your Mint transactions"),
                new Option<string>("--filter-path", "Path to your filter specification in csv format"),
                new Option<string>("--category-path", "Path to config for transaction categories"),
                new Option<string>("--driver-path", "Path to chrome driver"),
                new Option<string>("--mfa-secret", "MFA secret for one time password"),
                new Option<string>("--transactions-path", "Output file path and name for transactions")
            };

            cmd.Handler = CommandHandler.Create<string, string, string, string, string, string, string, string, string, string>(Startup);
            return cmd.Invoke(args);
        }

        static void Startup(string username, string password, string downloadPath, string filterPath,
                            string googleCredPath, string spreadsheetId, string categoryPath, string driverPath,
                            string mfaSecret, string transactionsPath)
        {
            // Get download path
            if (downloadPath == "")
            {
                string home = System.Environment.GetFolderPath(System.Environment.SpecialFolder.Personal);
                downloadPath = Path.Join(home, ".mintapi");
            }
            Console.WriteLine("Saving file to: " + downloadPath);

            // Get transactions from mint
            using (Mint.Scraper mint = new(downloadPath, driverPath))
            {
                mint.Login(username, password, mfaSecret);
                mint.DownloadTransactions();
            }

            // Parse transactions from mint
            Dictionary<string, string> mapping = Mint.Categorizer.GetCategorizer(categoryPath);
            Mint.Parser parser = new();
            List<Mint.Transaction> transactions = parser.GetTransactions(downloadPath, filterPath);

            // Update with transactions
            var sheet = new Sheet.Sheet(googleCredPath, spreadsheetId);
            sheet.Update(transactions, mapping);

            // Save to CSV local file
            if (transactionsPath != "")
            {
                var sheets = sheet.GetSheets();
                var csv = parser.BatchRangeToCSV(sheets);
                File.WriteAllLines(transactionsPath, csv);
            }

        }
    }
}
