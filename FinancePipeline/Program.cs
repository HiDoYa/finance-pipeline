using System;
using System.IO;
using System.CommandLine;
using System.CommandLine.Invocation;

namespace Financepipeline
{
    class Program
    {
        static int Main(string[] args)
        {
            Console.WriteLine("Hello World!");

            var cmd = new RootCommand
            {
                new Argument<string>("username", "Mint username"),
                new Argument<string>("password", "Mint password"),
                new Option<string>("--download-path", "Path to download your Mint transactions"),
                new Option<string>("--filter-path", "Path to your filter specification in csv format"),
                new Option<string>("--google-cred-path", "Path to your google credentials (service account)")
            };

            cmd.Handler = CommandHandler.Create<string, string, string, string, string>(Startup);
            return cmd.Invoke(args);
        }

        static void Startup(string username, string password, string downloadPath, string filterPath, string googleCredPath)
        {
            if (downloadPath == "")
            {
                string home = System.Environment.GetFolderPath(System.Environment.SpecialFolder.Personal);
                downloadPath = Path.Join(home, ".mintapi");
                Console.WriteLine("Saving file to: " + downloadPath);
            }

            //using (Mint.Scraper mint = new Mint.Scraper(downloadPath, debug: true))
            //{
            //    mint.Login(username, password);
            //    mint.DownloadTransactions();
            //}

            //Mint.Parser parser = new Mint.Parser(downloadPath);
            //List<Mint.Transactions> transactions = parser.GetTransactions(filterPath)

            string spreadsheetId = "1pNs9XrzAQsuizWVbvq4D5yDQ4nESZpw--V7kI6tT91E";
            Sheet.Sheet test = new Sheet.Sheet(googleCredPath, spreadsheetId);

            test.CreateNewSpreadsheet("Testasdf");
            test.SetLastUpdatedTime();
            Console.WriteLine(test.GetLastUpdatedTime());
        }
    }
}
