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
            Console.WriteLine("Hello World!");

            var cmd = new RootCommand
            {
                new Argument<string>("username", "Mint username"),
                new Argument<string>("password", "Mint password"),
                new Option<string>("--download-path", "Path to download your Mint transactions")
            };

            cmd.Handler = CommandHandler.Create<string, string, string>(Startup);
            return cmd.Invoke(args);
        }

        static void Startup(string username, string password, string downloadPath)
        {
            if (downloadPath == "")
            {
                string home = System.Environment.GetFolderPath(System.Environment.SpecialFolder.Personal);
                downloadPath = Path.Join(home, ".mintapi");
                Console.WriteLine("Saving file to: " + downloadPath);
            }

            using (Mint.Scraper mint = new Mint.Scraper(downloadPath, debug: true))
            {
                mint.Login(username, password);
                mint.DownloadTransactions();
            }

            Mint.Parser parser = new Mint.Parser(downloadPath);
            List<Mint.Transactions> transactions = parser.GetTransactions();


            Sheet.Sheet test = new Sheet.Sheet();
        }
    }
}
