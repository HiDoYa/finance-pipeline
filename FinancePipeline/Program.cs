﻿using System;
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
                new Option<string>("--google-cred-path", "Path to your google credentials (service account)"),
                new Option<string>("--spreadsheet-id", "Path to Google sheet id"),
                new Option<string>("--download-path", "Path to download your Mint transactions"),
                new Option<string>("--filter-path", "Path to your filter specification in csv format"),
                new Option<string>("--category-path", "Path to config for transaction categories")
            };

            cmd.Handler = CommandHandler.Create<string, string, string, string, string, string, string>(Startup);
            return cmd.Invoke(args);
        }

        static void Startup(string username, string password, string downloadPath, string filterPath, string googleCredPath, string spreadsheetId, string categoryPath)
        {
            if (downloadPath == "")
            {
                string home = System.Environment.GetFolderPath(System.Environment.SpecialFolder.Personal);
                downloadPath = Path.Join(home, ".mintapi");
                Console.WriteLine("Saving file to: " + downloadPath);
            }

            //using (Mint.Scraper mint = new Mint.Scraper(downloadPath))
            //{
            //    mint.Login(username, password);
            //    mint.DownloadTransactions();
            //}

            Mint.Categorizer categorizer = new Mint.Categorizer(categoryPath);
            Mint.Parser parser = new Mint.Parser(downloadPath);
            List<Mint.Transaction> transactions = parser.GetTransactions(filterPath);

            var sheet = new Sheet.Sheet(googleCredPath, spreadsheetId);
            sheet.Update(transactions, categorizer.Mapping);
        }
    }
}
