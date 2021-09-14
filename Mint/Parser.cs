using System;
using System.IO;
using System.Globalization;
using System.Collections.Generic;
using CsvHelper;

namespace Mint
{
    // Only includes fields I care about
    public class Transactions
    {
        public string Date { get; set; }
        public string Description { get; set; }
        public double Amount { get; set; }
        public string Category { get; set; }
    }

    public class Parser
    {
        private string DownloadFilePath;

        public Parser(string downloadPath)
        {
            DownloadFilePath = Path.Join(downloadPath, "transactions.csv");
        }

        public List<Transactions> GetTransactions()
        {
            List<Transactions> records = new List<Transactions>();
            using (StreamReader reader = new StreamReader(DownloadFilePath))
            using (CsvReader csv = new CsvReader(reader, CultureInfo.InvariantCulture))
            {
                csv.Read();
                csv.ReadHeader();
                while (csv.Read())
                {
                    if (csv.GetField<string>("Transaction Type") == "debit")
                    {
                        continue;
                    }

                    Transactions record = new Transactions
                    {
                        Date = csv.GetField<string>("Date"),
                        Description = csv.GetField<string>("Original Description"),
                        Amount = csv.GetField<double>("Amount"),
                        Category = csv.GetField<string>("Category")

                    };
                    records.Add(record);
                }
            }

            return records;
        }

        private bool Filter()
        {

            return false;
        }
    }
}
