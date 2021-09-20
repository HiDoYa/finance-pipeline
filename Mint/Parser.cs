﻿using System;
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
        private string _downloadFilePath;

        public Parser(string downloadPath)
        {
            _downloadFilePath = Path.Join(downloadPath, "transactions.csv");
        }

        public List<Transactions> GetTransactions(string filterPath)
        {
            Filter filter = new Filter(filterPath);
            List<Transactions> records = new List<Transactions>();
            using (StreamReader reader = new StreamReader(_downloadFilePath))
            using (CsvReader csv = new CsvReader(reader, CultureInfo.InvariantCulture))
            {
                csv.Read();
                csv.ReadHeader();
                while (csv.Read())
                {
                    if (csv.GetField<string>("Transaction Type").ToLower() != "debit")
                    {
                        continue;
                    }

                    Transactions record = new Transactions
                    {
                        Date = csv.GetField<string>("Date"),
                        Description = csv.GetField<string>("Original Description"),
                        Amount = csv.GetField<double>("Amount"),
                        Category = csv.GetField<string>("Category"),
                    };

                    if (filter.Keep(record))
                    {
                        // Check if any filter replacements need to be done
                        filter.Replace(ref record);

                        // Add record
                        records.Add(record);
                    }
                }
            }

            return records;
        }
    }
}
