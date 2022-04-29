﻿using System;
using System.IO;
using System.Globalization;
using System.Collections.Generic;
using CsvHelper;

namespace Mint
{
    // Stores transaction data from Mint
    // Only includes fields I care about
    public class Transaction
    {
        public string Date { get; set; }
        public string Description { get; set; }
        public double Amount { get; set; }
        public string Category { get; set; }
    }

    public class Parser
    {
        // Get list of transactions from downloaded file and filter them through our filter file
        public List<Transaction> GetTransactions(string downloadPath, string filterPath)
        {
            var downloadFilePath = Path.Join(downloadPath, "transactions.csv");

            Filter filter = new Filter(filterPath);
            List<Transaction> records = new List<Transaction>();
            using (StreamReader reader = new StreamReader(downloadFilePath))
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

                    Transaction record = new Transaction
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

            records.Sort((Transaction t1, Transaction t2) => DateTime.Parse(t1.Date).CompareTo(DateTime.Parse(t2.Date)));

            return records;
        }

        public List<String> BatchRangeToCSV(List<List<String>> sheets)
        {
            var csvDoc = new List<string>();
            foreach (var sheet in sheets)
            {
                // Remove header
                if (sheet[0] == "Date")
                {
                    continue;
                }

                var line = SingleCommaSepLine(sheet);
                csvDoc.Add(line);
            }
            return csvDoc;
        }

        private String SingleCommaSepLine(List<String> elements)
        {
            var res = "";

            for (int i = 0; i < elements.Count; i++)
            {
                var curElement = elements[i];
                curElement = curElement.Replace(",", "");
                res += curElement;
                if (i != elements.Count - 1)
                {
                    res += ",";
                }
            }

            return res;
        }
    }
}
