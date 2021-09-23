using System.IO;
using System.Globalization;
using System.Collections.Generic;
using CsvHelper;

namespace Mint
{
    public class Filter
    {
        List<string> _removeList;
        Dictionary<string, string> _renameList;

        public Filter(string filterPath)
        {
            // Don't populate anything if filePath is null
            if (filterPath == "")
            {
                return;
            }

            _removeList = new List<string>();
            _renameList = new Dictionary<string, string>();

            using (StreamReader reader = new StreamReader(filterPath))
            using (CsvReader csv = new CsvReader(reader, CultureInfo.InvariantCulture))
            {
                csv.Read();
                csv.ReadHeader();
                while (csv.Read())
                {
                    string type = csv.GetField<string>("Type");
                    string current, target;

                    switch (type.ToLower())
                    {
                        case "remove":
                            current = csv.GetField<string>("Current").ToLower();
                            _removeList.Add(current);
                            break;
                        case "rename":
                            current = csv.GetField<string>("Current").ToLower();
                            target = csv.GetField<string>("Target");
                            _renameList.Add(current, target);
                            break;
                    }
                }
            }
        }

        public bool Keep(Transaction transaction)
        {
            string category = transaction.Category.ToLower();
            return !(_removeList.Contains(category));
        }

        public void Replace(ref Transaction transaction)
        {
            string category = transaction.Category.ToLower();
            if (_renameList.ContainsKey(category))
            {
                transaction.Category = _renameList[category];
            }
        }
    }
}
