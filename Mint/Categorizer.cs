using System;
using System.IO;
using Newtonsoft.Json;
using System.Collections.Generic;

namespace Mint
{
    public class Categorizer
    {
        public Dictionary<string, string> Mapping { get; }

        public Categorizer(string categoryPath)
        {
            Mapping = new Dictionary<string, string>();

            Dictionary<string, List<string>> json;
            using (StreamReader r = new StreamReader(categoryPath))
            {
                string jsonRaw = r.ReadToEnd();
                json = JsonConvert.DeserializeObject<Dictionary<string, List<string>>>(jsonRaw);
            }

            foreach (KeyValuePair<string, List<string>> category in json)
            {
                foreach (string subcategory in category.Value)
                {
                    Mapping.Add(subcategory, category.Key);
                }
            }
        }
    }
}
