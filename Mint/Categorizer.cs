using System;
using System.IO;
using Newtonsoft.Json;
using System.Collections.Generic;

namespace Mint
{
    public class Categorizer
    {
        // Get mapping between subcategory to category based on json
        static public Dictionary<string, string> GetCategorizer(string categoryPath)
        {
            var mapping = new Dictionary<string, string>();

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
                    mapping.Add(subcategory, category.Key);
                }
            }

            return mapping;
        }
    }
}
