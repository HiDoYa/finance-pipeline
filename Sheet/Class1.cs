using System;
using System.IO;
using System.Collections.Generic;

using Google.Apis.Auth.OAuth2;
using Google.Apis.Sheets.v4;
using Google.Apis.Sheets.v4.Data;
using Google.Apis.Services;

namespace Sheet
{
    public class Sheet
    {
        public Sheet()
        {
            FileStream ts = File.OpenRead("/Users/hiroya.gojo/Downloads/finance-pipeline-325808-36b341a22811.json");
            GoogleCredential credential = GoogleCredential.FromServiceAccountCredential(ServiceAccountCredential.FromServiceAccountData(ts));

            var service = new SheetsService(new BaseClientService.Initializer()
            {
                HttpClientInitializer = credential,
                ApplicationName = "test"
            });

            string spreadsheetId = "1pNs9XrzAQsuizWVbvq4D5yDQ4nESZpw--V7kI6tT91E";
            string range = "Sheet1!A2:C";
            SpreadsheetsResource.ValuesResource.GetRequest request =
                    service.Spreadsheets.Values.Get(spreadsheetId, range);

            ValueRange response = request.Execute();
            IList<IList<Object>> values = response.Values;
            if (values != null && values.Count > 0)
            {
                Console.WriteLine("Name, Major");
                foreach (var row in values)
                {
                    // Print columns A and E, which correspond to indices 0 and 4.
                    Console.WriteLine("{0}, {1}", row[0], row[1]);
                }
            }
        }
    }
}
