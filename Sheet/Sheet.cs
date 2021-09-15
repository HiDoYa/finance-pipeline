using System;
using System.IO;
using System.ComponentModel;
using System.Collections.Generic;

using Google.Apis.Auth.OAuth2;
using Google.Apis.Sheets.v4;
using Google.Apis.Sheets.v4.Data;
using Google.Apis.Services;
using Newtonsoft.Json;


namespace Sheet
{
    // PersonalServiceAccountCred is used to serialize service account credentials
    public class PersonalServiceAccountCred
    {
        public string type;
        public string project_id;
        public string private_key_id;
        public string private_key;
        public string client_email;
        public string client_id;
        public string auth_uri;
        public string token_uri;
        public string auth_provider_x509_cert_url;
        public string client_x509_cert_url;
    }

    // Sheet class reads and writes to google sheet
    public class Sheet
    {
        SheetsService _service;
        string _spreadsheetId;

        public Sheet(string serviceAccountCredentialFile, string spreadsheetId)
        {
            string json = File.ReadAllText(serviceAccountCredentialFile);
            _spreadsheetId = spreadsheetId;

            var cr = JsonConvert.DeserializeObject<PersonalServiceAccountCred>(json); // "personal" service account credential

            ServiceAccountCredential credential = new ServiceAccountCredential(
               new ServiceAccountCredential.Initializer(cr.client_email)
               {
                   Scopes = new[] {
                    SheetsService.Scope.SpreadsheetsReadonly,
                    SheetsService.Scope.Spreadsheets
                   }
               }.FromPrivateKey(cr.private_key));

            _service = new SheetsService(
                new BaseClientService.Initializer()
                {
                    HttpClientInitializer = credential,
                });


        }

        public void Update()
        {
            _service.Spreadsheets.BatchUpdate();

        }

        public string GetLastUpdatedTime()
        {
            string range = "transactions!A2:E";
            SpreadsheetsResource.ValuesResource.GetRequest request =
                    _service.Spreadsheets.Values.Get(_spreadsheetId, range);

            ValueRange response = request.Execute();

            IList<IList<Object>> values = response.Values;
            if (values != null && values.Count > 0)
            {
                foreach (var row in values)
                {
                    // Print columns A and E, which correspond to indices 0 and 4.
                    Console.WriteLine("{0}, {1}, {2}, {3}, {4}", row[0], row[1], row[2], row[3], row[4]);
                }
            }

            return "";
        }

        public void SetLastUpdatedTime()
        {
            ValueRange valuerange = new ValueRange();
            valuerange.MajorDimension = "ROWS";
            valuerange.Values = new List<IList<object>>();
            valuerange.Values.Add(new List<object>());
            valuerange.Values[0].Add(DateTime.Now);

            var req = _service.Spreadsheets.Values.Update(valuerange, _spreadsheetId, "transactions!A2");
            req.ValueInputOption = SpreadsheetsResource.ValuesResource.UpdateRequest.ValueInputOptionEnum.USERENTERED;

            req.Execute();
        }
    }
}
