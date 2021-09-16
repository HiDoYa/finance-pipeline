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

            var cr = JsonConvert.DeserializeObject<PersonalServiceAccountCred>(json);

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
            // TODO
            var updateCellReq = new UpdateCellsRequest()
            {
                Fields = "*"
            };

            var batchUpdateReq = new BatchUpdateSpreadsheetRequest()
            {
                Requests = new List<Request>()
                {
                    new Request() {
                        UpdateCells = updateCellReq
                    }
                }
            };

            _service.Spreadsheets.BatchUpdate(batchUpdateReq, _spreadsheetId);
        }

        public DateTime GetLastUpdatedTime()
        {
            string range = "transactions!B1";
            var request = _service.Spreadsheets.Values.Get(_spreadsheetId, range);

            ValueRange response = request.Execute();
            string datetimeStr = (string)response.Values[0][0];
            return DateTime.Parse(datetimeStr);
        }

        public void SetLastUpdatedTime()
        {
            ValueRange valueRange = new ValueRange()
            {
                MajorDimension = "ROWS",
                Values = new List<IList<object>>()
                {
                    new List<object>() {
                        "Last Updated:",
                        DateTime.Now.ToString("MM/dd/yy H:mm")
                    }
                }
            };

            var req = _service.Spreadsheets.Values.Update(valueRange, _spreadsheetId, "transactions!A1");
            req.ValueInputOption = SpreadsheetsResource.ValuesResource.UpdateRequest.ValueInputOptionEnum.USERENTERED;

            req.Execute();
        }

        public void CreateSpreadsheetIfDNE(string title)
        {
            var test = _service.Spreadsheets.Get(_spreadsheetId).Execute();
            Console.WriteLine(test);
            // TODO

            CreateNewSpreadsheet(title);
        }

        public void CreateNewSpreadsheet(string title)
        {
            var addSheetReq = new AddSheetRequest()
            {
                Properties = new SheetProperties()
                {
                    Title = title
                }
            };

            var batchUpdateReq = new BatchUpdateSpreadsheetRequest()
            {
                Requests = new List<Request>()
                {
                    new Request() {
                        AddSheet = addSheetReq
                    }
                }
            };

            var req = _service.Spreadsheets.BatchUpdate(batchUpdateReq, _spreadsheetId);

            req.Execute();
        }
    }
}
