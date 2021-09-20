using System;
using System.IO;
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

        private string CreateSheetConditional(Dictionary<string, string> categoryDict)
        {

            string condition = "";
            foreach (KeyValuePair<string, string> item in categoryDict)
            {
                string temp = string.Join(",", "\"" + item.Key + "\"", "\"" + item.Value + "\"");
                condition = string.Join(",", condition, temp);
            }

            string categoryCond = String.Format("=SWITCH(INDIRECT(ADDRESS(ROW(),COLUMN()-1)){0}, \"Uncategorized\")", condition);
            return categoryCond;
        }

        public void Test()
        {
            var test = new Request();
            //test.SetDataValidation();
            //test.AddConditionalFormatRule();
            //test.AppendCells.

            ExtendedValue abc = new ExtendedValue()
            {
                StringValue = "Test"
            };

            var rowdata = new RowData()
            {
                Values = new List<CellData>()
                {
                    new CellData()
                    {
                        UserEnteredValue = abc
                    }
                }
            };

            var appendRequest = new AppendCellsRequest()
            {
                Fields = "*",
                Rows = new List<RowData>()
                {
                    rowdata
                }
            };

            var batchUpdateReq = new BatchUpdateSpreadsheetRequest()
            {
                Requests = new List<Request>()
                {
                    new Request() {
                        AppendCells = appendRequest,
                    }
                }
            };

            SpreadsheetsResource.BatchUpdateRequest request = _service.Spreadsheets.BatchUpdate(batchUpdateReq, _spreadsheetId);
            var response = request.Execute();

        }

        public void Update(List<Mint.Transactions> transactions, Dictionary<string, string> categoryDict)
        {
            string categoryCond = CreateSheetConditional(categoryDict);

            var transactionsVals = new List<IList<object>>();
            transactionsVals.Add(new List<object>()
            {
                "Date", "Description", "Sub Category", "Category", "Amount"
            });

            foreach (var transaction in transactions)
            {
                transactionsVals.Add(new List<object>()
                {
                    transaction.Date,
                    transaction.Description,
                    transaction.Category,
                    categoryCond,
                    transaction.Amount,
                });
            }

            ValueRange valueRange = new ValueRange()
            {
                MajorDimension = "ROWS",
                Values = transactionsVals
            };

            var req = _service.Spreadsheets.Values.Update(valueRange, _spreadsheetId, "transactions!A1");
            req.ValueInputOption = SpreadsheetsResource.ValuesResource.UpdateRequest.ValueInputOptionEnum.USERENTERED;

            req.Execute();
        }

        public DateTime GetLastUpdatedTime()
        {
            string range = "transactions!H1";
            var request = _service.Spreadsheets.Values.Get(_spreadsheetId, range);

            ValueRange response = request.Execute();
            if (response.Values == null)
            {
                return DateTime.MinValue;
            }

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

            var req = _service.Spreadsheets.Values.Update(valueRange, _spreadsheetId, "transactions!G1");
            req.ValueInputOption = SpreadsheetsResource.ValuesResource.UpdateRequest.ValueInputOptionEnum.USERENTERED;

            req.Execute();
        }

        public void CreateSpreadsheetIfDNE(string title)
        {
            var spreadsheetResource = _service.Spreadsheets.Get(_spreadsheetId).Execute();
            if (spreadsheetResource.Sheets[0].Properties.Title != title)
            {
                CreateNewSpreadsheet(title);
            }
        }

        private void CreateNewSpreadsheet(string title)
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
