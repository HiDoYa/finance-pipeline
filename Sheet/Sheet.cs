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
        Dictionary<string, int> _cachedSheetId;

        // Setup service account and sheets API
        public Sheet(string serviceAccountCredentialFile, string spreadsheetId)
        {
            string json = File.ReadAllText(serviceAccountCredentialFile);
            _spreadsheetId = spreadsheetId;
            _cachedSheetId = new Dictionary<string, int>();

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

        // The main method responsible for calling most other sheets functions
        public void Update(List<Mint.Transaction> transactions, Dictionary<string, string> categoryDict)
        {
            string categoryCond = CreateSheetConditional(categoryDict);
            Spreadsheet spreadsheetResource = GetSpreadsheetResource();

            string metadataSheetname = "Metadata";
            CreateSpreadsheetIfDNE(metadataSheetname, ref spreadsheetResource);
            DateTime lastUpdated = GetLastUpdatedTime(metadataSheetname);

            List<Request> reqList = new List<Request>();

            foreach (var transaction in transactions)
            {
                DateTime transactionDate = DateTime.Parse(transaction.Date);

                // Skip if the transaction date should've been added already
                if (lastUpdated.CompareTo(transactionDate) > 0)
                {
                    continue;
                }

                string monthYear = transactionDate.ToString("MMMM-yy");
                int spreadsheetId = CreateSpreadsheetIfDNE(monthYear, ref spreadsheetResource);

                ExtendedValue[] vals = new ExtendedValue[]
                {
                    new ExtendedValue() { StringValue = transaction.Date },
                    new ExtendedValue() { StringValue = transaction.Description } ,
                    new ExtendedValue() { StringValue = transaction.Category },
                    new ExtendedValue() { FormulaValue = categoryCond },
                    new ExtendedValue() { StringValue = transaction.Amount.ToString() },
                };

                reqList.Add(CreateAppendCellRequest(spreadsheetId, vals));
            }

            ApplyRequestList(reqList.ToArray());
            SetLastUpdatedTime(metadataSheetname);
        }

        // Create conditional statement formula for sheets
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

        // Get id from sheet name. Use cache if available
        public int GetIdFromSheetName(string sheetName)
        {
            if (_cachedSheetId.ContainsKey(sheetName))
            {
                return _cachedSheetId[sheetName];
            }

            var spreadsheetResource = _service.Spreadsheets.Get(_spreadsheetId).Execute();
            foreach (var sheet in spreadsheetResource.Sheets)
            {
                if (sheet.Properties.Title == sheetName)
                {
                    int sheetId = sheet.Properties.SheetId.GetValueOrDefault();
                    _cachedSheetId.Add(sheetName, sheetId);
                    return sheetId;
                }
            }

            return -1;
        }

        // Create append cell request but don't execute it yet
        public Request CreateAppendCellRequest(int sheetId, ExtendedValue[] values)
        {
            var celldata = new List<CellData>();
            foreach (ExtendedValue value in values)
            {
                celldata.Add(new CellData()
                {
                    UserEnteredValue = value
                });
            }

            var appendRequest = new AppendCellsRequest()
            {
                SheetId = sheetId,
                Fields = "*",
                Rows = new List<RowData>()
                {
                    new RowData()
                    {
                        Values = celldata
                    }
                }
            };

            return new Request() { AppendCells = appendRequest };
        }

        // Apply list of requests as batch
        public void ApplyRequestList(Request[] requestList)
        {
            var batchUpdateReq = new BatchUpdateSpreadsheetRequest()
            {
                Requests = requestList,
            };

            SpreadsheetsResource.BatchUpdateRequest request = _service.Spreadsheets.BatchUpdate(batchUpdateReq, _spreadsheetId);

            request.Execute();
        }

        // Get last updated time
        public DateTime GetLastUpdatedTime(string sheetName = "Metadata")
        {
            string range = sheetName + "!A1";
            var request = _service.Spreadsheets.Values.Get(_spreadsheetId, range);

            ValueRange response = request.Execute();
            if (response.Values == null)
            {
                return DateTime.MinValue;
            }

            string datetimeStr = (string)response.Values[0][0];
            return DateTime.Parse(datetimeStr);
        }

        // Set last updated time. If time is set, use that time.
        public void SetLastUpdatedTime(string sheetName = "Metadata", DateTime? time = null)
        {
            if (time == null)
            {
                time = DateTime.Now;
            }

            ValueRange valueRange = new ValueRange()
            {
                MajorDimension = "ROWS",
                Values = new List<IList<object>>()
                {
                    new List<object>() {
                        "Last Updated:",
                        time.Value.ToString("MM/dd/yy H:mm")
                    }
                }
            };

            var request = _service.Spreadsheets.Values.Update(valueRange, _spreadsheetId, sheetName + "!A1");
            request.ValueInputOption = SpreadsheetsResource.ValuesResource.UpdateRequest.ValueInputOptionEnum.USERENTERED;

            request.Execute();
        }

        // Spreadsheet resource is a deliberate call; an attempt to try to minimze sheet api calls
        public Spreadsheet GetSpreadsheetResource()
        {
            return _service.Spreadsheets.Get(_spreadsheetId).Execute();
        }

        public bool SpreadsheetExists(string title, Spreadsheet spreadsheetResource)
        {
            foreach (var spreadsheet in spreadsheetResource.Sheets)
            {
                if (spreadsheet.Properties.Title == title)
                {
                    return true;
                }
            }
            return false;
        }

        public int CreateSpreadsheetIfDNE(string title, ref Spreadsheet spreadsheetResource)
        {
            // Do nothing if spreadsheet already exists
            if (SpreadsheetExists(title, spreadsheetResource))
            {
                return GetIdFromSheetName(title);
            }

            // Create new spreadsheet and write headers
            int spreadsheetId = CreateNewSpreadsheet(title);
            ExtendedValue[] header = new ExtendedValue[]
            {
                    new ExtendedValue() { StringValue = "Date" },
                    new ExtendedValue() { StringValue = "Description" },
                    new ExtendedValue() { StringValue = "Sub Category" },
                    new ExtendedValue() { StringValue = "Category" },
                    new ExtendedValue() { StringValue = "Amount" }
            };
            var appendRequest = CreateAppendCellRequest(spreadsheetId, header);

            var updateRequest = new Request()
            {
                UpdateSheetProperties = new UpdateSheetPropertiesRequest()
                {
                    Fields = "*",
                    Properties = new SheetProperties()
                    {
                        Title = title,
                        SheetId = spreadsheetId,
                        GridProperties = new GridProperties()
                        {
                            FrozenRowCount = 1,
                            RowCount = 1000,
                            ColumnCount = 20,
                        }
                    }
                }
            };

            ApplyRequestList(new Request[] { appendRequest, updateRequest });


            // Update spreadsheet resource with new title
            // Dumb work around because of API call limits (essentially caching)
            spreadsheetResource.Sheets.Add(new Google.Apis.Sheets.v4.Data.Sheet()
            {
                Properties = new SheetProperties()
                {
                    Title = title,
                }
            });

            return spreadsheetId;
        }

        // Create new spreadsheet and execute it
        private int CreateNewSpreadsheet(string title)
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

            var request = _service.Spreadsheets.BatchUpdate(batchUpdateReq, _spreadsheetId);

            request.Execute();

            return GetIdFromSheetName(title);
        }
    }
}
