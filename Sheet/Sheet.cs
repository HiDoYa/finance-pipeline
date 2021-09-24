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

        // Consts
        private const int CATEGORY_ROW = 3;
        private const int SUBCAT_ROW = 2;

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
            CreateSpreadsheetIfDNE(metadataSheetname, ref spreadsheetResource, true);
            DateTime lastUpdated = GetLastUpdatedTime(metadataSheetname);

            List<Request> reqList = new List<Request>();
            HashSet<int> listOfModifiedSpreadsheets = new HashSet<int>();

            // Add each row of transactions
            foreach (var transaction in transactions)
            {
                DateTime transactionDate = DateTime.Parse(transaction.Date);

                // Skip if the transaction date should've been added already
                if (lastUpdated.CompareTo(transactionDate) > 0)
                {
                    continue;
                }

                string monthYear = transactionDate.ToString("MMMM yyyy");
                int sheetId = CreateSpreadsheetIfDNE(monthYear, ref spreadsheetResource);

                var vals = PopulateCellDataWithTransaction(transaction, categoryCond);

                // Set row data and format
                reqList.Add(CreateAppendCellRequest(sheetId, vals));

                listOfModifiedSpreadsheets.Add(sheetId);
            }

            // Autoresize newly added 
            var categories = new List<string>(categoryDict.Keys);
            foreach (var sheetId in listOfModifiedSpreadsheets)
            {
                reqList.Add(DataValidationRequest(sheetId, categories, SUBCAT_ROW));
                reqList.Add(AutoUpdateDimensionRequest(sheetId));
                reqList.Add(UpdateDimensionRequest(sheetId, CATEGORY_ROW));
            }

            ApplyRequestList(reqList.ToArray());
            SetLastUpdatedTime(metadataSheetname);
        }

        private List<CellData> PopulateCellDataWithTransaction(Mint.Transaction transaction, string categoryCond)
        {
            return new List<CellData>()
                {
                    new CellData()
                    {
                        UserEnteredFormat = new CellFormat()
                        {
                            NumberFormat = new NumberFormat()
                            {
                                Type = "DATE",
                                Pattern = "mm-dd-yyyy"
                            }
                        },
                        UserEnteredValue =  new ExtendedValue() { NumberValue = DateTime.Parse(transaction.Date).ToOADate() }
                    },
                    new CellData() { UserEnteredValue =  new ExtendedValue() { StringValue = transaction.Description } },
                    new CellData() { UserEnteredValue =  new ExtendedValue() { StringValue = transaction.Category } },
                    new CellData() { UserEnteredValue =  new ExtendedValue() { FormulaValue = categoryCond } },
                    new CellData()
                    {
                        UserEnteredFormat = new CellFormat()
                        {
                            NumberFormat = new NumberFormat()
                            {
                                Type = "CURRENCY"
                            }
                        },
                        UserEnteredValue =  new ExtendedValue() { NumberValue = transaction.Amount }
                    },
                };

        }

        // Create data validation request for a drop down in sheets
        private Request DataValidationRequest(int sheetId, List<string> values, int columnIndex)
        {
            var valueList = new List<ConditionValue>();
            foreach (string value in values)
            {
                valueList.Add(new ConditionValue()
                {
                    UserEnteredValue = value,
                });
            }

            var dataReq = new SetDataValidationRequest()
            {
                Range = new GridRange()
                {
                    StartRowIndex = 1,
                    StartColumnIndex = columnIndex,
                    EndColumnIndex = columnIndex + 1,
                    SheetId = sheetId,
                },
                Rule = new DataValidationRule()
                {
                    Condition = new BooleanCondition()
                    {
                        Type = "ONE_OF_LIST",
                        Values = valueList,
                    },
                    ShowCustomUi = true,
                    InputMessage = "Select a valid sub category",
                }
            };

            return new Request() { SetDataValidation = dataReq };
        }

        // Get request to update dimensions for all columns in spreadsheet
        private Request AutoUpdateDimensionRequest(int sheetId, int numCols = 10)
        {
            var resizeRequest = new AutoResizeDimensionsRequest()
            {
                Dimensions = new DimensionRange()
                {
                    SheetId = sheetId,
                    StartIndex = 0,
                    EndIndex = numCols,
                    Dimension = "COLUMNS",
                }
            };

            return new Request() { AutoResizeDimensions = resizeRequest };
        }

        // Get request to update dimension for a specific column in spreadsheet
        private Request UpdateDimensionRequest(int sheetId, int col, int size = 120)
        {
            var updateRequest = new UpdateDimensionPropertiesRequest()
            {
                Fields = "*",
                Range = new DimensionRange()
                {
                    SheetId = sheetId,
                    Dimension = "COLUMNS",
                    StartIndex = col,
                    EndIndex = col + 1
                },
                Properties = new DimensionProperties()
                {
                    PixelSize = size,
                }
            };

            return new Request() { UpdateDimensionProperties = updateRequest };
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
        private int GetIdFromSheetName(string sheetName)
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
        private Request CreateAppendCellRequest(int sheetId, List<CellData> values)
        {
            var appendRequest = new AppendCellsRequest()
            {
                SheetId = sheetId,
                Fields = "*",
                Rows = new List<RowData>()
                {
                    new RowData()
                    {
                        Values = values
                    }
                }
            };

            return new Request() { AppendCells = appendRequest };
        }

        // Apply list of requests as batch
        private void ApplyRequestList(Request[] requestList)
        {
            var batchUpdateReq = new BatchUpdateSpreadsheetRequest()
            {
                Requests = requestList,
            };

            SpreadsheetsResource.BatchUpdateRequest request = _service.Spreadsheets.BatchUpdate(batchUpdateReq, _spreadsheetId);

            request.Execute();
        }

        // Get last updated time
        private DateTime GetLastUpdatedTime(string sheetName = "Metadata")
        {
            string range = sheetName + "!B1";
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
        private void SetLastUpdatedTime(string sheetName = "Metadata", DateTime? time = null)
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
        private Spreadsheet GetSpreadsheetResource()
        {
            return _service.Spreadsheets.Get(_spreadsheetId).Execute();
        }

        private bool SpreadsheetExists(string title, Spreadsheet spreadsheetResource)
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

        // Wrapper for create new spreadsheet (create header, freeze header, check if already exists)
        private int CreateSpreadsheetIfDNE(string title, ref Spreadsheet spreadsheetResource, bool blank = false)
        {
            // Do nothing if spreadsheet already exists
            if (SpreadsheetExists(title, spreadsheetResource))
            {
                return GetIdFromSheetName(title);
            }

            // Create spreadsheet id
            int sheetId = CreateNewSpreadsheet(title);

            // Dictates whether spreadsheet should contain basic headers
            if (!blank)
            {
                // Create new spreadsheet and write headers
                var header = new List<CellData>()
                {
                    new CellData() { UserEnteredValue= new ExtendedValue() { StringValue = "Date" } },
                    new CellData() { UserEnteredValue= new ExtendedValue() { StringValue = "Description" } },
                    new CellData() { UserEnteredValue= new ExtendedValue() { StringValue = "Sub Category" } },
                    new CellData() { UserEnteredValue= new ExtendedValue() { StringValue = "Category" } },
                    new CellData() { UserEnteredValue= new ExtendedValue() { StringValue = "Amount" } },
                };
                var appendRequest = CreateAppendCellRequest(sheetId, header);

                // Freeze top row
                var updateRequest = new Request()
                {
                    UpdateSheetProperties = new UpdateSheetPropertiesRequest()
                    {
                        Fields = "*",
                        Properties = new SheetProperties()
                        {
                            Title = title,
                            SheetId = sheetId,
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
            }

            // Update spreadsheet resource with new title
            // Dumb work around because of API call limits (essentially caching)
            spreadsheetResource.Sheets.Add(new Google.Apis.Sheets.v4.Data.Sheet()
            {
                Properties = new SheetProperties()
                {
                    Title = title,
                }
            });

            return sheetId;
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
