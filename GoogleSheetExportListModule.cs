using Google;
using Google.Apis.Auth.OAuth2;
using Google.Apis.Drive.v3;
using Google.Apis.Services;
using Google.Apis.Sheets.v4;
using Google.Apis.Sheets.v4.Data;
using Microsoft.AspNetCore.Http;
using Sushi.Mediakiwi.Data;
using Sushi.Mediakiwi.Framework;
using Sushi.Mediakiwi.Framework.Interfaces;

namespace Sushi.Mediakiwi.Module.GoogleSheetsSync
{
    public class GoogleSheetExportListModule : IListModule
    {
        #region Properties

        public string ModuleTitle => "Google Sheets exporter";

        private SheetsService _sheetsService { get; set; }
        private DriveService _driveService { get; set; }

        public bool ShowInSearchMode { get; set; }

        public bool ShowInEditMode { get; set; }

        public string IconClass { get; set; }

        public string IconURL { get; set; }

        public string Tooltip { get; set; }

        public bool ConfirmationNeeded { get; set; }

        public string ConfirmationTitle { get; set; }

        public string ConfirmationQuestion { get; set; }

        private string ClientSecretsFileName { get; set; }

        #endregion Properties

        #region CTor

        public GoogleSheetExportListModule()
        {
            ShowInSearchMode = true;
            ShowInEditMode = false;
            Tooltip = "View this list in Google Sheets";
            IconClass = "icon-cloud-upload";
        }

        #endregion CTor

        #region Initialize Module Async

        public async Task InitAsync(string credentialsFileName)
        { 
             await InitAsync(credentialsFileName, null);
        }

        public async Task InitAsync(string credentialsFileName, string clientSecretsFileName)
        {
            GoogleCredential credential;
            ClientSecretsFileName = clientSecretsFileName;

            try
            {
                using (var stream = new FileStream(credentialsFileName, FileMode.Open, FileAccess.Read))
                {
                    credential = GoogleCredential.FromStream(stream).CreateScoped(new string[] { SheetsService.Scope.Spreadsheets, DriveService.Scope.DriveFile });
                }
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine(ex);
                throw ex;
            }

            await ModuleInstaller.InstallWhenNeededAsync();
            // Create Google Sheets API service.
            _sheetsService = new SheetsService(new BaseClientService.Initializer()
            {
                HttpClientInitializer = credential,
                ApplicationName = nameof(GoogleSheetExportListModule),
            });

            // Create Google Drive API service.
            _driveService = new DriveService(new BaseClientService.Initializer()
            {
                HttpClientInitializer = credential,
                ApplicationName = nameof(GoogleSheetExportListModule),
            });
        }

        #endregion Initialize Module Async

        #region Get Sheet Async

        private async Task<Spreadsheet> GetSheetAsync(string id)
        {
            if (string.IsNullOrWhiteSpace(id))
            {
                return null;
            }

            try
            {
                var getRequest = new SpreadsheetsResource.GetRequest(_sheetsService, id);
                getRequest.IncludeGridData = true;
                var responseGet = await getRequest.ExecuteAsync();
                if (responseGet != null)
                {
                    return responseGet;
                }
            }
            catch (GoogleApiException ex)
            {
                if (ex.HttpStatusCode == System.Net.HttpStatusCode.NotFound)
                {
                    return null;
                }
            }

            return null;
        }

        #endregion Get Sheet Async

        #region Update Sheet Async

        private async Task UpdateSheetAsync(IComponentListTemplate inList, Spreadsheet currentSpreadSheet, Data.GoogleSheetListLink listLink)
        {
            var valueRange = new ValueRange();
            valueRange.Values = new List<IList<object>>();
            if (currentSpreadSheet != null)
            {
                // Perform listsearch
                inList.wim.CurrentPage = 0;
                inList.wim.IsExportMode_XLS = true;
                inList.wim.Console.CurrentListInstance.wim.IsExportMode_XLS = true;
                inList.wim.GridDataCommunication.CurrentPage = -1;
                inList.wim.GridDataCommunication.ShowAll = true;
                inList.wim.GridDataCommunication.PageSize = 10000;

                inList.wim.DoListSearch();

                // Get first sheet
                var currentSheet = currentSpreadSheet.Sheets.FirstOrDefault();

                // All Additional Requests
                List<Request> requests = new List<Request>();

                // TODO: remove these lines, are for development only
                if (currentSheet?.ProtectedRanges?.Count > 0)
                {
                    foreach (var protectedRange in currentSheet.ProtectedRanges)
                    {
                        requests.Add(new Request()
                        {
                            DeleteNamedRange = new DeleteNamedRangeRequest()
                            {
                                NamedRangeId = protectedRange.NamedRangeId
                            }
                        });
                        
                        requests.Add(new Request()
                        {
                            DeleteProtectedRange = new DeleteProtectedRangeRequest()
                            {
                                ProtectedRangeId = protectedRange.ProtectedRangeId
                            }
                        });
                    }
                }

                // Delete all developer metadata
                requests.Add(new Request()
                {
                    DeleteDeveloperMetadata = new DeleteDeveloperMetadataRequest()
                    {
                        DataFilter = new DataFilter()
                        {
                            DeveloperMetadataLookup = new DeveloperMetadataLookup()
                            {
                                MetadataKey = "valueType"
                            }
                        }
                    }
                });

                requests.Add(new Request()
                {
                    DeleteDeveloperMetadata = new DeleteDeveloperMetadataRequest()
                    {
                        DataFilter = new DataFilter()
                        {
                            DeveloperMetadataLookup = new DeveloperMetadataLookup()
                            {
                                MetadataKey = "propertyName"
                            }
                        }
                    }
                });

                requests.Add(new Request()
                {
                    DeleteDeveloperMetadata = new DeleteDeveloperMetadataRequest()
                    {
                        DataFilter = new DataFilter()
                        {
                            DeveloperMetadataLookup = new DeveloperMetadataLookup()
                            {
                                MetadataKey = "propertyType"
                            }
                        }
                    }
                });
                
                requests.Add(new Request()
                {
                    DeleteDeveloperMetadata = new DeleteDeveloperMetadataRequest()
                    {
                        DataFilter = new DataFilter()
                        {
                            DeveloperMetadataLookup = new DeveloperMetadataLookup()
                            {
                                MetadataKey = "propertyIsKey"
                            }
                        }
                    }
                });


                // Check if we have a named range already
                var namedSheetId = currentSpreadSheet?.NamedRanges?.FirstOrDefault(x => x.Name == "MK.Columns")?.NamedRangeId;

                if (string.IsNullOrEmpty(namedSheetId))
                {
                    namedSheetId = "MK.Columns";
                    requests.Add(new Request()
                    {
                        AddNamedRange = new AddNamedRangeRequest()
                        {
                            NamedRange = new NamedRange()
                            {
                                Name = "MK.Columns",
                                Range = new GridRange()
                                {
                                    StartRowIndex = 0,
                                    EndRowIndex = 1,
                                    StartColumnIndex = 0,
                                    EndColumnIndex = 10000,
                                    SheetId = currentSheet.Properties.SheetId
                                },
                                NamedRangeId = namedSheetId
                            }
                        }
                    });
                }

                // Set first row as protected
                if (currentSheet.ProtectedRanges == null || currentSheet.ProtectedRanges.Count ==0)
                {
                    requests.Add(new Request()
                    {
                        AddProtectedRange = new AddProtectedRangeRequest()
                        {
                            ProtectedRange = new ProtectedRange()
                            {
                                NamedRangeId = namedSheetId,
                                Description ="These columns should match the ones exported from Mediakiwi",
                                Editors = new Editors()
                                {
                                    Users = new List<string>(),
                                    DomainUsersCanEdit = false,
                                    Groups = new List<string>(),
                                },
                            },
                        }
                    });
                }

                // Complete data collection
                IList<RowData> rowsData = new List<RowData>();

                // Add column headers
                RowData headerRow = new RowData() 
                { 
                    Values = new List<CellData>() 
                };

                // Add header columns
                foreach (var col in inList.wim.ListDataColumns.List)
                {
                    headerRow.Values.Add(new CellData()
                    {
                        UserEnteredValue = new ExtendedValue()
                        {
                            StringValue = col.ColumnName,
                        },
                        UserEnteredFormat = new CellFormat()
                        {
                            BackgroundColor = new Color()
                            {
                                Blue = 0.8f,
                                Green = 0.8f,
                                Red = 0.8f,
                                Alpha = 1
                            },
                            TextFormat = new TextFormat()
                            {
                                Bold = true,
                            },
                        },
                    });
                }

                // Add header row
                rowsData.Add(headerRow);

                List<DeveloperMetadata> devMetaData = new List<DeveloperMetadata>();

                // Retrieve Item Type, and set as Developer MetaData so it can later be reconstructed
                if (inList?.wim?.ListDataCollection?.Count > 0)
                {
                    var itemType = inList.wim.ListDataCollection[0].GetType().FullName;
                    devMetaData.Add(new DeveloperMetadata()
                    {
                        Location = new DeveloperMetadataLocation()
                        {
                            SheetId = currentSheet.Properties.SheetId
                        },
                        MetadataKey = "valueType",
                        MetadataValue = itemType,
                        Visibility = "DOCUMENT",
                    });
                }


                // Get the type of item we're interating
                var listItemType = inList?.wim?.ListDataCollection?.Count > 0 ? inList.wim.ListDataCollection[0].GetType() : null;

                // Add Developer MetaData for Header Columns
                foreach (var col in inList.wim.ListDataColumns.List)
                {
                    var colIdx = inList.wim.ListDataColumns.List.IndexOf(col) + 1;

                    devMetaData.Add(new DeveloperMetadata()
                    {
                        Location = new DeveloperMetadataLocation()
                        {
                            DimensionRange = new DimensionRange()
                            {
                                StartIndex = colIdx,
                                EndIndex = colIdx + 1,
                                SheetId = currentSheet.Properties.SheetId,
                                Dimension = "COLUMNS"
                            },
                        },
                        MetadataKey = "propertyName",
                        MetadataValue = col.ColumnValuePropertyName,
                        Visibility = "DOCUMENT",
                    });


                    // Add the type of the property this column represents
                    if (listItemType != null)
                    {
                        devMetaData.Add(new DeveloperMetadata()
                        {
                            Location = new DeveloperMetadataLocation()
                            {
                                DimensionRange = new DimensionRange()
                                {
                                    StartIndex = colIdx,
                                    EndIndex = colIdx + 1,
                                    SheetId = currentSheet.Properties.SheetId,
                                    Dimension = "COLUMNS"
                                },
                            },
                            MetadataKey = "propertyType",
                            MetadataValue = listItemType.GetProperty(col.ColumnValuePropertyName).PropertyType.FullName,
                            Visibility = "DOCUMENT",
                        });
                    }
                }

                foreach (var _developerMetadata in devMetaData)
                {
                    // Add developer metadata request
                    requests.Add(new Request()
                    {
                        CreateDeveloperMetadata = new CreateDeveloperMetadataRequest()
                        {
                            DeveloperMetadata = _developerMetadata,
                        }
                    });
                }

                // Loop through items
                foreach (var item in inList.wim.ListDataCollection)
                {
                    var rowData = new RowData()
                    {
                        Values = new List<CellData>()
                    };

                    // Get the type of item
                    var itemType = item.GetType();

                    // Create an instance of this type
                    var tempComp = Activator.CreateInstance(itemType);

                    // Reflect actual data to instance
                    Utils.ReflectProperty(item, tempComp);

                    // Loop through items in the search list
                    foreach (var col in inList.wim.ListDataColumns.List)
                    {
                        // Create new cell data for column value
                        var cellData = new CellData();

                        // Get data item from search list
                        var rawData = itemType.GetProperty(col.ColumnValuePropertyName).GetValue(tempComp);

                        if (rawData is string rawString)
                        {
                            cellData.UserEnteredValue = new ExtendedValue()
                            {
                                StringValue = rawString
                            };
                        }
                        else if (rawData is bool rawBool)
                        {
                            cellData.UserEnteredValue = new ExtendedValue()
                            {
                                BoolValue = rawBool
                            };
                        }
                        else if (rawData is int rawInt)
                        {
                            cellData.UserEnteredValue = new ExtendedValue()
                            {
                                NumberValue = rawInt
                            };
                        }
                        else if (rawData is double rawDouble)
                        {
                            cellData.UserEnteredValue = new ExtendedValue()
                            {
                                NumberValue = rawDouble
                            };
                        }
                        else if (rawData is decimal rawDecimal)
                        {
                            cellData.UserEnteredValue = new ExtendedValue()
                            {
                                NumberValue = (double)rawDecimal
                            };
                        }
                        else if (rawData is DateTime rawDateTime)
                        {
                            cellData.UserEnteredValue = new ExtendedValue()
                            {
                                NumberValue = rawDateTime.ToOADate()
                            };
                            cellData.UserEnteredFormat = new CellFormat()
                            {
                                NumberFormat = new NumberFormat()
                                {
                                    Type = "DATE_TIME"
                                }
                            };
                        }
                        else if (rawData != null)
                        {
                            cellData.UserEnteredValue = new ExtendedValue()
                            {
                                StringValue = rawData.ToString()
                            };
                        }

                        rowData.Values.Add(cellData);
                   
                    }
                    rowsData.Add(rowData);
                }

                // Update the SpreadSheet data
                requests.Add(new Request()
                {
                    UpdateCells = new UpdateCellsRequest()
                    {
                        Rows = rowsData,
                        Fields = "*",
                        Start = new GridCoordinate()
                        {
                            ColumnIndex = 0,
                            RowIndex = 0,
                            SheetId = currentSheet.Properties.SheetId,
                        }
                    }
                });


                // Auto fit columns
                requests.Add(new Request()
                {
                    AutoResizeDimensions = new AutoResizeDimensionsRequest()
                    {
                        Dimensions = new DimensionRange()
                        {
                            Dimension = "COLUMNS",
                            EndIndex = headerRow.Values.Count,
                            StartIndex = 0,
                            SheetId = currentSheet.Properties.SheetId
                        }
                    }
                });

                // Run all requests
                if (requests?.Count > 0)
                {
                    BatchUpdateSpreadsheetRequest requestBody = new BatchUpdateSpreadsheetRequest();
                    requestBody.Requests = requests;
                    SpreadsheetsResource.BatchUpdateRequest request = _sheetsService.Spreadsheets.BatchUpdate(requestBody, currentSpreadSheet.SpreadsheetId);
                    BatchUpdateSpreadsheetResponse response = await request.ExecuteAsync();
                }

                // Save the GoogleSheets URL to the list
                listLink.SheetId = currentSpreadSheet.SpreadsheetId;
                listLink.SheetUrl = currentSpreadSheet.SpreadsheetUrl;
                listLink.LastExport = DateTime.UtcNow;
                await listLink.SaveAsync();
            }
        }
        #endregion Update Sheet Async

        #region Create Sheet Async

        private async Task<Spreadsheet> CreateSheetAsync(IComponentListTemplate inList)
        {
            var newspreadSheet = new Spreadsheet()
            {
                Properties = new SpreadsheetProperties()
                {
                    Title = inList.wim.CurrentList.Name
                },
                Sheets = new List<Sheet>()
                {
                    new Sheet()
                    {
                        Properties = new SheetProperties()
                        {
                            Title = inList.wim.CurrentList.Name
                        }
                    }
                },
            };

            
            try
            {
                var createRequest = _sheetsService.Spreadsheets.Create(newspreadSheet);
                var responseCreate = await createRequest.ExecuteAsync();

                if (responseCreate != null)
                {
           
                    Google.Apis.Drive.v3.Data.Permission perms = new Google.Apis.Drive.v3.Data.Permission();
                    perms.Role = "writer";
                    perms.Type = "domain";
                    perms.Domain = "supershift.nl";

                    // Set permissions
                    await _driveService.Permissions.Create(perms, responseCreate.SpreadsheetId).ExecuteAsync();

                    return responseCreate;
                }
            }
            catch (GoogleApiException ex)
            {
                if (ex.HttpStatusCode == System.Net.HttpStatusCode.InternalServerError)
                {
                    // Internal error,
                    return null;
                }
            }

            return null;
        }

        #endregion Create Sheet Async

        #region Execute Module Async

        public async Task<ModuleExecutionResult> ExecuteAsync(IComponentListTemplate inList, IApplicationUser inUser, HttpContext context)
        {
            // When this module is User based, authorize the user
            await AuthorizeUser(inUser);

            // get existing List Link;
            var sheetListLink = await Data.GoogleSheetListLink.FetchSingleAsync(inList.wim.CurrentList.ID, inUser.ID);
            if (sheetListLink == null || sheetListLink.ID == 0)
            {
                sheetListLink = new Data.GoogleSheetListLink()
                {
                    ListID = inList.wim.CurrentList.ID,
                    UserID = inUser.ID,
                };

                await sheetListLink.SaveAsync();
            }

            // Get Sheet ID from db entity
            string sheetId = sheetListLink.SheetId;

            // Check for spreadsheet existence
            var existingSheet = await GetSheetAsync(sheetId);
            if (existingSheet != null)
            {
                await UpdateSheetAsync(inList, existingSheet, sheetListLink);
                return new ModuleExecutionResult()
                {
                    IsSuccess = true,
                    WimNotificationOutput = $"Google Sheet exists : <a href=\"{existingSheet.SpreadsheetUrl}\" target=\"_blank\">View here</a>",
                    RedirectUrl = existingSheet.SpreadsheetUrl
                };
            }
            // No spreadsheet exists
            else 
            {
                var newspreadSheet = await CreateSheetAsync(inList);
                if (newspreadSheet != null)
                {
                    await UpdateSheetAsync(inList, newspreadSheet, sheetListLink);

                    // Spreedsheet created,
                    return new ModuleExecutionResult()
                    {
                        IsSuccess = true,
                        WimNotificationOutput = $"Google Sheet created : <a href=\"{newspreadSheet.SpreadsheetUrl}\" target=\"_blank\">View here</a>",
                        RedirectUrl = newspreadSheet.SpreadsheetUrl
                    };
                }
                else
                {
                    return new ModuleExecutionResult()
                    {
                        IsSuccess = false,
                        WimNotificationOutput = $"Google Sheet could not be created."
                    };
                }
            }
        }

        #endregion Execute Module Async
        
        #region AuthorizeUser

        private async Task AuthorizeUser(IApplicationUser user)
        {
            // Only perform User OAuth when the client secrets file is set
            if (string.IsNullOrWhiteSpace(ClientSecretsFileName))
            {
                return;
            }

            UserCredential credential;

            try
            {
                using (var stream = new FileStream(ClientSecretsFileName, FileMode.Open, FileAccess.Read))
                {
                    credential = GoogleWebAuthorizationBroker.AuthorizeAsync(
                        (await GoogleClientSecrets.FromStreamAsync(stream)).Secrets,
                        new string[] { SheetsService.Scope.Spreadsheets, DriveService.Scope.DriveFile },
                        user.Email,
                        CancellationToken.None,
                        new GoogleTokenStore(user, nameof(GoogleSheetExportListModule))
                        ).Result;

                    Console.WriteLine("Credential file saved to user");
                }

            }
            catch (Exception ex)
            {
                Console.Error.WriteLine(ex);
                throw ex;
            }

            // Create Google Sheets API service.
            _sheetsService = new SheetsService(new BaseClientService.Initializer()
            {
                HttpClientInitializer = credential,
                ApplicationName = nameof(GoogleSheetExportListModule),
            });

            // Create Google Drive API service.
            _driveService = new DriveService(new BaseClientService.Initializer()
            {
                HttpClientInitializer = credential,
                ApplicationName = nameof(GoogleSheetExportListModule),
            });

        }

        #endregion AuthorizeUser

        #region Show On List
        public bool ShowOnList(IComponentListTemplate inList, IApplicationUser inUser)
        {
            return true;
        }

        #endregion Show On List
    }
}