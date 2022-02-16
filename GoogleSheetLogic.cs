using Google;
using Google.Apis.Auth.OAuth2;
using Google.Apis.Drive.v3;
using Google.Apis.Services;
using Google.Apis.Sheets.v4;
using Google.Apis.Sheets.v4.Data;
using Sushi.Mediakiwi.Data;
using Sushi.Mediakiwi.Framework;
using Sushi.Mediakiwi.Framework.EventArguments;

namespace Sushi.Mediakiwi.Module.GoogleSheetsSync
{
    internal class GoogleSheetLogic
    {
        private SheetsService _sheetsService { get; set; }
        private DriveService _driveService { get; set; }
        private string ClientSecretsFileName { get; set; }
        private string ServiceAccountSecretsFileName { get; set; }

        public GoogleSheetLogic(string credentialsFileName, string clientSecretsFileName)
        {
            ServiceAccountSecretsFileName = credentialsFileName;
            ClientSecretsFileName = clientSecretsFileName;
        }

        #region Initialize Async

        public async Task InitializeAsync()
        {
            if (string.IsNullOrWhiteSpace(ServiceAccountSecretsFileName))
            {
                throw new ArgumentNullException("ServiceAccountSecretsFileName", "A Filename should be provided for the serviceAcccount credentials file");
            }

            GoogleCredential credential;

            try
            {
                using (var stream = new FileStream(ServiceAccountSecretsFileName, FileMode.Open, FileAccess.Read))
                {
                    credential = GoogleCredential.FromStream(stream).CreateScoped(new string[] { SheetsService.Scope.Spreadsheets, DriveService.Scope.DriveFile });
                }
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine(ex);
                throw;
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

        #endregion Initialize Async

        #region Get Sheet Async

        public async Task<Spreadsheet> GetSheetAsync(string id)
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

        public async Task UpdateSheetAsync(IComponentListTemplate inList, Spreadsheet currentSpreadSheet, Data.GoogleSheetListLink listLink)
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
                if (currentSheet.ProtectedRanges == null || currentSheet.ProtectedRanges.Count == 0)
                {
                    requests.Add(new Request()
                    {
                        AddProtectedRange = new AddProtectedRangeRequest()
                        {
                            ProtectedRange = new ProtectedRange()
                            {
                                NamedRangeId = namedSheetId,
                                Description = "These columns should match the ones exported from Mediakiwi",
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

        public async Task<Spreadsheet> CreateSheetAsync(IComponentListTemplate inList)
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

        #region Convert Sheet to ListDataReceived Event

        public async Task<(bool success, string errorMessage, ComponentListDataReceived? listEvent)> ConvertSheetToListDataReceivedEvent(Data.GoogleSheetListLink sheetListLink, IComponentListTemplate inList)
        {
            // Collected Values Container
            List<Dictionary<string, object>> CollectedValues = new List<Dictionary<string, object>>();

            // Check if the link exists
            if (sheetListLink?.ID > 0)
            {
                // Get Sheet ID from db entity
                string sheetId = sheetListLink.SheetId;

                // Check for spreadsheet existence
                var existingSheet = await GetSheetAsync(sheetId);
                if (existingSheet != null)
                {
                    var currentSheet = existingSheet.Sheets.FirstOrDefault();
                    if (currentSheet?.Data?.Count > 0)
                    {
                        // Get TargetType
                        string fullTypeName = currentSheet?.DeveloperMetadata?.FirstOrDefault(x => x.MetadataKey.Equals("valueType", StringComparison.InvariantCulture)).MetadataValue;

                        var sheetData = currentSheet.Data.FirstOrDefault();

                        // Lookup table for property names
                        Dictionary<int, string> ColumnPropertyName = new Dictionary<int, string>();

                        // Lookup table for property types
                        Dictionary<int, Type> ColumnPropertyType = new Dictionary<int, Type>();

                        // First get a column index -> propertyName mapping
                        foreach (var metaData in sheetData?.ColumnMetadata?.Where(x => x.DeveloperMetadata?.Count > 0))
                        {
                            foreach (var devMetaData in metaData.DeveloperMetadata?.Where(x => x.MetadataKey.Equals("propertyName", StringComparison.InvariantCultureIgnoreCase)))
                            {
                                var targetIdx = devMetaData.Location.DimensionRange.StartIndex.GetValueOrDefault(1) - 1;
                                if (ColumnPropertyName.ContainsKey(targetIdx) == false)
                                {
                                    ColumnPropertyName[devMetaData.Location.DimensionRange.StartIndex.GetValueOrDefault(1) - 1] = devMetaData.MetadataValue;
                                }
                            }
                        }

                        // Then get a column index -> propertyType mapping
                        foreach (var metaData in sheetData?.ColumnMetadata?.Where(x => x.DeveloperMetadata?.Count > 0))
                        {
                            foreach (var devMetaData in metaData.DeveloperMetadata?.Where(x => x.MetadataKey.Equals("propertyType", StringComparison.InvariantCultureIgnoreCase)))
                            {
                                var targetIdx = devMetaData.Location.DimensionRange.StartIndex.GetValueOrDefault(1) - 1;
                                if (ColumnPropertyType.ContainsKey(targetIdx) == false)
                                {
                                    ColumnPropertyType[devMetaData.Location.DimensionRange.StartIndex.GetValueOrDefault(1) - 1] = Type.GetType(devMetaData.MetadataValue);
                                }
                            }
                        }


                        // Loop through rows, skipping the first (header) row
                        foreach (var row in sheetData.RowData.Skip(1))
                        {
                            Dictionary<string, object> props = new Dictionary<string, object>();

                            foreach (var cellValue in row.Values)
                            {
                                int lookupColidx = row.Values.IndexOf(cellValue);
                                object value = null;

                                if (cellValue?.EffectiveValue?.NumberValue.HasValue == true && cellValue?.EffectiveFormat?.NumberFormat?.Type?.Equals("DATE_TIME", StringComparison.InvariantCulture) == true)
                                {
                                    value = DateTime.FromOADate(cellValue.EffectiveValue.NumberValue.Value);
                                }
                                else if (cellValue?.EffectiveValue?.NumberValue.HasValue == true)
                                {
                                    value = cellValue.EffectiveValue.NumberValue.Value;
                                }
                                else if (cellValue?.EffectiveValue?.BoolValue.HasValue == true)
                                {
                                    value = cellValue.EffectiveValue.BoolValue.Value;
                                }
                                else if (string.IsNullOrWhiteSpace(cellValue?.EffectiveValue?.StringValue) == false)
                                {
                                    value = cellValue.EffectiveValue.StringValue;
                                }

                                value = ConvertValue(value, ColumnPropertyType[lookupColidx]);
                                props.Add(ColumnPropertyName[lookupColidx], value);
                            }

                            CollectedValues.Add(props);
                        }

                        return
                        (
                             success: true,
                             errorMessage: "",
                             listEvent: new ComponentListDataReceived()
                             {
                                 DataSource = nameof(GoogleSheetsImportListModule),
                                 ReceivedProperties = CollectedValues,
                                 FullTypeName = fullTypeName
                             }
                         );
                    }
                }
                
                return
                (
                    success: false,
                    errorMessage: "Could not retrieve Google Sheets data",
                    listEvent: null
                );
            }
            else
            {
                return
                (
                    success: false,
                    errorMessage: "Could not find a link to a Google Sheets file",
                    listEvent: null
                );
            }
        }
        #endregion Convert Sheet to ListDataReceived Event

        #region ConvertValue

        /// <summary>
        /// Converts the inputvalue to the supplied targetType
        /// </summary>
        /// <param name="inputValue">The value to convert</param>
        /// <param name="targetType">The output type to convert to</param>
        /// <returns>An object of type <c>targetType</c> or NULL when an empty value is supplied or the conversion failed</returns>
        private object ConvertValue(object inputValue, Type targetType)
        {
            object result = null;

            try
            {
                if (inputValue == null || inputValue == DBNull.Value)
                {
                    return result;
                }

                if (targetType.IsEnum == false)
                {
                    result = Convert.ChangeType(inputValue, targetType);
                }
            }
            catch
            {

            }

            return result;
        }

        #endregion ConvertValue

        #region AuthorizeUser

        public async Task AuthorizeUser(IApplicationUser user)
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
    }
}
