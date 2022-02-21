using Google;
using Google.Apis.Auth.OAuth2;
using Google.Apis.Auth.OAuth2.Flows;
using Google.Apis.Auth.OAuth2.Web;
using Google.Apis.Drive.v3;
using Google.Apis.Services;
using Google.Apis.Sheets.v4;
using Google.Apis.Sheets.v4.Data;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Http.Extensions;
using Microsoft.Extensions.Configuration;
using Sushi.Mediakiwi.Data;
using Sushi.Mediakiwi.Framework;
using Sushi.Mediakiwi.Framework.EventArguments;

namespace Sushi.Mediakiwi.Module.GoogleSheetsSync
{
    internal class GoogleSheetLogic
    {
        private SheetsService _sheetsService { get; set; }
        private DriveService _driveService { get; set; }

        public bool IsInitialized { get; set; }

        private GoogleSheetsConfig _config { get; set; }
        public GoogleSheetLogic(IConfiguration configuration)
        {
            _config = configuration.GetSection("GoogleSheetsSettings").Get<GoogleSheetsConfig>();
        }

        #region Initialize Async

        public async Task InitializeAsync()
        {
            if (IsInitialized == false)
            {
                if (string.IsNullOrWhiteSpace(_config.ServiceAccountFilename) && string.IsNullOrWhiteSpace(_config.ClientID) && string.IsNullOrWhiteSpace(_config.ClientSecret))
                {
                    throw new ApplicationException("At least one of 'credentialsFileName' or both 'clientId' and 'clientSecret' should be set");
                }

                GoogleCredential credential;

                if (string.IsNullOrWhiteSpace(_config.ServiceAccountFilename) == false)
                {
                    if (File.Exists(Path.Combine(AppDomain.CurrentDomain.BaseDirectory, _config.ServiceAccountFilename)))
                    {
                        try
                        {
                            using (var stream = new FileStream(_config.ServiceAccountFilename, FileMode.Open, FileAccess.Read))
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

                        IsInitialized = true;

                    }
                    else 
                    {
                        throw new ApplicationException($"The supplied file for 'credentialsFileName' ({_config.ServiceAccountFilename}) does not exist");
                    }
                }
            }
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

        #region Get DropdownValues Async

        private async Task<Dictionary<string, string>> GetDropdownValuesAsync(IComponentListTemplate inList, string propertyName)
        {
            Dictionary<string, string> temp = new Dictionary<string, string>();

            if (inList is ComponentListTemplate template)
            {
                var property = inList.GetType().GetProperty(propertyName);
                if (property != null)
                {
                    foreach (var att in property.GetCustomAttributes(true))
                    {
                        if (att is Framework.ContentListItem.Choice_DropdownAttribute dropdownAtt)
                        {
                            dropdownAtt.SenderInstance = template;
                            dropdownAtt.Property = property;

                            var apiField = await dropdownAtt.GetApiFieldAsync();
                            if (apiField?.Options?.Count > 0)
                            {
                                foreach (var option in apiField.Options)
                                {
                                    temp.Add(option.Value, option.Text);
                                }
                            }
                        }
                        else if (att is Framework.ContentListItem.Choice_RadioAttribute radioAtt)
                        {
                            radioAtt.SenderInstance = template;
                            radioAtt.Property = property;

                            var apiField = await radioAtt.GetApiFieldAsync();
                            if (apiField?.Options?.Count > 0)
                            {
                                foreach (var option in apiField.Options)
                                {
                                    temp.Add(option.Value, option.Text);
                                }
                            }
                        }
                    }
                }

            }

            return temp;
        }

        #endregion Get DropdownValues Async

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

                // All late requests
                List<Request> lateRequests = new List<Request>();

                #region Remove all developer info

                // Remove all existing protected ranges
                // Remove all existing named ranges
                if (currentSheet.ProtectedRanges?.Count > 0)
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

                #endregion Remove all developer info

                // Check if we have a named range already
                var namedSheetId = "MK.Columns";

                // Add named range
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


                // Set first row as protected
                requests.Add(new Request()
                {
                    AddProtectedRange = new AddProtectedRangeRequest()
                    {
                        ProtectedRange = new ProtectedRange()
                        {
                            NamedRangeId = namedSheetId,
                            Description = "These columns should match the ones exported from Mediakiwi",
                            WarningOnly = false
                        },
                    }
                });


                // Complete data collection
                IList<RowData> rowsData = new List<RowData>();

                // Add column headers
                RowData headerRow = new RowData()
                {
                    Values = new List<CellData>()
                };


                int colidx = 0;
                // Add header columns
                foreach (var col in inList.wim.ListDataColumns.List)
                {
                    var items = await GetDropdownValuesAsync(inList, col.ColumnValuePropertyName);
                    if (items?.Count > 0)
                    {
                        lateRequests.Add(new Request()
                        {
                            SetDataValidation = new SetDataValidationRequest()
                            {
                                Range = new GridRange()
                                {
                                    StartColumnIndex = colidx,
                                    EndColumnIndex = colidx + 1,
                                    StartRowIndex = 1,
                                    SheetId = currentSheet.Properties.SheetId
                                },
                                Rule = new DataValidationRule()
                                {
                                    Condition = new BooleanCondition()
                                    {
                                        Type = "ONE_OF_LIST",
                                        Values = items.Select(x => new ConditionValue()
                                        {
                                            UserEnteredValue = x.Key
                                        }).ToList()
                                    },
                                    ShowCustomUi = true,
                                    Strict = true,
                                }
                            }
                        });
                    }

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
                    colidx++;
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
                    var colIdx = inList.wim.ListDataColumns.List.IndexOf(col);

                    devMetaData.Add(new DeveloperMetadata()
                    {
                        Location = new DeveloperMetadataLocation()
                        {
                            DimensionRange = new DimensionRange()
                            {
                                StartIndex = colIdx + 1,
                                EndIndex = colIdx + 2,
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
                        Type propertyType = listItemType.GetProperty(col.ColumnValuePropertyName).PropertyType;
                        devMetaData.Add(new DeveloperMetadata()
                        {
                            Location = new DeveloperMetadataLocation()
                            {
                                DimensionRange = new DimensionRange()
                                {
                                    StartIndex = colIdx + 1,
                                    EndIndex = colIdx + 2,
                                    SheetId = currentSheet.Properties.SheetId,
                                    Dimension = "COLUMNS"
                                },
                            },
                            MetadataKey = "propertyType",
                            MetadataValue = propertyType.FullName,
                            Visibility = "DOCUMENT",
                        });

                        if (propertyType == typeof(bool) || propertyType == typeof(bool?))
                        {
                            lateRequests.Add(new Request()
                            {
                                RepeatCell = new RepeatCellRequest()
                                {
                                    Cell = new CellData()
                                    {
                                        DataValidation = new DataValidationRule()
                                        {
                                            Condition = new BooleanCondition()
                                            {
                                                Type = "BOOLEAN"
                                            }
                                        }
                                    },
                                    Range = new GridRange()
                                    {
                                        SheetId = currentSheet.Properties.SheetId,
                                        StartColumnIndex = colIdx,
                                        EndColumnIndex = colIdx + 1,
                                        StartRowIndex = 1,
                                    },
                                    Fields = "dataValidation"
                                }
                            });
                        }
                        else if (propertyType == typeof(DateTime) || propertyType == typeof(DateTime?))
                        {
                            lateRequests.Add(new Request()
                            {
                                RepeatCell = new RepeatCellRequest()
                                {
                                    Cell = new CellData()
                                    {
                                        DataValidation = new DataValidationRule()
                                        {
                                            Condition = new BooleanCondition()
                                            {
                                                Type = "DATE_IS_VALID"
                                            }
                                        }
                                    },
                                    Range = new GridRange()
                                    {
                                        SheetId = currentSheet.Properties.SheetId,
                                        StartColumnIndex = colIdx,
                                        EndColumnIndex = colIdx + 1,
                                        StartRowIndex = 1,
                                    },
                                    Fields = "dataValidation"
                                }
                            });
                        }
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
                            if (string.IsNullOrWhiteSpace(rawString) == false)
                            {
                                cellData.UserEnteredValue = new ExtendedValue()
                                {
                                    StringValue = rawString,
                                };
                            }
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

                    // Get row Index for inserting hash
                    int rowIdx = rowsData.IndexOf(rowData);

                    lateRequests.Add(new Request()
                    {
                        CreateDeveloperMetadata = new CreateDeveloperMetadataRequest()
                        {
                            DeveloperMetadata = new DeveloperMetadata()
                            {
                                Location = new DeveloperMetadataLocation()
                                {
                                    DimensionRange = new DimensionRange()
                                    {
                                        Dimension = "ROWS",
                                        EndIndex = rowIdx + 1,
                                        StartIndex = rowIdx,
                                        SheetId = currentSheet.Properties.SheetId
                                    }
                                },
                                MetadataKey = "rowHash",
                                MetadataValue = GoogleValueHasher.CreateHash(rowData.Values),
                                Visibility = "DOCUMENT",
                            }
                        }
                    });
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
                lateRequests.Add(new Request()
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
                    requestBody.Requests = requests.Concat(lateRequests).ToList();
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
            List<ComponentListDataReceivedItem> CollectedValues = new List<ComponentListDataReceivedItem>();

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

                        // Lookup table for value Hashes
                        Dictionary<int, string> originalRowHash = new Dictionary<int, string>();

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

                        // Then get a row index -> hash mapping
                        foreach (var metaData in sheetData?.RowMetadata?.Where(x => x.DeveloperMetadata?.Count > 0))
                        {
                            foreach (var devMetaData in metaData.DeveloperMetadata?.Where(x => x.MetadataKey.Equals("rowHash", StringComparison.InvariantCultureIgnoreCase)))
                            {
                                var targetIdx = devMetaData.Location.DimensionRange.StartIndex.GetValueOrDefault(1);
                                if (originalRowHash.ContainsKey(targetIdx) == false)
                                {
                                    originalRowHash[devMetaData.Location.DimensionRange.StartIndex.GetValueOrDefault(1)] = devMetaData.MetadataValue;
                                }
                            }
                        }


                        // Loop through rows, skipping the first (header) row
                        foreach (var row in sheetData.RowData.Skip(1))
                        {
                            // Get the current row index
                            int rowIdx = sheetData.RowData.IndexOf(row);

                            // Maintains a dictionary of the propertynames, coupled to a value
                            Dictionary<string, object> props = new Dictionary<string, object>();

                            // Maintains a dictionary of the cell values which we're included during export.
                            // this is needed to create a hash based on the correct columns
                            List<CellData> includedCells = new List<CellData>();

                            // Only extract values from columns we originally provided.
                            foreach (var cellValue in row.Values.Take(ColumnPropertyName.Count))
                            {
                                int lookupColidx = row.Values.IndexOf(cellValue);

                                // This column was included during export, so must be included in the hash
                                includedCells.Add(cellValue);

                                object value = null;
                                if (cellValue?.EffectiveValue?.NumberValue.HasValue == true && cellValue?.EffectiveFormat?.NumberFormat?.Type?.Equals("DATE_TIME", StringComparison.InvariantCulture) == true)
                                {
                                    var dateTime = DateTime.FromOADate(cellValue.EffectiveValue.NumberValue.Value);
                                    if (dateTime != DateTime.MinValue)
                                    {
                                        value = dateTime;
                                    }
                                    else
                                    {
                                        value = null;
                                    }
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

                            // Omit total empty rows
                            if (props?.Any(x => x.Value != null) == true)
                            {
                                // Determine the type
                                ReceivedItemTypeEnum itemType = ReceivedItemTypeEnum.NEW;

                                // Determine if this type was changed
                                if (originalRowHash.ContainsKey(rowIdx))
                                {
                                    // Create hash from current rows
                                    var currentRowHash = GoogleValueHasher.CreateHash(includedCells);

                                    if (currentRowHash.Equals(originalRowHash[rowIdx]))
                                    {
                                        itemType = ReceivedItemTypeEnum.UNCHANGED;
                                    }
                                    else
                                    {
                                        itemType = ReceivedItemTypeEnum.CHANGED;
                                    }
                                }

                                CollectedValues.Add(new ComponentListDataReceivedItem() 
                                { 
                                    PropertyValues = props,
                                    ItemType = itemType
                                });
                            }
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

        public async Task<bool?> AuthorizeUser(IApplicationUser user, HttpContext context)
        {
            // Only perform User OAuth when the client ID is set
            if (string.IsNullOrWhiteSpace(_config.ClientID) || string.IsNullOrWhiteSpace(_config.ClientSecret))
            {
                return null;
            }

            if (string.IsNullOrWhiteSpace(_config?.HandlerPath))
            {
                return false;
            }

            UserCredential credential;

            try
            {

                IAuthorizationCodeFlow flow = new GoogleAuthorizationCodeFlow(
                     new GoogleAuthorizationCodeFlow.Initializer
                     {
                         ClientSecrets = new ClientSecrets
                         {
                             ClientId = _config.ClientID,
                             ClientSecret = _config.ClientSecret
                         },
                         DataStore = new GoogleTokenStore(user, nameof(GoogleSheetLogic)),
                         Scopes = new string[] { SheetsService.Scope.Spreadsheets, DriveService.Scope.DriveFile }
                     }
                );

                var userId = nameof(GoogleSheetLogic);
                var uri = $"{context.Request.Scheme}://{context.Request.Host}{context.Request.PathBase}{_config.HandlerPath}";
                var code = context.Request.Query["code"].FirstOrDefault();
                if (code != null)
                {
                    var token = await flow.ExchangeCodeForTokenAsync(userId, code, uri.Substring(0, uri.IndexOf("?")), CancellationToken.None);

                    // Extract the right state.
                    var oauthState = await AuthWebUtility.ExtracRedirectFromState(flow.DataStore, userId, context.Request.Query["state"]);
                    await context.Response.WriteAsync($"<script>window.open('{oauthState}','_self');</script>");
                    //context.Response.Redirect(oauthState);
                }
                else
                {
                    var result = await new AuthorizationCodeWebApp(flow, uri, context.Request.GetDisplayUrl()).AuthorizeAsync(userId, CancellationToken.None);

                    if (result.RedirectUri != null)
                    {
                        // Redirect the user to the authorization server.
                        await context.Response.WriteAsync($"<script>window.open('{result.RedirectUri}','_self');</script>");
                        return false;
                    }
                    else
                    {

                        // Create Google Sheets API service.
                        _sheetsService = new SheetsService(new BaseClientService.Initializer()
                        {
                            HttpClientInitializer = result.Credential,
                            ApplicationName = nameof(GoogleSheetExportListModule),
                        });

                        // Create Google Drive API service.
                        _driveService = new DriveService(new BaseClientService.Initializer()
                        {
                            HttpClientInitializer = result.Credential,
                            ApplicationName = nameof(GoogleSheetExportListModule),
                        });
                    }
                }

                Console.WriteLine("Google OpenID set");
                return true;
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine(ex);
                throw ex;
            }
            return false;
        }

        #endregion AuthorizeUser
    }
}
