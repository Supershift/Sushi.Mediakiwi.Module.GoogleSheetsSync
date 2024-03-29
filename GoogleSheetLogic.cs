﻿using Google;
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
    public class GoogleSheetLogic
    {
        private const int MAX_DEVELOPERDATA_SIZE = 30000;

        public GoogleSheetsConfig GoogleSheetConfig { get; private set; }
        public SheetsService GoogleSheetsService { get; private set; }
        public DriveService GoogleDriveService { get; private set; }
        public CellFormat UserEnteredFormat = new CellFormat()
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
        };

        public bool IsInitialized { get; set; }

        #region CTor

        public GoogleSheetLogic(IConfiguration configuration)
        {
            GoogleSheetConfig = configuration.GetSection("GoogleSheetsSettings").Get<GoogleSheetsConfig>();

            if (IsInitialized == false)
            {
                if (string.IsNullOrWhiteSpace(GoogleSheetConfig.ServiceAccountFilename) && string.IsNullOrWhiteSpace(GoogleSheetConfig.ClientID) && string.IsNullOrWhiteSpace(GoogleSheetConfig.ClientSecret))
                {
                    throw new ApplicationException("At least one of 'credentialsFileName' or both 'clientId' and 'clientSecret' should be set");
                }

                GoogleCredential credential;

                if (string.IsNullOrWhiteSpace(GoogleSheetConfig.ServiceAccountFilename) == false)
                {
                    if (File.Exists(Path.Combine(AppDomain.CurrentDomain.BaseDirectory, GoogleSheetConfig.ServiceAccountFilename)))
                    {
                        try
                        {
                            using (var stream = new FileStream(GoogleSheetConfig.ServiceAccountFilename, FileMode.Open, FileAccess.Read))
                            {
                                credential = GoogleCredential.FromStream(stream).CreateScoped(new string[] { SheetsService.Scope.Spreadsheets, DriveService.Scope.DriveFile });
                            }
                        }
                        catch (Exception ex)
                        {
                            Notification.InsertOne("GoogleSheetLogic.CTOR", ex.Message);
                            Console.Error.WriteLine(ex);
                            throw;
                        }

                        // Create Google Sheets API service.
                        GoogleSheetsService = new SheetsService(new BaseClientService.Initializer()
                        {
                            HttpClientInitializer = credential,
                            ApplicationName = nameof(GoogleSheetExportListModule),
                        });

                        // Create Google Drive API service.
                        GoogleDriveService = new DriveService(new BaseClientService.Initializer()
                        {
                            HttpClientInitializer = credential,
                            ApplicationName = nameof(GoogleSheetExportListModule),
                        });

                        IsInitialized = true;

                    }
                    else
                    {
                        var errorMessage = $"The supplied file for 'credentialsFileName' ({GoogleSheetConfig.ServiceAccountFilename}) does not exist";
                        Notification.InsertOne("GoogleSheetLogic.CTOR", errorMessage);

                        throw new ApplicationException(errorMessage);
                    }
                }
            }
        }

        #endregion CTor

        #region Get Sheet Async

        public async Task<Spreadsheet> GetSheetAsync(string id)
        {
            if (string.IsNullOrWhiteSpace(id))
            {
                return null;
            }

            try
            {
                var getRequest = new SpreadsheetsResource.GetRequest(GoogleSheetsService, id);
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
            Dictionary<string, string> temp = new();

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

        public async Task<string> UpdateSheetAsync(IComponentListTemplate inList, Spreadsheet currentSpreadSheet, Data.GoogleSheetListLink listLink)
        {
            string returnMessage = "";
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

                // When the call is made from the MKAPI, clear the grid data first.
                if (inList.wim.IsMKApiCall == true)
                {
                    inList.wim.ClearGridData();
                }

                if (inList.wim.Console.Component == null)
                {
                    inList.wim.Console.Component = new Beta.GeneratedCms.Source.Component();
                }

                await inList.wim.Console.Component.CreateSearchListAsync(inList.wim.Console, 0);

                // Get first sheet
                var currentSheet = currentSpreadSheet.Sheets.FirstOrDefault();

                // Get all columns without the API ones.
                var listDataColumns = inList.wim.ListDataColumns.List.Where(x => x.Type != ListDataColumnType.APIOnly).ToList();

                // All Additional Requests
                List<Request> requests = new();

                // All late requests
                List<Request> lateRequests = new();

                #region Remove all developer info

                // Remove all existing protected ranges
                // Remove all existing named ranges
                if (currentSheet.ProtectedRanges?.Count > 0)
                {
                    foreach (var protectedRange in currentSheet.ProtectedRanges)
                    {
                        requests.Add(new Request
                        {
                            DeleteNamedRange = new DeleteNamedRangeRequest
                            {
                                NamedRangeId = protectedRange.NamedRangeId
                            }
                        });

                        requests.Add(new Request
                        {
                            DeleteProtectedRange = new DeleteProtectedRangeRequest
                            {
                                ProtectedRangeId = protectedRange.ProtectedRangeId
                            }
                        });
                    }
                }

                // Delete all developer metadata
                requests.Add(new Request
                {
                    DeleteDeveloperMetadata = new DeleteDeveloperMetadataRequest
                    {
                        DataFilter = new DataFilter
                        {
                            DeveloperMetadataLookup = new DeveloperMetadataLookup
                            {
                                MetadataKey = "valueType"
                            }
                        }
                    }
                });

                requests.Add(new Request
                {
                    DeleteDeveloperMetadata = new DeleteDeveloperMetadataRequest
                    {
                        DataFilter = new DataFilter
                        {
                            DeveloperMetadataLookup = new DeveloperMetadataLookup
                            {
                                MetadataKey = "propertyName"
                            }
                        }
                    }
                });

                requests.Add(new Request
                {
                    DeleteDeveloperMetadata = new DeleteDeveloperMetadataRequest
                    {
                        DataFilter = new DataFilter
                        {
                            DeveloperMetadataLookup = new DeveloperMetadataLookup
                            {
                                MetadataKey = "propertyType"
                            }
                        }
                    }
                });

                requests.Add(new Request
                {
                    DeleteDeveloperMetadata = new DeleteDeveloperMetadataRequest
                    {
                        DataFilter = new DataFilter
                        {
                            DeveloperMetadataLookup = new DeveloperMetadataLookup
                            {
                                MetadataKey = "propertyIsKey"
                            }
                        }
                    }
                });

                requests.Add(new Request
                {
                    DeleteDeveloperMetadata = new DeleteDeveloperMetadataRequest
                    {
                        DataFilter = new DataFilter
                        {
                            DeveloperMetadataLookup = new DeveloperMetadataLookup
                            {
                                MetadataKey = "rowHash"
                            }
                        }
                    }
                });

                #endregion Remove all developer info

                #region Remove existing cells

                requests.Add(new Request
                {
                    UpdateCells = new UpdateCellsRequest
                    {
                        Range = new GridRange
                        {
                            SheetId = currentSheet.Properties.SheetId,
                            StartRowIndex = 0,
                            EndRowIndex = inList.wim.ListDataCollection.Count + 1
                            
                        },
                        Fields = "userEnteredValue", 
                    }
                });

                #endregion Remove existing cells

                // Check if we have a named range already
                var namedSheetId = "MK.Columns";

                // Add named range
                requests.Add(new Request
                {
                    AddNamedRange = new AddNamedRangeRequest
                    {
                        NamedRange = new NamedRange
                        {
                            Name = "MK.Columns",
                            Range = new GridRange
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
                requests.Add(new Request
                {
                    AddProtectedRange = new AddProtectedRangeRequest
                    {
                        ProtectedRange = new ProtectedRange
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
                RowData headerRow = new()
                {
                    Values = new List<CellData>()
                };


                int colidx = 0;

                // Add header columns
                foreach (var col in listDataColumns)
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
                        UserEnteredFormat = UserEnteredFormat,
                    });

                    colidx++;
                }

                // Expand the sheet to the correct amount of rows
                requests.Add(new Request
                {
                    UpdateSheetProperties = new UpdateSheetPropertiesRequest
                    {
                        Fields = "*",
                        Properties = new SheetProperties
                        {
                            GridProperties = new GridProperties
                            {
                                RowCount = inList.wim.ListDataCollection.Count + 1,
                                ColumnCount = listDataColumns.Count + 1,
                            },
                            SheetId = currentSheet.Properties.SheetId,
                            Title = inList.wim.ListTitle
                        }
                    }
                });

                // Add header row
                rowsData.Add(headerRow);

                List<DeveloperMetadata> devMetaData = new();

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

                // Get all Properties from the object
                var listItemPropertyTypes = listItemType.GetProperties().Select(Property => new { Property.Name, Property }).ToDictionary(t => t.Name, t => t.Property);
   

                // Add Developer MetaData for Header Columns
                foreach (var col in listDataColumns)
                {
                    var colIdx = listDataColumns.IndexOf(col);

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
                        Type propertyType = listItemPropertyTypes[col.ColumnValuePropertyName].PropertyType;
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
                    requests.Add(new Request
                    {
                        CreateDeveloperMetadata = new CreateDeveloperMetadataRequest
                        {
                            DeveloperMetadata = _developerMetadata,
                        }
                    });
                }

                List<List<object>> itemsTempData = new();

                var dataEnumerator = inList.wim.ListDataCollection.GetEnumerator();
                while (dataEnumerator.MoveNext())
                {
                    List<object> tmp = listDataColumns.Select(x => FastReflectionHelper.GetProperty(dataEnumerator.Current, x.ColumnValuePropertyName)).ToList();
                    itemsTempData.Add(tmp);
                }

                foreach (var rowDataEnum in itemsTempData)
                {
                    var rowData = new RowData()
                    {
                        Values = new List<CellData>()
                    };

                    // Loop through items in the search list
                    foreach (var rawData in rowDataEnum)
                    {
                        // Create new cell data for column value
                        var cellData = new CellData();

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
                requests.Add(new Request
                {
                    UpdateCells = new UpdateCellsRequest
                    {
                        Rows = rowsData,
                        Fields = "*",
                        Start = new GridCoordinate
                        {
                            ColumnIndex = 0,
                            RowIndex = 0,
                            SheetId = currentSheet.Properties.SheetId
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

                // Determine if devdata can be added, depending on size
                int sizeLeft = MAX_DEVELOPERDATA_SIZE;
                sizeLeft -= requests.Where(x => x.CreateDeveloperMetadata?.DeveloperMetadata != null).Sum(x => x.CreateDeveloperMetadata.DeveloperMetadata.MetadataKey.Length);
                sizeLeft -= requests.Where(x => x.CreateDeveloperMetadata?.DeveloperMetadata != null).Sum(x => x.CreateDeveloperMetadata.DeveloperMetadata.MetadataValue.Length);
                sizeLeft -= lateRequests.Where(x => x.CreateDeveloperMetadata?.DeveloperMetadata != null).Sum(x => x.CreateDeveloperMetadata.DeveloperMetadata.MetadataKey.Length);
                sizeLeft -= lateRequests.Where(x => x.CreateDeveloperMetadata?.DeveloperMetadata != null).Sum(x => x.CreateDeveloperMetadata.DeveloperMetadata.MetadataValue.Length);

                // The developer metadata is too big, remove the rowhash
                if (sizeLeft < 1)
                {
                    lateRequests.RemoveAll(x => x.CreateDeveloperMetadata?.DeveloperMetadata?.MetadataKey.Equals("rowHash") == true);
                    returnMessage = "Change tracking is disabled, too much data to keep track off.";
                }

                // Run all requests
                if (requests?.Count > 0)
                {
                    BatchUpdateSpreadsheetRequest requestBody = new();
                    requestBody.Requests = requests.Concat(lateRequests).ToList();
                    SpreadsheetsResource.BatchUpdateRequest request = GoogleSheetsService.Spreadsheets.BatchUpdate(requestBody, currentSpreadSheet.SpreadsheetId);
                    BatchUpdateSpreadsheetResponse response = await request.ExecuteAsync();
                }

                // Save the GoogleSheets URL to the list
                listLink.SheetId = currentSpreadSheet.SpreadsheetId;
                listLink.SheetUrl = currentSpreadSheet.SpreadsheetUrl;
                listLink.LastExport = DateTime.UtcNow;
                await listLink.SaveAsync();
            }

            return returnMessage;
        }
        #endregion Update Sheet Async

        #region Create Sheet Async

        public async Task<Spreadsheet> CreateSheetAsync(string title, string userEmailAddress)
        {
            var newspreadSheet = new Spreadsheet()
            {
                Properties = new SpreadsheetProperties()
                {
                    Title = title
                },
                Sheets = new List<Sheet>()
                {
                    new Sheet()
                    {
                        Properties = new SheetProperties()
                        {
                            Title = title
                        }
                    }
                },
            };


            try
            {
                var createRequest = GoogleSheetsService.Spreadsheets.Create(newspreadSheet);
                var responseCreate = await createRequest.ExecuteAsync();

                if (responseCreate != null)
                {
                    if (GoogleSheetConfig?.PermissionBase == GoogleSheetsPermissionsEnum.USEREMAIL)
                    {
                        // Add all tasks to array
                        try
                        {
                            await CreatePermission(userEmailAddress, string.Empty, responseCreate.SpreadsheetId);
                        }
                        catch (GoogleApiException ex)
                        {
                            await Notification.InsertOneAsync("GoogleSheetLogic.CreateSheetAsync", ex);
                        }
                    }
                    else if (GoogleSheetConfig?.PermissionBase == GoogleSheetsPermissionsEnum.USERDOMAIN)
                    {
                        // Add all tasks to array
                        try
                        {
                            await CreatePermission(string.Empty, userEmailAddress.Substring(userEmailAddress.IndexOf('@') + 1), responseCreate.SpreadsheetId);
                        }
                        catch (GoogleApiException ex)
                        {
                            await Notification.InsertOneAsync("GoogleSheetLogic.CreateSheetAsync", ex);
                        }
                    }
                    else if (GoogleSheetConfig.PermissionBase == GoogleSheetsPermissionsEnum.SETDOMAINS)
                    {
                        List<string> allowedDomains = new List<string>();

                        if (GoogleSheetConfig?.AllowedDomains?.Length > 0)
                        {
                            allowedDomains = GoogleSheetConfig.AllowedDomains.ToList();
                        }
                        else
                        {
                            allowedDomains.Add("supershift.nl");
                        }

                        foreach (var domain in allowedDomains)
                        {
                            // Add all tasks to array
                            try
                            {
                                await CreatePermission(string.Empty, domain, responseCreate.SpreadsheetId);
                            }
                            catch (GoogleApiException ex)
                            {
                                await Notification.InsertOneAsync("GoogleSheetLogic.CreateSheetAsync", ex);
                            }
                        }
                    }

                    return responseCreate;
                }
            }
            catch (GoogleApiException ex)
            {
                if (ex.HttpStatusCode == System.Net.HttpStatusCode.InternalServerError)
                {
                    await Notification.InsertOneAsync("GoogleSheetLogic.CreateSheetAsync", ex);

                    // Internal error,
                    return null;
                }
            }

            return null;
        }

        #endregion Create Sheet Async

        #region Create Permission

        private async Task CreatePermission(string email, string domain, string spreadSheetId)
        {
            Google.Apis.Drive.v3.Data.Permission perms = new()
            {
                Role = "writer"
            };

            if (string.IsNullOrWhiteSpace(email) == false)
            {
                perms.Type = "user";
                perms.EmailAddress = email;
            }
            else
            {
                perms.Type = "domain";
                perms.Domain = domain;
            }

            await GoogleDriveService.Permissions.Create(perms, spreadSheetId).ExecuteAsync();

        }

        #endregion Create Permission

        #region Convert Sheet to ListDataReceived Event

        public async Task<(bool success, string errorMessage, ComponentListDataReceived? listEvent)> ConvertSheetToListDataReceivedEvent(Data.GoogleSheetListLink sheetListLink, IComponentListTemplate inList)
        {
            // Collected Values Container
            List<ComponentListDataReceivedItem> CollectedValues = new();

            // Check if the link exists
            if (sheetListLink?.ID > 0)
            {
                // Get Sheet ID from db entity
                string sheetId = sheetListLink.SheetId;

                // Whenever devdata 'rowHash' is not available for any row, the items are untracked.
                bool containsRowHash = false;

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
                        Dictionary<int, string> ColumnPropertyName = new();

                        // Lookup table for property types
                        Dictionary<int, Type> ColumnPropertyType = new();

                        // Lookup table for value Hashes
                        Dictionary<int, string> originalRowHash = new();

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
                                containsRowHash = true;
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
                            Dictionary<string, object> props = new();

                            // Maintains a dictionary of the cell values which we're included during export.
                            // this is needed to create a hash based on the correct columns
                            List<CellData> includedCells = new();

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
                                ReceivedItemTypeEnum itemType = (containsRowHash)? ReceivedItemTypeEnum.NEW : ReceivedItemTypeEnum.UNTRACKED;

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
            if (string.IsNullOrWhiteSpace(GoogleSheetConfig.ClientID) || string.IsNullOrWhiteSpace(GoogleSheetConfig.ClientSecret))
            {
                return null;
            }

            if (string.IsNullOrWhiteSpace(GoogleSheetConfig?.HandlerPath))
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
                             ClientId = GoogleSheetConfig.ClientID,
                             ClientSecret = GoogleSheetConfig.ClientSecret
                         },
                         DataStore = new GoogleTokenStore(user, nameof(GoogleSheetLogic)),
                         Scopes = new string[] 
                         { 
                             SheetsService.Scope.Spreadsheets, 
                             DriveService.Scope.DriveFile 
                         }
                     }
                );

                var userId = nameof(GoogleSheetLogic);
                var uri = $"{context.Request.Scheme}://{context.Request.Host}{context.Request.PathBase}{GoogleSheetConfig.HandlerPath}";
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
                        GoogleSheetsService = new SheetsService(new BaseClientService.Initializer()
                        {
                            HttpClientInitializer = result.Credential,
                            ApplicationName = nameof(GoogleSheetExportListModule),
                        });

                        // Create Google Drive API service.
                        GoogleDriveService = new DriveService(new BaseClientService.Initializer()
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
