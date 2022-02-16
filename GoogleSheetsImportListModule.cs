using Google;
using Google.Apis.Auth.OAuth2;
using Google.Apis.Drive.v3;
using Google.Apis.Services;
using Google.Apis.Sheets.v4;
using Google.Apis.Sheets.v4.Data;
using Microsoft.AspNetCore.Http;
using Sushi.Mediakiwi.Data;
using Sushi.Mediakiwi.Framework;
using Sushi.Mediakiwi.Framework.EventArguments;
using Sushi.Mediakiwi.Framework.Interfaces;

namespace Sushi.Mediakiwi.Module.GoogleSheetsSync
{

    internal class GoogleSheetsColumnValue
    {
        public string PropertyName { get; set; }
        public bool PropertyIsKey { get; set; }
        public object Value { get; set; }
        public GoogleSheetsColumnValue() { }

        public GoogleSheetsColumnValue(string propertyName, object value, bool isKey)
        {
            PropertyName = propertyName;
            Value = value;
            PropertyIsKey = isKey;
        }

        public GoogleSheetsColumnValue(string propertyName, object value)
        {
            PropertyName = propertyName;
            Value = value;
        }
    }

    public class GoogleSheetsImportListModule : IListModule
    {
        #region Properties

        public string ModuleTitle => "GoogleSheets exporter";

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

        public GoogleSheetsImportListModule()
        {
            ShowInSearchMode = true;
            ShowInEditMode = false;
            Tooltip = "Import data from GoogleSheets";
            IconClass = "icon-cloud-download";
            ConfirmationNeeded = true;
            ConfirmationTitle = "Are you sure ?";
            ConfirmationQuestion = "This will overwrite the data in this list with the values entered in GoogleSheets";
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

                var getSpreadsheetRequest = new SpreadsheetsResource.GetRequest(_sheetsService, id);
                getSpreadsheetRequest.IncludeGridData = true;

                var responseGetSpreadsheet = await getSpreadsheetRequest.ExecuteAsync();
                if (responseGetSpreadsheet != null)
                {
                    return responseGetSpreadsheet;
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

        #region Execute Module Async

        public async Task<ModuleExecutionResult> ExecuteAsync(IComponentListTemplate inList, IApplicationUser inUser, HttpContext context)
        {
            // When this module is User based, authorize the user
            await AuthorizeUser(inUser);

            // get existing List Link;
            var sheetListLink = await Data.GoogleSheetListLink.FetchSingleAsync(inList.wim.CurrentList.ID, inUser.ID);

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

                        // First get a column index -> propertyname mapping
                        int colIdx = 0;
                        foreach (var metaData in sheetData?.ColumnMetadata?.Where(x => x.DeveloperMetadata?.Count > 0))
                        {
                            foreach (var devMetaData in metaData.DeveloperMetadata?.Where(x => x.MetadataKey.Equals("propertyName", StringComparison.InvariantCultureIgnoreCase)))
                            {
                                var isKey = metaData?.DeveloperMetadata?.Any(x => x.MetadataKey.Equals("propertyIsKey") && x.MetadataValue.Equals("true")) == true;

                                var targetIdx = devMetaData.Location.DimensionRange.StartIndex.GetValueOrDefault(1) - 1;
                                if (ColumnPropertyName.ContainsKey(targetIdx) == false)
                                {
                                    ColumnPropertyName[devMetaData.Location.DimensionRange.StartIndex.GetValueOrDefault(1) - 1] = devMetaData.MetadataValue;
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

                                props.Add(ColumnPropertyName[lookupColidx], value);
                            }

                            CollectedValues.Add(props);
                        }

                        if (CollectedValues?.Count > 0 && inList is ComponentListTemplate template)
                        {
                            await template.OnListDataReceived(new ComponentListDataReceived()
                            {
                                DataSource = nameof(GoogleSheetsImportListModule),
                                ReceivedProperties = CollectedValues,
                                FullTypeName = fullTypeName
                            });
                        }
                    }
                }

                
                return new ModuleExecutionResult()
                {
                    IsSuccess = true,
                    WimNotificationOutput = "Updated"
                };
            }
            else 
            {
                return new ModuleExecutionResult()
                {
                    IsSuccess = false,
                    WimNotificationOutput = "Could not find a link to a Google Sheets file"
                };
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
            var listLink = Task.Run(async () => await Data.GoogleSheetListLink.FetchSingleAsync(inList.wim.CurrentList.ID, inUser.ID)).Result;
            var hasListLink = string.IsNullOrWhiteSpace(listLink?.SheetUrl) == false;

            if (inList is ComponentListTemplate template)
            {
                return (template.HasListDataReceived && hasListLink);
            }
            else 
            {
                return hasListLink;
            }
        }

        #endregion Show On List
    }
}
