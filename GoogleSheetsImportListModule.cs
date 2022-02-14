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
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Sushi.Mediakiwi.Module.GoogleSheetsSync
{
    internal class GoogleSheetsImportListModule : IListModule
    {
        #region Properties

        public string ModuleTitle => "GoogleSheets importer";
        public bool ShowInSearchMode { get; set; }
        public bool ShowInEditMode { get; set; }
        public string IconClass { get; set; }
        public string IconURL { get; set; }
        public string Tooltip { get; set; }
        public bool ConfirmationNeeded { get; set; }
        public string ConfirmationTitle { get; set; }
        public string ConfirmationQuestion { get; set; }

        private SheetsService _sheetsService { get; set; }
        private DriveService _driveService { get; set; }

        private string ClientSecretsFileName { get; set; }

        #endregion Properties

        #region CTor

        public GoogleSheetsImportListModule()
        {
            ShowInSearchMode = true;
            ShowInEditMode = false;
            Tooltip = "Import data from GoogleSheets";
            IconClass = "icon-file-text-o";
            ConfirmationNeeded = true;
            ConfirmationTitle = "Are you sure ?";
            ConfirmationQuestion = "This will overwrite the data in this list with the values entered in GoogleSheets";
        }

        #endregion CTor

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

        public async Task<ModuleExecutionResult> ExecuteAsync(IComponentListTemplate inList, IApplicationUser inUser, HttpContext context)
        {
            await AuthorizeUser(inUser);
            string sheetId = "";
            int rowCount = 0;

            if (inList?.wim?.CurrentList?.Settings?.HasProperty("GoogleSheetID") == true)
            {
                sheetId = inList.wim.CurrentList.Settings["GoogleSheetID"].Value;
            }

            //if (string.IsNullOrWhiteSpace(sheetId) == false)
            //{
            //    var spreadSheet = await GetSheetAsync(sheetId);
            //    var firstSheet = spreadSheet.Sheets.FirstOrDefault();
            //    if (firstSheet != null)
            //    {
            //        foreach (var row in firstSheet.Data)
            //        {
            //            rowCount++;
            //            foreach (var cell in row.ColumnMetadata[1].DeveloperMetadata.)
            //            {

            //            }
            //        }
            //    }
            //}

            return new ModuleExecutionResult()
            {
                IsSuccess = true,
                WimNotificationOutput = $"Successfully updated {rowCount} rows."
            };
        }

        #region Authorize User

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
        #endregion Authorize User

        #region Show on List

        public bool ShowOnList(IComponentListTemplate inList, IApplicationUser inUser)
        {
            string sheetId = "";
            if (inList?.wim?.CurrentList?.Settings?.HasProperty("GoogleSheetID") == true)
            {
                sheetId = inList.wim.CurrentList.Settings["GoogleSheetID"].Value;
            }

            return string.IsNullOrWhiteSpace(sheetId) == false;
        }

        #endregion Show on List
    }
}
