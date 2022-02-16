﻿using Microsoft.AspNetCore.Http;
using Microsoft.Extensions.DependencyInjection;
using Sushi.Mediakiwi.Data;
using Sushi.Mediakiwi.Framework;
using Sushi.Mediakiwi.Framework.Interfaces;

namespace Sushi.Mediakiwi.Module.GoogleSheetsSync
{
    public class GoogleSheetExportListModule : IListModule
    {
        #region Properties

        GoogleSheetLogic Converter { get; set; }

        public string ModuleTitle => "Google Sheets exporter";

        public bool ShowInSearchMode { get; set; } = true;

        public bool ShowInEditMode { get; set; } = false;

        public string IconClass { get; set; } = "icon-cloud-upload";

        public string IconURL { get; set; }

        public string Tooltip { get; set; } = "View this list in Google Sheets";

        public bool ConfirmationNeeded { get; set; }

        public string ConfirmationTitle { get; set; }

        public string ConfirmationQuestion { get; set; }

        #endregion Properties

        #region CTor

        public GoogleSheetExportListModule(IServiceProvider services)
        {
            Converter = services.GetService<GoogleSheetLogic>();
        }

        #endregion CTor

        #region Execute Module Async

        public async Task<ModuleExecutionResult> ExecuteAsync(IComponentListTemplate inList, IApplicationUser inUser, HttpContext context)
        {
            // When this module is User based, authorize the user
            await Converter.AuthorizeUser(inUser);

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
            var existingSheet = await Converter.GetSheetAsync(sheetId);
            if (existingSheet != null)
            {
                await Converter.UpdateSheetAsync(inList, existingSheet, sheetListLink);
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
                var newspreadSheet = await Converter.CreateSheetAsync(inList);
                if (newspreadSheet != null)
                {
                    await Converter.UpdateSheetAsync(inList, newspreadSheet, sheetListLink);

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

        #region Show On List

        public bool ShowOnList(IComponentListTemplate inList, IApplicationUser inUser)
        {
            return true;
        }

        #endregion Show On List
    }
}