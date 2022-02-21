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
    internal class GoogleSheetsViewModule : IListModule
    {
        #region Properties

        public string ModuleTitle => "Google Sheets Viewer";

        public bool ShowInSearchMode { get; set; } = true;

        public bool ShowInEditMode { get; set; } = false;

        public string IconClass { get; set; } = "icon-cloud";

        public string IconURL { get; set; }

        public string Tooltip { get; set; } = "View this list in Google Sheets";

        public bool ConfirmationNeeded { get; set; }

        public string ConfirmationTitle { get; set; }

        public string ConfirmationQuestion { get; set; }

        public async Task<ModuleExecutionResult> ExecuteAsync(IComponentListTemplate inList, IApplicationUser inUser, HttpContext context)
        {
            var listLink = Task.Run(async () => await Data.GoogleSheetListLink.FetchSingleAsync(inList.wim.CurrentList.ID, inUser.ID)).Result;
            //inList.wim.Redirect(listLink.SheetUrl);
            return new ModuleExecutionResult()
            {
                IsSuccess = true,
                RedirectUrl = listLink.SheetUrl
            };
        }

        public bool ShowOnList(IComponentListTemplate inList, IApplicationUser inUser)
        {
            var listLink = Task.Run(async () => await Data.GoogleSheetListLink.FetchSingleAsync(inList.wim.CurrentList.ID, inUser.ID)).Result;
            var hasListLink = string.IsNullOrWhiteSpace(listLink?.SheetUrl) == false;

            return hasListLink;
        }

        #endregion Properties
    }
}
