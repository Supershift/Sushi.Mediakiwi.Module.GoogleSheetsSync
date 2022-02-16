using Microsoft.AspNetCore.Http;
using Microsoft.Extensions.DependencyInjection;
using Sushi.Mediakiwi.Data;
using Sushi.Mediakiwi.Framework;
using Sushi.Mediakiwi.Framework.EventArguments;
using Sushi.Mediakiwi.Framework.Interfaces;

namespace Sushi.Mediakiwi.Module.GoogleSheetsSync
{

    public class GoogleSheetsImportListModule : IListModule
    {
        #region Properties

        GoogleSheetLogic Converter { get; set; }

        public string ModuleTitle => "Google Sheets importer";

        public bool ShowInSearchMode { get; set; } = true;

        public bool ShowInEditMode { get; set; } = false;

        public string IconClass { get; set; } = "icon-cloud-download";

        public string IconURL { get; set; }

        public string Tooltip { get; set; } = "Import data from Google Sheets";

        public bool ConfirmationNeeded { get; set; } = true;

        public string ConfirmationTitle { get; set; } = "Are you sure ?";
        public string ConfirmationQuestion { get; set; } = "This will overwrite the data in this list with the values entered in Google Sheets";

        #endregion Properties

        #region CTor

        public GoogleSheetsImportListModule(IServiceProvider services)
        {
            Converter = services.GetService<GoogleSheetLogic>();
        }

        #endregion CTor

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

        #region Execute Module Async

        public async Task<ModuleExecutionResult> ExecuteAsync(IComponentListTemplate inList, IApplicationUser inUser, HttpContext context)
        {
            // When this module is User based, authorize the user
            await Converter.AuthorizeUser(inUser);

            // get existing List Link;
            var sheetListLink = await Data.GoogleSheetListLink.FetchSingleAsync(inList.wim.CurrentList.ID, inUser.ID);

            // Convert the sheet to a list Event
            var convertSheetToListEventResult = await Converter.ConvertSheetToListDataReceivedEvent(sheetListLink, inList);

            if (convertSheetToListEventResult.success && inList is ComponentListTemplate template)
            {
                await template.OnListDataReceived(convertSheetToListEventResult.listEvent);

                return new ModuleExecutionResult()
                {
                    IsSuccess = true,
                    WimNotificationOutput = $"Received {convertSheetToListEventResult.listEvent.ReceivedProperties.Count} rows from Google Sheets"
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
