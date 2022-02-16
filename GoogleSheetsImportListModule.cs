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

            // Collected Values Container
            List<Dictionary<string, object>> CollectedValues = new List<Dictionary<string, object>>();

            // Check if the link exists
            if (sheetListLink?.ID > 0)
            {
                // Get Sheet ID from db entity
                string sheetId = sheetListLink.SheetId;

                // Check for spreadsheet existence
                var existingSheet = await Converter.GetSheetAsync(sheetId);
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
                    WimNotificationOutput = $"Received {CollectedValues.Count} rows from Google Sheets"
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
