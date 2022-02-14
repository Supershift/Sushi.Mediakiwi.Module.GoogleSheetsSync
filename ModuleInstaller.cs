namespace Sushi.Mediakiwi.Module.GoogleSheetsSync
{
    internal class ModuleInstaller
    {
        public static async Task InstallWhenNeededAsync()
        {
            var installNeeded = false;

            try
            {
                var listLinks = await Data.GoogleSheetListLink.FetchAllAsync();
                
            }
            catch (Exception)
            {
                installNeeded = true;
            }

            if (installNeeded == false)
            {
                return;
            }

            string sqlScript = "";

            try
            {
                using (Stream stream = typeof(ModuleInstaller).Assembly.GetManifestResourceStream("Sushi.Mediakiwi.Module.GoogleSheetsSync.Installer.sql"))
                using (StreamReader reader = new StreamReader(stream))
                {
                    sqlScript = reader.ReadToEnd();
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }

            try
            {
                var conn = new MicroORM.Connector<Data.GoogleSheetListLink>();

                await conn.ExecuteNonQueryAsync(sqlScript);
            }
            catch (Exception ex)
            {
                throw ex;
            }
            
        }
    }
}
