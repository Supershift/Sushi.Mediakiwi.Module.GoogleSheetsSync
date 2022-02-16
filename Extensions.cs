using Microsoft.Extensions.DependencyInjection;
using Sushi.Mediakiwi.Framework.Interfaces;

namespace Sushi.Mediakiwi.Module.GoogleSheetsSync
{
    public static class Extensions
    {
        /// <summary>
        /// Adds both the google sheets modules (Export and Import)
        /// </summary>
        /// <param name="services">The current service collection</param>
        /// <param name="serviceAccountCredentialsFileName">The filename to the file containing the serviceaccount credentials</param>
        public static void AddGoogleSheetsModules(this IServiceCollection services, string serviceAccountCredentialsFileName)
        { 
            AddGoogleSheetsModules(services, serviceAccountCredentialsFileName, "", true, true);
        }

        /// <summary>
        /// Adds the google sheets modules
        /// </summary>
        /// <param name="services"></param>
        /// <param name="services">The current service collection</param>
        /// <param name="serviceAccountCredentialsFileName">The filename to the file containing the serviceaccount credentials</param>
        /// <param name="clientSecretsFileName">The filename to the file containing the OAuth client credentials</param>
        /// <param name="enableExportModule">Enable the export module ?</param>
        /// <param name="enableImportModule">Enable the import module ?</param>
        public static void AddGoogleSheetsModules(this IServiceCollection services, string serviceAccountCredentialsFileName, string clientSecretsFileName, bool enableExportModule, bool enableImportModule)
        {
            // Do nothing
            if (enableExportModule == false && enableImportModule == false)
            {
                return;
            }

            GoogleSheetLogic logic = new GoogleSheetLogic(serviceAccountCredentialsFileName, clientSecretsFileName);

            // Run the installer when needed
            Task.Run(async () => await ModuleInstaller.InstallWhenNeededAsync());

            // Instantiate the Logic class
            Task.Run(async () => await logic.InitializeAsync());

            // Add the Sheets Logic
            services.AddSingleton(typeof(GoogleSheetLogic), logic);

            // Add the two modules
            if (enableExportModule)
            {
                services.AddSingleton(typeof(IListModule), typeof(GoogleSheetExportListModule));
            }

            if (enableImportModule)
            {
                services.AddSingleton(typeof(IListModule), typeof(GoogleSheetsImportListModule));
            }
        }
    }
}
