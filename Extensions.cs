using Microsoft.AspNetCore.Builder;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.DependencyInjection;
using Sushi.Mediakiwi.Framework.Interfaces;

namespace Sushi.Mediakiwi.Module.GoogleSheetsSync
{
    public static class Extensions
    {
        public static IApplicationBuilder UseGoogleOpenID(this IApplicationBuilder builder)
        {
            var config = builder.ApplicationServices.GetService<IConfiguration>();
            if (config != null)
            {
                GoogleSheetsConfig _config = config.GetSection("GoogleSheetsSettings").Get<GoogleSheetsConfig>();
                if (string.IsNullOrWhiteSpace(_config?.HandlerPath) == false)
                {
                    return builder.UseWhen(context => context.Request.Path.StartsWithSegments(_config.HandlerPath), appBuilder =>
                    {
                        appBuilder.UseMiddleware<GoogleAuthListener>();
                    });
                    
                }
                else
                {
                    throw new SystemException("When using 'UseGoogleOpenID' there must be a value added to configuration : 'GoogleSheetsSettings:handler-path' ");
                }
            }

            return builder;
        }

        /// <summary>
        /// Adds the google sheets modules
        /// </summary>
        /// <param name="services"></param>
        /// <param name="services">The current service collection</param>
        /// <param name="serviceAccountCredentialsFileName">The filename to the file containing the serviceaccount credentials</param>
        /// <param name="clientId">The OAuth client ID</param>
        /// <param name="clientSecret">The OAuth client Secret</param>
        /// <param name="enableExportModule">Enable the export module ?</param>
        /// <param name="enableImportModule">Enable the import module ?</param>
        public static void AddGoogleSheetsModules(this IServiceCollection services, bool enableExportModule, bool enableImportModule)
        {
            // Do nothing
            if (enableExportModule == false && enableImportModule == false)
            {
                return;
            }

            // Run the installer when needed
            Task.Run(async () => await ModuleInstaller.InstallWhenNeededAsync());

            // Add the Sheets Logic
            services.AddSingleton<GoogleSheetLogic>();

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
