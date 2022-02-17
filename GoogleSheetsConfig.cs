using Microsoft.Extensions.Configuration;

namespace Sushi.Mediakiwi.Module.GoogleSheetsSync
{
    internal class GoogleSheetsConfig 
    {
        /// <summary>
        /// The Google OpenID client ID
        /// </summary>
        [ConfigurationKeyName("client-id")]
        public string ClientID { get; set; }

        /// <summary>
        /// The Google OpenID client secret
        /// </summary>
        [ConfigurationKeyName("client-secret")]
        public string ClientSecret { get; set; }

        /// <summary>
        /// The (relative) filename to the service account credentials file
        /// </summary>
        [ConfigurationKeyName("service-account-filename")]
        public string ServiceAccountFilename { get; set; }

        /// <summary>
        /// On which path do we need to open the Google OpenID auth listener ?
        /// default = '/signin-google'
        /// </summary>
        [ConfigurationKeyName("handler-path")]
        public string HandlerPath { get; set; } = "/signin-google";
    }
}
