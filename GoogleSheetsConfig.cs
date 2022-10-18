using Microsoft.Extensions.Configuration;

namespace Sushi.Mediakiwi.Module.GoogleSheetsSync
{

    public class GoogleSheetsConfig 
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

        /// <summary>
        /// Which domains are allowed to edit the produced googlesheets file ?
        /// </summary>
        [ConfigurationKeyName("allowed-domains")]
        public string[] AllowedDomains { get; set; }


        /// <summary>
        /// What type of permissions are we using ?
        /// this can be one of :
        /// 'USEREMAIL' for User email based permission (only creating user can view).
        /// 'USERDOMAIN' for User email domain based permission (anyone in the same email domain as creating user can view).
        /// 'SETDOMAINS' to use the AllowedDomains collection (anyone in the same email domain as one of the items in the list can view) 
        /// </summary>
        [ConfigurationKeyName("permission-base")]
        public GoogleSheetsPermissionsEnum PermissionBase { get; set; } = GoogleSheetsPermissionsEnum.SETDOMAINS;

    }
}
