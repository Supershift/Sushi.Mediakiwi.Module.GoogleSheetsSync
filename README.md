# Sushi.Mediakiwi.Module.GoogleSheetsSync
A list module for use in Mediakiwi which allows for synchronizing data with GoogleSheets

Installation steps :
* Download the ServiceAccount credentials file from the Google Cloud API explorer.
* Download the OAuth client secrets file from the Google Cloud API explorer *(optional)*.
* Place these files in the Root of your project and set 'Copy to output directory' to always.
* Add these lines to your startup code :

```cs
// Get credential Files for Google Sheets
var serviceAccountCredentials = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "sheetsCredentials.json");
var clientSecretCredentials = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "sheetsClientSecret.json");

// Install both Import and Export modules
services.AddGoogleSheetsModules(serviceAccountCredentials, clientSecretCredentials);
```

Things to note :
* You can also enable only one Module, by setting _enableExportModule_ or _enableImportModule_
* You can omit the ClientSecrets file parameter, the module will then use the shared ServiceAccount for creating and uploading the Sheets
* When using the ClientSecrets param, each user will see it's own personal version of the created Sheet.
* When using only the ServiceAccount param, each user will see the same version of the created Sheet.

This will also create the database table if needed, so the database connectionstring must be known at this point.
this can be done with :
```cs
MicroORM.DatabaseConfiguration.SetDefaultConnectionString(connString);
```
