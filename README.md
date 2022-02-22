# Sushi.Mediakiwi.Module.GoogleSheetsSync
A list module for use in Mediakiwi which allows for synchronizing data with GoogleSheets

Installation steps :
* Download the ServiceAccount credentials file from the Google Cloud API explorer.
* Place this file in the Root of your project and set 'Copy to output directory' to always.
* Add this section to your configuration file (appsettings.json) :

```JSON
"GoogleSheetsSettings": {
  // Get this ClientID from the Google Cloud platform
  "client-id": "[GOOGLE-CLIENT-ID]",
  // Get this ClientSecret from the Google Cloud platform
  "client-secret": "[GOOGLE-CLIENT-SECRET]",
  // The relative filename for the AerviceAccount credentials file
  "service-account-filename": "sheetsCredentials.json",
  // What is the relative url path to listen to OpenID requests
  "handler-path": "/signin-google"
},
```

* Add these lines to your services startup (ConfigureServices) code :

```cs
// Install all included modules
services.AddGoogleSheetsModules(true, true, true);
```
This will also create the database table if needed, so the database connectionstring must be known at this point.
This can be done with :
```cs
MicroORM.DatabaseConfiguration.SetDefaultConnectionString(connString);
```

* Add these lines to your application startup (Configure) code, before the call to _app.UseMediakiwi()_ :

```cs
// Install the OpenID listener (only needed when ClientID and ClientSecret are used)
app.UseGoogleOpenID();
```

Things to note :
* You can also enable only one Module, by setting _enableExportModule_, _enableViewModule_ or _enableImportModule_.
* The **Import** module will only show up if the list has an implementation for the _ListDataReceived_ event.
* The **Export** module will only show up if the list has the setting for _XLS export_ enabled.
* The **View** module will only show up when a last has been exported at least once.
