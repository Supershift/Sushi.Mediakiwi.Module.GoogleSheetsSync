# Sushi.Mediakiwi.Module.GoogleSheetsSync
A list module for use in Mediakiwi which allows for synchronizing data with GoogleSheets

Installation steps :
* Download the ServiceAccount credentials file from the Google Cloud API explorer.
* Download the OAuth client secrets file from the Google Cloud API explorer.
* Place these files in the Root of your project and set 'Copy to output directory' to always.
* Add these lines to your startup code :

```cs
  GoogleSheetExportListModule mod = new GoogleSheetExportListModule();
  var serviceAccountCredentials = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "sheetsCredentials.json");
  var clientSecretCredentials = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "sheetsClientSecret.json");

  Task.Run(async () => await mod.InitAsync(serviceAccountCredentials));

  services.AddSingleton(typeof(IListModule), mod);
```

This will also create the database table if needed, so the database connection must be known at this point.
this can be done with :
```cs
  MicroORM.DatabaseConfiguration.SetDefaultConnectionString(connString);
```
