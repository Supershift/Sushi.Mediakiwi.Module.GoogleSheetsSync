# Using a Google Cloud Service Account #

This comes in handy if you want to enable quick use of the module and don't mind that a Google Sheet is the same for each user of your application. 
This means that:
* **User X** clicks on export.
* **User X** modifies the resulting sheet.
* **User Y** clicks on View Sheet.
* **User Y** sees the data that **User X** has modified.

So there is only one Google Sheet coupled to the list.

## 1. Creating a Google Cloud project ##

* Go to the [Google Cloud Platform](https://console.cloud.google.com/).
* Click on the dropdown in the topleft, next to the _Google Cloud Platform_ title.
* Click on the _New Project_ button in the modal that opens.
* Fill in the required fields and click the _Create_ button.
* Select the newly created project in the dropdown in the topleft, next to the _Google Cloud Platform_ title.

## 2. Add the required API's to the project ##

* Click on the hamburger menu icon in the topleft corner of the page.
* Click on _API's & Services_ in the opened menu.
* Click on _+ Enable APIs and Services_ on top of the page.
* Search for _'Google Sheets'_ in the searchbar.
* Click on the _'Google Sheets API'_ that appears as a result.
* Click on the _Enable_ button.
* **Do the same for _'Google Drive API'_.**

## 3. Adding a service account to the project ##

* Click on the hamburger menu icon in the topleft corner of the page.
* Click on _API's & Services_ in the opened menu.
* Click on _Credentials_ in the submenu.
* Click on _Create credentials_ on top of the page.
* Choose _Service account_ from the opened menu.
* Fill in the required fields and click on the _Create and continue_ button.
* We can skip the next step, so click the _Continue_ button.
* We can skip the next step as well.

A Service Account has now been created, next we need to download the credentials JSON file.

## 4. Download the service account credentials file ##

* Click on the hamburger menu icon in the topleft corner of the page.
* Click on _API's & Services_ in the opened menu.
* Click on _Credentials_ in the submenu.
* Click on the Service Account we've just created in step 3.
* Click on the tab _'Keys'_.
* There should be a dropdown _Add key_, click it and choose _Create new key_.
* In the modal that opens, choose the _JSON_ option and click _Create_.
* A JSON file will be downloaded containing the authentication information for our Service Account.

Continue with the apprioprate step [Readme](README.md)
