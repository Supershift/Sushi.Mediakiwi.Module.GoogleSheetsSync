# Using a Google Cloud OAuth OpenID #

This comes in handy if you want to enable personal use of the module. 
This means that:
* **User A** clicks on export.
* **User A** modifies the resulting sheet.
* **User B** clicks on View Sheet.
* **User B** does not see the data that **User X** has modified.

So there is a Google Sheet coupled a user for every list.

**The validation process can take up to 6 weeks to complete**

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

## 3. Adding a OAuth account to the project ##

* Click on the hamburger menu icon in the topleft corner of the page.
* Click on _API's & Services_ in the opened menu.
* Click on _Credentials_ in the submenu.
* Click on _Create credentials_ on top of the page.
* Choose _OAuth Client ID_ from the opened menu.
* Choose _Web Application_ from the opened menu.
* Fill in the required fields, the Redirect origins should contain the URLs from where you're accessing Google Sheets.
* Click on the _Create_ button.

An OAuth Client ID has now been created, next we need to configure the consent screen.

## 4. Configure OAth consent screen ##

* Click on the hamburger menu icon in the topleft corner of the page.
* Click on _API's & Services_ in the opened menu.
* Click on _OAuth consent screen_ in the submenu.
* Fill in the required information.
* The Authorized domains should match the ones added in step 3.
* For the Scopes step, pick the following scopes :
  * /auth/drive.file	
  * /auth/spreadsheets	
 
## 5. Submit the project for validation ##

Since you're using personal data from Google Users, the project you've created needs validation by google.
the steps needed for this are as follows :
* Submit the project for validation from within the OAuth consent screen.
* You need to have the following information ready for Google to allow submitting your project.
  * A URL to the homepage of your website.
  * A URL to the Privacy Policy of your website.
  * A YouTube video explaining the app you're building and the need for using Google Sheets.  
* Confirm domain ownership, you will receive an email from Google saying you need to confirm domain ownership.

When your project is validated continue with the apprioprate step [Readme](README.md)
