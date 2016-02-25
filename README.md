# google-analytics-auditor
This will become a public Add-on for Google Spreadsheets. It is a Google Apps Script (GAS) that connects to the [Google Analytics Management API](https://developers.google.com/analytics/devguides/config/mgmt/v3/) and pulls down information related to Goals, Filters, and View settings.

**BUGS**
 - select an account with no filters in it - cannot find length of undefined
 - issue when getting user limit - it doesn't continue submitting form and exits out
 - check for isBlank() could be checking too many rows and thus is slow - refactor to just check for a single row or cell

**Feature Requests**
 - add additional reports

**Todo**
 - review exceptions

**Prior to launch**
 - document code (JSDoc)
 - review [UI Style Guide for Add-ons](https://developers.google.com/apps-script/add-ons/style)
 - add onInstall dialog
 - add GA tracking


**Add to docs**
 - add all 20 goal steps?

**Short Description**
GA Auditor creates a report for Google Analytics filters, goals & view settings for all properties & views in an account.

**Post-install tip**
Congratulations, you've successfully installed GA Auditor. You can now begin creating reports. To start, select the "Create reports" sub-menu item under the GA Auditor menu. Then, select an account you'd like to audit, followed by the report type.