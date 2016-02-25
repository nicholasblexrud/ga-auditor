# GA Auditor
The GA Auditor spreadsheet add-on brings you the power of the Google Analytics Management API combined with the power of data manipulation in Google Spreadsheets. With this tool, you can create custom reports for:
- Filters
- Goals
- View settings

**BUGS**
 - check for isBlank() could be checking too many rows and thus is slow - refactor to just check for a single row or cell
 - add a counter for repeated calls when 'User Limit Exceeded' happens - timeout after 10 tries

**Feature Requests**
 - add additional reports

**Todo**
 - review exceptions

**Prior to launch**
 - document code (JSDoc)
 - review [UI Style Guide for Add-ons](https://developers.google.com/apps-script/add-ons/style)
 - add GA tracking

**Add to docs**
 - add all 20 goal steps?

**Short Description**
GA Auditor creates a report for Google Analytics filters, goals & view settings for all properties & views in an account.

**Post-install tip**
Congratulations, you've successfully installed GA Auditor. You can now begin creating reports. To start, select the "Create reports" sub-menu item under the GA Auditor menu. Then, select an account you'd like to audit, followed by the report type.