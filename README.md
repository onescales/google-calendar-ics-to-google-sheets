# Export Google Calendar ICS to Google Sheets (excel)
Simple steps to export Google Calendar, get a .ics file and import it to Google sheets.

These steps allow you to DIY instead of using 3rd party apps that cost money or may read your data. It is private to you only and easy to setup.

For a visual overview tutorial on ics to csv conversion and full details, see our article at https://onescales.com/blogs/main/export-google-calendar-ics-to-google-sheets-excel and our youtube video at https://www.youtube.com/watch?v=BRGHDS_-rDI

* now supports recurring events up to 5 years forward and ends in today's date.

# Steps

1. Create Google Sheet
- Open a google sheet
2. Create Apps Script
- In Google Sheet Top Menu, Click on Extensions -> App Script
- Edit the Code.gs file and copy paste this repository Code.gs (https://github.com/onescales/google-calendar-ics-to-google-sheets/blob/main/Code.gs)
- Create new file via + icon called "UploadForm.html" and copy paste this repository UploadForm.html (https://github.com/onescales/google-calendar-ics-to-google-sheets/blob/main/UploadForm.html)
- Click on Deploy on top right hand side. Set a "project title" (for example - ICS Import) and click on "Deploy" -> "New Deployment"
3. Set Permissions of Apps Script and Publish
Make sure to deploy as "Web app" and autorize permissions.
4. Go Back to Sheet
- Click on top menu "ICS Import" -> "Upload and Import ICS"
- Select All Dates or Specific Timeframe
- Select .ICS file from your computer and click on "Upload ICS"
5. Enjoy The Sheet!

# Additional Notes
If you would like to open .ics in Excel, then follow all steps above and in Google Sheets, click on "File" -> "Download" - "csv" and open in Excel the csv file.

# Support Us / Donate
If this helped you in any way, please consider supporting us at https://onescales.com/pages/support-us

# Suggestions, Comments and Contact
If you have any suggestions, comments, insight or just want to say hi, thanks or share your experience, you can contact us at:
- Our WebSite: https://onescales.com/
- Contact Us: https://onescales.com/pages/contact
- Youtube Channel: https://www.youtube.com/@onescales
- Twitter/X: https://twitter.com/one_scales
- LinkedIn: https://www.linkedin.com/company/one-scales/




