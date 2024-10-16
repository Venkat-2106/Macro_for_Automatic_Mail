Instructions for Using the VBA Scripts
Prerequisites
Microsoft Excel: Ensure you have Microsoft Excel installed.
Microsoft PowerPoint: Make sure you have PowerPoint installed.
Microsoft Outlook: Outlook should be installed and configured for email sending.
Charts in Excel: You should have charts created in your Excel workbook.
Setting Up the Excel Workbook
Open Excel: Launch Microsoft Excel and open the workbook containing your charts.

Create a Sheet for Email Information:

Create a new sheet (e.g., "Sheet6") that will contain:
Column A: Recipient Email Addresses
Column B: CC Email Addresses (if any)
Column C: Email Body/Text
Column D: File Paths for Attachments (if applicable)
Example:

scss
Copy code
| A                | B          | C                | D                |
|------------------|------------|------------------|------------------|
| recipient1@example.com | cc1@example.com | Body text for email | C:\path\to\file1 |
| recipient2@example.com | cc2@example.com | Body text for email | C:\path\to\file2 |
Adding the VBA Code
Open the VBA Editor:

Press ALT + F11 in Excel to open the Visual Basic for Applications (VBA) editor.
Insert a New Module:

Right-click on any of the items in the Project Explorer.
Select Insert -> Module.
Copy and Paste the Code:

Copy the provided VBA code (both ExportChartsToPowerPoint and MultibleMail subroutines) and paste it into the new module.
Close the VBA Editor:

Save your work and close the VBA editor.
Running the Scripts
Export Charts to PowerPoint:

To export charts, press ALT + F8, select ExportChartsToPowerPoint, and click Run.
This will create a PowerPoint presentation with charts from all sheets and send it as an email.
Send Emails:

To send emails, press ALT + F8, select MultibleMail, and click Run.
This will send emails to the recipients specified in "Sheet6" with the specified body and attachments.
Customizing the Code
Email Recipient: Modify the recipient addresses in "Sheet6" to your desired email addresses.
Subject and Body: Change the subject and body text in the script or directly in "Sheet6".
Attachments: Ensure that the file paths for any attachments are correct and accessible.
Important Notes
Error Handling: If any errors occur during the execution, a message box will display the error details.
Testing: Test with a small number of emails and charts first to ensure everything works as expected.
Outlook Security: You may receive security prompts from Outlook when sending emails programmatically.
Conclusion
Following these instructions will allow you to successfully export charts from Excel to PowerPoint and send emails with the relevant information. If you encounter any issues, please consult your IT support or the documentation for Microsoft Office applications. Feel free to reach out for further assistance!
