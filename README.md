# mailchimpMacro
Based on MailChimp's RESTFUL API <br>
By ArKits - arkits@outlook.com

This is a small macro that is designed to add all entries in an Excel file to MailChimp, without using the web-ui.

* The macro picks column 1 for emails and column 3 for names. 
* The macro checks to see if an email address is already subscribed or not, and ignores those emails.
* Subscribes an email address to the list (with additional details) 
* DeleteRow() Function is to delete the rows which have "TRUE" in columns 14 to 16
