# ServerProcessingStats
This script uses Powershell, SQLite, the PSSQLite Powershell module by RamblingCookieMonster, and Excel to calculate server processing times and generate an Excel workbook.

Here's the long explanation:

This script grabs server names from a SQLite table, then uses a ForEach loop to iterate through each server.  For each server,  a check is performed to see if a folder exists containing a duplicate of my test data.  If not, a copy is created to perform my processing on.  The script then stores the time the script started running my processing files, and also stores when the process file completed.  Then subtracts the two to calculate how many seconds have passed.

As it stands this script creats a total of 12 worksheets, 1 worksheet for the Overview page which contains a bar graph of average, today, and yesterday's times, and then a worksheet for each server that the script was run on.
On each server worksheet, there are 2 line graphs.  One that displays the last 7 days of data, along with the numerical values.  And one that displays the last 30 days of data, also with the numerical values.

After all the worksheets have been populated, an email containing the workbook is sent out using the SMTP server.  If an SMTP server is not required, then that bit of code can be commented out and a local copy will still be saved.
