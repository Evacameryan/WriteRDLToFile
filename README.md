# WriteRDLToFile
This project writes a SQL Server Report Server RDL to file.

We removed our originally references so for the program to build correctly you must fix a couple of things.
1.	Change the Web Reference URL of the Web References RE2005 and RS2005 to your report server’s report service and report execution. (Most likely you just need to change the \<reportserver\> to your report server url)
2.	Settings.settings file needs to have “ReportServer” replaced with the URL of your report server 
3.	In SaveRDL you need to replace certain things
  
    •	<windowsUserName> - windows username with server privileges
  
    •	<windowsUserPassword> - passwords of windows username
  
    •	<Domain> - domain of username
  
    •	<server> - server name
  
4.	In SaveRDLToFileStream replace <server> with your report server name.
