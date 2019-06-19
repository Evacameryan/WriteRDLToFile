# WriteRDLToFile
This project writes a SQL Server Report Server RDL to file.

We removed our originally references so for the program to build correctly you must fix a couple of things.
1.	Change the Web Reference URL of the Web References RE2005 and RS2005 to your report server’s report service and report execution. (Most likely you just need to change the \<reportserver\> to your report server url)
2.	Settings.settings file needs to have “ReportServer” replaced with the URL of your report server (if for some reason it hasn't changed when updating the Web References.
3.	In SaveRDL you need to replace certain things
  
    •	\<windowsUserName\> - windows username with server privileges
  
    •	\<windowsUserPassword\> - passwords of windows username
  
    •	\<Domain\> - domain of username
  
    •	\<server\> - server name
  
4.	In SaveRDLToFileStream replace <server> with your report server name.

5.  After this is done you can either add this to a method as a class library and remove the main method or use it as a console application to test. 

Find the name of your report folder and the report you would like to print. Place them in the \<reportFolder\> and \<reportName\> locations in the main method. Replace \<location\> with the location to save the report.

Also add any parameters in a dictionary to be passed to as a parameter.
