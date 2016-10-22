# Meeting_List_Database_CSharp
##Summary
* Meeting list that has a SQL Server database
* Used Visual Studio 2012 ASP.NET, C-Sharp
* Template used: Web Form Application
* Membership provider used to enable update, inserts, and deletes to selected individuals
* The list can be filtered by day, time, and / or town.
* The list can be downloaded as a csv file.
* The C-Sharp code is used for the default day on startup.
* C-Sharp code is used to filter the data from the user selection.
* C-Sharp code is used to export the list to a csv file. 

##Code notes
* The schema that is used for the Meeting List is in the MeetinglistDbSchema.sql.
* Run the MeetinglistDbSchema.sql in Sql Server Management Studio to create the database schema.
* Use the Web Form Application template in VS 2012.
* The files that were added or changed are in this repository. 
* Set up the connection strings in the webconfig file for the membership provider and for the meeting list data.
* Use the aspx files for the page layouts and use css for styles.
* Use the C-Sharp files for code examples. 
