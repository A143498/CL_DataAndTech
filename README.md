# CL Expense Saving Trending Dashboard Information

Hello Everyone! Welcome to the location where I will be placing information related to the Automation Process for the CL Expense Reduction Display Project being worked on currently. 

In this repo, you will see information about each project, the breakdown of how to get information from each project, the start of the automation for the project backend, and the steps that have been going on to get through this project.


Setting up Python to do automated part of the Project
----------------------------------------------------
You need to make sure python is installed. 
run the command
pip install python

Once that runs successfully and it is downloaded, now you need to install pandas and numpy. I use both of those data dictionaries in my code. run these commands below.
pip install pandas
pip install numpy

Once those have run and are installed you need to make sure your computer DSN has a connection to the different data sources. Even though you have connection through SQL, or DB2, you need to make sure the connection is in your data source network. You can use the links below to connect to the different data sources using python. 

connect to DB2https://githubprod.prci.com/progressive/progressive-python/blob/master/Connecting%20to%20Data/Connect%20to%20Data%20with%20Windows/Query%20DB2%20from%20Windows.py

connect to DB2. https://progressiveinsurance.sharepoint.com/:w:/r/sites/BIDS%20Technical%20Forum/Python/_layouts/15/Doc.aspx?sourcedoc=%7B5D40D2B2-555B-4022-9310-73F186CD42CB%7D&file=Python%20Connecting%20to%20DB2%20Server%20on%20Windows.docx&action=default&mobileredirect=true&DefaultItemOpen=1&cid=9ae16055-aa3c-4017-a45e-ece3bbec07dd

connect to different servers https://progressiveinsurance.sharepoint.com/sites/analytics/SiteAssets/SQL-Data-Connections.pdf#search=set%20up%20db2%20connection%20in%20odbc%20data%20sources

One caveat to the connections, I have not connected to DB2 without error so this will need to be looked into to be able to get Docuflash's data into the excel. The other data sources all connect to one server and that is listed in the python file. 

Background information on this project 
----------------------------------------

For this project, you need to understand that 4 of the projects come from one data source. Location: MSS-P1-PCA-06
The projects may come from different tables within PCA-06 but you still use just one connection. 
The one project not in this location is docuflash. That is located in DB2. More background information on these projects are located at this file path location: "K:\CVOWB\Martha\TableauAutomation". 

In the python document file, you will need to connect some of the queries to the prog month table still but this shouldn't be too difficult.

A few things that need to be done within the python code still.
1. bring in Docuflash data
2. verify the numbers are correct across each project
3. connect each of the tables to the progressive month table. Make sure when connecting that you are at the daily level and then switch to prog month after. I saw this was an issue that could happen and this should be verified before moving on and assuming it is correct!
4. Verify that the data already brought in is capturing the correct calculations and the correct numbers for those calculations.
5. merge all of the data together 
6. export the data to an excel file.
7. verify the export looks like the current excel being used. If not, you need to figure a way to make them similair.


I did not put in the quarter-year calculation that is in the excel so this needs to be done as well.
