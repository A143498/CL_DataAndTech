# Martha Hause
# This documentation incorporates multiple different sources being joined together to get one central area for the projects expense savings to be represented.
# The projects in this file is Docuflash, COI, Self-service, Loss Run, and Session Time Reduced.
import pyodbc as p
import pandas as pd
import numpy as np

# python "K:\CVOWB\Martha\TableauAutomation.py"        <- run this to get the code to work. for some reason my code doesn't run unless you add this.

# this creates the connection to PCA-06. make sure you have the DSN connected on your computer or accessing the data source will not work.
PCA_connect = p.connect("DSN=MSS-P1-PCA-06")

# Bringing in Loss Run
print("Loss Run Information")
# the only thing this needs is to connect to the prog months data frame and sort out the months into progressive months and we should be good with this connection. and to add in the calculation for the quarters.


LR_Table = """
Select top 100
      FORMAT(EmailSentDate, 'MM-yy') As 'Date'
     ,Sum(Case When EmailSubject = 'Progressive Agent Contact Us' Then 1 Else 0 End) As 'Auto - Form'
     ,Sum(Case When EmailSubject Like 'PGR_FAO_LossRunRequest%' Then 1 Else 0 End)   As 'Auto - Widget'
     ,Sum( Case
             When Not(
                     EmailSubject = 'Progressive Agent Contact Us'
                     Or EmailSubject Like 'PGR_FAO_LossRunRequest%'
                     ) Then 1 Else 0 End)  As 'Auto - Email'
     ,Count(*)  As 'Auto - Total'
From  PCAStage.dbo.vwEmailClassifierProcessed
Where
      EmailSentDate > '2022-01-01' AND ProcessStatus = '1'
Group By FORMAT(EmailSentDate, 'MM-yy')
"""

LR_df = pd.read_sql(LR_Table, PCA_connect)

LR_df["Year"] = LR_df["Date"].str[-2:]


# creates the calculation for Money Saved. this needs to be relooked at.
def SavingsCalc(row):
    if row["Year"] == "22":
        return row["Auto - Total"] * 1.6666666666666666666666666666667 * 7.3
    elif row["Year"] == "23":
        return row["Auto - Total"] * 1.78436759319112 * 7.3


LR_df["Money Saved"] = LR_df.apply(SavingsCalc, axis=1)

# creates the Year for the excel.
LR_df["Year"] = "20" + LR_df["Year"]

# defining the projects here.
LR_df["Project"] = "Loss Run"
LR_df["Sub-project"] = "Loss Run"

# need to add the quarter part to this still. need to work with months part of the date here and turn it into an int. to get the month value I think
LR_df["Quarter Year"] = LR_df["Year"]

# displaying and capturing only certain Columns for the LR DataFrame.
LR_df = LR_df[["Project", "Sub-project", "Date", "Year", "Quarter Year", "Money Saved"]]
print(LR_df)


# Bringing in Docuflash
print("Docuflash Information")
# you have to connect to DB2 here. i am having trouble connecting to it so this may be another battle for the next person to attempt. Documentation on what I have done will be attached in the GitHub section.

# Bringing in COI
print("COI Information")

COI_Table = """
SELECT count(distinct(INFOID)) as dis_INFOID, FORMAT(ClickDateTime, 'yyyy-MM-dd') as date
FROM FSScoreCard.CLCQ.CLEXP_MLTIPRDTCOIData
GROUP BY FORMAT(ClickDateTime, 'yyyy-MM-dd');
"""
# Top 100 count(INFOID) as dis_INFOID,  distinct FORMAT(ClickDateTime, 'MM-yy') as date where ClickDateTime > '02-27-23'
# execute the statement
COI_df = pd.read_sql(COI_Table, PCA_connect)
print(COI_df)

# this is where I am bringing in the Progressive accounting months. because it uses PCA-06 as well, you don't need to create a new connection to the data source. You can use the one from the beginning of the python script. I only brought in accounting months from jan 2022 to now because the projects started then. we don't need the months before

ProgMonths_Table = """
SELECT DT_VAL, ACCT_CCYYMM
FROM [FSScoreCard].[CLCQ].[DimDate]
  where ACCT_CCYY >= 2022
"""

# this statement below reads the query and associates it to the connect on line ----------------------------.
ProgMonths_df = pd.read_sql(ProgMonths_Table, PCA_connect)
print(ProgMonths_df)

# below is a merge. A merge is very similar to a join in SAS. Look up a merge to understand more.
COI = pd.merge(ProgMonths_df, COI_df, how="inner", left_on="DT_VAL", right_on="date")

print(COI)

COI_grouped = COI.groupby("ACCT_CCYYMM").sum()
# need to figure out how to only grab the dis_infoid and the ACCT_CCYYMM. Below, that isn't the best way so it needs to be reworked a little bit.

COI_groupe = COI_grouped[["ACCT_CCYYMM", "dis_INFOID"]]
print(COI_groupe)

# after you get that worked out, you can add in the calculation to work with the numbers above. I am going to comment out the stuff below but this is close to what I was thinking could be added to create the calc for the project saved each month

# you need to go to the https://progressiveinsurance.sharepoint.com/:x:/r/sites/CLExperience/_layouts/15/Doc.aspx?sourcedoc=%7Bb84ed224-681c-48f2-b046-ed9e902b19e0%7D&action=edit&activeCell=%27Summary%20used%20in%20Monthly%27!B92&wdinitialsession=716985b7-9a79-4fcf-9cba-d5365a3b2b13&wdrldsc=16&wdrldc=1&wdrldr=AccessTokenExpiredWarning%2CRefreshingExpiredAccessT&cid=da31ddc6-b889-44fd-98e8-559dbd6cadd2 url
# once you are there, go to the COI tab and look at the calculations done. You can mimic what is already done in here to grab the calculations.


# Bringing in Self-Service

print("Self - Service Project Below")
# the only thing this needs is to connect to the prog months data frame and sort out the months into progressive months and we should be good with this connection. and to add in the calculation for the quarters.


# this is the SQL query to bring in the different pagenames for the Self-Service project and the Session Time Reduced project
SessionTime_Service_Table = """
Select A.eventmmyy,
a.pagename,
a.applicantgroups,
A.PRODUCT_GROUP,
SUM(a.pagenamecount) as pagehits FROM
(SELECT
	[EventDate] as eventmmyy,
	CASE WHEN [pagetype] = 'BASE_PAGE' THEN concat([pagename],'@')
	 ELSE [pagename] end AS [pagename],

	CASE WHEN [userdata userroletxt] LIKE '%Agent%' AND [geocarrier] <>  'Progressive Casualty Insurance Companies' AND [userdata AgentCode] <> 'null' then 'Agent'
         WHEN [userdata userroletxt] = 'null' AND [geocarrier] <>  'Progressive Casualty Insurance Companies' AND [userdata AgentCode] = 'null' AND [pagename] LIKE '%foragentsonly%' then 'Agent'
         WHEN [userdata UserRoleTxt] <> 'Agent' and  [userdata PartyLogonId] IN ('holder', 'proxy') AND [geocarrier] <> 'Progressive Casualty Insurance Companies' then 'Customer'
         WHEN [userdata userroletxt] = 'null' AND [geocarrier] <>  'Progressive Casualty Insurance Companies' AND [userdata PartyLogonId] = 'null' AND [pagename] LIKE '%progressivecommercial%' then 'Customer'
         WHEN  [userdata UserRoleTxt] = 'null' AND [geocarrier] <>  'Progressive Casualty Insurance Companies' then 'Unknown'
         WHEN [geocarrier] = 'Progressive Casualty Insurance Companies' then 'Internal'
         ELSE  'Internal' END AS applicantgroups,

	CASE WHEN  [userdata ProductCode] = 'CV' then 'Quoting'
	WHEN [userdata ProductCode] = 'CA' then 'Servicing'
	ELSE 'NA' END AS PRODUCT_GROUP,

	count([pagename]) as pagenamecount
	FROM [CLExperience].[dbo].[APPDCSV]
	WHERE
	Eventdate between '2022-12-31' and '2023-02-03' -- need to change this for a monthly update.
	AND [pagename] IN ('EndorseRBS', 'VehicleGaragingInformation', 'VehicleGaragingInformation-Reload', 'AdditionalDetails', 'OrderResults')
	AND [userdata ProductCode] = 'CA'
	Group By
	[Eventdate],
	[pagename],
	[pagetype],
	[userdata userroletxt], [geocarrier], [userdata AgentCode], [pagename], [userdata PartyLogonId], [userdata ProductCode]) A

WHERE applicantgroups='Internal' and pagename not in ('VehicleGaragingInformation@','OrderResults@')
Group by A.eventmmyy,
a.pagename,
a.applicantgroups,

A.PRODUCT_GROUP
Order By A.eventmmyy """

# above in like 74, there are 5 pagename values listed I have broken them out into which projects they associate to below.
# pagenames for Self-Service Project = 'EndorseRBS', 'VehicleGaragingInformation', 'VehicleGaragingInformation-Reload'
# pagenames for Session Time Reduced Project = 'AdditionalDetails', 'OrderResults'

# this statement below reads the query and associates it to the connect on line 93 currently.
SessionTime_df = pd.read_sql(SessionTime_Service_Table, PCA_connect)
print(SessionTime_df)

# this is where I am bringing in the Progressive accounting months. because it uses PCA-06 as well, you don't need to create a new connection to the data source. You can use the one from the beginning of the python script. I only brought in accounting months from jan 2022 to now because the projects started then. we don't need the months before
# ProgMonths_Table = """
# SELECT DT_VAL, ACCT_CCYYMM
# FROM [FSScoreCard].[CLCQ].[DimDate]
#   where ACCT_CCYY >= 2022
# """

# # this statement below reads the query and associates it to the connect on line 93 currently.
# ProgMonths_df = pd.read_sql(ProgMonths_Table, PCA_connect)
# I am commenting this out for the time being. this variable is located on line 91 so you shoulldn't have to re-establish the variable again (unless it has changed but in this case, it has not.)

# below is a merge. A merge is very similar to a join in SAS. Look up a merge to understand more.
print("Prog Self Service With Prog Months.")
ProgSelfService_SessionReduced_df = pd.merge(
    ProgMonths_df, SessionTime_df, how="inner", left_on="DT_VAL", right_on="eventmmyy"
)

# pivoting the data on the pagehits field and making the pagename column new column titles. think of a pivot from excel or SAS and it is doing the same thing
SelfService = pd.pivot_table(
    ProgSelfService_SessionReduced_df,
    index=["ACCT_CCYYMM", "applicantgroups", "PRODUCT_GROUP"],
    columns="pagename",
    values="pagehits",
    aggfunc=np.sum,
).reset_index()

# I reprinted the data here to make sure the merge went through. You can uncomment this to see the change from before the pivot to after!
# print(SelfService)
# print("")

# making a new dataframe for Session Time Reduced querying calculations
SessionTimeReduced = SelfService
# print("Session Time Reduced Data Frame")
# print(SessionTimeReduced) # you can print this out if you'd like to see how it looks, but I also printed it out in the Session Time Reduced Section.

print("")

# calculations to produce the information for Self Service.
SelfService["EndorseRBS_Internal_Pg_Count"] = (
    SelfService["EndorseRBS"] + SelfService["EndorseRBS@"]
)
SelfService["Vehicle_Garaging_Internal_Pg_Cnt"] = (
    SelfService["VehicleGaragingInformation"]
    + SelfService["VehicleGaragingInformation-Reload"]
)
SelfService["Endorse_RBS_Internal_Pg_Cnt"] = (
    SelfService["EndorseRBS_Internal_Pg_Count"]
    - SelfService["Vehicle_Garaging_Internal_Pg_Cnt"]
)
SelfService["Endorsement_HairCut"] = SelfService["Endorse_RBS_Internal_Pg_Cnt"] * 0.8

SelfService["ACCT_CCYYMM"] = SelfService["ACCT_CCYYMM"].astype(
    str
)  # converting the column to a string data type
SelfService["Year"] = SelfService["ACCT_CCYYMM"].str[
    2:4
]  # grabbing only the 22 or 23 or in the future 24, 25, 26 from the ACCT_CCYYMM field

# creating a date field for the excel version below
SelfService["Date"] = (
    SelfService["ACCT_CCYYMM"].str[2:4] + "-" + SelfService["ACCT_CCYYMM"].str[4:]
)


# creates the calculation for Minimum Flow. needs to be changed each year. to do that, you just change the number.
def MinimumFlow(row):
    if row["Year"] == "22":
        return 7.52629901179471
    elif row["Year"] == "23":
        return 7.52629901179471


SelfService["MinFlow"] = SelfService.apply(
    MinimumFlow, axis=1
)  # this statement right here lets the lines 149-153 run. Look up "creating a function in python". This will help you understand why this has to be done this way


# get avg user for the year. change if it changes each year
def AvgUser(row):
    if row["Year"] == "22":
        return 2.814383234
    elif row["Year"] == "23":
        return 2.814383234


SelfService["AvgUser"] = SelfService.apply(AvgUser, axis=1)

# creating the Benefit Hours Column Number.
SelfService["Benefit_Hrs"] = (
    (SelfService["Endorsement_HairCut"] * SelfService["MinFlow"])
    + (SelfService["Endorsement_HairCut"] * SelfService["AvgUser"])
) / 3600


# get avg user for the year. change if it changes each year
def InternalRate(row):
    if row["Year"] == "22":
        return 99.992
    elif row["Year"] == "23":
        return 106  # this number needs to be changed to the number correct for 2023.


SelfService["Internal_Rate"] = SelfService.apply(InternalRate, axis=1)


# creates the calculation for Money Saved. this will need to be updated each year if the .9 changes.
def Money(row):
    if row["Year"] == "22":
        return SelfService["Benefit_Hrs"] * SelfService["Internal_Rate"] * 0.9
    elif row["Year"] == "23":
        return SelfService["Benefit_Hrs"] * SelfService["Internal_Rate"] * 0.9


SelfService["Money Saved"] = SelfService.apply(Money, axis=1)

# dropping fields we don't need and adding fields we do to push over to the excel output.
# creates the Year for the excel.
SelfService["Year"] = "20" + SelfService["Year"]

# defining the projects here.
SelfService["Project"] = "Servicing"
SelfService["Sub-project"] = "Servicing"

# need to add the quarter part to this still. need to work with months part of the date here and turn it into an int. to get the month value I think
SelfService["Quarter Year"] = SelfService["Year"]

# displaying and capturing only certain Columns for the LR DataFrame.
SelfService = SelfService[
    ["Project", "Sub-project", "Date", "Year", "Quarter Year", "Money Saved"]
]

# prints final output for the Self Service project.
print(SelfService)
print("")

# ------------------------------------------------------------------------------------------------
# Starting the Session Time Reduced Calculation Below this point

# this is not a normal DataFrame set up. Just so you know. I am trying not to add too many columns with similar data to the SelfService Data Frame. if this is going to change by year, we need to restructure this. the structure need would be similiar to function from above and below.
print("Session Time Reduced Project Below")
Baseline = {"OrderBase": [0.313570414102314], "AdditionalBase": [0.175184323782372]}
Baseline_df = pd.DataFrame(Baseline)
# print(Baseline_df)  # prints the baseline dataframe

SessionTimeReduced["Expected_Order_Results"] = (
    SessionTimeReduced["AdditionalDetails"] / Baseline_df["AdditionalBase"]
)
SessionTimeReduced["Expected_Additional_Details"] = (
    SessionTimeReduced["OrderResults"] / Baseline_df["OrderBase"]
)

SessionTimeReduced["OrderResults_AffectedPages"] = (
    SessionTimeReduced["Expected_Order_Results"] - SessionTimeReduced["OrderResults"]
)
SessionTimeReduced["AdditionalDetails_AffectedPages"] = (
    SessionTimeReduced["Expected_Additional_Details"]
    - SessionTimeReduced["AdditionalDetails"]
)
# from here, you need to figure out how to get the calc saved and add that as a column for each month to make a final calc on how much was saved because of this project.
# this will entail a restructure of the second excel used to get the money saved.

print("Session Time Reduced Table")
print(SessionTimeReduced)

PCA_connect.close()
