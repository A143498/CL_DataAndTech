# Martha Hause
# This documentation incorporates multiple different sources being joined together to get one central area for the projects expense savings to be represented.
# The projects in this file is Docuflash, COI, Self-service, Loss Run, and Session Time Reduced.
import pyodbc as p
import pandas as pd

# python "u:/Github Training/trialfile.py"        <- run this to get the code to work. for some reason my code doesn't run unless you add this.

# Bringing in Loss Run
print("Loss Run Information")

LR_connect = p.connect("DSN=MSS-P1-PCA-06")

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

LR_df = pd.read_sql(LR_Table, LR_connect)
print(LR_df)

LR_connect.close()

# Bringing in Docuflash
print("Docuflash Information")


# Bringing in COI
print("COI Information")

COI_connect = p.connect("DSN=MSS-P1-PCA-06")

COI_Table = """
SELECT count(distinct(INFOID)) as dis_INFOID, FORMAT(ClickDateTime, 'MM-yy') as date
FROM FSScoreCard.CLCQ.CLEXP_MLTIPRDTCOIData
GROUP BY FORMAT(ClickDateTime, 'MM-yy');
"""
# Top 100 count(INFOID) as dis_INFOID,  distinct FORMAT(ClickDateTime, 'MM-yy') as date where ClickDateTime > '02-27-23'
# execute the statement
COI_df = pd.read_sql(COI_Table, COI_connect)
print(COI_df)

# close connection
COI_connect.close()


# Bringing in Self-Service
print("Self-Service Information")


# Bringing in Session Time Reduced
print("Session Time Reduced Information")
