# -*- coding: utf-8 -*-
"""
Created on Wed Sep 18 16:39:32 2024
@author: JohnK
"""
import pyodbc
import pandas as pd
from datetime import datetime
import os

# Read Excel files
Account_List_Report = pd.read_excel("C:\\Users\\YamagantiMuraliKrish\\Downloads\\New folder\\Account List Report.xlsx")
Missing_Invoices_Report = pd.read_excel("C:\\Users\\YamagantiMuraliKrish\\Downloads\\Missing Invoice Report\\Missing Invoice Report.xlsx")

# Database connection
server = "wmssql.database.windows.net"
database = "Vectorproddb"
username = "John"
password = "Quadrant@12$"
driver = "{SQL Server}"
cnxn = pyodbc.connect(
    f"DRIVER={driver};SERVER=tcp:{server};PORT=1433;DATABASE={database};UID={username};PWD={password}"
)

# Execute SQL query
OpenInvoicesReport_query = """SET NOCOUNT ON;
select BD.ImageFileName, I.InvoiceNumber, I.InvoiceStatus, isnull(convert(varchar, I.ServiceStartDate, 101),'None') as ServiceStartDate
, isnull(convert(varchar,I.Invoicedate,101), 'None') as InvoiceDate
, isnull(convert(varchar, I.ServiceEnddate, 101), 'None') as ServiceEndDate, isnull(I.CurrentCharges, 0) as CurrentCharges,
isnull(E.ExceptionDescription, 'Blank') ExceptionDescription,
A.AccountNumber, P.propertyName, C.ClientName, CT.ContractNo, V.VendorName
from Invoice I                                    
    inner join BatchDetails BD on BD.BatchDetailId =I.BatchDetailId                                    
    inner join Batch B on B.BatchId = BD.BatchId                                    
    inner join Accounts A on A.AccountId = I.AccountId                                    
    inner join Contracts CT on CT.ContractId = A.ContractId                                    
    inner join Properties P on P.PropertyId = A.PropertyId                                    
    inner join Clients C on C.clientid = p.Clientid     
    inner join Vendors V on V.VendorId = A.VendorId           
    left join Exceptions E on E.BatchDetailId = I.BatchDetailId    
where 1=1 
and (I.InvoiceStatus not in ('Audited') and I.CreatedDate >= '01/01/2021')
or (I.InvoiceStatus in ('Audited') and I.CreatedDate >= '06/01/2021')
"""
Open_Invoices_Report = pd.read_sql(OpenInvoicesReport_query, cnxn)
cnxn.close()

# Add new columns to Account_List_Report
New_Columns_Acct_List = pd.DataFrame({
    'Age': [None] * len(Account_List_Report),
    'Today': [datetime.now().date()] * len(Account_List_Report),
    'LastInvoiceDate': [None] * len(Account_List_Report),
    'Missing': [None] * len(Account_List_Report)
})
Account_List_Report = pd.concat([New_Columns_Acct_List, Account_List_Report], axis=1)

if 'Account #' in Account_List_Report.columns:
    Account_List_Report['Missing'] = ~Account_List_Report['Account #'].isin(Missing_Invoices_Report['AccountNumber'])
else:
    print("Column 'Account #' not found in 'Account List Report'.")

# Filter for missing accounts
Account_List_Report = Account_List_Report[Account_List_Report['Missing'] == True]

# Prepare Open_Invoices_Report for merging
Open_Invoices_Report['InvoiceDate'] = pd.to_datetime(Open_Invoices_Report['InvoiceDate'], format='%m/%d/%Y', errors='coerce')
Open_Invoices_Report = Open_Invoices_Report.dropna(subset=['InvoiceDate']).sort_values(['AccountNumber', 'InvoiceDate'], ascending=[True, False])
Open_Invoices_Report = Open_Invoices_Report.drop_duplicates(subset='AccountNumber', keep='first').reset_index(drop=True)


if 'Account #' in Account_List_Report.columns and 'AccountNumber' in Open_Invoices_Report.columns:
    Account_List_Report = pd.merge(
        Account_List_Report,
        Open_Invoices_Report[['AccountNumber', 'InvoiceDate']],
        left_on='Account #',
        right_on='AccountNumber',
        how='left'
    )
    Account_List_Report['LastInvoiceDate'] = Account_List_Report['InvoiceDate']
    Account_List_Report = Account_List_Report.drop(columns=['AccountNumber', 'InvoiceDate'])

# Convert LastInvoiceDate to datetime and handle errors
try:
    Account_List_Report['LastInvoiceDate'] = pd.to_datetime(Account_List_Report['LastInvoiceDate'], errors='coerce')
except Exception as e:
    print(f"Error converting to datetime: {e}")
    sample_date = Account_List_Report['LastInvoiceDate'].iloc[0]
    if '/' in str(sample_date):
        Account_List_Report['LastInvoiceDate'] = pd.to_datetime(Account_List_Report['LastInvoiceDate'], format='%m/%d/%Y', errors='coerce')
    elif '-' in str(sample_date):
        Account_List_Report['LastInvoiceDate'] = pd.to_datetime(Account_List_Report['LastInvoiceDate'], format='%Y-%m-%d', errors='coerce')
    else:
        print("Unable to determine date format. Please provide the correct format.")

# Filter by valid years
valid_years = [2022, 2023, 2024]
Account_List_Report['Year'] = Account_List_Report['LastInvoiceDate'].dt.year
Account_List_Report = Account_List_Report[Account_List_Report['Year'].isin(valid_years)]

# Convert Today to datetime for Age calculation
Account_List_Report['Today'] = pd.to_datetime(Account_List_Report['Today'])
Account_List_Report['Age'] = (Account_List_Report['Today'] - Account_List_Report['LastInvoiceDate']).dt.days

# Exclude rows where Age is between 0 and 29 days
Account_List_Report = Account_List_Report[(Account_List_Report['Age'] < 0) | (Account_List_Report['Age'] > 29)]

# Save to Excel file
output_path = "C:\\Users\\YamagantiMuraliKrish\\Downloads\\Account_List_Report\\Hello123.xlsx"
Account_List_Report.to_excel(output_path, index=False)
print(f"Output saved to {output_path}")
