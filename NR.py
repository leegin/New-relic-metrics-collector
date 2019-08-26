#!/usr/local/bin/python3
Author : Leegin Bernads T.S
#Script to fetch the details from New Relic

from nrql.api import NRQL
import pandas as pd
import gspread
from oauth2client.service_account import ServiceAccountCredentials


#Credentials for running the insight query.
nrql = NRQL()
nrql.api_key = '<NRQL-API-KEY>'
nrql.account_id = '<NR-ACCOUNT-ID>'

#Insight query to retieve the required information about the application

print("Running NRQL query and fetching the data........")

req = nrql.query("FROM NrDailyUsage SELECT count(*) WHERE productLine='APM' AND usageType='Application' FACET consumingAccountId, apmAppId, apmAppName, apmAgentVersion, apmLanguage SINCE 1 day ago LIMIT 2000")

#empty list to store all the lists(list of lists)
item = []
print("Preparing the sheet from which all the application details can be seen.......")

for k in req['facets']:
	item.append(k['name'])
	df = pd.DataFrame(item, columns =['Account_id','App_id','App_name','Agent_version','App_Language'])
	df['New_relic URL']=df.apply(lambda x:'https://rpm.newrelic.com/accounts/%s/applications/%s' % (x['Account_id'],x['App_id']),axis=1)

# Function to convert numbers to letters (1->A, 2->B, ... 26->Z, 27->AA, 28->AB...)
def numberToLetters(q):
    q = q - 1
    result = ''
    while q >= 0:
        remain = q % 26
        result = chr(remain+65) + result;
        q = q//26 - 1
    return result

print("Connecting with the google sheet...")

#Credentials to access google sheet through google api
scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
credentials = ServiceAccountCredentials.from_json_keyfile_name('<PATH OF THE JSON CREDENDTIAL FILE>', scope) #provide the path of the credential file here.
gc = gspread.authorize(credentials)

ws = gc.open_by_url('<GOOGLE-SHEET-URL>')
wks = ws.worksheet('<SHEET-NAME>')

print("Started populating the data to the google sheet....")
# columns names
columns = df.columns.values.tolist()

# selection of the range that will be updated
cell_list = wks.range('A1:'+numberToLetters(len(columns))+'1')

# modifying the values in the range
for cell in cell_list:
	val = columns[cell.col-1]
	cell.value = val

# update column names in batch
wks.update_cells(cell_list)

# Select the range that will be updated.
# Get number of lines and columns
num_lines, num_columns = df.shape

# selection of the range that will be updated
cell_list = wks.range('A2:'+numberToLetters(num_columns)+str(num_lines+1))

# modifying the values in the range
for cell in cell_list:
	val = df.iloc[cell.row-2,cell.col-1]
	cell.value = val

# Update the cell values in batch
wks.update_cells(cell_list)

print("Finished updating the sheet!")
