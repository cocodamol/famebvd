# convert exports from famebvd into relational tables and upload to bigquery

import os
import pandas as pd
import re
import argparse
from googleapiclient.discovery import build
import httplib2
from oauth2client import client
from oauth2client import file
from oauth2client import tools
import datetime
from google.cloud import bigquery
import json


path = os.getcwd()
files = os.listdir(path)
files_xlsx = [f for f in files if f[-4:] == 'xlsx']


def renamer(col):
    # for columns with the same name pandas appends '.1',
    # but this causes issue for later melting then splitting
    # rename the .1 columns with prefix instead, in this case they are all cash flow figures so I call them'CF_'
    if any(col.endswith(suffix) for suffix in ['.1']):
        col = 'CF_' + col[:-2]
        return col
    else:
        return col


def convert(df):
    # melt, dropna when value is None so it doesn't create too many rows
    m = df.melt(['Company name', 'Inactive', 'Quoted', 'Branch', 'OwnData', 'Woco',
                 'R/O Full Postcode', 'Registered number',
                 'Primary UK SIC (2007) code', 'Latest accounts date',
                 'Stock exchange(s) listed','GUO - Country ISO code']).dropna(subset = ['value'])
    # extract year by spliting variable string by '\n'
    n = m['variable'].str.split('\n', expand=True).loc[:, [0, 1]]
    # name the two columns with the splitted results variable[0] and year[1]
    n.columns = ['variable', 'Year']
    # remove the redundant th GBP from the Year column also strip leading space
    n['Year'] = n['Year'].str.replace("th GBP", '').str.strip()
    # drop the original variable column and replace with dataframe n (variable, year)
    m = m.drop('variable', 1).assign(**n)
    # fill empty values in indexes otherwise pivot_table will fail
    m.fillna('n.a.', inplace = True)
    m = m.pivot_table(index = ['Company name', 'Inactive', 'Quoted', 'Branch', 'OwnData', 'Woco',
                 'R/O Full Postcode', 'Registered number',
                 'Primary UK SIC (2007) code', 'Latest accounts date',
                 'Stock exchange(s) listed','GUO - Country ISO code', 'Year'], columns = 'variable', dropna = True, values = 'value', aggfunc='first')
    # append empty series if does not exist
    for col in cols:
        if col not in m.columns:
            m[col] = pd.np.nan
    # reset index post pivot
    m = m.reset_index()
    # replace n.a. with None type, otherwise schema for those fields will become String
    m.replace(['n.a.'], [None], inplace = True)
    # remove special characters from column names
    m.columns = m.columns.str.replace(r"[^a-zA-Z\d\_]+", '')
    clean(m)
    return m


def clean(df):
    # in case Latestaccountsdate is not full date, append '01/01/'to it
    df['Latestaccountsdate'] = df['Latestaccountsdate'].apply(str)
    df['Latestaccountsdate'] = df.Latestaccountsdate.apply(lambda s: '01/01/' + s if len(s) == 4 else s)
    return df


def bq_upload(df):
    df.to_gbq('fame.all_Dec-2020', project_id = 'baines-hager-research', chunksize= 10000, if_exists='append', table_schema=table_schema)
    print('Uploading #{} to gbq...'.format(out_id))

cols = ['CF_Dividends','CF_Taxation','Cons./Uncons.','Corporation Tax','Current Assets','Current Liabilities','Deferred Tax','Deferred Taxation','Depreciation','''Directors' Remuneration''','Dividends','EBITDA','Fixed Assets','Highest Paid Director','Intangible Assets','Interest Paid','Long Term Debt','National Turnover','Number of employees','Overseas Turnover','Profit (Loss) after Tax','Profit (Loss) before Tax','Profit (Loss) for period [=Net income]','Remuneration','Shareholders Funds','Short Term Loans & Overdrafts','Taxation','Turnover']

SCOPES = ['https://www.googleapis.com/auth/cloud-platform]']
CLIENT_SECRETS_PATH = 'client_secrets.json' # Path to client_secrets.json file.

# build client, authenticate Analytics, Cloud Storage, BigQuery
os.environ["GOOGLE_APPLICATION_CREDENTIALS"] = "service_account.json" # TODO: set this to path of Google Cloud Storage service account json
bq_client = bigquery.Client.from_service_account_json('service_account.json')

table_schema = [{'name':'Companyname','type':'STRING'},
                {'name':'Inactive','type':'STRING'},
                {'name':'Quoted','type':'STRING'},
                {'name':'Branch','type':'STRING'},
                {'name':'OwnData','type':'STRING'},
                {'name':'Woco','type':'STRING'},
                {'name':'ROFullPostcode','type':'STRING'},
                {'name':'Registerednumber','type':'STRING'},
                {'name':'PrimaryUKSIC2007code','type':'INTEGER'},
                {'name':'Latestaccountsdate','type':'STRING'},
                {'name':'Stockexchangeslisted','type':'STRING'},
                {'name':'GUOCountryISOcode','type':'STRING'},
                {'name':'Year','type':'Integer'},
                {'name':'CF_Dividends','type':'FLOAT'},
                {'name':'CF_Taxation','type':'FLOAT'},
                {'name':'ConsUncons','type':'STRING'},
                {'name':'CorporationTax','type':'FLOAT'},
                {'name':'CurrentAssets','type':'FLOAT'},
                {'name':'CurrentLiabilities','type':'FLOAT'},
                {'name':'DeferredTax','type':'FLOAT'},
                {'name':'DeferredTaxation','type':'FLOAT'},
                {'name':'Depreciation','type':'FLOAT'},
                {'name':'DirectorsRemuneration','type':'FLOAT'},
                {'name':'Dividends','type':'FLOAT'},
                {'name':'EBITDA','type':'FLOAT'},
                {'name':'FixedAssets','type':'FLOAT'},
                {'name':'HighestPaidDirector','type':'FLOAT'},
                {'name':'IntangibleAssets','type':'FLOAT'},
                {'name':'InterestPaid','type':'FLOAT'},
                {'name':'LongTermDebt','type':'FLOAT'},
                {'name':'NationalTurnover','type':'FLOAT'},
                {'name':'Numberofemployees','type':'INTEGER'},
                {'name':'OverseasTurnover','type':'FLOAT'},
                {'name':'ProfitLossafterTax','type':'FLOAT'},
                {'name':'ProfitLossbeforeTax','type':'FLOAT'},
                {'name':'ProfitLossforperiodNetincome','type':'FLOAT'},
                {'name':'Remuneration','type':'FLOAT'},
                {'name':'ShareholdersFunds','type':'FLOAT'},
                {'name':'ShortTermLoansOverdrafts','type':'FLOAT'},
                {'name':'Taxation','type':'FLOAT'},
                {'name':'Turnover','type':'FLOAT'}]

if __name__ == "__main__":
    text_file = open("output_ids.txt", "w")
    for f in files_xlsx:
        out_id = int(f.split("-")[1].split('.')[0])//1730
        df = pd.read_excel(f, 'Results', index_col=0)
        df = df.rename(mapper=renamer, axis ='columns')
        df = convert(df)
        print('Successfully converted file #{}... '.format(out_id))
        bq_upload(df)
        text_file.write('{}\n'.format(out_id))
        text_file.flush()
        # delete local xlsx file, optional
        os.remove(f)
