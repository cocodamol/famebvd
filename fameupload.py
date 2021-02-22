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
import proto



path = os.getcwd()
files = os.listdir(path)
files_xlsx = [f for f in files if f[-4:] == 'xlsx']


def renamer(col):
    if any(col.endswith(suffix) for suffix in ['.1']):
        col = 'CF_' + col[:-2]
        return col
    else:
        return col


def convert(df):
    m = df.melt(['Company name', 'Inactive', 'Quoted', 'Branch', 'OwnData', 'Woco',
                 'R/O Full Postcode', 'Registered number',
                 'Primary UK SIC (2007) code', 'Latest accounts date',
                 'Stock exchange(s) listed','GUO - Country ISO code']).dropna(subset = ['value'])
    n = m['variable'].str.split('\n', expand=True).loc[:, [0, 1]]
    n.columns = ['variable', 'year']
    n['year'] = n['year'].str.replace("th GBP", '').str.strip()
    m = m.drop('variable', 1).assign(**n)
    m.fillna('n.a.', inplace = True)
    m = m.pivot_table(index = ['Company name', 'Inactive', 'Quoted', 'Branch', 'OwnData', 'Woco',
                 'R/O Full Postcode', 'Registered number',
                 'Primary UK SIC (2007) code', 'Latest accounts date',
                 'Stock exchange(s) listed','GUO - Country ISO code', 'year'], columns = 'variable', dropna = True, values = 'value', aggfunc='first')
    for col in cols:
        if col not in m.columns:
            m[col] = pd.np.nan
    return m

# def bq_upload(df):





cols = ['CF_Dividends','CF_Taxation','Cons./Uncons.','Corporation Tax','Current Assets','Current Liabilities','Deferred Tax','Deferred Taxation','Depreciation','''Directors' Remuneration''','Dividends','EBITDA','Fixed Assets','Highest Paid Director','Intangible Assets','Interest Paid','Long Term Debt','National Turnover','Number of employees','Overseas Turnover','Profit (Loss) after Tax','Profit (Loss) before Tax','Profit (Loss) for period [=Net income]','Remuneration','Shareholders Funds','Short Term Loans & Overdrafts','Taxation','Turnover']

SCOPES = ['https://www.googleapis.com/auth/cloud-platform]']
CLIENT_SECRETS_PATH = 'client_secrets.json' # Path to client_secrets.json file.

# build client, authenticate Analytics, Cloud Storage, BigQuery
os.environ["GOOGLE_APPLICATION_CREDENTIALS"] = "service_account.json" # TODO: set this to path of Google Cloud Storage service account json
bq_client = bigquery.Client.from_service_account_json('service_account.json')


def proto_message_to_dict(message: proto.Message) -> dict:
    """Helper method to parse protobuf message to dictionary."""
    return json.loads(message.__class__.to_json(message))


if __name__ == "__main__":
    print('?')
    text_file = open("output_ids.txt", "w")
    for f in files_xlsx:
        out_id = int(f.split("-")[1].split('.')[0])//1730
        outname = 'unfiltered_{}.csv'.format(out_id)
        df = pd.read_excel(f, 'Results', index_col=0)
        df = df.rename(mapper=renamer, axis ='columns')
        df = convert(df)


        df.to_csv(outname, mode='w', header='column_names')
        print('{} converted'.format(outname))

        text_file.write('{}\n'.format(out_id))
