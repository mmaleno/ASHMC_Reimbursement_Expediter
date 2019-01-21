#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
File: expediter.py

Author: Max Maleno [mmaleno@hmc.edu]

Last Updated: 01-21-2019

Transfers information from Reimbursement spreadsheet to the
master budget spreadsheet. To run this, you must first follow
the instructions at https://developers.google.com/sheets/api/quickstart/python
to get credentials.json and token.pickle in your cloned directory.

References:
https://developers.google.com/sheets/api/quickstart/python

"""

from __future__ import print_function
import pickle
import os.path
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request

# If modifying these scopes, delete the file token.pickle.
SCOPES = ['https://www.googleapis.com/auth/spreadsheets.readonly']

# The ID and range of a sample spreadsheet.
REIMBURSEMENT_SPREADSHEET_ID = '1Wb6eLQsM6nP8Jbwz8N1Bz-eUskrxn8ZJ4Wjr6xHYMN8'
REIMBURSEMENT_RANGE_NAME = 'Form Responses 1!B2:I'

def main():
    """Shows basic usage of the Sheets API.
    Prints values from a sample spreadsheet.
    """
    creds = None
    # The file token.pickle stores the user's access and refresh tokens, and is
    # created automatically when the authorization flow completes for the first
    # time.
    if os.path.exists('token.pickle'):
        with open('token.pickle', 'rb') as token:
            creds = pickle.load(token)
    # If there are no (valid) credentials available, let the user log in.
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                'credentials.json', SCOPES)
            creds = flow.run_local_server()
        # Save the credentials for the next run
        with open('token.pickle', 'wb') as token:
            pickle.dump(creds, token)

    service = build('sheets', 'v4', credentials=creds)

    # Call the Sheets API
    sheet = service.spreadsheets()
    result = sheet.values().get(spreadsheetId=REIMBURSEMENT_SPREADSHEET_ID,
                                range=REIMBURSEMENT_RANGE_NAME).execute()
    values = result.get('values', [])

    if not values:
        print('No data found.')
    else:
        print('Here is some data')
        for row in values:
            print(row)

if __name__ == '__main__':
    main()