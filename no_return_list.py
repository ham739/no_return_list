#!/usr/bin/env python
# -*- coding: utf-8 -*-

import gspread
import configparser
from oauth2client import file, client, tools
from googleapiclient.discovery import build
import base64
from email.mime.text import MIMEText

no_return_list = []
config_file = "config.ini"

config = configparser.ConfigParser()
config.read(config_file, 'UTF-8')

def get_credentials():

    SCOPES = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive','https://www.googleapis.com/auth/gmail.send']
    store = file.Storage('token.json')
    credentials = store.get()
    if not credentials or credentials.invalid:
        flow = client.flow_from_clientsecrets('test.json', SCOPES)
        credentials = tools.run_flow(flow, store)

    return credentials

def get_value(credentials):

    gc = gspread.authorize(credentials)
    workbook = gc.open_by_url(config['config']["sheet_url"])
    worksheet = workbook.worksheet(config['config']["sheet_name"])
    cell_list = worksheet.get_all_values()

    return cell_list

def make_message(sheet_value):

    message_text = ""

    for a in sheet_value:

        no_return = a[2]
        if no_return == "未返却":
            no_return_list.append(a)

    if no_return_list:

        for b in no_return_list:
            for c in b:
                message_text = message_text + c
            message_text = message_text + "\n"

    return message_text

def send_message(credentials,message_text):

    message = MIMEText(message_text)
    message['to'] = config['config']["mail_send"]
    message['subject'] = config['config']["mail_title"]
    encode_message = base64.urlsafe_b64encode(message.as_bytes())
    encode_message = {'raw': encode_message.decode()}

    service = build('gmail', 'v1', credentials=credentials)

    message_run = (service.users().messages().send(userId="me", body=encode_message)
                   .execute())

if __name__ == '__main__':

    credentials = get_credentials()

    sheet_value = get_value(credentials)

    message_text = make_message(sheet_value)

    send_message(credentials,message_text)
