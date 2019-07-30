import datetime, imaplib, re, pprint, pickle, os.path
from googleapiclient.discovery import build
from google.oauth2 import service_account
import requests

now = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
month_year = re.findall(r'\d{4}-\d{2}', now)[0]
date = re.findall(r'\d{4}-\d{2}-\d{2}', now)[0]
day = re.findall(r'\d{2}$', date)[0]

SCOPES = ['https://www.googleapis.com/auth/spreadsheets']
SERVICE_ACCOUNT_FILE = '/home/pi/Desktop/accounting/project-1.json'
creds = service_account.Credentials.from_service_account_file(SERVICE_ACCOUNT_FILE, scopes=SCOPES)

with open('/home/pi/Desktop/accounting/spreadsheet.txt', 'r') as f:
    lines = f.readlines()
    spreadsheet_id = lines[0].rstrip()
f.close()

default_sheet_id = '2042501166'
range_ = '{}!A2:C6'.format(month_year)
PREV_MONTH_INCOME_range = '{}!G1'.format(month_year)

def create_sheet():
    service = build('sheets', 'v4', credentials=creds)
    all_sheet_data = service.spreadsheets().get(spreadsheetId=spreadsheet_id).execute()
    
    last_month_title = all_sheet_data["sheets"][-1]["properties"]["title"]
    past_month_income_range = "{}!G3".format(last_month_title)
    num_indices = len(all_sheet_data["sheets"])
    
    requests = []
    requests.append({
        "duplicateSheet": {
            "sourceSheetId": 2042501166,
            "insertSheetIndex": num_indices,
            "newSheetName": month_year
            }
        })
    body = {'requests': requests}
        
    # duplicate the "default" sheet so that a sheet for the new month is created
    response = service.spreadsheets().batchUpdate(spreadsheetId=spreadsheet_id, body=body).execute()
    
    # get the income from last month, store in body variable to pass into the new month's sheet later
    response2 = service.spreadsheets().values().get(spreadsheetId=spreadsheet_id, range=past_month_income_range).execute()
    last_month_income = [[response2["values"][0][0]]]
    last_month_income_body = {"values": last_month_income}
    
    # put last month's income into the "PREV_MONTH_INCOME" for the new month
    response3 = service.spreadsheets().values().update(spreadsheetId=spreadsheet_id, range=PREV_MONTH_INCOME_range, body=last_month_income_body, valueInputOption="USER_ENTERED").execute()


def read():
    with open("/home/pi/Desktop/accounting/credentials.txt", "r") as f:
        lines = f.readlines()
        username = lines[0].rstrip()
        password = lines[1].rstrip()
    f.close()
    
    # Login to INBOX
    imap = imaplib.IMAP4_SSL("imap.gmail.com", 993)
    imap.login(username, password)
    imap.select('"[Gmail]/All Mail"')

    # Print all unread messages from a certain sender of interest
    status, response = imap.search(None, '(UNSEEN)', '(FROM "%s")' % ("1800USBanks@alerts.usbank.com"))
    unread_msg_nums = response[0].split()
    da = []
    for e_id in unread_msg_nums:
        _, response = imap.fetch(e_id, '(UID BODY[TEXT])')
        da.append(response[0][1])

    # extract relevant data from emails
    matchesList = []
    for i in da:
        x = i.decode("utf-8")
        try:
            matchesList.append(re.findall(r'charged \$\d+\.\d\d at[^\.]+?\.', x)[0])
        except:
            pass
        try:
            matchesList.append(re.findall(r'Your Deposit of \$\d+\.\d\d is complete\.', x)[0])
        except:
            pass     
        try: 
            matchesList.append(re.findall(r'Your transaction of \$\d+\.\d\d is complete\.', x)[0])
        except:
            pass
    
    # extract more precise, relevant data
    matchesDict = {}
    matchesListCounter = 0
    for i in matchesList:
        transaction_amount = ''
        transaction_location = ''
        transaction_type = ''
        transaction_amount = re.findall(r'\$\d+\.\d\d', i)[0]
        if "at" in i:
            transaction_location = re.findall(r'at [^\.]+?\.', i)[0]
            transaction_location = transaction_location[3:-1]
        elif "transaction" in i:
            transaction_location = 'Automatic withdrawal'
        else:
            transaction_location = 'Deposit'
            
        if "Deposit" in i:
            transaction_type = 'credit'
        elif "charged" in i:
            transaction_type = 'debit'
        elif "transaction" in i:
            transaction_type = 'debit'
        
        matchesDict[matchesListCounter] = [transaction_amount, transaction_location, transaction_type]
        matchesListCounter += 1
        
    return(matchesDict)


def sheet():
    # get the length of the range of transactions thus far in the month, store in range_update so we can append to it
    service = build('sheets', 'v4', credentials=creds)
    sheet = service.spreadsheets()
    result = sheet.values().get(spreadsheetId=spreadsheet_id, range=range_).execute()
    values = result.get('values', [])
    len_transactions = len(values)
    range_update = '%s!A%s:D%s' % (month_year, len_transactions+2, len_transactions+2)    
    
    # call read() to get new transaction data from emails    
    matchesDict = read()

    # parse through transaction data, store relevant data in value_range_body for the request later
    for i in matchesDict:
        transaction_amount = matchesDict[i][0]
        transaction_location = matchesDict[i][1]
        transaction_type = matchesDict[i][2]
        if transaction_type == 'credit':
            value_range_body = {
                "range": range_update,
                "majorDimension": "ROWS",
                "values": [[date, transaction_amount, '', transaction_location]]
            }
        else:
            value_range_body = {
                "range": range_update,
                "majorDimension": "ROWS",
                "values": [[date, '', transaction_amount, transaction_location]]
            }         
            
        # make the request to append the new transaction data to the Google Sheet  
        request = service.spreadsheets().values().append(spreadsheetId=spreadsheet_id, range=range_update, valueInputOption='USER_ENTERED', body=value_range_body).execute()
    

def main():
    try:
        if day == "01":
            create_sheet()
        sheet()
    except Exception as E:
        print(E)        
if __name__ == "__main__":
    main()