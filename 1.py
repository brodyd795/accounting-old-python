import datetime, imaplib, re, pprint, pickle, os.path
from googleapiclient.discovery import build
from google.oauth2 import service_account

now = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
month_year = re.findall(r'\d{4}-\d{2}', now)[0]
date = re.findall(r'\d{4}-\d{2}-\d{2}', now)[0]

SCOPES = ['https://www.googleapis.com/auth/spreadsheets']
SERVICE_ACCOUNT_FILE = 'project-1.json'
creds = service_account.Credentials.from_service_account_file(SERVICE_ACCOUNT_FILE, scopes=SCOPES)

with open('spreadsheet.txt', 'r') as f:
	lines = f.readlines()
	spreadsheet_id = lines[0].rstrip()
f.close()

range_get_existing = '{}!A2:C6'.format(month_year)

def read():
	with open("credentials.txt", "r") as f:
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
    service = build('sheets', 'v4', credentials=creds)

    # Call the Sheets API
    sheet = service.spreadsheets()
    result = sheet.values().get(spreadsheetId=spreadsheet_id,
                                range=range_get_existing).execute()
    values = result.get('values', [])
    len_transactions = len(values)
    range_update = '%s!A%s:D%s' % (month_year, len_transactions+2, len_transactions+2)

    value_input_option = 'USER_ENTERED'
    insert_data_option = 'OVERWRITE'
    
    
    
    matchesDict = read()

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
    	request = service.spreadsheets().values().append(spreadsheetId=spreadsheet_id, range=range_update, valueInputOption=value_input_option, body=value_range_body)
    	response = request.execute()
    

def main():
	try:
		sheet()
	except Exception as E:
		print(E)
		
if __name__ == "__main__":
	main()
	
	
# TODO
# automatically create new sheet for a new month
# automatically get income from past month and insert into budget for next month