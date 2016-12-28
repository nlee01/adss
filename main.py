import gspread, pprint, json
from datetime import datetime
from oauth2client.service_account import ServiceAccountCredentials

scope = ['https://spreadsheets.google.com/feeds']

credentials = ServiceAccountCredentials.from_json_keyfile_name('service_account_id.json', scope)

gc = gspread.authorize(credentials)

sh = gc.open_by_key("18XWeVV0Mnsupg6b6tZ6hC6kr7a0LGueT_UpBIriWjNQ")

sheets = sh.worksheets()
nathan = sh.worksheet("Nathan").get_all_values()

def parseDate(string):
	return datetime.strptime(string, '%m/%d/%Y')

new = []

for row in range(1, len(nathan)):
	new.append(dict(zip(nathan[0], nathan[row])))

print json.dumps(new, indent=4)
print (new[0]["Company"])