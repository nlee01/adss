import gspread, pprint, json
from datetime import datetime, date
from oauth2client.service_account import ServiceAccountCredentials

scope = ['https://spreadsheets.google.com/feeds']

credentials = ServiceAccountCredentials.from_json_keyfile_name('setup_json/service_account_id.json', scope)

gc = gspread.authorize(credentials)

universal = gc.open_by_key("18XWeVV0Mnsupg6b6tZ6hC6kr7a0LGueT_UpBIriWjNQ")
analytics = gc.open_by_key("1sHl3wmBSU1DjS2WYY3e4WP41kbAJv33gdDt_sM3HHNk")
start_index = 3

def parseDate(string):
	return datetime.strptime(string, '%m/%d/%Y')

def create_email_dict(spreadsheet):
	json_all = []
	email_dict = {}
	analytics_all = spreadsheet.worksheet("ALL").get_all_values()
	for row in range(1, len(analytics_all)):
		json_all.append(dict(zip(analytics_all[0], analytics_all[row])))
	for item in json_all:
		email_dict[item["Name"]] = item["Email"]
	return email_dict

def create_json_from_sheet(spreadsheet, sheet):
	sheet_data = spreadsheet.get_worksheet(sheet).get_all_values()
	new_json = []
	for row in range(1, len(sheet_data)):
		new_json.append(dict(zip(sheet_data[0], sheet_data[row])))
	return new_json

def create_attention_json(json, leniency, emails):
	name = json[1]["Name"]
	missing_json = []
	past_due_json = []
	for row in json:
		if row["Last Contact"] == "":
			missing_json.append(row)
		elif (datetime.now() - parseDate(row["Last Contact"])).days > leniency:
			past_due_json.append(row)
	attention_json = [name, emails[name], past_due_json, missing_json]
	return attention_json

def execute():
	all_data = []
	email_dict = create_email_dict(analytics)
	for s in range(start_index, len(universal.worksheets())):
		all_data.append(create_attention_json(create_json_from_sheet(universal, s), 5, email_dict))
	print json.dumps(all_data, indent=4)
	print all_data
execute()
