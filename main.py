print ">>> Don't forget to activate the [adss] venv. <<<"
import gspread, pprint, json, smtplib, sys
from datetime import datetime, date
from oauth2client.service_account import ServiceAccountCredentials
from email.MIMEMultipart import MIMEMultipart
from email.MIMEText import MIMEText
from string import Template

###############################################################################
###############################################################################
###############################################################################

scope = ['https://spreadsheets.google.com/feeds']

# relative location of the google developer service account credential (downloaded from google developer page)
credentials = ServiceAccountCredentials.from_json_keyfile_name('setup_json/service_account_id.json', scope)

gc = gspread.authorize(credentials)

# UNIVERSAL SPREADSHEET 2017 linked by the key (unique identifier in the URL)
universal = gc.open_by_key("18XWeVV0Mnsupg6b6tZ6hC6kr7a0LGueT_UpBIriWjNQ")
# Pitching Analytics 2017 linked by the key
analytics = gc.open_by_key("1sHl3wmBSU1DjS2WYY3e4WP41kbAJv33gdDt_sM3HHNk")
# NOTE: start index is 3, as in the index of the first client management spreadsheet
start_index = 4
# month_goal is the target number of clients a given associate should have listed on his/her personal sheet
month_goal = 20
# initializing global data variables for later use
final_json = []
digest_json = []
breakdown_json = []

# function to take a date in string format and return a datetime object, and returns the current time on exception
def parseDate(string):
	try:
		return datetime.strptime(string, '%m/%d/%Y')
	except:
		return datetime.now()
	
# function to create an email dictionary matching department members to their respective emails
def create_email_dict(spreadsheet, sheet, arg3):
	json_all = []
	email_dict = {}
	analytics_all = spreadsheet.worksheet(sheet).get_all_values()
	for row in range(1, len(analytics_all)):
		json_all.append(dict(zip(analytics_all[0], analytics_all[row])))
	# arg3 is either "test" or "run", with "test" specifying to use a dummy email list for testing purposes
	for item in json_all:
		if arg3 == "test":
			email_dict[item["Name"]] = item["Test Email"]
		else:
			email_dict[item["Name"]] = item["Email"]
	return email_dict

# function to take an individual client management sheet and return a zipped dictionary, with each row becoming a 
# separate item and each cell value mapped to the column definition
def create_json_from_sheet(spreadsheet, sheet):
	try:
		sheet_data = spreadsheet.get_worksheet(sheet).get_all_values()
	except:
		sheet_data = spreadsheet.worksheet(sheet).get_all_values()
	new_json = []
	for row in range(1, len(sheet_data)):
		new_json.append(dict(zip(sheet_data[0], sheet_data[row])))
	return new_json

# function to take a json (derived from the universal spreadsheet), a leniency amount (number of days with no contact),
# and an email dictionary to return a json item with the name, email, and clients of an associate
def create_attention_json(json, leniency, emails):
	name = json[1]["Name"]
	missing_json = []
	past_due_json = []
	hold_json = []
	for row in json:
		if row["Last Contact"] == "":
			if row["Company"] != "":
				missing_json.append(row)
		elif (datetime.now() - parseDate(row["Last Contact"])).days > leniency:
			if (row["Status"] != "SOLD") and (row["Status"] != "HOLD") and (row["Status"] != "hold"):
				past_due_json.append(row)
			if (row["Status"] == "HOLD") or (row["Status"] == "hold"):
				hold_json.append(row)
	attention_json = [name, emails[name], past_due_json, missing_json, hold_json]
	if attention_json[2] == [] and attention_json[3] == []:
		return []
	else:
		return attention_json
# function to take the pitching analytics sheets and return html email code for the personal breakdown
def create_breakdown_json(json_all, arg3):
	header = "<table style='margin:0px;' width='100%'><tr>"
	for associate in json_all:
		if int(associate["Companies"]) > 20:
			total = int(associate["Companies"])
		else:
			total = month_goal
		green_cell = "<td style='width:" + str(100/total) + "%; height:20px; background-color:#ADFFA8'></td>"
		green_cell_sold = "<td style='width:" + str(100/total) + "%; height:20px; background-color:#0CFF00'></td>"
		yellow_cell = "<td style='width:" + str(100/total) + "%; height:20px; background-color:#FFEAA8'></td>"
		red_cell = "<td style='width:" + str(100/total) + "%; height:20px; background-color:#FFA8A8'></td>"
		green_cell_sm = "<td style='width:" + str(100/total) + "%; height:4px; background-color:#ADFFA8'></td>"
		green_cell_sold_sm = "<td style='width:" + str(100/total) + "%; height:4px; background-color:#0CFF00'></td>"
		yellow_cell_sm = "<td style='width:" + str(100/total) + "%; height:4px; background-color:#FFEAA8'></td>"
		red_cell_sm = "<td style='width:" + str(100/total) + "%; height:4px; background-color:#FFA8A8'></td>"
		purple_cell_sm = "<td style='width:" + str(100/total) + "%; height:4px; background-color:#F000FF'></td>"
		text = header
		for item in range(0, int(associate["SOLD"])):
			text += green_cell_sold
			total += -1
		for item in range(0, (int(associate["Companies"]) - int(associate["SOLD"]))):
			text += green_cell
			total += -1
		for rest in range(0, total):
			text += red_cell
		text += "</tr><tr>"
		for item in range(0, int(associate["SOLD"])):
			text += green_cell_sold_sm
		for item in range(0, int(associate["HOLD"])):
			text += purple_cell_sm
		for item in range(0, int(associate["<5 Days"])):
			text += green_cell_sm
		for item in range(0, int(associate["5+ Days"])):
			text += red_cell_sm
		for item in range(0, int(associate["Missing"])):
			text += red_cell_sm
		for rest in range(0, total):
			text += red_cell_sm
		text += "</tr></table>"


		if arg3 == "test":
			breakdown_json.append([associate["Name"], associate["Test Email"], text])
		else:
			breakdown_json.append([associate["Name"], associate["Email"], text])
	return breakdown_json
# function to take the pitching analytics sheets and return html email code for the daily digest
def create_digest_json(json_all, json_managers, arg3):
	text = "<table style='font-size:12px; text-align:center; margin:0px;' width='100%'><tr><th>Name</th><th>Total</th><th><5</th><th>5+</th><th>||</th><th> ? </th><th>SOLD</th><th>Revenue</th></tr>"
	for associate in json_all:
		text += "<tr><td>" + associate["Name"] + "</td>" + color_total(associate["Companies"]) + color_less_5(associate["<5 Days"]) + color_greater_5(associate["5+ Days"]) + color_hold(associate["HOLD"]) + color_unknown(associate["Missing"]) + color_sold(associate["SOLD"]) + "<td><i>" + associate["Sold Revenue"] + "</i></td></tr>"
	for manager in json_managers:
		if arg3 == "test":
			digest_json.append([manager["Name"], manager["Test Email"], text + "</table>"])
		else:
			digest_json.append([manager["Name"], manager["Email"], text + "</table>"])
	return digest_json

# custom functions for coloring table cells
def green(string):
	return "<td style='background-color:#ADFFA8; text-align: center'>" + string + "</td>"
def yellow(string):
	return "<td style='background-color:#FFEAA8; text-align: center'>" + string + "</td>"
def red(string):
	return "<td style='background-color:#FFA8A8; text-align: center'>" + string + "</td>"
def purple(string):
	return "<td style='background-color:#F887FF; text-align: center'>" + string + "</td>"
def none(string):
	return "<td style='text-align: center'>" + string + "</td>"

def color_total(string):
	try:
		if int(string) >= 20:
			return green(string)
		elif int(string) > 15:
			return yellow(string)
		else:
			return red(string)
	except:
		return none(string)
def color_less_5(string):
	try:
		if int(string) > 15:
			return green(string)
		else:
			return yellow(string)
	except:
		return none(string)
def color_greater_5(string):
	try:
		if int(string) > 3:
			return red(string)
		elif int(string) > 0:
			return yellow(string)
		else:
			return green(string)
	except:
		return none(string)
def color_hold(string):
	try:
		if int(string) > 8:
			return red(string)
		elif int(string) > 0:
			return purple(string)
		else:
			return none(string)
	except:
		return none(string)
def color_unknown(string):
	try:
		if int(string) > 0:
			return red(string)
		else:
			return none(string)
	except:
		return none(string)
def color_sold(string):
	try:
		if int(string) > 3:
			return green(string)
		else:
			return none(string)
	except:
		return none(string)

###############################################################################
###############################################################################
###############################################################################

# checks if cells are empty and returns "<empty>" if empty
def check_empty(cell):
	if cell == "":
		return "<i>&#60;empty&#62;</i>"
	else:
		# str.encord('utf-8') protects against unicode type errors (/u2018 and /u2019 in particular)
		return cell.encode('utf-8')

# takes the past due clients and returns html email code for the reminder email
def return_past_due(clients):
	if clients == []:
		return ""
	else:
		text = "<h2 style='color:#A90000'>Past-Due Clients</h2>"
		for row in clients:
			company = check_empty(row["Company"])
			contact_name = check_empty(row["Contact Name"])
			contact_email = check_empty(row["Email"])
			contact_phone = check_empty(row["Phone"])
			last_contact = check_empty(row["Last Contact"])
			text += "<strong>" + company + "</strong><ul><li type='square'>Contact: " + contact_name + " (" + contact_email + ", " + contact_phone + ")</li><li type='square'>Last Contacted: " + last_contact + "</li></ul><br>"
		return text

# takes the unlisted last contact clients and returns html email code for the reminder email
def return_unlisted(clients):
	if clients == []:
		return "<br>"
	else:
		text = "<h3 style='color:#A90000'>No last contact listed for:</h3><p>"
		for row in clients:
			company = check_empty(row["Company"])
			text += "<li type='square'><strong>" + company + "</strong></li>"
		return text + "</p><br>"

# sends all reminder emails according to the json given
def send_personal_breakdown(json, server, fromaddr):
	today = datetime.now().strftime('%m/%d/%y %H:%M:%S')
	print "...sending personal breakdowns..."
	for associate in json:
		toaddr = associate[1]
		msg = MIMEMultipart()
		msg['From'] = fromaddr
		msg['To'] = toaddr
		msg['Subject'] = "[ADSS PERFORMANCE BREAKDOWN ***BETA] %s" % today
		breakdown = associate[2]
		body = """
		<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01//EN" "http://www.w3.org/TR/html4/strict.dtd">
		<html><head><meta http-equiv="Content-Type" content="text/html; charset=utf-8"><title></title>
		  <style type="text/css">
		    #outlook a {padding:0;}
		    body{width:100% !important; -webkit-text-size-adjust:100%; -ms-text-size-adjust:100%; margin:0; padding:0;} /* force default font sizes */
		    .ExternalClass {width:100%;} .ExternalClass, .ExternalClass p, .ExternalClass span, .ExternalClass font, .ExternalClass td, .ExternalClass div {line-height: 100%;} /* Hotmail */
		    a:active, a:visited, a[href^="tel"], a[href^="sms"] { text-decoration: none; color: #000001 !important; pointer-events: auto; cursor: default;}
		    table td {border-collapse: collapse;}
		  </style>
		</head>
		<body leftmargin="0" topmargin="0" marginwidth="0" marginheight="0" style="margin: 0px; padding: 0px; background-color: #FFFFFF;" bgcolor="#FFFFFF"><table bgcolor="#800000" width="100%" border="0" align="center" cellpadding="0" cellspacing="0"><tr><td><table width="100%" border="0" align="center" cellpadding="0" cellspacing="0"><tr><td valign="top" style="padding-top:4px; padding-bottom:4px; padding-left:4px; padding-right:4px;">
		<table width="100%" border="0" cellpadding="35px" cellspacing="0" bgcolor="#FFFFFF">
		  <tr>
		    <h3>ADSS&#09;&#09;<span style="font-size:11px; color:gray;">| <span style='color:#800000'>FOR ASSOCIATES</span> ***BETA</span></h3>
		    <table width="100%" height="2px" bgcolor="#800000"></table>
		    <h3 style='color:#800000; text-align: center'>PERFORMANCE BREAKDOWN: <i>$name</i></h3>
		    $breakdown
		    <p style='font-size:10px; color:gray; text-align: center'>as evaluated on $today</p>
		    <p style='text-align:center'><a style='font-size:9px' href='https://docs.google.com/spreadsheets/d/18XWeVV0Mnsupg6b6tZ6hC6kr7a0LGueT_UpBIriWjNQ' target='_blank'>UNIVERSAL SPREADSHEET</a> | <a style='font-size:9px' href='https://docs.google.com/spreadsheets/d/1sHl3wmBSU1DjS2WYY3e4WP41kbAJv33gdDt_sM3HHNk' target='_blank'>PITCHING ANALYTICS</a></p>
		    <p style='text-align: center'><a href='https://docs.google.com/a/thecrimson.com/document/d/1V0VBJYNnKoZ0kn3pZD4GUgTqFF9MXuhBkjIPoptS-ao/edit?usp=sharing' style='font-size:8px; color:gray;'><i>what's this?</i></a><p>
		    <table width="100%" height="2px" bgcolor="#800000"></table>
		    <br>
		    <span style="font-size:9px; color:gray;">[This message was sent by the Advertising Department Spreadsheet System. Let me know if this email was sent incorrectly.&#09;&#09;- Nathan]</span></p>
		  </tr>
		</table>

		</td></tr></table></td></tr></table></body></html>
		"""
		html = Template(body).substitute(name = associate[0], today = today, breakdown = breakdown)
		msg.attach(MIMEText(html, 'html'))
		text = msg.as_string()
		server.sendmail(fromaddr, toaddr, text)
		print "#"

# sends all reminder emails according to the json given
def send_reminders(json, server, fromaddr):
	today = datetime.now().strftime('%m/%d/%y %H:%M:%S')
	print "...sending reminders..."
	for item in json:
		print item[0]
		toaddr = item[1]
		msg = MIMEMultipart()
		msg['From'] = fromaddr
		msg['To'] = toaddr
		msg['Subject'] = "[ADSS REMINDER ***BETA] %s" % today
		past_due_clients = return_past_due(item[2])
		unlisted_clients = return_unlisted(item[3])
		body = """
		<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01//EN" "http://www.w3.org/TR/html4/strict.dtd">
		<html><head><meta http-equiv="Content-Type" content="text/html; charset=utf-8"><title></title>
		  <style type="text/css">
		    #outlook a {padding:0;}
		    body{width:100% !important; -webkit-text-size-adjust:100%; -ms-text-size-adjust:100%; margin:0; padding:0;} /* force default font sizes */
		    .ExternalClass {width:100%;} .ExternalClass, .ExternalClass p, .ExternalClass span, .ExternalClass font, .ExternalClass td, .ExternalClass div {line-height: 100%;} /* Hotmail */
		    a:active, a:visited, a[href^="tel"], a[href^="sms"] { text-decoration: none; color: #000001 !important; pointer-events: auto; cursor: default;}
		    table td {border-collapse: collapse;}
		  </style>
		</head>
		<body leftmargin="0" topmargin="0" marginwidth="0" marginheight="0" style="margin: 0px; padding: 0px; background-color: #FFFFFF;" bgcolor="#FFFFFF"><table bgcolor="#EDEDED" width="100%" border="0" align="center" cellpadding="0" cellspacing="0"><tr><td><table width="100%" border="0" align="center" cellpadding="0" cellspacing="0"><tr><td valign="top" style="padding-top:8px; padding-bottom:8px; padding-left:8px; padding-right:8px;">
		<table width="100%" border="0" cellpadding="35px" cellspacing="0" bgcolor="#FFFFFF">
		  <tr>
		    <h3>ADSS&#09;&#09;<span style="font-size:11px; color:gray;">| <span style='color:#800000'>FOR ASSOCIATES</span> ***BETA</span></h3>
		    <table width="100%" height="2px" bgcolor="#A90000"></table>
		    <p>Hey $name, here are clients for you to bump up and contact/recontact. Don't forget to <a href='https://docs.google.com/spreadsheets/d/18XWeVV0Mnsupg6b6tZ6hC6kr7a0LGueT_UpBIriWjNQ' target='_blank'>update the spreadsheet after</a>!
		    $past_due_clients
		    $unlisted_clients
		    <table width="100%" height="2px" bgcolor="#A90000"></table>
		    <br>
		    <span style="font-size:9px; color:gray;">[This message was sent by the Advertising Department Spreadsheet System. Let me know if this email was sent incorrectly.&#09;&#09;- Nathan]</span></p>
		  </tr>
		</table>

		</td></tr></table></td></tr></table></body></html>
		"""
		html = Template(body).substitute(name = item[0], past_due_clients = past_due_clients, unlisted_clients = unlisted_clients)
		msg.attach(MIMEText(html, 'html'))
		text = msg.as_string()
		server.sendmail(fromaddr, toaddr, text)
		print "x"

# sends the daily digest email to the managers only
def send_digest(json, server, fromaddr):
	today = datetime.now().strftime('%m/%d/%y %H:%M:%S')
	print "...sending digest..."
	for manager in json:
		toaddr = manager[1]
		msg = MIMEMultipart()
		msg['From'] = fromaddr
		msg['To'] = toaddr
		msg['Subject'] = "[ADSS DAILY DIGEST ***BETA] %s" % today
		digest = manager[2]
		body = """
		<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01//EN" "http://www.w3.org/TR/html4/strict.dtd">
		<html><head><meta http-equiv="Content-Type" content="text/html; charset=utf-8"><title></title>
		  <style type="text/css">
		    #outlook a {padding:0;}
		    body{width:100% !important; -webkit-text-size-adjust:100%; -ms-text-size-adjust:100%; margin:0; padding:0;} /* force default font sizes */
		    .ExternalClass {width:100%;} .ExternalClass, .ExternalClass p, .ExternalClass span, .ExternalClass font, .ExternalClass td, .ExternalClass div {line-height: 100%;} /* Hotmail */
		    a:active, a:visited, a[href^="tel"], a[href^="sms"] { text-decoration: none; color: #000001 !important; pointer-events: auto; cursor: default;}
		    table td {border-collapse: collapse;}
		  </style>
		</head>
		<body leftmargin="0" topmargin="0" marginwidth="0" marginheight="0" style="margin: 0px; padding: 0px; background-color: #FFFFFF;" bgcolor="#FFFFFF"><table bgcolor="#800000" width="100%" border="0" align="center" cellpadding="0" cellspacing="0"><tr><td><table width="100%" border="0" align="center" cellpadding="0" cellspacing="0"><tr><td valign="top" style="padding-top:4px; padding-bottom:4px; padding-left:4px; padding-right:4px;">
		<table width="100%" border="0" cellpadding="35px" cellspacing="0" bgcolor="#FFFFFF">
		  <tr>
		    <h3>ADSS&#09;&#09;<span style="font-size:11px; color:gray;">| <span style='color:#800000'>FOR MANAGERS</span> ***BETA</span></h3>
		    <table width="100%" height="2px" bgcolor="#800000"></table>
		    <p>Hey $name, here's your Ads daily digest.<br><a style='font-size:9px' href='https://docs.google.com/spreadsheets/d/18XWeVV0Mnsupg6b6tZ6hC6kr7a0LGueT_UpBIriWjNQ' target='_blank'>UNIVERSAL SPREADSHEET</a> | <a style='font-size:9px' href='https://docs.google.com/spreadsheets/d/1sHl3wmBSU1DjS2WYY3e4WP41kbAJv33gdDt_sM3HHNk' target='_blank'>PITCHING ANALYTICS</a>
		    <h3 style='color:#800000; text-align: center'>DAILY DIGEST</h3><p>
		    $digest
		    <br>
		    <table width="100%" height="2px" bgcolor="#800000"></table>
		    <br>
		    <span style="font-size:9px; color:gray;">[This message was sent by the Advertising Department Spreadsheet System. Let me know if this email was sent incorrectly.&#09;&#09;- Nathan]</span></p>
		  </tr>
		</table>

		</td></tr></table></td></tr></table></body></html>
		"""
		html = Template(body).substitute(name = manager[0], digest = digest)
		msg.attach(MIMEText(html, 'html'))
		text = msg.as_string()
		server.sendmail(fromaddr, toaddr, text)
		print "#"

###############################################################################
###############################################################################
###############################################################################

def main(argv):
	fromaddr = argv[0]
	ep = argv[1]
	test_or_run = argv[3]
	print "...creating email dictionary..."
	email_dict = create_email_dict(analytics, "ALL", test_or_run)
	print "...creating final json..."
	for s in range(start_index, len(universal.worksheets())):
		j = create_attention_json(create_json_from_sheet(universal, s), 5, email_dict)
		if j != []:
			final_json.append(j)
			print "x"
		print "."
	print "...starting email server..."
	server = smtplib.SMTP('smtp.gmail.com', 587)
	server.starttls()
	print "...logging in..."
	server.login(fromaddr, ep)
	if argv[2] == "remind":
		send_reminders(final_json, server, fromaddr)
	elif argv[2] == "digest":
		send_digest(create_digest_json(create_json_from_sheet(analytics, "ALL"), create_json_from_sheet(analytics, "Managers"), test_or_run), server, fromaddr)
	elif argv[2] == "breakdown":
		send_personal_breakdown(create_breakdown_json(create_json_from_sheet(analytics, "ALL"), test_or_run), server, fromaddr)
	elif argv[2] == "all":
		send_reminders(final_json, server, fromaddr)
		send_digest(create_digest_json(create_json_from_sheet(analytics, "ALL"), create_json_from_sheet(analytics, "Managers"), test_or_run), server, fromaddr)
		send_personal_breakdown(create_breakdown_json(create_json_from_sheet(analytics, "ALL"), test_or_run), server, fromaddr)
	else:
		print "invalid argv[2] provided. no emails sent."
	server.quit()
	print "...done."

if __name__ == "__main__":
	main(sys.argv[1:])
	# try:
	# 	main(sys.argv[1:])
	# except:
	# 	print ">>> script failed to execute. did you pass the right console arguments?"
	# 	print ">>> usage: python main.py [your email] [your email password] [remind/digest/all] [test/run]"
	# 	print ">>> exiting..."