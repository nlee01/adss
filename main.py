import gspread, pprint, json, smtplib
from datetime import datetime, date
from oauth2client.service_account import ServiceAccountCredentials
from email.MIMEMultipart import MIMEMultipart
from email.MIMEText import MIMEText
from string import Template

###############################################################################
###############################################################################
###############################################################################

scope = ['https://spreadsheets.google.com/feeds']
credentials = ServiceAccountCredentials.from_json_keyfile_name('setup_json/service_account_id.json', scope)
gc = gspread.authorize(credentials)

universal = gc.open_by_key("18XWeVV0Mnsupg6b6tZ6hC6kr7a0LGueT_UpBIriWjNQ")
analytics = gc.open_by_key("1sHl3wmBSU1DjS2WYY3e4WP41kbAJv33gdDt_sM3HHNk")
start_index = 3
final_json = []

def parseDate(string):
	try:
		return datetime.strptime(string, '%m/%d/%Y')
	except:
		return datetime.now()
	

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
			if row["Company"] != "":
				missing_json.append(row)
		elif (datetime.now() - parseDate(row["Last Contact"])).days > leniency:
			past_due_json.append(row)
	attention_json = [name, emails[name], past_due_json, missing_json]
	return attention_json


###############################################################################
###############################################################################
###############################################################################

fromaddr = "nathan.lee@thecrimson.com"

def check_empty(cell):
	if cell == "":
		return "<i>&#60;empty&#62;</i>"
	else:
		return cell

def return_past_due(clients):
	text = ""
	for row in clients:
		company = check_empty(row["Company"])
		contact_name = check_empty(row["Contact Name"])
		contact_email = check_empty(row["Email"])
		contact_phone = check_empty(row["Phone"])
		text += "<strong>" + company + "</strong> [contact: <i>" + contact_name + "</i>, " + contact_email + ", " + contact_phone + "]<br>"
	return text
def return_unlisted(clients):
	if clients == []:
		return "<br>"
	else:
		text = "<h3 style='color:#A90000'>No last contact listed for:</h3><p>"
		for row in clients:
			company = check_empty(row["Company"])
			text += "<strong>" + company + "</strong>"
		return text + "</p><br>"

def send_emails(json, email_password):
	today = datetime.now().strftime('%m/%d/%y %H:%M:%S')
	print "...starting email server..."
	server = smtplib.SMTP('smtp.gmail.com', 587)
	server.starttls()
	print "...logging in..."
	server.login(fromaddr, email_password)
	print "...sending emails..."
	for item in json:
		toaddr = item[1]
		msg = MIMEMultipart()
		msg['From'] = fromaddr
		msg['To'] = toaddr
		msg['Subject'] = "[ADSS*] %s" % today
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
		    <h3>ADSS&#09;&#09;<span style="font-size:11px; color:gray;">| Advertising Department Spreadsheet System</span></h3>
		    <table width="100%" height="2px" bgcolor="#A90000"></table>
		    <p>Hey $name, here are clients for you to bump up and contact/recontact. Don't forget to <a href='https://docs.google.com/spreadsheets/d/18XWeVV0Mnsupg6b6tZ6hC6kr7a0LGueT_UpBIriWjNQ' target='_blank'>update the spreadsheet after</a>!
		    <h2 style="color:#A90000">Past-Due Clients</h2>
		    <p>$past_due_clients</p>
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
		print "."
	server.quit()




def execute():
	ep = "Crimsonpass1"
	print "...creating email dictionary..."
	email_dict = create_email_dict(analytics)
	print "...creating final json..."
	for s in range(start_index, len(universal.worksheets())):
		final_json.append(create_attention_json(create_json_from_sheet(universal, s), 5, email_dict))
		print "."
	send_emails(final_json, ep)
	print "...done."

execute()
