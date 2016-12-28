import gspread, pprint, json, smtplib
from datetime import datetime, date
from oauth2client.service_account import ServiceAccountCredentials
from email.MIMEMultipart import MIMEMultipart
from email.MIMEText import MIMEText
from string import Template

scope = ['https://spreadsheets.google.com/feeds']

credentials = ServiceAccountCredentials.from_json_keyfile_name('setup_json/service_account_id.json', scope)

gc = gspread.authorize(credentials)

universal = gc.open_by_key("18XWeVV0Mnsupg6b6tZ6hC6kr7a0LGueT_UpBIriWjNQ")
analytics = gc.open_by_key("1sHl3wmBSU1DjS2WYY3e4WP41kbAJv33gdDt_sM3HHNk")
start_index = 3
final_json = []

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
	email_dict = create_email_dict(analytics)
	for s in range(start_index, len(universal.worksheets())):
		final_json.append(create_attention_json(create_json_from_sheet(universal, s), 5, email_dict))
	print json.dumps(final_json, indent=4)

# execute()

###############################################################################
###############################################################################
###############################################################################

# def send_emails(json, host_email, host_password):

fromaddr = "nathan.lee@thecrimson.com"
toaddr = "nlee01@college.harvard.edu"
msg = MIMEMultipart()
msg['From'] = fromaddr
msg['To'] = toaddr
msg['Subject'] = "Test CSS3"
variable = "test variable here"
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
    <h3>ADSS</h3>
    <table width="100%" height="2px" bgcolor="#A90000"></table>
    <h2 style="color:#A90000">Past-Due Clients</h2>
    <p>$test</p>
    <br>
    <h3 style="color:#A90000">No last contact listed for:</h3>
    <p>$test2</p>
  </tr>
</table>

</td></tr></table></td></tr></table></body></html>
"""
html = Template(body).substitute(test = "test variable here", test2 = variable)
msg.attach(MIMEText(html, 'html'))
 
server = smtplib.SMTP('smtp.gmail.com', 587)
server.starttls()
server.login(fromaddr, "Crimsonpass1")
text = msg.as_string()
server.sendmail(fromaddr, toaddr, text)
server.quit()
