import sys
import time
import datetime
try:
	from bs4 import BeautifulSoup
except:
	sys.exit("Error importing BeautifulSoup module")
try:
	import xlsxwriter
except (ImportError):
	sys.exit("Error importing xlsxwriter module")
try:
	from selenium import webdriver
except (ImportError):
	sys.exit("Error importing selenium module")
try:
	import getpass
except (ImportError):
	sys.exit("Error importing getpass module")
def create_excel_file():
	# Ask user for name of file
	print("Enter full path name for excel file including .xlsx extension\nIf full path isn't entered file will save in same directory as the script\n>>", end="")
	fileName = input()
	spreadsheet = {}
	try:
		# Create workbook
		workbook = xlsxwriter.Workbook(fileName)
		spreadsheet["workbook"] = workbook
		# Create worksheet
		worksheet = workbook.add_worksheet("PI Requests")
		spreadsheet["worksheet"] = worksheet
		# Print first row to spreadsheet
		create_header(workbook, worksheet)
		return spreadsheet
	except (OSError):
		sys.exit("Error creating excel file")
# Creates and formats first row of spreadsheet
def create_header(workbook, worksheet):
	# Create cell format
	cellFormat = workbook.add_format()
	cellFormat.set_fg_color("#BABABA")
	cellFormat.set_align("center")
	cellFormat.set_border_color("black")
	# Create list with header items
	headerText = ["Quote", "Filter", "Tag#", "Old Tag#", "Client", "SAP ID", "Email DSA", "Request Date"]
	# Create list with column witdths
	widths = [17, 17, 13, 15, 45, 14, 23, 14.5]
	for i in range(len(headerText)):
		worksheet.set_column(i, i, widths[i])
		worksheet.write_string(0, i, headerText[i], cellFormat)
def collect_login_info():
	print("Enter username for HP Service Manager\n>>", end='')
	username = input()
	password = getpass.getpass("Enter password (No text will show up for security of passowrd)\n>>")
	return ({"username" : username, "password": password})
def perform_login(driver, info):
	# Enter info and login
	driver.find_element_by_id("LoginUsername").send_keys(info["username"])
	driver.find_element_by_id("LoginPassword").send_keys(info["password"])
	driver.find_element_by_id("loginBtn").click()
	# allow page to load
	time.sleep(10)
def open_and_process_tickets(driver):
	soup = BeautifulSoup(driver.page_source, "html.parser")
	# Find iframe ID
	iframe = soup.find("iframe")
	iframeID = iframe["id"]
	# Switch to iframe
	driver.switch_to.frame(driver.find_element_by_id(iframeID))
	time.sleep(3)
	soup = BeautifulSoup(driver.page_source, "html.parser")
	# Find PI Support ID
	tags = soup.find_all("a")
	for group in tags:
		if group.string != None:
			if "PI SUPPORT" in group.string:
				PISupportID = group["id"]
				break
	# Open PI support group
	driver.find_element_by_id(PISupportID).click()
	# Allow tickets to load
	time.sleep(3)
	# Find Tickets
	driver.switch_to_default_content()
	ticketIDs = find_ticket_IDs(driver)
	numOfTickets = len(ticketIDs)
	ticketsDict = {}
	for ticket in range(numOfTickets):
		# Re-find ticket IDs
		# Must re-find ticket IDs because they change everytime a ticket is opened and closed
		driver.switch_to_default_content()
		ticketIDs = find_ticket_IDs(driver)
		process_ticket(driver, ticketIDs[ticket], ticketsDict)
	return ticketsDict
def process_ticket(driver, ticketID, ticketsDict):
	# Click on ticket
	driver.find_element_by_id(ticketID).click()
	time.sleep(5)
	# Get html of ticket page
	soup = BeautifulSoup(driver.page_source, "html.parser")
	# collect description
	descriptionArea = soup.find("textarea", {"name": "instance/description/description"})
	description = descriptionArea.string.lower()
	if software_request(description):
		process_software_request(soup, description, ticketsDict)
	# Close Ticket
	driver.switch_to_default_content()
	soup = BeautifulSoup(driver.page_source, "html.parser")
	buttons = soup.find_all("button")
	for b in buttons:
		if b.string != None:
			if b.string == "Cancel":
				cancelID = b["id"]
	driver.find_element_by_id(cancelID).click()
	time.sleep(5)
def find_ticket_IDs(driver):
	ticketIDs = []
	soup = BeautifulSoup(driver.page_source, "html.parser")
	# Find iframe ID
	iframe = soup.find("iframe")
	iframeID = iframe["id"]
	# Switch to iframe
	driver.switch_to.frame(driver.find_element_by_id(iframeID))
	time.sleep(3)
	soup = BeautifulSoup(driver.page_source, "html.parser")
	tags = soup.find_all("a")
	# Process each ticket
	for ticket in tags:
		if ticket.string != None:
			if "Q" in ticket.string:
				ticketIDs.append(ticket["id"])
	return ticketIDs
def software_request(description):
	isSoftwareRequest = False
	if "processbook" in description or "process book" in description or "pi process" in description or "data link" in description or "datalink" in description or "pi data" in description:
		if "tag" in description:
			isSoftwareRequest = True
		elif "request" in description:
			isSoftwareRequest = True
		elif "transfer" in description:
			isSoftwareRequest = True
	return isSoftwareRequest
def process_software_request(soup, description, ticketsDict):
	# Collect parent Quote number
	quoteTag = soup.find("input", {"name": "instance/parent.quote"})
	parentQuote = quoteTag["value"]
	requestDate = str(datetime.date.today())
	ticketsDict[parentQuote] = {}
	# Check if its a processbook request, data link request, or both
	if "data link" in description or "datalink" in description or "pi data" in description:
		ticketsDict[parentQuote]["data link filter"] = "Data Link 2013"
		tag = find_tag(description)
		ticketsDict[parentQuote]["data link tag"] = tag
		oldTag = find_old_tag(description)
		ticketsDict[parentQuote]["data link old tag"] = oldTag
		client = find_client(soup)
		ticketsDict[parentQuote]["data link client"] = client
		sapID = find_sap_id(soup)
		ticketsDict[parentQuote]["data link SAP ID"] = sapID
		ticketsDict[parentQuote]["request date"] = requestDate
	if "processbook" in description or "process book" in description or "pi process" in description:
		ticketsDict[parentQuote]["process book filter"] = "Win7 - ProcessBook"
		tag = find_tag(description)
		ticketsDict[parentQuote]["process book tag"] = tag
		oldTag = find_old_tag(description)
		ticketsDict[parentQuote]["process book old tag"] = oldTag
		client = find_client(soup)
		ticketsDict[parentQuote]["process book client"] = client
		sapID = find_sap_id(soup)
		ticketsDict[parentQuote]["process book SAP ID"] = sapID
		ticketsDict[parentQuote]["request date"] = requestDate
def find_tag(description):
	tag = "No tag found"
	whitespace = [" ", ",", "\n", "\t", ":", ";"]
	try:
		if description.count("tag") == 1:
			tag = ""
			index = description.index("tag") + 3
			while index < len(description) and description[index] in whitespace:
				index += 1
			while  index < len(description) and description[index].isdecimal():
				tag = tag + description[index]
				index += 1
			tag = "TAG" + tag
		elif "new tag" in description:
			tag = ""
			index = description.index("new tag") + 7
			while index < len(description) and description[index] in whitespace:
				index += 1
			while  index < len(description) and description[index].isdecimal():
				tag = tag + description[index]
				index += 1
			tag = "TAG" + tag
		elif "to tag" in description:
			tag = ""
			index = description.index("to tag") + 6
			while index < len(description) and description[index] in whitespace:
				index += 1
			while  index < len(description) and description[index].isdecimal():
				tag = tag + description[index]
				index += 1
			tag = "TAG" + tag
	except:
		tag = "Error finding tag"
	return tag
def find_old_tag(description):
	oldTag = "No old tag found"
	whitespace = [" ", ",", "\n", "\t", ":", ";"]
	try:
		if "old tag" in description:
			oldTag = ""
			index = description.index("old tag") + 7
			while index < len(description) and description[index] in whitespace:
				index += 1
			while  index < len(description) and description[index].isdecimal():
				oldTag = oldTag + description[index]
				index += 1
			oldTag = "TAG" + oldTag
		elif "from tag" in description:
			oldTag = ""
			index = description.index("from tag") + 8
			while index < len(description) and description[index] in whitespace:
				index += 1
			while  index < len(description) and description[index].isdecimal():
				oldTag = oldTag + description[index]
				index += 1
			oldTag = "TAG" + oldTag
	except:
		oldTag = "Error finding tag"
	return oldTag
def find_client(soup):
	client = "no client found"
	tag = soup.find("input", {"id": "X26"})
	try:
		client = tag["value"]
	except:
		pass
	return client
def find_sap_id(soup):
	sapID = "no SAP ID found"
	tag = soup.find("input", {"id": "X24"})
	try:
		sapID = tag["value"]
	except:
		pass
	return sapID
def populate_worksheet(worksheet, ticketsDict):
	quotes = ticketsDict.keys()
	# Print all data link requests
	row = 1
	for quote in quotes:
		if "data link filter" in ticketsDict[quote]:
			worksheet.write_string(row, 0, quote)
			worksheet.write_string(row, 1, ticketsDict[quote]["data link filter"])
			worksheet.write_string(row, 2, ticketsDict[quote]["data link tag"])
			worksheet.write_string(row, 3, ticketsDict[quote]["data link old tag"])
			worksheet.write_string(row, 4, ticketsDict[quote]["data link client"])
			worksheet.write_string(row, 5, ticketsDict[quote]["data link SAP ID"])
			worksheet.write_string(row, 7, ticketsDict[quote]["request date"])
			row += 1
	# Print all process book requests
	for quote in quotes:
		if "process book filter" in ticketsDict[quote]:
			worksheet.write_string(row, 0, quote)
			worksheet.write_string(row, 1, ticketsDict[quote]["process book filter"])
			worksheet.write_string(row, 2, ticketsDict[quote]["process book tag"])
			worksheet.write_string(row, 3, ticketsDict[quote]["process book old tag"])
			worksheet.write_string(row, 4, ticketsDict[quote]["process book client"])
			worksheet.write_string(row, 5, ticketsDict[quote]["process book SAP ID"])
			worksheet.write_string(row, 7, ticketsDict[quote]["request date"])
			row += 1
def main():
	print("\n\nTicket Helper to parse PI ProcessBook and DataLink requests into excel spreadsheet\nWritten by Eric Thomas\n\n")
	spreadsheet = create_excel_file()
	credentials = collect_login_info()
		# Open session on website
	try:
		driver = webdriver.Edge()
	except:
		sys.exit("Error with Edge driver\n\n\nRemember to close all instances of Edge before running program!\n")
	try:
		driver.get("https://itsm.fenetwork.com/HPSM9.33_PROD/index.do")
		time.sleep(2)
	except:
		sys.exit("Error accessing Service Manager")
	perform_login(driver, credentials)
	ticketsDict = open_and_process_tickets(driver)
	populate_worksheet(spreadsheet["worksheet"], ticketsDict)
	quotes = ticketsDict.keys()
	print("Tickets processed:")
	for quote in quotes:
		print (quote, end=", ")
	spreadsheet["workbook"].close()
	driver.close()
if __name__ == "__main__":
 	main()