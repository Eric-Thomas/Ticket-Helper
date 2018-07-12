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
	try:
		# Create workbook
		workbook = xlsxwriter.Workbook(fileName)
		# Create worksheet
		worksheet = workbook.add_worksheet("PI Requests")
		# Print first row to spreadsheet
		create_header(workbook, worksheet)
		return worksheet
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
	headerText = ["Filter", "Tag#", "Old Tag#", "Client", "SAP ID", "Email DSA", "Request Date"]
	# Create list with column witdths
	widths = [17, 13, 13, 45, 11, 23, 14.5]
	for i in range(7):
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
	iframe = soup.find_all("iframe")
	iframeID = iframe[0]["id"]
	# Switch to iframe
	driver.switch_to.frame(driver.find_element_by_id(iframeID))
	time.sleep(3)
	soup = BeautifulSoup(driver.page_source, "html.parser")
	# Find PI Support ID
	tags = soup.find_all("a")
	for group in range(len(tags)):
		if tags[group].string != None:
			if "PI SUPPORT" in tags[group].string:
				PISupportID = tags[group]["id"]
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
	# Find iframe ID
	iframe = soup.find_all("iframe")
	iframeID = iframe[0]["id"]
	# Switch to iframe
	driver.switch_to.frame(driver.find_element_by_id(iframeID))
	time.sleep(3)
	# collect description
	descriptionArea = soup.find("textarea", {"name": "instance/description/description"})
	description = descriptionArea.string
	if software_request(description.lower()):
		process_software_request(soup, description, ticketsDict)
	# Close Ticket
	driver.switch_to_default_content()
	soup = BeautifulSoup(driver.page_source, "html.parser")
	buttons = soup.find_all("button")
	for i in range(len(buttons)):
		if buttons[i].string != None:
			if buttons[i].string == "Cancel":
				cancelID = buttons[i]["id"]
	driver.find_element_by_id(cancelID).click()
	time.sleep(5)
def find_ticket_IDs(driver):
	ticketIDs = []
	soup = BeautifulSoup(driver.page_source, "html.parser")
	# Find iframe ID
	iframe = soup.find_all("iframe")
	iframeID = iframe[0]["id"]
	# Switch to iframe
	driver.switch_to.frame(driver.find_element_by_id(iframeID))
	time.sleep(3)
	soup = BeautifulSoup(driver.page_source, "html.parser")
	tags = soup.find_all("a")
	# Process each ticket
	for ticket in range(len(tags)):
		if tags[ticket].string != None:
			if "Q" in tags[ticket].string:
				ticketIDs.append(tags[ticket]["id"])
	return ticketIDs
def software_request(description):
	isSoftwareRequest = False
	if "processbook" in description or "process book" in description or "pi process" in description or "data link" in description or "datalink" in description or "pi data" in description:
		if "software request" in description:
			if "tag" in description:
				isSoftwareRequest = True
		elif "transfer request" in description:
			if "tag" in description:
				isSoftwareRequest = True
		elif "install request" in description:
			if "tag" in description:
				isSoftwareRequest = True
	return isSoftwareRequest
def process_software_request(soup, description, ticketsDict):
	# Collect parent Quote number
	quoteTag = soup.find("button", {"name", "instance/parent.quote"})
	parentQuote = quoteTag["value"]
	ticketsDict[parentQuote] = {}
	# Check if its a processbook request, data link request, or both
	if "processbook" in description or "process book" in description or "pi process" in description:
		ticketsDict[parentQuote]["data link filter"] = "Data Link 2013"
		tag = find_tag(description)
		ticketsDict[parentQuote]["data link tag"] = tag
		oldTag = find_old_tag(description)
		ticketsDict[parentQuote]["data link old tag"] = oldTag
		client = find_client(soup)
		ticketsDict[parentQuote]["data link client"] = client
		sapID = find_sap_id(soup)
		ticketsDict[parentQuote]["data link SAP ID"] = sapID
		requestDate = datetime.date.today()
		ticketsDict[parentQuote]["request date"] = requestDate
	if "data link" in description or "datalink" in description or "pi data" in description:
		ticketsDict[parentQuote]["process book filter"] = "Win7 - ProcessBook"
		tag = find_tag(description)
		ticketsDict[parentQuote]["process book tag"] = tag
		oldTag = find_old_tag(description)
		ticketsDict[parentQuote]["process book old tag"] = oldTag
		client = find_client(soup)
		ticketsDict[parentQuote]["process book client"] = client
		sapID = find_sap_id(soup)
		ticketsDict[parentQuote]["process book SAP ID"] = sapID
		requestDate = datetime.date.today()
		ticketsDict[parentQuote]["request date"] = requestDate
def find_tag(description):
	tag = ""
	return tag
def find_old_tag(description):
	oldTag = ""
	return oldTag
def find_client(soup):
	client = ""
	return client
def find_sap_id(soup):
	sapID = ""
	return sapID
def populate_worksheet(worksheet, ticketsDict):
	quotes = ticketsDict.keys()
	# Print all data link requests
	row = 1
	for quote in quotes:
		if "data link filter" in ticketsDict[quote]:
			worksheet.write_string(row, 0, ticketsDict[quote]["data link filter"])
			worksheet.write_string(row, 1, ticketsDict[quote]["data link tag"])
			worksheet.write_string(row, 2, ticketsDict[quote]["data link old tag"])
			worksheet.write_string(row, 3, ticketsDict[quote]["data link client"])
			worksheet.write_number(row, 4, ticketsDict[quote]["data link SAP ID"])
			worksheet.write_string(row, 6, ticketsDict[quote]["data link request date"])

	# Print all process book requests
	for quote in quotes:
		if "process book filter" in ticketDicts[quote]:
			worksheet.write_string(row, 0, ticketsDict[quote]["process book filter"])
			worksheet.write_string(row, 1, ticketsDict[quote]["process book tag"])
			worksheet.write_string(row, 2, ticketsDict[quote]["process book old tag"])
			worksheet.write_string(row, 3, ticketsDict[quote]["process book client"])
			worksheet.write_number(row, 4, ticketsDict[quote]["process book SAP ID"])
			worksheet.write_string(row, 6, ticketsDict[quote]["process book request date"])


def main():
	print("\n\nTicket Helper to parse PI ProcessBook and DataLink requests into excel spreadsheet\nWritten by Eric Thomas\n\n")
	worksheet = create_excel_file()
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
	populate_worksheet(worksheet, ticketsDict)
	worksheet.close()
	driver.close()
if __name__ == "__main__":
 	main()