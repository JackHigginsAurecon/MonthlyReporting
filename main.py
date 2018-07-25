import time, os
import textwrap
import pandas as pd
from pandas import ExcelWriter
from pandas import ExcelFile


from reportlab.lib import colors
from reportlab.lib.enums import TA_JUSTIFY, TA_LEFT, TA_CENTER, TA_RIGHT
from reportlab.lib.pagesizes import letter
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Image, Frame, TableStyle, PageBreak, PageTemplate, NextPageTemplate
from reportlab.platypus.tables import Table
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import inch, mm

from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.desired_capabilities import DesiredCapabilities
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC





# --- UNIQUE INFORMATION ---
# FILE NAME
file_name = "Aurecon Monthly Report.pdf"

# PROJECT NAME
project_name = "Waigani National Courts Complex"

# USERNAME/PASSWORD
_username = "ellie.hubbard"
_password = "Nn5759625"

# DEFAULT DOWNLOAD DIRECTORY
download_dir = os.path.join("C:\\", "Users", "Jack.Higgins", "Downloads")

# MONTH OF REPORT
reporting_month = "May"

# RETURN ADDRESS DETAILS
company_name = "Aurecon Australasia Pty Ltd"
abn_number = "ABN 54 005 139 873"
address_parts = ["Level 14, 32 Turbot Street", "Brisbane QLD 4000"]
po_box_parts = ["Locked Bag 331", "Brisbane QLD 4001", "Australia"]

# CONTACT DETAILS
telephone_number = "<b>T   </b> +61 7 3173 8000"
fax_number = "<b>F   </b> +61 7 3173 8001"
email_address = "<b>E   </b> brisbane@aurecongroup.com"
website_url = "<b>W   </b> aurecongroup.com"

# FOOTER DETAILS
_pagenum = 1 # INITAL PAGE NUMBER
project_number = "<b>Project</b> 240594"
file_number = "<b>File</b> %s" % file_name
date_issued = "<b>31 May 2018</b>"
revision_number = "Revision 0"
page_number = "<b>Page %s</b>" % _pagenum
footer_concat = ' '.join([project_number, file_number, date_issued, revision_number, page_number])

# TABLE DETAILS
client_name = "Peddle Thorp"
copy_name = "-"
date = "31 May 2018"
subject_parts = ["Waigani National Courts Complex", "Building Services Monthly Report - May 2018"]
from_name = "Aurecon"
reference_number = "240594"
pages_number = "14"

# HEADING 1 - GENERAL DETAILS
heading1_title = "General"
heading1_line1 = "Construction visit by Aurecon to Waigani"
heading1_line2 = "Ongoing responses to RFIs and Workflows"
heading1_line3 = "Subcontractor engagement status:"
heading1_line3_sub1 = "Mechanical, electrical, hydraulics:  XYZ Construction"
heading1_line3_sub2 = "Dry Fire subcontractor: QMEC"
heading1_line3_sub3 = "Lifts: Lift subcontractor: Lift Technologies (PNG Agent for Kone)"
heading1_line3_sub4 = "Security subcontractor: Ruswin (TBC)"
heading1_line3_sub5 = "AV: XYZ Construction"

# HEADING 2 - SITE VISITS
heading2_title = "Waigani Site Visits"
heading2_line1 = "The following visits to Waigani were undertaken by Aurecon: "
heading2_line1_sub1 = "Jason MacKander"
heading2_line1_sub1_1 = "Tuesday May 15th to Thursday May 17th, 2018"
heading2_line2 = "The following meetings and workshops were attended:"
heading2_line2_sub1 = "Site inspections"
heading2_line2_sub2 = "Site meeting"
heading2_line2_sub3 = "Project reviews with I.F. Neheja & Associates (local building services consultant)"
heading2_line3 = "Site visit summary during construction phase:"
heading2_line3_sub1 = "Aurecon trips allowed in fee proposal:  70 (plus 5 during Defects Liability Period)"
heading2_line3_sub2 = "Aurecon trips taken to date: 20 (29%% of total allowance)"

# HEADING 3 - AUDIO VISUAL
heading3_title = "Audio Visual"
heading3_subtitle1 = "General"
heading3_subtitle1_line1 = "AV subcontractor: XYZ Construction"
heading3_subtitle2 = "Design " 
heading3_subtitle2_line1 = "Nil"
heading3_subtitle3 = "Construction Progress"
heading3_subtitle3_line1 = "Cast in conduits/ cable pathways have commenced on site"
heading3_subtitle4 = "Site inspections"
heading3_subtitle4_line1 = "Nil"
heading3_subtitle5 = "Contractor Submissions Status (Shop Drawings, Samples, Data)"








# GLOBAL VARIABLE FOR PAGES SIZING / MARGINS
PAGESIZE = (210 * mm, 297 * mm)
BASE_MARGIN = 10 * mm


# SET UP DOCUMENT USING SIMPLEDOCTEMPLATE
Story = []






#------------------------------------------------
# ~ function to run the Selenium WebCrawler, open up the url to Aconex, and input login details.
# driver ::	returns the driver at the current state.
# index :: the individual item (found in dl_manager that is the focus element)
def connect_to_aconex():

	url = "https://au1.aconex.com/Logon"

	
	# Selenium Script to Web Crawl Begins here
	chrome_opts = webdriver.ChromeOptions()
	prefs = {"download.default_directory" : download_dir}
	chrome_opts.add_experimental_option("prefs", prefs)

	chrome_path = os.path.join('C:\\', 'Users', 'Jack.Higgins', 'Desktop', 'Automated Aconex Workflows', 'chromedriver.exe')
	#chrome_path = os.path.join('F:\\' 'Monthly Reporting - Digital Dashboards', 'chromedriver.exe')

	driver = webdriver.Chrome(executable_path = chrome_path, chrome_options = chrome_opts) 
	#self.driver.minimize_window()
	#self.driver.set_window_position(2000, 0)

	# Navigate to the URL
	driver.get(url)

	wait = WebDriverWait(driver, 30)
	if wait:
		loginname_xpath = '//*[@id="userName"]'
		username = driver.find_element_by_xpath(loginname_xpath)
		username.send_keys(_username)

		username = driver.find_element_by_id('password')
		username.send_keys(_password)

		loginButton = driver.find_element_by_id('login')
		loginButton.click()
	
	time.sleep(1)

	# Wait until prescence of Project Changer container
	project_changer_container_xpath = '/html/body/div[1]/div/div[1]/span[2]/span'

	wait = WebDriverWait(driver, 30).until(EC.element_to_be_clickable((By.XPATH, project_changer_container_xpath)))
	if wait:
		project_changer_container = driver.find_element_by_xpath(project_changer_container_xpath)
		project_changer_container.click()

	# Find the project in the list
	_project_xpath = '//*[@title="{}"]'.format(project_name)
	_project = driver.find_element_by_xpath(_project_xpath)
	_project.click()

	time.sleep(1)

	_mail_xpath = '//*[@id="nav-bar-CORRESPONDENCE"]'
	wait = WebDriverWait(driver, 30).until(EC.element_to_be_clickable((By.XPATH, _mail_xpath)))
	if wait:
		# Find the Project Mail
		_mailbtn = driver.find_element_by_xpath(_mail_xpath)
		_mailbtn.click()

	time.sleep(0.1)

	_sent_xpath = '//*[@id="nav-bar-CORRESPONDENCE-CORRESPONDENCE-SEARCHINBOX"]'
	_sentbtn = driver.find_element_by_xpath(_sent_xpath)
	_sentbtn.click()

	time.sleep(1)

	# Switch to the main frame
	driver.switch_to.frame('main')

	select_dates = True

	# LOOP THROUGH ALL THE POSSIBLE ATTRIBUTES - CHAPTERS FOR THE REPORT.
	for attribute2_name in ['Audio Visual', 'Electrical']:


		# Open the advanced search button
		advanced_search_xpath = '/html/body/div[2]/div/div[2]/div[1]/div[1]/div[4]/button'
		wait = WebDriverWait(driver, 30).until(EC.element_to_be_clickable((By.XPATH, advanced_search_xpath)))
		if wait:
			advanced_search_btn = driver.find_element_by_xpath(advanced_search_xpath)
			advanced_search_btn.click()

		time.sleep(0.5)

		# ONLY PICK THE DATE RANGE ONCE
		if select_dates == True:

			# Date Input Menu
			date_menu_xpath = '/html/body/div[1]/div/div/div[2]/div[2]/div[4]/div[1]/div[2]/div[1]/input'
			wait = WebDriverWait(driver, 30).until(EC.element_to_be_clickable((By.XPATH, date_menu_xpath)))
			if wait:
				date_menu_field = driver.find_element_by_xpath(date_menu_xpath)
				date_menu_field.click()

			time.sleep(0.5)

			# Pick between as the function to choose dates:
			between_fn_xpath = '/html/body/div[1]/div/div/div[2]/div[2]/div[4]/div[1]/div[2]/div[2]/div/ul/li[1]/ul/li[1]'
			wait = WebDriverWait(driver, 30).until(EC.element_to_be_clickable((By.XPATH, between_fn_xpath)))
			if wait:
				between_fn = driver.find_element_by_xpath(between_fn_xpath)
				between_fn.click()

			time.sleep(0.5)

			# Select the correct month/day (ONLY FOR TESTING WITH JUNE)
			goback_month_left_xpath = '/html/body/div[1]/div/div/div[2]/div[2]/div[4]/div[1]/div[2]/div[2]/div[2]/ul[1]/li[1]/div/table/thead/tr[1]/th[2]'
			goback_month_right_xpath = '/html/body/div[1]/div/div/div[2]/div[2]/div[4]/div[1]/div[2]/div[2]/div[2]/ul[2]/li[1]/div/table/thead/tr[1]/th[2]'
			wait = WebDriverWait(driver, 30).until(EC.element_to_be_clickable((By.XPATH, goback_month_left_xpath)))

			if wait:
				goback_month_left = driver.find_element_by_xpath(goback_month_left_xpath)
				goback_month_right = driver.find_element_by_xpath(goback_month_right_xpath)
				for i in range(0, 2):
					goback_month_left.click()
					goback_month_right.click()
					time.sleep(0.1)

			time.sleep(0.5)

			# Day will always be the 1st
			dayone_xpath = '/html/body/div[1]/div/div/div[2]/div[2]/div[4]/div[1]/div[2]/div[2]/div[2]/ul[1]/li[1]/div/table/tbody/tr[1]/td[2]/p'
			lastday_xpath = '/html/body/div[1]/div/div/div[2]/div[2]/div[4]/div[1]/div[2]/div[2]/div[2]/ul[2]/li[1]/div/table/tbody/tr[5]/td[4]/p/span'

			wait = WebDriverWait(driver, 30).until(EC.element_to_be_clickable((By.XPATH, dayone_xpath)))
			if wait:
				dayone = driver.find_element_by_xpath(dayone_xpath)
				dayone.click()

				lastday = driver.find_element_by_xpath(lastday_xpath)
				lastday.click()

				time.sleep(0.5)

			#select_dates = False

		else:
			pass


		# Select the Attribute based on the heading
		# Remove the existing attribute if it exists:
		remove_attribute_xpath = '/html/body/div[1]/div/div/div[2]/div[2]/div[7]/div[1]/div[2]/div/div[2]'
		try:
			remove_attribute = driver.find_element_by_xpath(remove_attribute_xpath)
			remove_attribute.click()
		except Exception as NoSuchElementException:
			pass


		#attribute2_name = 'Audio Visual'
		attribute2_xpath = '/html/body/div[1]/div/div/div[2]/div[2]/div[7]/div[1]/div[2]/div/div[1]/div/div'
		wait = WebDriverWait(driver, 30).until(EC.element_to_be_clickable((By.XPATH, attribute2_xpath)))
		if wait:
			attribute2 = driver.find_element_by_xpath(attribute2_xpath)
			attribute2.click()

			time.sleep(0.5)

			attribute2_selection_xpath  = '//*[@data-value="{}"]'.format(attribute2_name)
			attribute2_selection = driver.find_element_by_xpath(attribute2_selection_xpath)
			attribute2_selection.click()

		# Click the search button
		search_button_xpath = '/html/body/div[1]/div/div/div[3]/button[1]'
		search_button = driver.find_element_by_xpath(search_button_xpath)
		search_button.click()

		time.sleep(0.5)


		if select_dates == True:
			# Add the closed_out_by_org column before generating the report, do ONLY ONCE.
			add_remove_columns_xpath = '/html/body/div[2]/div/div[2]/div[1]/div[1]/div[3]/button'
			add_remove_columns = driver.find_element_by_xpath(add_remove_columns_xpath)
			add_remove_columns.click()

			time.sleep(0.5)

			closed_out_xpath = '/html/body/div[1]/div/div/div[2]/div/div[2]/div[1]/select/option[5]'
			closed_out = driver.find_element_by_xpath(closed_out_xpath)
			closed_out.click()

			time.sleep(0.5)

			add_to_list_xpath = '/html/body/div[1]/div/div/div[2]/div/div[2]/div[2]/button[1]'
			add_to_list = driver.find_element_by_xpath(add_to_list_xpath)
			add_to_list.click()

			time.sleep(0.5)

			ok_btn_xpath = '//*[@id="ok"]'
			ok_btn = driver.find_element_by_xpath(ok_btn_xpath)
			ok_btn.click()

			time.sleep(0.5)

			select_dates = False

		# Generate a Report of the Workflow Inbox data.
		report_button_xpath = '/html/body/div[2]/div/div[2]/div[1]/div[1]/div[2]/button'
		wait = WebDriverWait(driver, 30).until(EC.element_to_be_clickable((By.XPATH, report_button_xpath)))
		if wait:
			report_button = driver.find_element_by_xpath(report_button_xpath)
			report_button.click()
			time.sleep(0.5)

		# Export to Excel
		export_to_excel_xpath = '/html/body/div[2]/div/div[2]/div[1]/div[1]/div[2]/ul/li[1]/a'
		export_to_excel = driver.find_element_by_xpath(export_to_excel_xpath)
		export_to_excel.click()

		time.sleep(0.5)

		row_per_mail_xpath = '/html/body/div[1]/div/div/div[2]/div[3]/button[2]'
		row_per_mail = driver.find_element_by_xpath(row_per_mail_xpath)
		row_per_mail.click()

		time.sleep(0.5)


		if attribute2_name != 'Electrical': # CLOSE DOWN THE EXPORT WINDOW IF NOT AT THE LAST ATTRIBUTE
			close_btn_xpath = '//*[@id="close"]'
			close_btn = driver.find_element_by_xpath(close_btn_xpath)
			close_btn.click()
		else:
			print('LAST ITERATION...')
			go_to_temp_files_xpath = '/html/body/div[1]/div/div/div[3]/div[2]/div/a[1]'
			go_to_temp_files = driver.find_element_by_xpath(go_to_temp_files)
			go_to_temp_files.click()

		time.sleep(0.5)


	# Go to the Temporary Files List once all Attributes/Disciplines are created...




	time.sleep(10)
	return driver



# RUN THE FUNCTION TO GET THE EXCEL DATA FOR EACH ATTRIBUTE/DISCIPLINE
connect_to_aconex()




def read_excel_report(file_in_name):

	file_in_path = os.path.join(download_dir, file_in_name)

	# READ THE EXCEL REPORT GENERATED FOR A SPECIFIC ATTRIBUTE AND RETURN THE DATA FORMATTED AS A LIST OF LISTS [[...],[...]]
	pandas_df = pd.read_excel(file_in_path, sheet_name = 'Mail')

	#print(pandas_df)
	# TABLE FORMAT IS [DATE, MAIL NO, ATTRIBUTE 2, SUBJECT, FROM ORGANISATION, CLOSED OUT BY ORGANISATION]:
	date_column = pandas_df[pandas_df.columns[3]]
	mailno_column = pandas_df[pandas_df.columns[1]]
	subject_column = pandas_df[pandas_df.columns[2]]
	attribute_label = pandas_df.iat[2, 1]
	attribute2 = attribute_label[attribute_label.find("Attribute 2:")+len("Attribute 2")+1 : (attribute_label.find(",Recipient")-len("Recipient Type")-5)]
	fromorg_column = pandas_df[pandas_df.columns[5]]

	print(fromorg_column)


	type_column = pandas_df[pandas_df.columns[7]]
	


	return


read_excel_report("ExportMailIn-20180725_07-55.xls")






doc = SimpleDocTemplate(
	file_name,
	pagesize = PAGESIZE,
	topMargin = BASE_MARGIN,
	leftMargin = BASE_MARGIN,
	rightMargin = BASE_MARGIN,
	bottomMargin = BASE_MARGIN,
	showBoundary = 0)

# SET UP STYLES FOR JUSTIFICATION
styles = getSampleStyleSheet()
styles.add(ParagraphStyle(name = 'Generic', alignment = TA_JUSTIFY, leading = 10))
styles.add(ParagraphStyle(name = 'TableHeaders', alignment = TA_JUSTIFY, leading = 6, fontSize = 6))
styles.add(ParagraphStyle(name = 'Justify', alignment = TA_JUSTIFY, leading = 10))
styles.add(ParagraphStyle(name = 'Bullets', alignment = TA_JUSTIFY, leading = 10, bulletColor = colors.green))
styles.add(ParagraphStyle(name = 'Headings', alignment = TA_JUSTIFY, leading = 10, bulletFontSize = 12, bulletFontName = "Helvetica-Bold"))

# LOGO
aurecon_logo = "aurecon_logo_black.png"

# --- HELPER FUNCTIONS ---
def write_line(input_text, font_size, spacing_distance, other_bullet, bold = False, prim_bullet_list = False, sec_bullet_list = False, thir_bullet_list = False):
	if bold == True:
		ptext = '<font size=%d><b>%s</b></font>' % (font_size, input_text)
	else:
		ptext = '<font size=%d>%s</font>' % (font_size, input_text)

	if prim_bullet_list == True and sec_bullet_list == False:
		Story.append(Paragraph(ptext, styles["Bullets"], bulletText = '■   '))
	elif prim_bullet_list == False and sec_bullet_list == True:
		Story.append(Paragraph(ptext, styles["Generic"], bulletText = '    -   '))
	elif thir_bullet_list == True:
		Story.append(Paragraph(ptext, styles["Generic"], bulletText = '        -   '))
	else:
		Story.append(Paragraph(ptext, styles["Headings"], bulletText = other_bullet))

	Story.append(Spacer(1, spacing_distance))
	return




# -----------------------------------------------------------------


# --- FRAME SETUP --- (TEMPLATE HEADER/FOOTER PAGE 1)
# SET UP THE 3 FRAMES FOR THE HEADER
frame_1 = Frame(doc.leftMargin, 						(doc.bottomMargin) + (doc.height)-(30*mm), 				(doc.width/3)-(6*mm), 	(doc.height-(240*mm)), 	id = 'column_1')
frame_2 = Frame(doc.leftMargin+(doc.width/3)+(3*mm), 	(doc.bottomMargin) + (doc.height)-(30*mm)+(15*mm), 		(doc.width/3)-(6*mm), 	(doc.height-(255*mm)),	id = 'column_2')
frame_3 = Frame(doc.leftMargin+(2*doc.width/3)+(6*mm),  (doc.bottomMargin) + (doc.height)-(30*mm)+(22.5*mm), 	(doc.width/3)-(6*mm), 	(doc.height-(260*mm)), 	id = 'column_3')

# SET UP THE FRAME FOR THE FOOTER
frame_4 = Frame(doc.leftMargin+(doc.width/2)+(12.5*mm),	(doc.bottomMargin),			 							(doc.width/2)-(6*mm), 	(doc.height-(267.5*mm)), id = 'footer')

# SET UP THE FRAME FOR THE TITLE
frame_title = Frame(doc.leftMargin, (doc.bottomMargin) + (doc.height)-(50*mm), (doc.width), (doc.height-(260*mm)), id = 'title')

# SET UP THE FRAME FOR MAIN TABLE
frame_table = Frame(doc.leftMargin, (doc.bottomMargin) + (doc.height)-(95*mm), (doc.width), (doc.height-(235*mm)), id = 'table')

# SET UP THE FRAME FOR HEADING 1 - GENERAL
frame_heading1 = Frame(doc.leftMargin, (doc.bottomMargin) + (doc.height)-(175*mm), (doc.width), (doc.height-(200*mm)), id = 'heading1')

# SET UP THE FRAME FOR HEADING 2 - SITE VISITS
frame_heading2 = Frame(doc.leftMargin, (doc.bottomMargin) + (doc.height)-(255*mm), (doc.width), (doc.height-(195*mm)), id = 'heading2')


# ADD THE TEMPLATE FRAMES FOR THE HEADER AND FOOTER
doc.addPageTemplates([PageTemplate(id = 'HEADFOOTPAGE1', pagesize = PAGESIZE, frames = [frame_1, frame_2, frame_3, frame_4, frame_title, frame_table, frame_heading1, frame_heading2])])





# SET UP THE FRAME FOR HEADING 3 - AUDIO VISUAL
frame_heading3 = Frame(doc.leftMargin, (doc.bottomMargin) + (doc.height)-(200*mm), (doc.width), (doc.height)-(100*mm), id = 'heading3')


# ADD THE FRAMES TO THE TEMPLACE - ON THE SECOND PAGE
doc.addPageTemplates([PageTemplate(id = 'HEADFOOTPAGE2', pagesize = PAGESIZE, frames = [frame_3, frame_4, frame_heading3])])



# ADD THE ADDRESS DETAILS TO THE STORY - IN THE FIRST FRAME
ptext = '<font size=8><b>%s</b></font>' % company_name
Story.append(Paragraph(ptext, styles["Generic"]))
ptext = '<font size=8>%s</font>' % abn_number
Story.append(Paragraph(ptext, styles["Generic"]))
Story.append(Spacer(1, 8))

for part in address_parts:
	ptext = '<font size=8>%s</font>' % part.strip()
	Story.append(Paragraph(ptext, styles["Generic"]))
Story.append(Spacer(1, 8))

for part in po_box_parts:
	ptext = '<font size=8>%s</font>' % part.strip()
	Story.append(Paragraph(ptext, styles["Generic"]))

# ADD THE CONTACT DETAILS TO THE STORY - IN THE SECOND FRAME
ptext = '<font size=8>%s</font>' % telephone_number
Story.append(Paragraph(ptext, styles["Generic"]))
ptext = '<font size=8>%s</font>' % fax_number
Story.append(Paragraph(ptext, styles["Generic"]))
ptext = '<font size=8>%s</font>' % email_address
Story.append(Paragraph(ptext, styles["Generic"]))
ptext = '<font size=8>%s</font>' % website_url
Story.append(Paragraph(ptext, styles["Generic"]))

# ADD THE LOGO TO THE STORY - IN THE THIRD FRAME
aurecon_image = Image(aurecon_logo, (doc.width/3)-(10*mm), (doc.height-(240*mm))-(25*mm))
Story.append(aurecon_image)

# ADD THE FOOTER TO THE STORY - IN THE FOOTER FRAME
ptext = '<font size=6>%s</font>' % footer_concat
Story.append(Paragraph(ptext, styles["Generic"]))











# ADD THE TITLE TO THE STORY - IN THE TITLE FRAME
ptext = '<font size=20><b>Monthly Report - Building Services</b></font>'
Story.append(Paragraph(ptext, styles["Generic"]))
Story.append(Spacer(1, 20))

# ADD THE TABLE TO THE STORY - IN THE TABLE FRAME
# SET UP THE BOLD PARAGRAPHS PRIOR TO ESTABLISHING THE DATA
client_paragraph = Paragraph('<b>%s</b>' % client_name, styles["Generic"])
from_paragraph = Paragraph('<b>%s</b>' % from_name, styles["Generic"])
copy_paragraph = Paragraph('<b>%s</b>' % copy_name, styles["Generic"])
date_paragraph = Paragraph('<b>%s</b>' % date, styles["Generic"])
reference_paragraph = Paragraph('<b>%s</b>' % reference_number, styles["Generic"])
pages_paragraph = Paragraph('<b>%s</b>' % pages_number, styles["Generic"])
subject_paragraph_1 = Paragraph('<b>%s</b>' % subject_parts[0], styles["Generic"])
subject_paragraph_2 = Paragraph('<b>%s</b>' % subject_parts[1], styles["Generic"])
			   
table_data = [
["To", 		client_paragraph,	 	"From", 		from_paragraph],
["Copy",	copy_paragraph, 		"Reference", 	reference_paragraph],
["Date", 	date_paragraph, 		"Pages", 		pages_paragraph],
["Subject", subject_paragraph_1,		"",				""],
["",		subject_paragraph_2,		"",				""],
]

table = Table(table_data, colWidths = (47.5*mm), rowHeights = (7.5*mm))
table.setStyle(TableStyle([('INNERGRID', (0, 0), (-1, -1), 0.25, colors.black),
						   ('BOX', (0, 0), (-1, -1), 0.25, colors.black),
						   ('BACKGROUND', (0, 0), (0, 4), colors.gray), 
						   ('BACKGROUND', (2, 0), (2, 2), colors.gray), 
						   ('SPAN', (1, 3), (3, 3)),
						   ('SPAN', (1, 4), (3, 4)),
						   ('SPAN', (0, 3), (0, 4)),
						  ]))
Story.append(table)

# ADD THE HEADING 1 - GENERAL TO THE STORY - IN THE HEADING 1_GENERAL FRAME
write_line(input_text = heading1_title, 		font_size = 12, spacing_distance = 12, other_bullet = "(1)    ", bold = True, prim_bullet_list = False, sec_bullet_list = False)
write_line(input_text = heading1_line1, 		font_size = 10, spacing_distance = 10, other_bullet = "", bold = False, prim_bullet_list = True, sec_bullet_list = False)
write_line(input_text = heading1_line2, 		font_size = 10, spacing_distance = 10, other_bullet = "", bold = False, prim_bullet_list = True, sec_bullet_list = False)
write_line(input_text = heading1_line3, 		font_size = 10, spacing_distance = 10, other_bullet = "", bold = False, prim_bullet_list = True, sec_bullet_list = False)
write_line(input_text = heading1_line3_sub1, 	font_size = 8, spacing_distance = 8, other_bullet = "", bold = False, prim_bullet_list = False, sec_bullet_list = True)
write_line(input_text = heading1_line3_sub2, 	font_size = 8, spacing_distance = 8, other_bullet = "", bold = False, prim_bullet_list = False, sec_bullet_list = True)
write_line(input_text = heading1_line3_sub3, 	font_size = 8, spacing_distance = 8, other_bullet = "", bold = False, prim_bullet_list = False, sec_bullet_list = True)
write_line(input_text = heading1_line3_sub4, 	font_size = 8, spacing_distance = 8, other_bullet = "", bold = False, prim_bullet_list = False, sec_bullet_list = True)
write_line(input_text = heading1_line3_sub5, 	font_size = 8, spacing_distance = 8, other_bullet = "", bold = False, prim_bullet_list = False, sec_bullet_list = True)


# DRAW A VERTICAL LINE ONCE HEADING 1 - GENERAL, IS FINISHED
table_data = [
["",	"",		"",		""],
]

table = Table(table_data, colWidths = (47.5*mm), rowHeights = (3*mm))
table.setStyle(TableStyle([('LINEBELOW', (0, 0), (-1, -1), 1, colors.black),
						  ]))
Story.append(table)
Story.append(Spacer(1, 12))

# ADD HEADING 2 - SITE VISITS TO THE STORY - IN THE HEADING 2_SITE VISITS FRAME
write_line(input_text = heading2_title, 		font_size = 12, spacing_distance = 12, other_bullet = "(2)    ", bold = True, prim_bullet_list = False, sec_bullet_list = False)
write_line(input_text = heading2_line1, 		font_size = 10, spacing_distance = 10, other_bullet = "", bold = False, prim_bullet_list = True, sec_bullet_list = False)
write_line(input_text = heading2_line1_sub1, 	font_size = 8, spacing_distance = 8, other_bullet = "", bold = False, prim_bullet_list = False, sec_bullet_list = True)
write_line(input_text = heading2_line1_sub1_1, 	font_size = 8, spacing_distance = 8, other_bullet = "", bold = False, prim_bullet_list = False, sec_bullet_list = False, thir_bullet_list = True)
write_line(input_text = heading2_line2, 		font_size = 10, spacing_distance = 10, other_bullet = "", bold = False, prim_bullet_list = True, sec_bullet_list = False)
write_line(input_text = heading2_line2_sub1, 	font_size = 8, spacing_distance = 8, other_bullet = "", bold = False, prim_bullet_list = False, sec_bullet_list = True)
write_line(input_text = heading2_line2_sub2, 	font_size = 8, spacing_distance = 8, other_bullet = "", bold = False, prim_bullet_list = False, sec_bullet_list = True)
write_line(input_text = heading2_line2_sub3, 	font_size = 8, spacing_distance = 8, other_bullet = "", bold = False, prim_bullet_list = False, sec_bullet_list = True)
write_line(input_text = heading2_line3, 		font_size = 10, spacing_distance = 10, other_bullet = "", bold = False, prim_bullet_list = True, sec_bullet_list = False)
write_line(input_text = heading2_line3_sub1, 	font_size = 8, spacing_distance = 8, other_bullet = "", bold = False, prim_bullet_list = False, sec_bullet_list = True)
write_line(input_text = heading2_line3_sub2, 	font_size = 8, spacing_distance = 8, other_bullet = "", bold = False, prim_bullet_list = False, sec_bullet_list = True)

# DRAW A VERTICAL LINE ONCE HEADING 1 - GENERAL, IS FINISHED
table_data = [
["",	"",		"",		""],
]

table = Table(table_data, colWidths = (47.5*mm), rowHeights = (3*mm))
table.setStyle(TableStyle([('LINEBELOW', (0, 0), (-1, -1), 1, colors.black),
						  ]))
Story.append(table)
Story.append(Spacer(1, 12))

Story.append(NextPageTemplate('HEADFOOTPAGE2'))
Story.append(PageBreak())
_pagenum = 2
page_number = "<b>Page %s</b>" % _pagenum
footer_concat = ' '.join([project_number, file_number, date_issued, revision_number, page_number])



# ADD THE LOGO TO THE STORY - IN THE THIRD FRAME
aurecon_image = Image(aurecon_logo, (doc.width/3)-(10*mm), (doc.height-(240*mm))-(25*mm))
Story.append(aurecon_image)

# ADD THE FOOTER TO THE STORY - IN THE FOOTER FRAME
ptext = '<font size=6>%s</font>' % footer_concat
Story.append(Paragraph(ptext, styles["Generic"]))
#Story.append(Spacer(1, 10))

# ADD HEADING 3 - AUDIO VISUAL TO THE STORY - IN THE HEADING 3_AUDIO VISUAL FRAME
write_line(input_text = heading3_title, 			font_size = 12, spacing_distance = 12, 	other_bullet = "(3)    ", 	bold = True, prim_bullet_list = False, sec_bullet_list = False)
write_line(input_text = heading3_subtitle1, 		font_size = 10, spacing_distance = 10, 	other_bullet = "(3.1)  ", 	bold = True, prim_bullet_list = False, sec_bullet_list = False)
write_line(input_text = heading3_subtitle1_line1, 	font_size = 8, spacing_distance = 8, 	other_bullet = "", 			bold = False, prim_bullet_list = True, sec_bullet_list = False)
write_line(input_text = heading3_subtitle2, 		font_size = 10, spacing_distance = 10, 	other_bullet = "(3.2)  ",	bold = True, prim_bullet_list = False, sec_bullet_list = False)
write_line(input_text = heading3_subtitle2_line1, 	font_size = 8, spacing_distance = 8, 	other_bullet = "", 			bold = False, prim_bullet_list = True, sec_bullet_list = False)
write_line(input_text = heading3_subtitle3, 		font_size = 10, spacing_distance = 10, 	other_bullet = "(3.3)  ", 	bold = True, prim_bullet_list = False, sec_bullet_list = False)
write_line(input_text = heading3_subtitle3_line1, 	font_size = 8, spacing_distance = 8, 	other_bullet = "", 			bold = False, prim_bullet_list = True, sec_bullet_list = False)
write_line(input_text = heading3_subtitle4, 		font_size = 10, spacing_distance = 10, 	other_bullet = "(3.4)  ", 	bold = True, prim_bullet_list = False, sec_bullet_list = False)
write_line(input_text = heading3_subtitle4_line1, 	font_size = 8, spacing_distance = 8, 	other_bullet = "", 			bold = False, prim_bullet_list = True, sec_bullet_list = False)
write_line(input_text = heading3_subtitle5, 		font_size = 10, spacing_distance = 10, 	other_bullet = "(3.5)  ", 	bold = True, prim_bullet_list = False, sec_bullet_list = False)

# ADD THE CONTRACTOR SUBMISSIONS TABLE
temp = "(WF-000329) Material Submittal for ELECTRICINEMA Motorized Screen"
temp2 = "China Railway Construction Engineering (PNG) Ltd"

temp = textwrap.fill(temp, width = 15)
temp2 = textwrap.fill(temp2, width = 15)

date_paragraph = Paragraph('<b><fontcolor="red">Date</fontcolor></b>', styles["TableHeaders"])




table_data = [
[date_paragraph,		"Mail No",					"Attribute 2",		"Subject",		"From Organisation",	  "Type",				   "Closed Out by Org"],
["25/05/2018",	"CRCELTD-\nWTRAN-\n000319",	"Audio Visual",		temp,			temp2,					  "Workflow Transmittal",  ""				  ],
[""		 	 ,	""						  ,	""			  ,		""			 ,	""					,	  "",					   ""				  ],
]

table = Table(table_data, colWidths = (27*mm), rowHeights = (7.5*mm))
table.setStyle(TableStyle([('INNERGRID', (0, 0), (-1, -1), 0.25, colors.black),
						   ('BOX', (0, 0), (-1, -1), 0.25, colors.black),
						   ('SPAN', (0, 1), (0, 2)),
						   ('SPAN', (1, 1), (1, 2)),

						   ('SPAN', (2, 1), (2, 2)),
						   ('SPAN', (3, 1), (3, 2)),
						   ('SPAN', (4, 1), (4, 2)),
						   ('SPAN', (5, 1), (5, 2)),
						   ('SPAN', (6, 1), (6, 2)),

						   ('FONTSIZE', (0, 0), (-1, -1), 6),
						   ('LEADING', (0, 0), (-1, -1), 6),
						  ]))
Story.append(table)





# BUILD THE DOCUMENT
doc.build(Story)




