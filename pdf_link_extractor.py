"""
This script extracts links from a PDF file, then prints those links to an excel file.

Author: Kai Street
Date: 18 October 2020
"""

import datetime as dt, fitz, random as rdm, time, xlsxwriter as xw
from multiprocessing.pool import ThreadPool

#Gets current date; will use this to attach current date to excel file's name
today = dt.datetime.now()
file_date = today.strftime('_%b_%d_%Y_%H_%M_%p')


def pdf_link_extractor():
	"""
	This function takes a PDF file (as a user input), extracts all URLs from the file, prints those links to an excel file and saves the file.

	Author: Kai Street
	Date: 18 October 2020
	
	Precondition: ask_pdf user input must be valid .pdf file name, not including '.pdf'

	"""

	print('Program running...')
	time.sleep(rdm.uniform(0.0,0.3))
	
	#User inputs PDF name
	ask_pdf = str(input('Enter the PDF name (DO NOT include ".pdf" at the end: '))
	ask_pdf_file = ask_pdf.replace(' ','_')
	pdf_document = fitz.open('/Users/mrkaistreet/Desktop/'+ask_pdf+'.pdf')
	
	#Empty accumulators to store all links in
	all_url_dict_lists = []
	all_urls = []

	#Search PDF for all links
	for page in pdf_document:
		page_urls = []
		page_urls = page.getLinks()
		for page_link in page_urls:
			all_url_dict_lists.append(page_link)

	#Append all links to accumulators
	for actual_url in all_url_dict_lists:
		all_urls.append(actual_url['uri'])
	
	#Create excel workbook and worksheet
	workbook = xw.Workbook('/Desktop/'+ask_pdf_file+'_pdf_links'+file_date+'.xlsx',{'strings_to_urls':False})
	worksheet1 = workbook.add_worksheet('pdf_links')
	row = 0
	col = 0
	bold = workbook.add_format({'bold':True})
	worksheet1.write(row,col,'Links',bold)
	worksheet1.freeze_panes(1,0)

	#Write links to file, then close file
	worksheet1.write_column(row+1,col,all_urls)
	workbook.close()
	time.sleep(rdm.uniform(0.0,0.3))

	#Status update on the number of links added to the file
	print(str('Number of links captured: ')+str(len(all_urls)))
	time.sleep(rdm.uniform(0.0,0.3))
	print('Program complete.')
	time.sleep(rdm.uniform(0.0,0.3))

#Ask if user wants to run extractor tool, then loop if user has multiple PDFs to scan; leverages multiprocessing to speed up output.
ask_run = str(input('Run the PDF link extraction tool? Enter y if yes or n if no: '))
while ask_run == 'y' or ask_run == 'Y':	
	pool = ThreadPool(processes=4)
	pool.apply(pdf_link_extractor)
	pool.close()
	ask_run = str(input('Want to run the PDF link extraction tool again? Enter y if yes or n if no: '))
else:
	print('Program exiting...')
	time.sleep(rdm.uniform(0.0,0.3))
	print('Program exited.')
