import argparse
import xlrd
import csv
import os
import sys
from eagle import eagle

class eagle_yardi:

	def __init__(self, directory, output, args):
		self.args = args
		self.headers = ["Unit type", "Unit", "Applicant/Prospect", "Name", "Agent Name", "Event Date", "Source", "Status", "Notes", "Result/ Reason"]
		self.section_titles = ['Application Approved', 'Application Denied', 'Appointment  (Qualified)', 'Call  (Qualified)', 'Canceled Guest  (Qualified)', 'Email  (Qualified)', 'Other  (Qualified)', 'Return Visit  (Qualified)', 'Show  (Qualified)', 'Submit Application', 'Walk-In  (Qualified)']
		self.header_fields_raw = ["Property_Name", "Name", "Source", "Event Date", "Action"]
		self.header_fields = ["Property_Name", "Name", "First Name", "Last Name", "Source", "Event Date", "Action"]
		self.excel_files = []
		#self.dirpath = directory
		self.output = output
		for dirpath, dirnames, files in os.walk(directory):
			self.dirpath = dirpath
			for f in files:
				self.excel_files.append(os.path.join(dirpath, f))
		#print self.dirpath
		#for excel_file in self.excel_files:
		#	print excel_file

	def prop_name_parse(self):
		with open(self.output, 'wb') as out_file:
			writer = csv.writer(out_file)
			for excel_file in self.excel_files:
				workbook = xlrd.open_workbook(excel_file)
				worksheet = workbook.sheet_by_index(0)
				prop_name_cell = worksheet.cell(3,0).value
				header_row = worksheet.row_values(7,0)
				header_row_filtered = filter(None, header_row)
				#header_row2 = worksheet.row(7)
				print  prop_name_cell
				#print header_row
				print repr(header_row[3])
				cell = header_row[3].replace('\n','')
				print cell
				#print filter(None, header_row)
				#print header_row2
				#if worksheet.nrows >= 10:
				#	print worksheet.cell(10,5).value
				writer.writerow([prop_name_cell])
				#writer.writerow(header_row)
				#writer.writerow(header_row_filtered)
				#writer.writerow(header_row2)

	def headers_parse(self):
		with open(self.output, 'wb') as out_file:
			writer = csv.writer(out_file)
			wb_num = 0
			for excel_file in self.excel_files:
				workbook = xlrd.open_workbook(excel_file)
				worksheet = workbook.sheet_by_index(0)

				for row_index in range(worksheet.nrows):
					cell = worksheet.cell(row_index, 0).value
					cell = cell.strip()
					cell = cell.replace('\n','')
					if cell in self.headers:
						header_row = worksheet.row_values(row_index,0)
						print header_row
						print header_row[5]
						wb_num += 1
						print wb_num
						writer.writerow(header_row)

	def section_titles_parse(self):
		with open(self.output, 'wb') as out_file:
			writer = csv.writer(out_file)
			for excel_file in self.excel_files:
				workbook = xlrd.open_workbook(excel_file)
				worksheet = workbook.sheet_by_index(0)
				prop_name_cell = ""
				for row_index in range(worksheet.nrows):
					cell = worksheet.cell(row_index, 0).value
					cell = cell.strip()
					cell = cell.replace('\n','')
					row = worksheet.row_values(row_index,0)
					if 'property:' in cell.lower():
						prop_name_cell = worksheet.cell(row_index,0).value
					elif  'unit' in row[1].lower() and 'type' in row[1].lower() and 'total' in (c.lower() for c in row):
						print row
						writer.writerow([prop_name_cell] + row)

	def names_parse(self):
		#header_fields = ["Property_Name", "Name", "Source"]
		people = []
		for excel_file in self.excel_files:
			workbook = xlrd.open_workbook(excel_file)
			worksheet = workbook.sheet_by_index(0)

			name_index = 0
			event_date_index = 0
			source_index = 0
			start = False
			end = False
			prop_name_cell = ""
			person = {}

			for row_index in range(worksheet.nrows):
				cell = worksheet.cell(row_index, 0).value
				cell = cell.strip()
				cell = cell.replace('\n','')
				if 'property:' in cell.lower():
					prop_name_cell = worksheet.cell(row_index,0).value
				elif cell in self.headers:
					start = True
					header_row = worksheet.row_values(row_index,0)
					#Find the column index of the different header row names
					for i, c in enumerate(header_row):
						if 'name' in c.lower() and 'agent' not in c.lower():
							print c, i
							name_index = i
							#start = True
						elif 'event' in c.lower() and 'date' in c.lower():
							event_date_index = i
						elif 'source' in c.lower():
							source_index = i
				#The row with cell "Grand Total" is used as the flag to signal the end of name collecting.
				elif 'grand' in cell.lower() and 'total' in cell.lower():
					end = True
				elif end:
					continue
				#Begin collecting names after getting past the header row
				elif start:
					if cell in self.section_titles:
						section_title_cell = cell
					name_cell = worksheet.cell(row_index, name_index).value
					event_date_cell = worksheet.cell(row_index, event_date_index).value
					source_cell = worksheet.cell(row_index, source_index).value
					if name_cell != "":
						#writer.writerow([prop_name_cell, name_cell, source_cell])
						if self.args.raw == True:
							person = {'Property_Name': prop_name_cell, 'Name': name_cell, 'Source': source_cell, 'Event Date': event_date_cell,'Action': section_title_cell}
						else:
							person = {'Property_Name': prop_name_cell, 'Name': name_cell, 'Source': source_cell, 'Event Date': event_date_cell,'Action': section_title_cell}
							processed_name = eagle.namer(name_cell)
							person['First Name'] = processed_name[0]
							person['Last Name'] = processed_name[1]
							#person = {'Property_Name': prop_name_cell, 'Name': name_cell, 'Source': source_cell, 'Event Date': event_date_cell,'Action': section_title_cell}
						#Add person dict to people list
						people.append(person)

		with open(self.output, 'wb') as out_file:
			if self.args.raw == True:
				writer = csv.DictWriter(out_file, fieldnames=self.header_fields_raw)
				writer.writeheader()
				for person in people:
					writer.writerow(person)
			else:
				writer = csv.DictWriter(out_file, fieldnames=self.header_fields)
				writer.writeheader()
				for person in people:
					writer.writerow(person)



def main():
	parser = argparse.ArgumentParser(description='')
	parser.add_argument('-p', '--prop_name', action='store_true', default=False, help='Extract the cell containing the property name.')
	parser.add_argument('--header_row', action='store_true', default=False, help='Extract the header_row cells.')
	parser.add_argument('-s', '--section_titles', action='store_true', default=False, help='Extract the section titles row.')
	parser.add_argument('-r', '--raw', action='store_true', default=False, help='Extract the renter names unprocessed.')
	parser.add_argument('-n', '--names', action='store_true', default=False, help='Extract the renter names and processes them.')
	parser.add_argument('infile', nargs='?', type=str, help='Enter the folder path that contains the Move-In reports.')
	parser.add_argument('outfile', nargs='?', type=str, default=sys.stdout)
	args = parser.parse_args()

	if os.path.exists(args.infile):
		#eagle_yardi(args.infile, args.outfile).prop_name_parse()
		if args.prop_name:
			eagle_yardi(args.infile, args.outfile, args).prop_name_parse()
		elif args.header_row:
			eagle_yardi(args.infile, args.outfile, args).headers_parse()
		elif args.section_titles:
			eagle_yardi(args.infile, args.outfile, args).section_titles_parse()
		elif args.raw or args.names:
			eagle_yardi(args.infile, args.outfile, args).names_parse()
	else:
		print "The folder path does not exist!"
		print "Exiting the script."
		sys.exit()

if __name__ == '__main__':
	main()