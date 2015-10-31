import argparse
import xlrd
import csv
import os
import sys

class eagle_yardi:

	def __init__(self, directory, output):
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

	def name_collector(self):
		headers = ["Unit type", "Unit", "Applicant/Prospect", "Name", "Agent Name", "Event Date", "Source", "Status", "Notes", "Result/ Reason"]
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
					if cell in headers:
						header_row = worksheet.row_values(row_index,0)
						print header_row
						print header_row[5]
						wb_num += 1
						print wb_num
						break


def main():
	parser = argparse.ArgumentParser(description='')
	parser.add_argument('infile', nargs='?', type=str, help='Enter the folder path that contains the Move-In reports.')
	parser.add_argument('outfile', nargs='?', type=str, default=sys.stdout)
	args = parser.parse_args()

	if os.path.exists(args.infile):
		#eagle_yardi(args.infile, args.outfile).prop_name_parse()
		eagle_yardi(args.infile, args.outfile).name_collector()
	else:
		print "The folder path does not exist!"
		print "Exiting the script."
		sys.exit()

if __name__ == '__main__':
	main()