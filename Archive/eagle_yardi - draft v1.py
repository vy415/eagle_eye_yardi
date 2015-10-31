import argparse
import xlrd
import csv
import os
import sys


def prop_name_parse(dir, output):
	excel_files = []
	for dirpath, dirnames, files in os.walk(dir):
		for f in files:
			excel_files.append(os.path.join(dirpath, f))
	with open(output, 'wb') as out_file:
		writer = csv.writer(out_file)
		for excel_file in excel_files:
			workbook = xlrd.open_workbook(os.path.join(dirpath, excel_file))
			worksheet = workbook.sheet_by_index(0)
			prop_name_cell = worksheet.cell(3,0).value
			header_row = worksheet.row_values(7,0)
			header_row_filtered = filter(None, header_row)
			#header_row2 = worksheet.row(7)
			print  prop_name_cell
			print header_row
			print filter(None, header_row)
			#print header_row2
			if worksheet.nrows >= 10:
				print worksheet.cell(10,5).value
			writer.writerow([prop_name_cell])
			writer.writerow(header_row)
			writer.writerow(header_row_filtered)
			#writer.writerow(header_row2)


def main():
	parser = argparse.ArgumentParser(description='')
	parser.add_argument('infile', nargs='?', type=str, help='Enter the folder path that contains the Move-In reports.')
	parser.add_argument('outfile', nargs='?', type=str, default=sys.stdout)
	args = parser.parse_args()

	if os.path.exists(args.infile):
		prop_name_parse(args.infile, args.outfile)
	else:
		print "The folder path does not exist!"
		print "Exiting the script."
		sys.exit()

if __name__ == '__main__':
	main()