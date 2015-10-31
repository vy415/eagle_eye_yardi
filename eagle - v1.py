import csv
import re
from nameparser import HumanName
import argparse
from sets import Set
import lxml.html as lxml
import re
import sys
import os
import traceback
from dateutil import parser
from nameparser.config import CONSTANTS

CONSTANTS.titles.remove('guru', 'bart', 'do', 'marquis')


class eagle:
	def __init__(self, directory, args):
		self.files = []
		self.args = args
		self.keys = {}
		for root, dirs, files in os.walk(directory):
			
			for f in files:
				
				if not f.endswith("txt"): self.files.append(os.path.join(root,f))
		if args.keys != '':
			try:
				with open(self.args.keys, 'r') as r_file:
					reader = csv.DictReader(r_file, delimiter="\t")
					for row in reader:
						w_row = {}
						for key, val in dict.iteritems(row):
							w_row[key.strip().lower()] = val.strip().lower()
						if w_row['property id'] != '' or w_row['property id'] != 'x':
							self.keys[w_row['pmc-prop']] = w_row['property id']
			except Exception as inst:
				sys.stderr.write("FATAL ERROR %s. Failure to read input file %s\n" % (inst, args.keys))
		

	@staticmethod
	def namer(field):
		#pre
		if type(field) == tuple:
			w_name = re.sub('[\t\r\n]', '', ", ".join([x.encode('ascii', 'ignore') for x in field])).upper()
		else:
			w_name = re.sub('[\t\r\n]', '', field.encode('ascii', 'ignore')).upper()
		if 'ANONYMOUS' not in w_name:
			if ' FORMER ' not in w_name:
				w_name = re.split(";", w_name)[0]
			else:
				w_name = re.split(";", w_name)[1]

			w_name = re.sub("(?<=[`'/+]) | (?=['`/+])", '', w_name) #6A, 4A-C
			
			out = HumanName(w_name)
			out.middle = re.sub("^[A-Z] |^[A-Z]\. ", '', out.middle)
			if " " in out.last:
				out.last = re.sub("^[A-Z] |^[A-Z]\. ", '', out.last)
			if re.sub("^[A-Z]\.|^[A-Z]", '', out.first) == '' and len(out.middle) != 0:
				out.first, out.middle = out.middle, ""
			else:
				out.first = re.sub("^[A-Z] |^[A-Z]\. ", '', out.first)
			
			#post
			
			if out.middle.startswith("FOR ") or out.middle.startswith("- "): #7A, 1B, 3E
				out.middle = "" 

			if " FOR " in out.last:
				out.last = re.sub(" FOR .*", '', out.last)

			if len(out.last) == 0 and len(out.title) != 0: #9A
				if " " in out.first:
					out = HumanName(out.first)
				else:
					out.first, out.last = "", out.first

			if " AND " in out.middle or " & " in out.middle:
				out.last = re.split("( AND )|( & )", out.middle)[0]
				out.middle = ""
 			if "AND" in out.last or "&" in out.last:

				if out.last.startswith("AND ") or out.last.startswith("& "): #3F
					out.last = HumanName(out.last).last
				elif " AND " in out.last or " & " in out.last:
					out.last = re.sub("( AND ).*|( & ).*", '', out.last)
			out.first = re.split("( AND )|&|/|\+", out.first)[0]
			out.last = re.split("/", out.last)[0].strip()
			if len(re.sub("[^A-Z]", '', out.first)) == 1 and " " in out.last:
				out.first = out.last.split(" ")[0]
				out.last = out.last.split(" ")[1]
			out.capitalize()
			first, last = out.first, out.last
			if len(out.middle) > 0:
				if re.sub("^[A-Z]\.|^[A-Z]", '', out.middle) == '':
					out.middle = ""
				elif first.endswith("-") or out.middle.startswith("-"):
					first += out.middle
				else:
					first += " %s" % out.middle #8A-B
			if len(out.suffix) > 0:
				last += " %s" % out.suffix #2A
			return (first, last)
		else:
			name = HumanName(w_name)
			return (name.first, name.last)




	def run(self):
		headers = ["Property ID", "PMC-Prop", "First Name", "Last Name", "Phone #s", "Email", "Move-In Date", "Source 1", "Source 2"]
		writer = csv.DictWriter(self.args.outfile, fieldnames=headers, delimiter="\t")

		r_head = False
		count = 0
		for i, f in enumerate(self.files, 1):
			people = []
			
			try:
				node = lxml.parse(f)
				sys.stderr.write("Working on %s: %s of %s\n" % (f, i, len(self.files)))
				
				prop = node.xpath("//div[@class='CompanySite']//text()")[0].encode('ascii', 'ignore')

				for subnode in node.xpath("//tr[not(@class='DetailsHeader')]"):
					firsts = subnode.xpath("./td[@class='LeaseVarianceFirst_x0020_Name']/text()")
					lasts = subnode.xpath("./td[@class='LeaseVarianceLast_x0020_Name']/text()")
					f_name = subnode.xpath("./td[@class='LeaseVarianceHousehold_x0020_name']/text()")
					person = {}
					for item in zip(lasts, firsts):
						person = {'First Name': '', 'Last Name': '', 'Phone #s': '', 'Email': '',
									'Move-In Date': self.args.default, 'Source 1': '', 'Source 2': '', 'PMC-Prop': prop,
									'Property ID': self.keys.get(prop.lower(), '')}
						if self.args.raw:
							person['raw'] = re.sub('[\t\r\n]', '', item[1].strip())
							person['raw'] += ", %s" % re.sub('[\t\r\n]', '', item[0].strip())
						name = self.namer(item)
						person['First Name'] = name[0]
						person['Last Name'] = name[1]

						try:
							number = subnode.xpath("./td[@class='LeaseVarianceHome_x0020_Phone']/text()")[0].encode('ascii', 'ignore').strip()
							if len(number) > 0: person['Phone #s'] = re.sub("\D", '', number)
						except:
							pass

						if len(person['Phone #s']) == 0:
							try:
								number = subnode.xpath("./td[@class='LeaseVarianceCell_x0020_Phone']/text()")[0].encode('ascii', 'ignore').strip()
								if len(number) > 0: person['Phone #s'] = re.sub("\D", '', number)
							except:
								pass
						
						if len(person['Phone #s']) == 0:	
							try:
								number = subnode.xpath("./td[@class='LeaseVarianceWork_x0020_Phone']/text()")[0].encode('ascii', 'ignore').strip()
								if len(number) > 0: person['Phone #s'] = re.sub("\D", '', number)
							except:
								pass
						try:
							source = subnode.xpath("./td[@class='LeaseVariance_x0031_st_x0020_Phone_x0020_Type']/text()")[0].encode('ascii', 'ignore').strip().lower()
							number = subnode.xpath("./td[@class='LeaseVariance_x0031_st_x0020_Phone_x0020_Number']/text()")[0].encode('ascii', 'ignore').strip()
							if len(number) > 0 and len(source) > 0:
								if len(person['Phone #s']) > 0 and type(person['Phone #s']) == list:
									if "home" in source:
										person['Phone #s'] = [number, "home"]
									elif "cell" in source and "home" not in person['Phone #s'][1]:
										person['Phone #s'] = [number, "cell"]
								else:
									person['Phone #s'] = [number, source]


						except:
							pass
						try:
							source = subnode.xpath("./td[@class='LeaseVariance_x0032_nd_x0020_Phone_x0020_Type']/text()")[0].encode('ascii', 'ignore').strip().lower()
							number = subnode.xpath("./td[@class='LeaseVariance_x0032_nd_x0020_Phone_x0020_Number']/text()")[0].encode('ascii', 'ignore').strip()
							if len(number) > 0 and len(source) > 0:
								if len(person['Phone #s']) > 0 and type(person['Phone #s']) == list:
									if "home" in source:
										person['Phone #s'] = [number, "home"]
									elif "cell" in source and "home" not in person['Phone #s'][1]:
										person['Phone #s'] = [number, "cell"]
								else:
									person['Phone #s'] = [number, source]

				
						except:
							pass
						if type(person['Phone #s']) == list:
							person['Phone #s'] = re.sub("\D", '', person['Phone #s'][0])


						#if person['Phone #s'].endswith(";"): person['Phone #s'] = person['Phone #s'][:-1]
						try:
							email = subnode.xpath("./td[@class='LeaseVarianceE-mail']/text()")[0].encode('ascii', 'ignore').strip()
							if len(email) > 0: person['Email'] = email
						except:
							pass
						try:
							mid = subnode.xpath("./td[@class='LeaseVarianceMove-in_x0020_date']/text()")[0].encode('ascii', 'ignore').strip()
							if len(mid) > 0: person['Move-In Date'] = mid
						except:
							pass
						try:
							s1 = subnode.xpath("./td[@class='LeaseVariance_x0031_st_x0020_advertising_x0020_source']/text()")[0].encode('ascii', 'ignore').strip()
							if len(s1) > 0: person['Source 1'] = s1
						except:
							pass
						try:
							s2 = subnode.xpath("./td[@class='LeaseVariance_x0032_nd_x0020_advertising_x0020_source']/text()")[0].encode('ascii', 'ignore').strip()
							if len(s2) > 0: person['Source 2'] = s2
						except:
							pass
						try:
							s1 = subnode.xpath("./td[@class='LeaseVariancePrimary_x0020_advertising_x0020_source']/text()")[0].encode('ascii', 'ignore').strip()
							if len(s1) > 0: person['Source 1'] = s1
						except:
							pass
						try:
							s2 = subnode.xpath("./td[@class='LeaseVarianceSecondary_x0020_advertising_x0020_source']/text()")[0].encode('ascii', 'ignore').strip()
							if len(s2) > 0: person['Source 2'] = s2
						except:
							pass
						
					
					for item in f_name:
						person = {'First Name': '', 'Last Name': '', 'Phone #s': '', 'Email': '',
									'Move-In Date': self.args.default, 'Source 1': '', 'Source 2': '', 'PMC-Prop': prop,
									'Property ID': self.keys.get(prop.lower(), '')}
						
						text = item
						if self.args.raw:
							person['raw'] = re.sub('[\t\n\r]', '', text.strip())

						name = self.namer(text)
						person['First Name'] = name[0]
						person['Last Name'] = name[1]
						try:
							number = subnode.xpath("./td[@class='LeaseVarianceHome_x0020_Phone']/text()")[0].encode('ascii', 'ignore').strip()
							if len(number) > 0: person['Phone #s'] = re.sub("\D", '', number)
						except:
							pass

						if len(person['Phone #s']) == 0:
							try:
								number = subnode.xpath("./td[@class='LeaseVarianceCell_x0020_Phone']/text()")[0].encode('ascii', 'ignore').strip()
								if len(number) > 0: person['Phone #s'] = re.sub("\D", '', number)
							except:
								pass
						
						if len(person['Phone #s']) == 0:	
							try:
								number = subnode.xpath("./td[@class='LeaseVarianceWork_x0020_Phone']/text()")[0].encode('ascii', 'ignore').strip()
								if len(number) > 0: person['Phone #s'] = re.sub("\D", '', number)
							except:
								pass
						try:
							source = subnode.xpath("./td[@class='LeaseVariance_x0031_st_x0020_Phone_x0020_Type']/text()")[0].encode('ascii', 'ignore').strip().lower()
							number = subnode.xpath("./td[@class='LeaseVariance_x0031_st_x0020_Phone_x0020_Number']/text()")[0].encode('ascii', 'ignore').strip()
							if len(number) > 0 and len(source) > 0:
								if len(person['Phone #s']) > 0 and type(person['Phone #s']) == list:
									if "home" in source:
										person['Phone #s'] = [number, "home"]
									elif "cell" in source and "home" not in person['Phone #s'][1]:
										person['Phone #s'] = [number, "cell"]
								else:
									person['Phone #s'] = [number, source]


						except:
							pass
						try:
							source = subnode.xpath("./td[@class='LeaseVariance_x0032_nd_x0020_Phone_x0020_Type']/text()")[0].encode('ascii', 'ignore').strip().lower()
							number = subnode.xpath("./td[@class='LeaseVariance_x0032_nd_x0020_Phone_x0020_Number']/text()")[0].encode('ascii', 'ignore').strip()
							if len(number) > 0 and len(source) > 0:
								if len(person['Phone #s']) > 0 and type(person['Phone #s']) == list:
									if "home" in source:
										person['Phone #s'] = [number, "home"]
									elif "cell" in source and "home" not in person['Phone #s'][1]:
										person['Phone #s'] = [number, "cell"]
								else:
									person['Phone #s'] = [number, source]

				
						except:
							pass
						if type(person['Phone #s']) == list:
							person['Phone #s'] = re.sub("\D", '', person['Phone #s'][0])


						#if person['Phone #s'].endswith(";"): person['Phone #s'] = person['Phone #s'][:-1]

						try:
							email = subnode.xpath("./td[@class='LeaseVarianceE-mail']/text()")[0].encode('ascii', 'ignore').strip()
							if len(email) > 0: person['Email'] = email
						except:
							pass
						try:
							mid = subnode.xpath("./td[@class='LeaseVarianceMove-in_x0020_date']/text()")[0].encode('ascii', 'ignore').strip()
							if len(mid) > 0: person['Move-In Date'] = mid
						except:
							pass
						try:
							s1 = subnode.xpath("./td[@class='LeaseVariance_x0031_st_x0020_advertising_x0020_source']/text()")[0].encode('ascii', 'ignore').strip()
							if len(s1) > 0: person['Source 1'] = s1
						except:
							pass
						try:
							s2 = subnode.xpath("./td[@class='LeaseVariance_x0032_nd_x0020_advertising_x0020_source']/text()")[0].encode('ascii', 'ignore').strip()
							if len(s2) > 0: person['Source 2'] = s2
						except:
							pass
						try:
							s1 = subnode.xpath("./td[@class='LeaseVariancePrimary_x0020_advertising_x0020_source']/text()")[0].encode('ascii', 'ignore').strip()
							if len(s1) > 0: person['Source 1'] = s1
						except:
							pass
						try:
							s2 = subnode.xpath("./td[@class='LeaseVarianceSecondary_x0020_advertising_x0020_source']/text()")[0].encode('ascii', 'ignore').strip()
							if len(s2) > 0: person['Source 2'] = s2
						except:
							pass
						
					try:
						date = parser.parse(person['Move-In Date'])
						person['Move-In Date'] = "%d-%02d-%02d" % (date.year, date.month, date.day)
					except:
						pass
					
					if len(person) > 0: people.append(person)
			
			except Exception as inst:
				traceback.print_exception(sys.exc_info()[0], sys.exc_info()[1], sys.exc_info()[2])

			if not self.args.raw:
				for person in people:
					if not r_head:
						writer.writeheader()
						r_head = True

					s1 = person['Source 1'].lower()
					s2 = person['Source 2'].lower()
					if self.args.alist:
						if ('apartment' in s1 and 'list' in s1) or ('apartment' in s2 and 'list' in s2):
							writer.writerow(person)
					else:
						writer.writerow(person)
			else:
				 with open("RAW_(%s).csv" % (count/100000), 'a') as a_file:
				 	writer = csv.writer(a_file, delimiter="\t")
			 		for person in people:
			 			
			 			s1 = person['Source 1']
			 			s2 = person['Source 2']
			 			source = ''
			 			if len(s1) > 0: source = s1
			 			elif len(s2) > 0: source = s2

			 			writer.writerow([person['raw'].encode('ascii', 'ignore'), source, person['Property']])
			 			count += 1


			people = []


def main():
	parser = argparse.ArgumentParser(description='')

	parser.add_argument('infile', nargs='?', type=str)

	parser.add_argument('-d', '--default', nargs='?', type=str, default='')

	parser.add_argument('-k', '--keys', nargs='?', type=str, default='')

	parser.add_argument('-r', '--raw', action='store_true', default=False)

	parser.add_argument('--name', action='store_true', default=False)

	parser.add_argument('-a', '--alist', action='store_true', default=False)

	parser.add_argument('outfile', nargs='?', type=argparse.FileType('w'), default=sys.stdout)

	args = parser.parse_args()

	if args.name:
		test = eagle(args.infile, args)
		size = 0
		with open(args.infile, 'r') as r_file:
			size = len(r_file.readlines())
		with open(args.infile, 'r') as r_file:
			reader = csv.reader(r_file, delimiter="\t")
			writer = csv.writer(args.outfile)
			out = []

			for i, row in enumerate(reader, 1):
				name = test.namer(row[0])
				sys.stderr.write("Working on: %s of %s\n" % (i, size))
				writer.writerow([row[0], name[1], name[0]])

		
	else:
		eagle(args.infile, args).run()


if __name__ == '__main__':
	main()