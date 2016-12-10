import xlsxwriter
from os import listdir


class DmpfileExtractor(object):
	
	def __init__(self, dmpfile=None):
		self.b3_object_types = [
			"Report",
			"EventEnrollment",
			"EventNotification",
			"Schedule",
			"InfinityNumeric",
			"InfinityInput",
			"InfinityOutput",
			"InfinityString",
			"Group",
			"InfinityProgram",
			"InfinitySystemVariable"
		]
		self.dmpfile = dmpfile
		if self.dmpfile is not None:
			self.set_dmpfile(dmpfile)
		
	def set_dmpfile(self, dmpfile):
		self.dmpfile = dmpfile
		with open(dmpfile, 'rb') as f:
			# self.data = f.readlines()
			self.data = [line.strip() for line in f]
	
	def get_b3_objects(self, verbose=False):
		self.objects = {}
		for object_type in self.b3_object_types:
				self.objects[object_type] = self.get_b3_object_by_type(object_type)
		if verbose is True:
			self.get_b3_objects_attr()

	def get_b3_object_by_type(self, object_type):
		"""
		When passed the dump file as a list of rows,
		return a list of all objects that match a certain type
		"""
		objects = []
		for row in self.data:
			# check if row not empty
			if len(row.split()) > 0:
				# for object list, object type is first item after splitting whitespace
				if row.split()[0] == object_type:					
					objects.append(self.get_b3_object_name(row))
		return objects
		
	def get_b3_object_name(self, row):
		return row.split()[4]
		
	def get_b3_objects_attr(self):
		for object_type in self.objects:
			for object_name in self.objects[object_type]:
				self.get_b3_object_attr_by_name_type(object_name, object_type)
	
	def get_b3_object_attr_by_name_type(self, object_name, object_type):
		# print(object_name)
		object_attr = {"name": object_name}
		try:
			# get range of data representing the object attributes
			i1 = self.data.index("Object : " + object_name)
			i2 = i1 + self.data[i1:].index("EndObject")
		except ValueError:
			i1 = i2 = None
		if i1 is not None:
			if object_type in ["InfinityNumeric", "InfinityInput", "InfinityOutput"]:
				self.add_object_attr(object_attr, self.data[i1:i2], "ElecType")
				self.add_object_attr(object_attr, self.data[i1:i2], "Value")
				if object_type in ["InfinityInput", "InfinityOutput"]:
					self.add_object_attr(object_attr, self.data[i1:i2], "Channel")
				print(object_attr)
		self.objects[object_type][self.objects[object_type].index(object_name)] = object_attr
				
	def get_variable_subtype(self, object_type, is_digital):
		"""
		Returns a guess of the I/O type, eg BV (binary value), AO (analog output)
		based on the Infinity object type and whether or not the signal is digital.
		"""
		is_digital_switch = {
			True: "B",
			False: "A"
		}
		object_type_switch = {
			"InfinityNumeric": "V",
			"InfinityInput": "I",
			"InfinityOutput": "O"
		}
		return is_digital_switch[is_digital] + object_type_switch[object_type]
		
	
	def get_attr_value(self, object_data, attr_key):
		attr_value = ""
		for attr in object_data:
			if attr_key + " :" in attr:
				attr_value = attr.split(":")[-1].strip()
				break
		return attr_value
		
	def add_object_attr(self, object_attr, object_data, attr_key):
		object_attr[attr_key] = self.get_attr_value(object_data, attr_key)
		
	
	def to_excel(self, workbook="dmpfile_extraction.xlsx"):
		wb = xlsxwriter.Workbook(workbook)
		for object_type in self.b3_object_types:
			sheet = wb.add_worksheet(object_type)
			self.write_objects_to_sheet(object_type, sheet)
		wb.close()
	
	def write_objects_to_sheet(self, object_type, sheet):
		if len(self.objects[object_type]) < 1:
			pass
		else:
			headers = self.objects[object_type][0].keys()
			first_row = 0
			# write headers
			for header in headers:
				col = headers.index(header) # we are keeping order.
				sheet.write(0, col, header) # we have written first row which is the header of worksheet also.

			for n, object in enumerate(self.objects[object_type]):
				for header, value in object.items():
					col = headers.index(header)
					sheet.write(n+1, col, value)
			

if __name__ == '__main__':
	# define dump file locations
	def get_dmpfiles(path="."):
		files = listdir(path)
		dmpfiles = [file for file in files if ".dmp" in file.lower()]
		return dmpfiles
		
	dumpfiles = get_dmpfiles()
	print(dumpfiles)

	for dumpfile in dumpfiles:
		xl_outfile = 'ddc_objects_' + dumpfile.split(".")[0] + '_.xlsx'
		extractor = DmpfileExtractor(dmpfile=dumpfile)
		extractor.get_b3_objects(verbose=True)
		extractor.to_excel(workbook=xl_outfile)

