import xlsxwriter

# define dump file locations
dumpfile1 = 'CHW1 text export.dmp'
dumpfile2 = 'CHW2 text export.dmp'
xl_outfile = 'ddc_objects.xlsx'


# define Continuum object types
alarm_object_type = 'EventEnrollment'
variable_object_type = 'InfinityNumeric'
schedule_object_type = 'Schedule'


def get_b3_object_name(row):
	return row.split()[4]

def get_b3_object_by_type(type, rows):
	"""
	When passed the dump file as a list of rows,
	return a list of all objects that match a certain type
	"""
	objects = []
	for row in rows:
		# check if row not empty
		if len(row.split()) > 0:
			# for object list, object type is first item after splitting whitespace
			if row.split()[0] == type:
				objects.append(get_b3_object_name(row))
	return objects
	
def write_list_to_sheet(worksheet, items):
	for n, item in enumerate(items):
		worksheet.write(n, 0, item)


variables = []
alarms = []
schedules = []

for dumpfile in [dumpfile1, dumpfile2]:	
	# read b3 dump file
	print(dumpfile)
	with open(dumpfile, 'rb') as f:
		data = f.readlines()
		variables.extend(get_b3_object_by_type(variable_object_type, data))
		alarms.extend(get_b3_object_by_type(alarm_object_type, data))
		schedules.extend(get_b3_object_by_type(schedule_object_type, data))

# create xl worksbook
wb = xlsxwriter.Workbook(xl_outfile)
ws_variables = wb.add_worksheet('Variables')
ws_alarms = wb.add_worksheet('Alarms')
ws_schedules = wb.add_worksheet('Schedules')

write_list_to_sheet(ws_variables, variables)
write_list_to_sheet(ws_alarms, alarms)
write_list_to_sheet(ws_schedules, schedules)
		
wb.close()
# print(data)