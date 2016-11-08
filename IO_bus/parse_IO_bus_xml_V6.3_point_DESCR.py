##########################
# parse_IO_bus_xml_V6.3.py
##########################
# Author: CG
# Date: 7/08/2015
# Version: 6.2
#
# export must contain a single AS only
# Excel file must have following naming:
# must contain string 'tblSXWRdata' or "_io_bus"
# must be either .xlsx or .xls extension
# eg
# "[AS_name]_io_bus.xls" or 'tblSXWRdata.xlsx'
import xml.etree.ElementTree as etree
from xml.dom import minidom
import pandas as pd
import datetime as dt
import os
import sys

#################
# Define funcions
# Generate xml element from row in IO module dataframe
def create_iomodule_element(iomodule):
	"""Return an xml element from row in IO module dataframe.
    """	
	# create element 'OI'
	element = etree.Element("OI")
	# set element attributes
	element.set('NAME', iomodule['module_name'])
	element.set('TYPE', iomodule['module_type'])
	# create subelement 'PI'
	subelement = etree.Element("PI")
	# set subelement attributes
	subelement.set('Name','ModuleID')
	subelement.set('Value',str(iomodule['module_id']))
	# insert subelement into element
	element.append(subelement)
	return element

# Generate xml element from row in points dataframe
def create_point_element(point):
	"""Return an xml element from row in IO dataframe.
    """
	# create element 'OI'
	element = etree.Element("OI")
	# set element attributes
	element.set('NAME', point['point_name'])
	element.set('TYPE', point['point_type'])
	element.set('DESCR', point['point_description'])
	# create subelement 'PI'
	subelement = etree.Element("PI")
	# subelement attributes vary if point is input or output
	if 'Output' in point['point_type']: #point["point_type"].find('Output'):
		pi_name = "OutputChannelNumber"
		pi_value = point["OutputChannelNumber"]
	else:
		pi_name = "InputChannelNumber"
		pi_value = point["InputChannelNumber"]
	# set subelement attributes
	subelement.set('Name', pi_name)
	subelement.set('Value', str(int(pi_value)))
	# insert subelement into element
	element.append(subelement)
	return element

def insert_elements(parent, points, iomodules):
	"""Insert IO module and points elements from DataFrames into parent xml element.
    """
	# migrate points list to SBO IO bus xml
	for n, iomodule in iomodules.iterrows():
		# create XML element for each IO module
		try:
			iomodule_element = create_iomodule_element(iomodule)
		except:
			print("Error: create XML element for "+iomodule['module_name']+" failed, skipping!")
			continue
		
		print(iomodule['module_name'])
		# create dataframe of points in IO module only
		module_points = df.loc[df['module_name'] == iomodule['module_name']]
		
		for n, point in module_points.iterrows():
			# create XML element for each point in IO module
			try:
				point_element = create_point_element(point)
			except:
				print("Error: create XML element for "+point['point_name']+" failed, skipping!")
				continue
			iomodule_element.append(point_element)
		# insert IO module element into parent XML
		parent.append(iomodule_element)

def get_sbo_xl_files(dir='.'):
	"""Get filenames of all matching Excel files in directory.
	"""
	xlfiles = []
	for n, filename in enumerate(next(os.walk(dir))[2]):
		if '_io_bus' in filename and '.xls' in filename:
			xlfiles.append(filename)
		elif 'tblSXWRdata'in filename and '.xls' in filename:
			xlfiles.append(filename)
	return xlfiles

# Return a pretty-printed XML string for an element.
def prettify_xml(element):
    """Return a pretty-printed XML string for an element.
    """
    rough_string = etree.tostring(element, 'utf-8')
    reparsed = minidom.parseString(rough_string)
    return reparsed.toprettyxml(encoding="utf-8", indent="\t")

def  create_iomodules_df(df, server):
	'''Return DataFrame of IO modules on AS IO bus.
	'''
	iomodules = df.copy()
	# drop unrelated cols and keep only unique io modules
	iomodules = iomodules[['as_id','module_id','module_name','module_type']].drop_duplicates()
	# keep only io modules belonging to this AS
	iomodules = iomodules.loc[iomodules['as_id'] == server['as_id']]

	return iomodules

def  create_points_df(df, server):
	'''Return DataFrame of points on AS IO bus.
	'''	
	points = df.copy()
	# filter for points only specified AS
	points = df.loc[df['as_id'] == server['as_id']]
	# Cleanup dataframe structure
	# drop duplicate IO module columns from points dataframe
	drop_cols = ['module_name', 'module_type', 'wire1', 'as_id', 'io1', 'system1']
	points = drop_col_if_exist(points, drop_cols)

	return points

def drop_col_if_exist(df, col_names):
	'''Drops dataframe columns if they exist.
		df (pandas.dataframe) : dataframe
		col_names (LIST) : list of df column nakmes to drop
	'''
	for col_name in col_names:
		if col_name in df:
			df = df.drop(col_name,1)
	return df

def prettyprint_data(server,points,iomodules):
	'''Print to console translated points data for reference. 
	'''
	print('\n\n*******************************')
	print('*** Automation Server '+server['as_name']+' ***')
	print('** IO modules **')
	print(iomodules)
	print('\n** points **')
	print(points[['module_id','point_name','point_type']])

#################
# Define script variables
xml_root='<?xml version="1.0" encoding="utf-8"?><ObjectSet ExportMode="Standard" Version="_VERSION_" Note="TypesFirst"><MetaInformation><ExportMode Value="Standard" /><RuntimeVersion Value="_VERSION_" /><SourceVersion Value="_VERSION_" /><ServerFullPath Value="/Clive-ES/Servers/Clive-AS" /></MetaInformation><ExportedObjects></ExportedObjects></ObjectSet>'

# SBO version
# THIS IS SHITTY ARGUMENT PARSING, FIX IT!
version = "1.6.1.5000"
if len(sys.argv) > 1:
	if len(sys.argv) == 3:
		if sys.argv[1] in ['-v','--version']:
			version = str(sys.argv[2])
		else:
			print("Invalid argument. Usage:\n\tpython parse_io_bus_xml.py [-v <SBO version>]\n\n")
	else:
		print("Invalid argument. Usage:\n\tpython parse_io_bus_xml.py [-v <SBO version>]\n\n")
print("SBO version "+version)
# get filenames of all matching Excel files in directory.
xlfiles = get_sbo_xl_files()

###############
# Execution
for xlfile in xlfiles:
	try:			
		# START

		# import Excel worksheet as pandas dataframe
		df = pd.read_excel(xlfile, sheetname='tblSXWRdata')
		
		# create AS dataframe
		df_as = df.copy()
		df_as = df_as[['as_id']].drop_duplicates()
		df_as['as_name'] = 'AS'+df_as['as_id'].astype(str).apply(lambda x: x.zfill(3))

		# Cleanup up dataframe data
		# drop empty IO
		df = df.dropna(subset=['point_name'])

		# loop through AS dataframe
		for n, server in df_as.iterrows():
			# create IO modules dataframe
			df_iomodules = create_iomodules_df(df, server)
			# create points dataframe
			df_points = create_points_df(df, server)
			# print to console translated points data
			prettyprint_data(server, df_points, df_iomodules)

			# generate and export xml
			root = etree.fromstring(xml_root)

			# define parent element of IO module and points subelements
			parent = root[1]

			# migrate points list to into correct SBO IO bus xml
			insert_elements(parent, df_points, df_iomodules)

			# insert correct version to xml
			root.attrib['Version'] = version
			root[0][1].attrib['Value'] = version
			root[0][2].attrib['Value'] = version

			# write to new xml file
			# define XML IO bus import filename
			xmlfile = server['as_name']+'_io_bus'+"_"+version+"_{:%Y%m%d_%H%M%S}".format(dt.datetime.now())+'.xml'
			with open(xmlfile, "w") as outfile:
				outfile.write(prettify_xml(root))

	except:
		print("Error: create XML IO bus file for "+xlfile+" failed, skipping!")
		print("Make sure file "+xlfile+" is still in this directory and is not open. Check data in "+xlfile+" is valid.")
		continue



