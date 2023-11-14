########################
# Advanced example
########################
from xml_app_factory import ApplicationFactoryManager

# declare filenames/paths here
# xl_sorted_in_file = 'examples/sorted apps example.xlsx'
xl_sorted_in_file = 'examples/irc_apps_list.xlsx'
template_map = {    
    'L00-DC1p1': {'templateFilename': 'examples/T3-L05M-C2-IRC-B1 app Export 2023-10-03.xml'},
    'L00-DC1p2': {'templateFilename': 'examples/T3-L05M-C2-IRC-B1 app Export 2023-10-03.xml'},
	'IT2-L05M': {'templateFilename': 'examples/T3-L05M-C2-IRC-B1 app Export 2023-10-03.xml'},
}
# instantiate ApplicationFactoryManager object and make xml files
app_factory_manager = ApplicationFactoryManager(
	template_map=template_map,
	xlfile=xl_sorted_in_file,
	# sheetname='L33-33A-All3StgHtg',
	# xml_out_file_prefix='examples/example_ebo_apps',
	xml_out_file_prefix='examples/queens_wharf_irc_ebo_apps',
)
app_factory_manager.make_documents()
