########################
# Advanced example
########################
from xml_app_factory import ApplicationFactoryManager

# declare filenames/paths here
# xl_sorted_in_file = 'examples/sorted apps example.xlsx'
xl_sorted_in_file = 'examples/escalators_apps_list.xlsx'
template_map = {    
    'crisscross': {'templateFilename': 'examples/E24 E25 crisscross UPS App Export 2023-10-19 201917.xml'},
    'crisscross-2': {'templateFilename': 'examples/E24 E25 crisscross UPS App Export 2023-10-19 201917.xml'},
    'crisscross-3': {'templateFilename': 'examples/E24 E25 crisscross UPS App Export 2023-10-19 201917.xml'},
    'parallel': {'templateFilename': 'examples/E24 E25 crisscross UPS App Export 2023-10-19 201917.xml'},
    'parallel-2': {'templateFilename': 'examples/E24 E25 crisscross UPS App Export 2023-10-19 201917.xml'},
	'parallel-no-ups': {'templateFilename': 'examples/E24 E25 crisscross UPS App Export 2023-10-19 201917.xml'},
    'parallel-no-ups-2': {'templateFilename': 'examples/E24 E25 crisscross UPS App Export 2023-10-19 201917.xml'},
}
# instantiate ApplicationFactoryManager object and make xml files
app_factory_manager = ApplicationFactoryManager(
	template_map=template_map,
	xlfile=xl_sorted_in_file,
	# sheetname='L33-33A-All3StgHtg',
	# xml_out_file_prefix='examples/example_ebo_apps',
	xml_out_file_prefix='examples/queens_wharf_escalators_ebo_apps',
)
app_factory_manager.make_documents()
