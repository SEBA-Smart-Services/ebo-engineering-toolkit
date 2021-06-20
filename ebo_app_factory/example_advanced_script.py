########################
# Advanced example
########################
from xml_app_factory import ApplicationFactoryManager

# declare filenames/paths here
# xl_sorted_in_file = 'examples/sorted apps example.xlsx'
xl_sorted_in_file = 'examples/apps sorted.xlsx'
template_map = {
	'L2-3-All3StgHtg': {'templateFilename': 'examples/VAV-L21-NW2 application special.xml'},
	'L4-12-3StgHtg': {'templateFilename': 'examples/VAV-L21-NW2 application special.xml'},
	'L13-15-3StgHtg': {'templateFilename': 'examples/VAV-L21-NW2 application special.xml'},
	'L10-3StgHtg': {'templateFilename': 'examples/VAV-L21-NW2 application special.xml'},
	'L16-27-3StgHtg': {'templateFilename': 'examples/VAV-L21-NW2 application special.xml'},
	'L28-30-3StgHtg': {'templateFilename': 'examples/VAV-L21-NW2 application special.xml'},
	'L20-3StgHtg': {'templateFilename': 'examples/VAV-L21-NW2 application special.xml'},
	'L31-32-3StgHtg': {'templateFilename': 'examples/VAV-L21-NW2 application special.xml'},
	'L33-33A-All3StgHtg': {'templateFilename': 'examples/VAV-L21-NW2 application special.xml'},
	'1StgHtg': {'templateFilename': 'examples/VAV-L04-INT09 application special.xml'},
	'L2-3NoHtg': {'templateFilename': 'examples/VAV-L21-INT4 application special.xml'},
	'L4-12NoHtg': {'templateFilename': 'examples/VAV-L21-INT4 application special.xml'},
	'L10-NoHtg': {'templateFilename': 'examples/VAV-L21-INT4 application special.xml'},
	'L13-15NoHtg': {'templateFilename': 'examples/VAV-L21-INT4 application special.xml'},
	'L16-27NoHtg': {'templateFilename': 'examples/VAV-L21-INT4 application special.xml'},
	'L28-30NoHtg': {'templateFilename': 'examples/VAV-L21-INT4 application special.xml'},
	'L20NoHtg':{'templateFilename': 'examples/VAV-L21-INT4 application special.xml'},
	'L31-32NoHtg': {'templateFilename': 'examples/VAV-L21-INT4 application special.xml'}
}
# instantiate ApplicationFactoryManager object and make xml files
app_factory_manager = ApplicationFactoryManager(
	template_map=template_map,
	xlfile=xl_sorted_in_file,
	# sheetname='L33-33A-All3StgHtg',
	# xml_out_file_prefix='examples/example_ebo_apps',
	xml_out_file_prefix='examples/ebo_apps',
)
app_factory_manager.make_documents()
