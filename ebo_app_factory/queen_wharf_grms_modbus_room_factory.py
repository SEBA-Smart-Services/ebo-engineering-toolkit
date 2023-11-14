########################
# Advanced example
########################
from xml_app_factory import ApplicationFactoryManager

# declare filenames/paths here
# xl_sorted_in_file = 'examples/sorted apps example.xlsx'
xl_sorted_in_file = 'examples/grms_apps_list.xlsx'
template_map = {    
    'IT2Device9': {'templateFilename': 'examples/Modbus GRMS Room 852 Export 2023-11-14.xml'},
    'IT3Device9': {'templateFilename': 'examples/Modbus GRMS Room 852 Export 2023-11-14.xml'},
}
# instantiate ApplicationFactoryManager object and make xml files
app_factory_manager = ApplicationFactoryManager(
	template_map=template_map,
	xlfile=xl_sorted_in_file,
	xml_out_file_prefix='examples/queens_wharf_grms_modbus_rooms',
)
app_factory_manager.make_documents()
