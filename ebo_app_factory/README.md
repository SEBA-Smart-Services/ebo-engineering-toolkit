# EBO Application Factory

Mass produce EBO applications, objects or graphics based on a single or multiple template exported xml files and a list of copies.

## Overview

This tool allows an exported EBO application or objects to be used as a template for mass producing EBO applications or objects.

Each 'copy' and its bindings (if exported Special) are updated using a list of placeholders strings in the original template and a list of replacement strings for each copy. The list of placeholders and copies are stored in an Excel spreadsheet. These strings could be anything unique to each instance, such as equipment names, relative bind paths.

All 'copies' of the template are written into an xml document for importing into EBO. Each copy is presented in parallel in the root of the import file.

## Usage

The *xml_app_factory* module can either be imported into a script or for quick usage, the examples in the __main__ section of the module itself can be edited to suit the use case.

### Basic usage

The application requires two inputs in its most basic use case:

1. An EBO exported xml containing the objects to be used as a template making copies.
1. An Excel workbook with at least one sheet containing a list of replacement strings to replace placeholder strings within the template xml.

Instructions for basic use:

1. From EBO WorkStation, export the application, folder, or objects to be used as a template for copies to be made. The export can be either Standard or Special. Save the template xml file in the working directory of your EBO Application Factory program.

![EBO export step 1]('images/ebo export step 1.png')

![EBO export step 2]('./images/ebo export step 2.png')

![EBO export step 2]('./images/ebo export step 3.png')

1. Identify strings within the template application that should be used as placeholders for replacement with new strings for each copy. These strings may be equipment names, relative bind paths or any text that is unique to each copy of the template. Each substring should be placed in a cell in the first row of the Excel sheet. In this example:
 - "VAV-L21-INT4" is the name of the equipment and it should be replaced with the equipment name for each copy. This string has been placed in cell A1.
 - The substring "L21-INT4" is common in all bindings and should be replaced with the equivalent substring for each copy's bindings. This string has been placed in cell B1.
1. Update the spreadsheet with the equivalent replacement strings for each copy. Each row in the sheet from row 2 onward should represent a different copy of the template application. The columns should line up with the placeholder strings from the template application in the first row. For example:
 - "VAV-L04-INT09" is placed in column A to line up with template placeholder string "VAV-L21-INT4".
 - "L04_INT09" is placed in column B to line up with template placeholder string "L21-INT4".

![Excel basic workbook]('./images/excel_basic_example_markup.png')

1.  Save the Excel workbook in the working directory of your EBO Application Factory program.
1. In your EBO Application Factory program, import the EBO Application Factory module.
1. Set the filename and path of the input Excel spreadsheet and template xml file and the filename and path of the xml file to be created containing the application copies to be imported into EBO.
1. Instantitate the ApplicationTemplate, FactoryInputsFromSpreadsheet and ApplicationFactory objects, passing in the inputs as shown in the example below.

```python
########################
# Basic example
########################
from ebo_app_factory.xml_app_factory import ApplicationTemplate, FactoryInputsFromSpreadsheet, ApplicationFactory

# declare filenames/paths here
xl_in_file = 'examples/basic apps example.xlsx'
xml_in_file = 'examples/VAV-L21-INT4 application special.xml'
xml_out_file = 'examples/generated_ebo_apps_basic_example.xml'

# instantiate AppTemplate object object
app_template = ApplicationTemplate(xml_in_file, print_result=False)
# instantiate FactoryInputsFromSpreadsheet object
factory_inputs = FactoryInputsFromSpreadsheet(xl_in_file, print_result=False)
# instantiate ApplicationFactory object and make xml
app_factory = ApplicationFactory(
  template_child_elements_dict=app_template.template_child_elements_dict,
  factory_placeholders=factory_inputs.factory_placeholders,
  factory_copy_substrings=factory_inputs.factory_copy_substrings,
  xml_out_file=xml_out_file,
)
```

1. Call the make_document method of your ApplicationFactory object to generate the xml file containing copies of your application, ready for import into EBO.

```python
# app_factory.make_copies()
app_factory.make_document()

```

![basic usage terminal output]('./images/basic usage output.png')

1. Import the created xml file into EBO.

![EBO import step 1]('./images/ebo import step 1.png')
![EBO import step 2]('./images/ebo import step 2.png')

1. Inspect the imported copies for correct naming and bindings.

### Advanced usage

In situtaions where there are mulitple groups of EBO applications that will either be imported into difference locations in the System Tree or use different template xml files, or both, a helper class called `ApplicationFactoryManager` has been created. A unique xml file is created for each sheet in the Excel workbook. In this situation, in addition to the two Basic Usage requirements, as additional input is required:

- A dictionary of Excel workbook sheet names and corresponding  EBO exported xml template file to use. I'm sure there is a more elegant way of mapping sheets to template xml files but that is a future problem. Example:

```python
template_map = {
  'L2-3-All3StgHtg': {'templateFilename': 'examples/VAV-L21-NW2 application special.xml'},
  'L4-12-3StgHtg': {'templateFilename': 'examples/VAV-L21-NW2 application special.xml'},
  'L13-15-3StgHtg': {'templateFilename': 'examples/VAV-L21-NW2 application special.xml'},
  'L16-27-3StgHtg': {'templateFilename': 'examples/VAV-L21-NW2 application special.xml'},
  '1StgHtg': {'templateFilename': 'examples/VAV-L04-INT09 application special.xml'},
  'L2-3NoHtg': {'templateFilename': 'examples/VAV-L21-INT4 application special.xml'},
  'L4-12NoHtg': {'templateFilename': 'examples/VAV-L21-INT4 application special.xml'},
  'L13-15NoHtg': {'templateFilename': 'examples/VAV-L21-INT4 application special.xml'},
  'L16-27NoHtg': {'templateFilename': 'examples/VAV-L21-INT4 application special.xml'},
}

```

Instructions for advanced use:

1. From EBO WorkStation, export the application, folder, or objects to be used as a template for each 'application group' of copies to be made. Unlike the basic usage, an unlimited number of template applications can be exported. Save the template xml files in the working directory of your EBO Application Factory program.
1. Create the Excel workbook as per the basic usage but split each 'application group' into a different sheet. For each sheet the same rule applies, first row contains template placeholder strings, subsequent rows represent replacement strings for each copy.
1.  Save the Excel workbook in the working directory of your EBO Application Factory program.
1. In your EBO Application Factory program, import the EBO Application Factory Manager module.
1. Set the filename and path of the input Excel spreadsheet.
1. Create or import the dictionary mapping Excel workbook sheet names to template xml file paths containing the application copies to be imported into EBO.
1. Instantitate the ApplicationFactoryManager objects, passing in the inputs as shown in the example below. Include the file path/prefix for the application group xml files to be written to. The suffix of each file will be the Excel workbook sheet name.

```python
########################
# Advanced example
########################
from ebo_app_factory.xml_app_factory import ApplicationFactoryManager

# declare filenames/paths here
xl_sorted_in_file = 'examples/sorted apps example.xlsx'
# create dictionary mapping Excel sheet names to template xml files
template_map = {
  'L2-3-All3StgHtg': {'templateFilename': 'examples/VAV-L21-NW2 application special.xml'},
  'L4-12-3StgHtg': {'templateFilename': 'examples/VAV-L21-NW2 application special.xml'},
  'L13-15-3StgHtg': {'templateFilename': 'examples/VAV-L21-NW2 application special.xml'},
  'L16-32-3StgHtg': {'templateFilename': 'examples/VAV-L21-NW2 application special.xml'},
  '1StgHtg': {'templateFilename': 'examples/VAV-L04-INT09 application special.xml'},
  'L2-3NoHtg': {'templateFilename': 'examples/VAV-L21-INT4 application special.xml'},
  'L4-12NoHtg': {'templateFilename': 'examples/VAV-L21-INT4 application special.xml'},
  'L13-15NoHtg': {'templateFilename': 'examples/VAV-L21-INT4 application special.xml'},
  'L16-32NoHtg': {'templateFilename': 'examples/VAV-L21-INT4 application special.xml'},
}
# instantiate ApplicationFactoryManager object
app_factory_manager = ApplicationFactoryManager(
  template_map=template_map,
  xlfile=xl_sorted_in_file,
  xml_out_file_prefix='examples/example_ebo_apps'
)

```

1. Call the make_documents method of your ApplicationFactoryManager object to generate the xml file containing copies of your application, ready for import into EBO.

```python
# create xml files
app_factory_manager.make_documents()

```

![advanced usage terminal output]('./images/advanced usage output.png')
![advanced usage output files]('advanced usage output files.png')

1. As per Basic Usage, import the created xml files into EBO and inspect the imported copies for correct naming and bindings.
