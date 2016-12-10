# Continuum to SBO Conversion Tools

The following tools work in conjunction to convert Continuuum dumpfiles (eg VAV-1234.dmp) to either SBO b3 applications or SBO 3rd party BACnet device applications.

## Continuum dumpfile extractor
Reads a dumpfile and converts it to an Excel workbook, where each sheet contains a list of Continuum objects:
- Report
- EventEnrollment
- EventNotification
- Schedule
- InfinityNumeric
- InfinityInput
- InfinityOutput
- InfinityString
- Group
- Program
- InfinitySystemVariable

As well as relevant attributes of the object (incomplete).

Example usage

```python
# This is a very long and ugly import call. I should probably clean it up one day
from SmartStruxure-engineering-toolkit.legacy_system.continuum.converter import DmpfileExtractor

# define dump file locations
# Assume you have a bunch of .dmp files exported from Continuum sitting in the same directory
def get_dmpfiles(path="."):
	files = listdir(path)
	dmpfiles = [file for file in files if ".dmp" in file.lower()]
	return dmpfiles

dumpfiles = get_dmpfiles()

# loop through .dmp files and spit out Excel workbook with the same name
for dumpfile in dumpfiles:
	my_xlfile = 'ddc_objects_' + dumpfile.split(".")[0] + '_.xlsx'
	extractor = DmpfileExtractor(dmpfile=dumpfile)
	extractor.get_b3_objects(verbose=True)
	extractor.to_excel(workbook=my_xlfile)

```

## Continuum b3 application to SBO converter
Takes an Excel workbook contining Continuum objects (ie the output of DmpfileExtractor) and converts it into a b3 or BACnet application.

Currently only supports the following object types:
- Analog Input
- Digital Input
- Multistate Input
- Analog Output
- Digital Output
- Multistate Output
- Analog Value
- Digital Value
- Multistate Value
- Datetime Value
- String Value
- Alarms (creates object only, does not preserve bind)
- Schedules (creates object only, does not preserve event objects)

Example usage:

```python
# This is a very long and ugly import call. I should probably clean it up one day
from SmartStruxure-engineering-toolkit.legacy_system.continuum.converter import b3ApplicationBuilder

# The Excel workbook you just created
my_xlfile = 'Example Continuum DDC objects.xlsx'

# The XML filename you want to output to for importing into SBO BACnet device Application
my_xmlfile = 'Example Continuum DDC objects output.xml'

# instantiate ModbusSlaveTransition object
sbo_modbus_converter = ModbusSlaveTransition(xlfile=my_xlfile, xmlfile=my_xmlfile)

# create SBO xml for importing
b3xmlbuilder = b3ApplicationBuilder()
b3xmlbuilder.make_xml()

```

NOTE: Does not support Continuum programs yet.