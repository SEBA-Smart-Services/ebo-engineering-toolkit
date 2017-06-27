# Make StruxureWare IO bus from points list

# Usage
1. Create points list, either in Excel or elsewhere.
2. Convert points list to Excel workbook or csv if not already.
3. Close workbook/csv file and run the script from within the same directory.
4. An IO bus xml is created in the same directory, import xml directly onto the IO bus in StruxureWare.
5. Configure points in StruxureWare, eg engineering units, electrical units.

TODO:
- provide excel points list samples
- provide excel points list schema

# Dependencies
- python 3
- pandas
- numpy
- lxml

This could probably be rewritten to work without any extra python packages, but... I'm busy.
