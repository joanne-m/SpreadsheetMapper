# Spreadsheet Mapper

Simple class for converting a spreadsheet from one format to another.

[v1.0](https://github.com/joanne-m/SpreadsheetMapper/releases/tag/v1.0) - Modify column names.

## Prerequisites
``` 
Python 3.x
pip install pandas pyyaml xlrd XlsxWriter
```

## Installation
1. Clone this repository.
2. Update the config.yml file to point to the input, output, and map files.

```
mapfile : mapping.xlsx  # Filename of mapping spreadsheet. Should contain at least two headers: "from", which contains headers in input file, 
                        # and "to", which is the corresponding header in the output file. Unmapped columns in "from" are disregarded.
mapsheet : ~            # (optional) Specific sheet within the mapfile containing the mapping. If set to null, will default to first sheet.

inputfile: input.csv    # Filename of input spreadsheet. Headers should correspond to headers in mapfile's "from" column.
inputsheet: ~           # (optional) Specific sheet within the inputfile containing the mapping. If set to null, will default to first sheet.

outputfile: output.xlsx # Filename of input spreadsheet. Resulting headers would correspond to headers in mapfile's "to" column.
outputsheet: ~          # (optional) Specific sheet within the inputfile containing the mapping. If set to null, will default to first sheet.
```
3. Sample usage:

```
import mapper

# Creating an instance of mapper would automatically load the config.yml file and load the mapping from the set mapfile.
m = mapper.Mapper()

# Simplest usage is to call the convert() function. This would just use the file and sheetnames from the configuration file.
m.convert()



# A different configuration file could also be used.
m = mapper.Mapper('otherconfig.yml')

# Or instead of updating the configuration file, a different mapfile (and sheet) could be used.
# m.setmapping(mapfile = "mapping.csv")

# Can also convert additional files using the previously set mapping.
# m.convert(inputfile = "input.csv", outputfile = "output.xlsx")

```

## Authors

* [**Joanne Mendoza**](https://github.com/joanne-m)
