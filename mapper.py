import pandas as pd
import yaml
import pathlib

class Mapper():
    ''' Maps a CSV/Excel input file's columns to those specified on a mapping file, and outputs an excel file.
        Usage:
            m = Mapper()
            m.convert()
    '''
    def __init__(self, configfile = "config.yml"):
        ''' Loads config file.
            Preferred usage: supply mapfile, inputfile, and outputfile.
        '''
        self.config = yaml.safe_load(open(configfile))
        print("Successfully loaded config.yml.")
        self.setmapping()
        self.valueset = {}

    def __opendataframe(self, filename, sheetname=None):
        ''' Helper function to load a csv / excel file into a pandas DataFrame.
        '''
        if filename:
            extension = pathlib.Path(filename).suffix
            if extension in ['.xlsx','.xls','.xlsm', '.xlsb', '.odf', '.ods', '.odt']:
                if sheetname:
                    df = pd.read_excel(filename, sheet_name=sheetname)
                else:
                    df = pd.read_excel(filename)
            elif extension == '.csv':
                df = pd.read_csv(filename)
            else:
                raise TypeError("Unsupported file type. Please use xlsx or csv.")
        else:
            raise Exception("Filename is required.")
        return df

    def setmapping(self, mapfile=None, mapsheet=None, include_unmapped=None):
        ''' mapfile (and mapsheet) is a csv / excel file required to have two headers: "from", which contains headers in input file, and "to", the corresponding header in the outputfile.
            Entries in either the "from" or the "to" field could be empty (no corresponding mapping), which would mean that those unmapped columns would be omitted in the output file (default behavior).
            But if include_unmapped is set to True, all fields in "to" are included in the output file (with unmapped fields set to empty).
        '''
        mapfile = self.config['mapfile'] if mapfile is None else mapfile
        mapsheet = self.config['mapsheet'] if mapsheet is None else mapsheet
        self.include_unmapped = self.config['include_unmapped'] if include_unmapped is None and 'include_unmapped' in self.config else include_unmapped

        mappingdf = self.__opendataframe(mapfile,mapsheet )

        # Disregard fields with no mapping
        self.mappingdf = mappingdf[mappingdf['to'].notna() & mappingdf['from'].notna()]

        # Keep original mapping (for include_unmapped = True)
        self.rawmappingdf = mappingdf[mappingdf['to'].notna()]

        print("\nMapping the following fields (based on "+mapfile+"):")
        if self.include_unmapped: print(self.rawmappingdf[['from', 'to']])
        else: print(self.mappingdf[['from', 'to']])

    def convert(self, inputfile = None, inputsheet = None, outputfile = None, outputsheet = None):
        ''' Main function. Maps input file's columns to output file's using previously set mappings.
        '''
        inputfile = self.config['inputfile'] if inputfile is None else inputfile
        inputsheet = self.config['inputsheet'] if inputsheet is None else inputsheet
        outputfile = self.config['outputfile'] if outputfile is None else outputfile
        if outputsheet is None:
            outputsheet = self.config['outputsheet'] if ('outputsheet' in self.config and self.config['outputsheet']) else 'Sheet1'

        # Open input file
        indf = self.__opendataframe(inputfile, inputsheet )

        # Create output dataframe containing only fields with existing mapping.
        outdf = indf[self.mappingdf['from']]

        # Replace column headers according to mapping.
        outdf.columns = self.mappingdf['to']

        # Include unmapped fields if required
        if self.include_unmapped: outdf = outdf.reindex(columns = self.rawmappingdf[self.rawmappingdf['to'].notna()]['to'])

        date_format = self.config['date_format'] if ('date_format' in self.config and self.config['date_format']) else 'mm/dd/yyyy'

        # Write output data to an excel file
        writer = pd.ExcelWriter(outputfile, engine='xlsxwriter', date_format = date_format, datetime_format=date_format)
        outdf.to_excel(writer, outputsheet, index=False)
        writer.save()

        print('\nSuccessfully mapped', inputfile, 'to', outputfile+'.')