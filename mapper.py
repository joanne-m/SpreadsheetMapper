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

    def setmapping(self, mapfile=None, mapsheet=None):
        ''' mapfile (and mapsheet) is a csv / excel file required to have two headers: from, which contains headers in input file, and to, the corresponding header in the outputfile.
            Entries in the from field could be empty (no corresponding mapping), which would mean that those columns would be omitted in the output file.
        '''
        mapfile = self.config['mapfile'] if mapfile is None else mapfile
        mapsheet = self.config['mapsheet'] if mapsheet is None else mapsheet

        mappingdf = self.__opendataframe(mapfile,mapsheet )

        # Disregard fields with no mapping
        self.mappingdf = mappingdf[mappingdf['to'].notna()]
        print("\nMapping the following fields (based on "+mapfile+"):")
        print(self.mappingdf[['from', 'to']])

    def convert(self, inputfile = None, inputsheet = None, outputsheet = None, outputfile = None):
        ''' Main function. Maps input file's columns to output file's using previously set mappings.
        '''
        inputfile = self.config['inputfile'] if inputfile is None else inputfile
        inputsheet = self.config['inputsheet'] if inputsheet is None else inputsheet
        outputfile = self.config['outputfile'] if outputfile is None else outputfile
        if outputsheet is None:
            outputsheet = self.config['outputsheet'] if ('outputsheet' in self.config and self.config['outputsheet']) else 'Sheet1'

        # Open inputfile
        indf = self.__opendataframe(inputfile, inputsheet )

        # Create output dataframe containing only fields with existing mapping.
        outdf = indf[self.mappingdf['from']]

        # Replace column headers according to mapping.
        outdf.columns = self.mappingdf['to']

        # Write output data to an excel file
        writer = pd.ExcelWriter(outputfile, engine='xlsxwriter')
        outdf.to_excel(writer, outputsheet, index=False)
        writer.save()

        print('\nSuccessfully mapped', inputfile, 'to', outputfile+'.')

