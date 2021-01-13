import pandas as pd
import numpy as np

class DemographicFileMaker:

    def __init__(self, **args):
        # initialize the object by specifying input and output files.
        self.raw_data_file = args['raw_data']
        self.item_code_file = args['item_code']
        self.survey_expert_file = args['survey_expert']
        self.output_file = args['output']
    
    def readAllFiles(self):
        # read files and save it in object data.
        self.raw_data_pd = pd.read_excel(self.raw_data_file, engine="openpyxl")
        self.item_code_pd = pd.read_excel(self.item_code_file, engine="openpyxl")
        self.survey_expert_pd = pd.read_excel(self.survey_expert_file, engine="openpyxl")

        print("dataframe", self.raw_data_pd)
        print("item_code", self.item_code_pd)
        print("survey_expert", self.survey_expert_pd)

    def writeOutput(self):
        pass

    def makeReport(self):
        pass

if __name__ == "__main__":

    # Create a object.
    init_data = {
        'raw_data': "./Demographics File Sample 2021-01-10.xlsx",
        'item_code': "./Item Code 2021-01-10.xlsx",
        'survey_expert': "./Qualtrics Survey Export Sample 2021-01-10.xlsx",
        'output': "./Mockup STAR.xlsx"
    }

    dfm = DemographicFileMaker(**init_data)

    # read all needed files.
    dfm.readAllFiles()

    # do main processes
    dfm.makeReport()

    # write output file (Mockup STAR.xlsx)
    dfm.writeOutput()