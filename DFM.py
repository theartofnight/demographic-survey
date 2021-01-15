import pandas as pd
import numpy as np
from tqdm import tqdm

class DemographicFileMaker:

    def __init__(self, **args):
        ## initialize the object by specifying input and output files.
        self.raw_data_file = args['raw_data']
        self.item_code_file = args['item_code']
        self.demographics_file = args['demographics']
        self.output_path = args['output']
        self._leader_id = args['leader_id']
        # self.heatmap_color_file = args['heatmap_color']
    
    def readAllFiles(self):
        ## read files and save it in object data.
        print("reading files...")
        self.raw_data_pd = pd.read_excel(self.raw_data_file, engine="openpyxl")
        self.item_code_pd = pd.read_excel(self.item_code_file, engine="openpyxl")
        self.demographics_pd = pd.read_excel(self.demographics_file, engine="openpyxl")
        # self.heatmap_color_pd = pd.read_excel(self.heatmap_color_file, engine="openpyxl")
        print("done!")

        self._preProcess()
       
    def writeOutput(self):
        pass

    def makeReport(self):

        print("wow")

    def calculateValues(self):
        
        self._prepareColumnsForID()

        item_list = []
        item_dict = {}
        for key in self._group_dict:
            item_list.append([0, key])
            temp_list = []
            for item in self._item_pd.iloc[self._group_dict[key], 0]:
                item_list.append([1, item])
                temp_list.append(item)
            item_dict.update({key:temp_list})

        for _ in tqdm(item_list):
            criteria, item = _
            if criteria == 0:
                sub_item_list = item_dict[item]

                self._filterResource(sub_item_list)

                self._calcualteEachRow(item)
        
    
    def _get_names_from_field(self, field_list):

        keys = [field[0] for field in field_list]
        return keys


    def _filterResource(self, filter_item):

        filter_item.insert(0, 'ExternalReference')
        self._filtered_raw_data = self.raw_data_pd[filter_item]

    ## do pre-process the data to be prepared in order to calculate.
    def _preProcess(self):

        print("data pre-processing...")
        self._item_pd = self.item_code_pd.iloc[:, :]
        self._item_pd = self._item_pd[~self._item_pd[self._item_pd.columns.values[0]].str.endswith("_c")].reset_index(drop=True)
        self._item_pd = self._item_pd[self._item_pd[self._item_pd.columns.values[0]].isin(self.raw_data_pd.columns.values)].reset_index(drop=True)

        self._group_dict = self._item_pd.groupby([self._item_pd.columns.values[1]]).groups

        self.raw_data_pd = self.raw_data_pd.iloc[2:]

        ## Convert numeric values into favorable or not.
        ## [1, 2, 3] -> 0, [4, 5] -> 1, [6, -99, ''] -> ''
        for field in self._item_pd.iloc[:, 0]:
            new_list = []
            for item in self.raw_data_pd[field].tolist():
                if item < 4 and item > 0:
                    new_list.append(0)
                elif item >= 4 and item <= 5:
                    new_list.append(1)
                else:
                    new_list.append('')
            self.raw_data_pd[field] = new_list
        
        ## convert the demographics data to include only answered entries.
        self._answered_demographics_data = self.demographics_pd[self.demographics_pd.iloc[:, 0].isin(self.raw_data_pd['ExternalReference'].tolist())].reset_index(drop=True)

        print("done!")
    
    def _init_dict(self, column_fields):

        _dict = {}
        keys = self._get_names_from_field(column_fields)

        for key in keys:
            _dict.update({key: {}})

        return _dict

    def _prepareColumnsForID(self):
        ## find the supervisor level of the given leader id.

        print("preparing data by leader ID...")

        leader_entry = self.demographics_pd[self.demographics_pd.iloc[:, 0] == self._leader_id]
        for index, level in enumerate(leader_entry.iloc[:, 2 : 11]):
            if (leader_entry[level] == self._leader_id).tolist()[0]:
                leader_level = index + 2
                break
        try:
            self._supervisor_id = leader_entry.iloc[0, leader_level - 1]
        except:
            print("the leader level is", leader_level)

        ## get Your Org data
        self._your_org = self._answered_demographics_data[self._answered_demographics_data.iloc[:, leader_level] == self._leader_id].reset_index(drop=True)
        # self._your_org = self.demographics_pd[self.demographics_pd.iloc[:, leader_level] == self._leader_id].reset_index(drop=True)
        
        ## make Parent group.
        self._parent_org = self._answered_demographics_data[self._answered_demographics_data.iloc[:, leader_level - 1] == self._supervisor_id].reset_index(drop=True)

        ## make direct report fields.
        self._direct_report_field = []
        temp_org = self._your_org[~(self._your_org.iloc[:, 0] == self._leader_id)]
        _dict = temp_org.groupby(temp_org.columns.values[leader_level + 1]).groups
        for key in _dict:
            self._direct_report_field.append([self._answered_demographics_data[self._answered_demographics_data.iloc[:, 0] == key].iloc[0, 1], _dict[key]])

        ## make grade group fields.
        self._grade_group_fields = []
        _dict = self._your_org.groupby("Pay Grade Group").groups
        for key in _dict:
            self._grade_group_fields.append([key, _dict[key]])

        ## make tenure group fields.
        self._tenure_group_fields = []
        _dict = self._your_org.groupby("Length of Service Group").groups
        for key in _dict:
            self._tenure_group_fields.append([key, _dict[key]])
        
        ## make performance rating fields.
        self._performance_rating_fields = []
        _dict = self._your_org.groupby("2019 Performance Rating").groups
        for key in _dict:
            self._performance_rating_fields.append([key, _dict[key]])

        ## make talent cordinate fields.
        self._talent_cordinate_fields = []
        _dict = self._your_org.groupby("2020 Talent Coordinate").groups
        for key in _dict:
            self._talent_cordinate_fields.append([key, _dict[key]])
        
        ## make gender fields.
        self._gender_fields = []
        _dict = self._your_org.groupby("Gender").groups
        for key in _dict:
            self._gender_fields.append([key, _dict[key]])

        ## make Ethnicity fields.
        self._ethnicity_fields = []
        _dict = self._your_org.groupby("Ethnicity (US)").groups
        for key in _dict:
            self._ethnicity_fields.append([key, _dict[key]])

        ## make age group fields.
        self._age_fields = []
        _dict = self._your_org.groupby("Age Group").groups
        for key in _dict:
            self._age_fields.append([key, _dict[key]])

        ## make country fields.
        self._country_fields = []
        _dict = self._your_org.groupby("Country").groups
        for key in _dict:
            self._country_fields.append([key, _dict[key]])

        ## make kite fields.
        self._kite_fields = []
        _dict = self._your_org.groupby("Kite Employee Flag").groups
        for key in _dict:
            self._kite_fields.append([key, _dict[key]])
        
        ## init dict to save all rows data
        self._gilead_overall = {}
        self._parent_group = {}
        self._your_org_2018 = {}

        ## below are deep dictionaries
        self._direct_reports = self._init_dict(self._direct_report_field)
        self._grade_group = self._init_dict(self._grade_group_fields)
        self._tenure_group = self._init_dict(self._tenure_group_fields)
        self._performance_rating = self._init_dict(self._performance_rating_fields)
        self._talent_cordinate = self._init_dict(self._talent_cordinate_fields)
        self._gender = self._init_dict(self._gender_fields)
        self._ethnicity = self._init_dict(self._ethnicity_fields)
        self._age_group = self._init_dict(self._age_fields)
        self._country = self._init_dict(self._country_fields)
        self._kite = self._init_dict(self._kite_fields)

        self.index_match = {
            "": ["Gilead Overall", "Parent Group", "Your Org (2018)"],
            "Direct Reports (as of April 24, 2018)": self._get_names_from_field(self._direct_report_field),
            "Grade Group": self._get_names_from_field(self._grade_group_fields),
            "Tenure Group": self._get_names_from_field(self._tenure_group_fields),
            "2019 Performance Rating": self._get_names_from_field(self._performance_rating_fields),
            "2020 Talent Coordinate": self._get_names_from_field(self._talent_cordinate_fields),
            "Gender": self._get_names_from_field(self._gender_fields),
            "Ethnicity (US)": self._get_names_from_field(self._ethnicity_fields),
            "Age Group": self._get_names_from_field(self._age_fields),
            "Country": self._get_names_from_field(self._country_fields),
            "Kite": self._get_names_from_field(self._kite_fields),
        }
        self.precious_dict = {}
        for first_index in self.index_match:
            _ = {}
            for item in self.index_match[first_index]:
                _.update({item: {}})
            self.precious_dict.update({first_index: _})

        print("done!")

    def _get_sum(self, data, nums, item):

        _dict = {}
        determine_parent_na = False
        total = 0
        len_total = len(data.columns) * nums

        for ind in range(len(data.columns)):
            _ = data.iloc[:, ind]
            is_empty_column = True
            lens = nums
            sub = 0
            for __ in _:
                if __ != '':
                    is_empty_column = False
                    sub += __
                    total += __
                else:
                    lens -= 1
                    len_total -= 1
            if is_empty_column:
                _dict.update({data.columns.values[ind]: "N/A"})
                determine_parent_na = True
            else:
                _dict.update({data.columns.values[ind]: str(round(sub / lens * 100)) + "%"})
        
        if determine_parent_na:
            _dict.update({item: "N/A"})
        else:
            _dict.update({item: str(round(total / len_total * 100)) + "%"})
        return _dict

    def _calcualteEachRow(self, item):
        
        ## calculate Gilead overall %s
        _dict = self._calculateOverall(self._answered_demographics_data, item)
        # self._gilead_overall.update(_dict)
        self.precious_dict[""]["Gilead Overall"].update(_dict)

        ## calculate Parent Group %s
        _dict = self._calculateOverall(self._parent_org, item)
        # self._parent_group.update(_dict)
        self.precious_dict[""]["Parent Group"].update(_dict)

        ## calculate Your Org(2018) %s
        _dict = self._calculateOverall(self._your_org, item)
        # self._your_org_2018.update(_dict)
        self.precious_dict[""]["Your Org (2018)"].update(_dict)

        ## calculate Direct reports %s
        self._calculateSubFields(self._direct_report_field, self.precious_dict["Direct Reports (as of April 24, 2018)"], item)
        # self._calculateSubFields(self._direct_report_field, self._direct_reports, item)
        
        ## calcualte Grade Group %s
        self._calculateSubFields(self._grade_group_fields, self.precious_dict["Grade Group"], item)
        # self._calculateSubFields(self._grade_group_fields, self._grade_group, item)

        ## calcualte Tenure Group %s
        self._calculateSubFields(self._tenure_group_fields, self.precious_dict["Tenure Group"], item)
        # self._calculateSubFields(self._tenure_group_fields, self._tenure_group, item)

        ## calculate Performance Rating %s
        self._calculateSubFields(self._performance_rating_fields, self.precious_dict["2019 Performance Rating"], item)
        # self._calculateSubFields(self._performance_rating_fields, self._performance_rating, item)

        ## calculate Talent Coordinate %s
        self._calculateSubFields(self._talent_cordinate_fields, self.precious_dict["2020 Talent Coordinate"], item)
        # self._calculateSubFields(self._talent_cordinate_fields, self._talent_cordinate, item)

        ## calculate Gender %s
        self._calculateSubFields(self._gender_fields, self.precious_dict["Gender"], item)
        # self._calculateSubFields(self._gender_fields, self._gender, item)

        ## calculate Ethnicity (US) %s
        self._calculateSubFields(self._ethnicity_fields, self.precious_dict["Ethnicity (US)"], item)
        # self._calculateSubFields(self._ethnicity_fields, self._ethnicity, item)

        ## calculate Age Group %s
        self._calculateSubFields(self._age_fields, self.precious_dict["Age Group"], item)
        # self._calculateSubFields(self._age_fields, self._age_group, item)

        ## calculate Country %s
        self._calculateSubFields(self._country_fields, self.precious_dict["Country"], item)
        # self._calculateSubFields(self._country_fields, self._country, item)

        ## calculate Kite %s
        self._calculateSubFields(self._kite_fields, self.precious_dict["Kite"], item)
        # self._calculateSubFields(self._kite_fields, self._kite, item)

    def _calculateOverall(self, dataframe, item):

        ids = dataframe.iloc[:, 0]
        nums = len(ids)
        working_pd = self._filtered_raw_data[self._filtered_raw_data["ExternalReference"].isin(ids)].reset_index(drop=True)
        return self._get_sum(working_pd.iloc[:, 1:], nums, item)
    
    def _calculateSubFields(self, dataframe, dictionary, item):

        for field in dataframe:
            _list = field[1]
            ids = self._your_org.iloc[_list, 0].tolist()
            nums = len(ids)
            working_pd = self._filtered_raw_data[self._filtered_raw_data["ExternalReference"].isin(ids)].reset_index(drop=True)
            _dict = self._get_sum(working_pd.iloc[:, 1:], nums, item)
            dictionary[field[0]].update(_dict)
        return dictionary

if __name__ == "__main__":

    ## Create a object.
    init_data = {
        'raw_data': "./Qualtrics Survey Export Sample 2021-01-13.xlsx",
        'item_code': "./Item Code 2021-01-10.xlsx",
        'demographics': "./Demographics File Sample 2021-01-13.xlsx",
        # 'heatmap_color': "Heatmap Colors.xlsx",
        'output': "./output for rest of leaders",
        'leader_id': 112372,
    }

    dfm = DemographicFileMaker(**init_data)

    ## read all needed files.
    dfm.readAllFiles()

    ## do main process to calculate report.
    dfm.calculateValues()

    ## make report dataframe to output.
    dfm.makeReport()

    ## write output file (Mockup STAR.xlsx)
    dfm.writeOutput()