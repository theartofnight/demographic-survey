import pandas as pd
import os
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

        print("writing...")
        path = self.output_path + "/" + str(self._leader_id) + " STAR.xlsx"
        if os.path.exists(path):
            os.remove(self.output_path + "/" + str(self._leader_id) + " STAR.xlsx")
        os.makedirs(self.output_path, exist_ok=True)
        self.whole_frame.to_excel(path, engine="openpyxl")
        print("done!")

    def makeReport(self):

        field_picture_postion = str(round(self._participated / self._invited * 100)) + "% Participation Rate" + \
            "\n" + str(self._participated) + '/' + str(self._invited) + "\n" + "(Participated / Invited)"
        whole_frame_list = []
        for key in self.precious_dict:
            sub_dict = self.precious_dict[key]
            frames = []

            if key == '':
                _list = []
                _list.append("Number of Respondents (incl. N/A)")
                for criteria, item in self._item_list:
                    if criteria == 0:
                        _list.append(item)
                    elif criteria == 1:
                        _list.append(self._item_pd[self._item_pd.iloc[:, 0] == item].iloc[0, 3])
                frames.append(pd.DataFrame({field_picture_postion: _list}))

            for sub_key in sub_dict:
                sub_item = sub_dict[sub_key]
                _list = []

                _list.append(self.first_row[sub_key])
                for criteria, item in self._item_list:
                    _list.append(sub_item[item])

                frames.append(pd.DataFrame({sub_key: _list}))
            
            whole_frame_list.append(pd.concat({key: pd.concat(frames, axis=1)}, axis=1))

        self.whole_frame = pd.concat(whole_frame_list, axis=1)

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

        self._item_list = item_list

        for _ in tqdm(item_list):
            criteria, item = _
            if criteria == 0:
                sub_item_list = item_dict[item]

                self._filterResource(sub_item_list)

                self._calcualteEachRow(item)
        
        print("complete!")
    
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

        self.first_row = {}

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
        
        ## calculate nums of participated and invitied
        self._participated = len(self._your_org.iloc[:, 0])
        self._invited = len(self.demographics_pd[self.demographics_pd.iloc[:, leader_level] == self._leader_id].reset_index(drop=True).iloc[:, 0])
        
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
        _dict, lens = self._calculateOverall(self._answered_demographics_data, item)
        self.precious_dict[""]["Gilead Overall"].update(_dict)
        self.first_row.update({"Gilead Overall": lens})

        ## calculate Parent Group %s
        _dict, lens = self._calculateOverall(self._parent_org, item)
        self.precious_dict[""]["Parent Group"].update(_dict)
        self.first_row.update({"Parent Group": lens})

        ## calculate Your Org(2018) %s
        _dict, lens = self._calculateOverall(self._your_org, item)
        self.precious_dict[""]["Your Org (2018)"].update(_dict)
        self.first_row.update({"Your Org (2018)": lens})

        ## calculate Direct reports %s
        self._calculateSubFields(self._direct_report_field, self.precious_dict["Direct Reports (as of April 24, 2018)"], item)
        
        ## calcualte Grade Group %s
        self._calculateSubFields(self._grade_group_fields, self.precious_dict["Grade Group"], item)

        ## calcualte Tenure Group %s
        self._calculateSubFields(self._tenure_group_fields, self.precious_dict["Tenure Group"], item)

        ## calculate Performance Rating %s
        self._calculateSubFields(self._performance_rating_fields, self.precious_dict["2019 Performance Rating"], item)

        ## calculate Talent Coordinate %s
        self._calculateSubFields(self._talent_cordinate_fields, self.precious_dict["2020 Talent Coordinate"], item)

        ## calculate Gender %s
        self._calculateSubFields(self._gender_fields, self.precious_dict["Gender"], item)

        ## calculate Ethnicity (US) %s
        self._calculateSubFields(self._ethnicity_fields, self.precious_dict["Ethnicity (US)"], item)

        ## calculate Age Group %s
        self._calculateSubFields(self._age_fields, self.precious_dict["Age Group"], item)

        ## calculate Country %s
        self._calculateSubFields(self._country_fields, self.precious_dict["Country"], item)

        ## calculate Kite %s
        self._calculateSubFields(self._kite_fields, self.precious_dict["Kite"], item)

    def _calculateOverall(self, dataframe, item):

        ids = dataframe.iloc[:, 0]
        nums = len(ids)
        working_pd = self._filtered_raw_data[self._filtered_raw_data["ExternalReference"].isin(ids)].reset_index(drop=True)
        return self._get_sum(working_pd.iloc[:, 1:], nums, item), len(ids)
    
    def _calculateSubFields(self, dataframe, dictionary, item):

        for field in dataframe:
            _list = field[1]
            ids = self._your_org.iloc[_list, 0].tolist()
            nums = len(ids)
            working_pd = self._filtered_raw_data[self._filtered_raw_data["ExternalReference"].isin(ids)].reset_index(drop=True)
            _dict = self._get_sum(working_pd.iloc[:, 1:], nums, item)
            dictionary[field[0]].update(_dict)
            self.first_row.update({field[0]: len(ids)})

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