import pandas as pd
import os
import openpyxl
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

        self.performance_order = [
            "Exceptional",
            "Exceeded",
            "Achieved",
            "Improvement Needed",
            "On Leave",
            "No Rating",
        ]
    
    def readAllFiles(self):
        ## read files and save it in object data.
        print("reading files...")
        self.raw_data_pd = pd.read_excel(self.raw_data_file, engine="openpyxl")
        self.item_code_pd = pd.read_excel(self.item_code_file, engine="openpyxl", sheet_name="ItemCodeSTAR")
        self.category_pd =  pd.read_excel(self.item_code_file, engine="openpyxl", sheet_name="CurrentCategorySTAR")
        self.demographics_pd = pd.read_excel(self.demographics_file, engine="openpyxl")
        # self.heatmap_color_pd = pd.read_excel(self.heatmap_color_file, engine="openpyxl")
        print("done!")

        self._preProcess()
       
    def writeOutput(self):

        print("writing...")
        ## specify the path of output file
        path = self.output_path + "/" + str(self._leader_id) + " STAR.xlsx"

        ## if the output file already exists, remove it.
        if os.path.exists(path):
            os.remove(self.output_path + "/" + str(self._leader_id) + " STAR.xlsx")

        ## make a folder to involve the output file.
        os.makedirs(self.output_path, exist_ok=True)

        ## and write the output file.
        self.whole_frame.to_excel(path, engine="openpyxl")
        print("done!")

    def makeReport(self):

        field_picture_postion = str(round(self._participated / self._invited * 100)) + "% Participation Rate" + \
            "\n" + str(self._participated) + '/' + str(self._invited) + "\n" + "(Participated / Invited)"
        
        whole_frame_list = []
        for key in self.precious_dict:
            sub_dict = self.precious_dict[key]
            frames = []
            
            ## fill the first column.
            if key == '':
                _list = []
                _list.append("Number of Respondents (incl. N/A)")
                for criteria, item in self._item_list:
                    if criteria == 0:
                        _list.append(item)
                    elif criteria == 1:
                        _list.append(self._item_pd[self._item_pd["Item ID"] == item]["Short Text [2020 onward]"].values[0])
                frames.append(pd.DataFrame({field_picture_postion: _list}))

            ## fill rows.
            for sub_key in sub_dict:
                sub_item = sub_dict[sub_key]
                _list = []

                _list.append(self.first_row[key + sub_key])
                for criteria, item in self._item_list:
                    if criteria == 1:
                        item = self._item_pd[self._item_pd["Item ID"] == item]["Unique Item Code"].values[0]

                    _list.append(sub_item[item])

                frames.append(pd.DataFrame({sub_key: _list}))
            
            whole_frame_list.append(pd.concat({key: pd.concat(frames, axis=1)}, axis=1))

        self.whole_frame = pd.concat(whole_frame_list, axis=1)

    def calculateValues(self):
        
        self._prepareColumnsForID()

        item_list = []
        item_dict = {}

        for key in self.order_category:
            item_list.append([0, key])
            temp_list = []
            for item in self.category_pd.iloc[self._group_dict[key], 0]:
                item_list.append([1, item])
                temp_list.append(item)
            item_dict.update({key:temp_list})

        self._item_list = item_list

        for _ in tqdm(item_list):
            criteria, item = _
            if criteria == 0:
                sub_id_list = item_dict[item]

                self._filterResource(sub_id_list)

                self._calcualteEachRow(item)
        
        print("complete!")
    
    def _get_names_from_field(self, field_list):

        keys = [field[0] for field in field_list]
        return keys


    def _filterResource(self, id_list):

        filter_item = self._item_pd[self._item_pd["Item ID"].isin(id_list)]["Unique Item Code"].tolist()
        filter_item.insert(0, 'ExternalReference')
        self._filtered_raw_data = self.raw_data_pd[filter_item]

    ## do pre-process the data to be prepared in order to calculate.
    def _preProcess(self):

        print("data pre-processing...")

        ## make a item group and filter the able source.
        self._item_pd = self.item_code_pd[self.item_code_pd["Type ID"] == "T01"]
        self._item_pd = self._item_pd[self._item_pd["Unique Item Code"].isin(self.raw_data_pd.columns.values)].reset_index(drop=True)
        
        self.category_pd = self.category_pd[self.category_pd["Item ID in 2020 Survey"].isin(self._item_pd["Item ID"].tolist())]
        self.category_pd = self.category_pd.drop_duplicates(subset=["Item ID in 2020 Survey"]).reset_index(drop=True)
        self.order_category = self.category_pd.drop_duplicates(subset=["2020 Category"]).iloc[:, 1].tolist()

        self._group_dict = self.category_pd.groupby(["2020 Category"]).groups

        self.raw_data_pd = self.raw_data_pd.iloc[2:].reset_index(drop=True)

        ## Convert numeric values into favorable or not.
        ## [1, 2, 3] -> 0, [4, 5] -> 1, [6, -99, ''] -> ''
        for field in self._item_pd["Unique Item Code"]:
            new_list = []
            for item in self.raw_data_pd[field].tolist():
                if item < 4 and item > 0:
                    new_list.append(0)
                elif item >= 4 and item <= 5:
                    new_list.append(1)
                else:
                    new_list.append('')
            self.raw_data_pd[field] = new_list

        ## new feature -> process A/B pair to calculate easily.
        pairs = self._item_pd[self._item_pd["AB Code"].isin(["A", "B"])].reset_index(drop=True)
        pairs_dict = pairs.groupby(["Item ID"]).groups

        for key in pairs_dict:

            _list = pairs_dict[key]
            pairs_row = pairs.iloc[_list, :]
            item_list = pairs_row["Unique Item Code"].tolist()
            text_list = pairs_row["Short Text [2020 onward]"].tolist()
            new_text = "/".join(text_list)
            self._item_pd = self._item_pd[~(self._item_pd["Unique Item Code"] == item_list[1])].reset_index(drop=True)
            _index = self._item_pd[self._item_pd["Unique Item Code"] == item_list[0]].index
            self._item_pd.loc[_index, "Short Text [2020 onward]"] = new_text

            column_pair = self.raw_data_pd[item_list]
            _list = []
            for index in range(len(column_pair.index)):
                _row = column_pair.iloc[index, :].tolist()
                for val in _row:
                    if type(val) == type(0):
                        _list.append(val)
                        break
                else:
                    _list.append('')
            _pd = pd.DataFrame(_list, columns=[item_list[0]])
            self.raw_data_pd = self.raw_data_pd.drop(columns=[item_list[1]])
            self.raw_data_pd[item_list[0]] = _pd

        ## convert the demographics data to include only answered entries.
        self._answered_demographics_data = self.demographics_pd[self.demographics_pd.iloc[:, 0].isin(self.raw_data_pd['ExternalReference'].tolist())].reset_index(drop=True)
 
        print("done!")

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
        temp_org = self._your_org[~(self._your_org.iloc[:, 0] == self._leader_id)].reset_index(drop=True)
        _dict = temp_org.groupby(temp_org.columns.values[leader_level + 1]).groups
        for key in _dict:
            self._direct_report_field.append([self.demographics_pd[self.demographics_pd.iloc[:, 0] == key].iloc[0, 1], _dict[key]])

        ## make grade group fields.
        self._grade_group_fields = []
        _dict = self._your_org.groupby("Pay Grade Group").groups
        for key in _dict:
            self._grade_group_fields.append([key, _dict[key]])

        ## make tenure group fields.
        self._tenure_group_fields = []
        _dict = self._your_org.groupby("Length of Service Group").groups
        for key in _dict:
            if not key == "15+ Years":
                self._tenure_group_fields.append([key, _dict[key]])
        self._tenure_group_fields.append(["15+ Years", _dict[key]])
        
        ## make performance rating fields.
        self._performance_rating_fields = []
        _dict = self._your_org.groupby("2019 Performance Rating").groups
        _list = _dict.keys()
        for key in self.performance_order:
            try:
                self._performance_rating_fields.append([key, _dict[key]])
                _list.remove(key)
            except:
                pass
        for key in _list:
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

        for ind in range(len(data.columns)):
            _ = data.iloc[:, ind]
            is_empty_column = True
            lens = nums
            sub = 0
            for __ in _:
                if __ != '':
                    is_empty_column = False
                    sub += __
                else:
                    lens -= 1
            if is_empty_column:
                _dict.update({data.columns.values[ind]: "N/A"})
                determine_parent_na = True
            else:
                _dict.update({data.columns.values[ind]: str(round(sub / lens * 100)) + "%"})
        
        if determine_parent_na:
            _dict.update({item: "N/A"})
        else:
            total_lens = nums
            total = 0
            cols = len(data.columns)
            for ind in range(nums):
                _series = data.iloc[ind, :]
                sub_sum = 0
                _is_nan = False
                for val in _series:
                    if val != '':
                        sub_sum += val
                    else:
                        _is_nan = True
                        total_lens -= 1
                        break

                if not _is_nan:
                    total += sub_sum / cols
            if total_lens == 0:
                _dict.update({item: "N/A"})
            else:
                _dict.update({item: str(round(total / total_lens * 100)) + "%"})
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
        self._calculateSubFields(self._direct_report_field, "Direct Reports (as of April 24, 2018)", item)
        
        ## calcualte Grade Group %s
        self._calculateSubFields(self._grade_group_fields, "Grade Group", item)

        ## calcualte Tenure Group %s
        self._calculateSubFields(self._tenure_group_fields, "Tenure Group", item)

        ## calculate Performance Rating %s
        self._calculateSubFields(self._performance_rating_fields, "2019 Performance Rating", item)

        ## calculate Talent Coordinate %s
        self._calculateSubFields(self._talent_cordinate_fields, "2020 Talent Coordinate", item)

        ## calculate Gender %s
        self._calculateSubFields(self._gender_fields, "Gender", item)

        ## calculate Ethnicity (US) %s
        self._calculateSubFields(self._ethnicity_fields, "Ethnicity (US)", item)

        ## calculate Age Group %s
        self._calculateSubFields(self._age_fields, "Age Group", item)

        ## calculate Country %s
        self._calculateSubFields(self._country_fields, "Country", item)

        ## calculate Kite %s
        self._calculateSubFields(self._kite_fields, "Kite", item)

    def _calculateOverall(self, dataframe, item):

        ids = dataframe.iloc[:, 0]
        nums = len(ids)
        working_pd = self._filtered_raw_data[self._filtered_raw_data["ExternalReference"].isin(ids)].reset_index(drop=True)
        return self._get_sum(working_pd.iloc[:, 1:], nums, item), len(ids)
    
    def _calculateSubFields(self, dataframe, column_name, item):

        for field in dataframe:
            dictionary = self.precious_dict[column_name]
            _list = field[1]
            ids = self._your_org.iloc[_list, 0].tolist()
            nums = len(ids)
            working_pd = self._filtered_raw_data[self._filtered_raw_data["ExternalReference"].isin(ids)].reset_index(drop=True)
            _dict = self._get_sum(working_pd.iloc[:, 1:], nums, item)
            dictionary[field[0]].update(_dict)
            self.first_row.update({column_name + field[0]: len(ids)})

if __name__ == "__main__":

    ## Create a object.
    init_data = {
        'raw_data': "./Qualtrics Survey Export Sample New.xlsx",
        'item_code': "./Item Code New.xlsx",
        'demographics': "./Demographics File Sample 2021-01-17.xlsx",
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