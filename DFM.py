import pandas as pd
import os
import openpyxl
from openpyxl.styles import Font, Alignment, PatternFill, Color, Border, Side, numbers
from openpyxl.drawing.image import Image
from openpyxl.utils import cell as ce
from tqdm import tqdm

class DemographicFileMaker:

    def __init__(self, **args):
        ## initialize the object by specifying input and output files.
        self.raw_data_file = args['raw_data']
        self.item_code_file = args['item_code']
        self.demographics_file = args['demographics']
        self.output_path = args['output']
        self._leader_id = args['leader_id']
        self.heatmap_color_file = args['heatmap_color']
        self.image_src = args['image']


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
        self.heatmap_color_pd = pd.read_excel(self.heatmap_color_file, engine="openpyxl")
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
        self.book.save(path)
        print("done!")

    def makeReport(self):

        ## make a content to be displayed in the picture position.
        field_picture_postion = str(round(self._participated / self._invited * 100)) + "% Participation Rate" + \
            "\n" + str(self._participated) + '/' + str(self._invited) + "\n" + "(Participated / Invited)"

        ## prepare image.
        img  = Image(self.image_src)
        img.height = 90
        img.width = 110

        ## calculate total rows that will be placed in our output file.
        total_rows = 4 + len(self._item_list)

        ## make a workbook and sheet.
        self.book = openpyxl.Workbook()
        sheet = self.book.active

        ## set styles like font, color, direction, border...
        ft = Font(name="Arial", size=8)
        ft_bold = Font(name="Arial", size=8, bold=True)
        light_na_font = Font(name="Arial", size=8, color="999999")
        bold_na_font = Font(name="Arial", size=8, color="999999", bold=True)
        vertical = Alignment(textRotation=90, horizontal='center')
        center_alignment = Alignment(horizontal='center')
        right_alignment = Alignment(horizontal='right')
        grey_back = PatternFill("solid", fgColor="EEEEEE")
        white_back = PatternFill("solid", fgColor="FFFFFF")
        side = Side(style='thin', color="CCCCCC")
        thin_border = Border(left=side,
                     right=side,
                     top=side,
                     bottom=side)

        ## merge all needed columns and rows.
        col_num = 1
        skip_list = []
        for key in self.precious_dict:
            delta = len(self.precious_dict[key]) - 1
            if col_num == 1:
                delta += 1
            end_col = delta + col_num
            cell = sheet.cell(row=1, column=col_num)
            cell.font = ft
            cell.value = key
            sheet.merge_cells(start_row=1, start_column=col_num, end_row=2, end_column=end_col)
            col_num += delta + 2

            ## make a skip_list which will be used to make a empty column between groups.
            skip_list.append(col_num - 1)

        ## set empty cells white.
        for row in range(1, total_rows + 1 + 1):
            for col in range(1, col_num - 2 + 1 + 1):
                sheet.cell(row=row, column=col).fill = white_back

        ## prepare the whole data to be placed in the sheet.
        frames = []
        for key in tqdm(self.precious_dict, desc="prepare the whole data"):
            sub_dict = self.precious_dict[key]
            
            ## prepare and write the first column data.
            if key == '':
                _list = []
                _list.append([field_picture_postion, 1])
                _list.append(["Number of Respondents (incl. N/A)", 1])
                for criteria, item in self._item_list:
                    if criteria == 0:
                        _list.append([item, 0])
                    elif criteria == 1:
                        _list.append([self._item_pd[self._item_pd["Item ID"] == item]["Short Text [2020 onward]"].values[0], 1])
                
                ## write the first column data and set styles.
                for index, item in enumerate(_list):
                    cell = sheet.cell(row=3 + index, column=1)
                    if item[1] == 0:
                        cell.font = ft_bold
                    else:
                        cell.font = ft

                    cell.border = thin_border
                    cell.fill = grey_back

                    ## set value to cell.
                    cell.value = item[0]

            ## prepare the rest of data to fill other columns
            for sub_key in sub_dict:
                sub_item = sub_dict[sub_key]
                _list = []
                _list.append([sub_key, 2])
                _list.append([self.first_row[key + sub_key], 1])
                for criteria, item in self._item_list:
                    if criteria == 1:
                        item = self._item_pd[self._item_pd["Item ID"] == item]["Unique Item Code"].values[0]

                    _list.append(sub_item[item])

                frames.append(_list)
        
        ## this method is used to make a empty column when write columns.
        def get_column_number(number):

            for num in skip_list:
                if number >= num:
                    number += 1
            return number
        
        ## below method is used to calculate delta.
        def get_delta(first, second):
            return round(second * 100) - round(first * 100)

        gilead_org = frames[0]
        parent_org = frames[1]
        your_org = frames[2]

        ## write the rest of the columns and set styles.
        for col_index, column in enumerate(tqdm(frames, desc="formating and styling")):
            for row_index, item in enumerate(column):
                cell = sheet.cell(row=3 + row_index, column = get_column_number(2 + col_index))

                ## set borders of all cells.
                cell.border = thin_border

                ## set background grey of the row->4.
                if row_index == 1:
                    cell.fill = grey_back

                ## if the cell is placed in category row, set bold font style. Otherwise set general font style.
                if item[1] == 0:
                    cell.font = ft_bold
                else:
                    cell.font = ft

                ## set corresponding styles to row->3
                if item[1] == 2:
                    cell.alignment = center_alignment
                    cell.alignment = vertical
                    cell.fill = grey_back
                
                ## set value to cell.
                cell.value = item[0]

                ## make "N/A" cell lightgrey.
                if row_index >= 1:
                    if item[0] == "N/A":
                        cell.alignment = right_alignment
                        cell.font = bold_na_font if item[1] == 0 else light_na_font

                ## set background color and set format of percentage to rows below 5.
                if row_index >= 2:

                    ## set format percentage.
                    cell.number_format = numbers.FORMAT_PERCENTAGE

                    ## compare the rest of the columns with your org and set background.
                    if col_index >= 3:
                        try:
                            cell.fill = PatternFill("solid", fgColor=self._get_color(get_delta(your_org[row_index][0], item[0])))
                        except:
                            ## this skip the case of N/A
                            pass

                    ## compare your org (2018) with parent org and set background.
                    elif col_index == 2:
                        try:
                            cell.fill = PatternFill("solid", fgColor=self._get_color(get_delta(parent_org[row_index][0], item[0])))
                        except:
                            ## this skip the case of N/A
                            pass

                    ## compare parent org with gilead org and set background.
                    elif col_index == 1:
                        try:
                            cell.fill = PatternFill("solid", fgColor=self._get_color(get_delta(gilead_org[row_index][0], item[0])))
                        except:
                            ## this skip the case of N/A
                            pass

        ## set width of the first column.
        sheet.column_dimensions['A'].width = 25

        ## set columns' width and empty columns' width.
        for i in range(2, col_num - 2 + 1):
            name = ce.get_column_letter(i)
            if i in skip_list:
                sheet.column_dimensions[name].width = 0.8
            else:
                sheet.column_dimensions[name].width = 3.8
        
        ## set all rows' height.
        for i in range(1, total_rows + 1):
            if i == 3:
                sheet.row_dimensions[i].height = 95
            else:
                sheet.row_dimensions[i].height = 10

        ## freeze panes.
        sheet.freeze_panes = 'E5'

        ## push content right in A3, A4.
        sheet['A3'].alignment = right_alignment
        sheet['A4'].alignment = right_alignment

        ## insert picture and set background white
        sheet['A3'].fill = white_back
        sheet.add_image(img, 'A3')

    def calculateValues(self):
        
        self._prepareColumnsForID()

        item_list = []
        item_dict = {}
        category_list = []

        for key in self.order_category:
            category_list.append(key)
            item_list.append([0, key])
            temp_list = []
            for item in self.category_pd.iloc[self._group_dict[key], 0]:
                item_list.append([1, item])
                temp_list.append(item)
            item_dict.update({key:temp_list})

        self._item_list = item_list

        for category in tqdm(category_list, desc="iterating over the list of category"):
            sub_id_list = item_dict[category]

            self._filterResource(sub_id_list)

            self._calcualteEachRow(category)
        
        print("complete!")
    
    def _get_names_from_field(self, field_list):

        keys = [field[0] for field in field_list]
        return keys


    def _filterResource(self, id_list):

        filter_item = self._item_pd[self._item_pd["Item ID"].isin(id_list)]["Unique Item Code"].tolist()
        filter_item.insert(0, 'ExternalReference')
        self._filtered_raw_data = self.raw_data_pd[filter_item]

    def _get_color(self, delta):
        if delta > 25:
            delta = 25
        elif delta < -25:
            delta = -25
        series = self.heatmap_color_pd[self.heatmap_color_pd["Delta"] == delta]
        return str(hex(series["R"].values[0]))[2:] + str(hex(series["G"].values[0]))[2:] + str(hex(series["B"].values[0]))[2:]

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
                _dict.update({data.columns.values[ind]: ["N/A", 1]})
                determine_parent_na = True
            else:
                _dict.update({data.columns.values[ind]: [sub / lens, 1]})
        
        if determine_parent_na:
            _dict.update({item: ["N/A", 0]})
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
                _dict.update({item: ["N/A", 0]})
            else:
                _dict.update({item: [total / total_lens, 0]})
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
        'heatmap_color': "Heatmap Colors.xlsx",
        'output': "./output for rest of leaders",
        'image': "./image.png",
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