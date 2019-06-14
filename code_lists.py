import os

import xlrd
from collections import OrderedDict


class DataRetrieval:

    def __init__(self):
        self.download_folder_url = os.path.join(
            os.path.expanduser('~'), 'Downloads/')
        self.document_extension = ' .xlsx'
        self.worksheet_data_retrieval = WorksheetDataRetrival()

    def open_workbook(self):
        workbook_name = str(input('Enter the full name of the workbook: '))
        workbook = xlrd.open_workbook(
            self.download_folder_url + workbook_name + self.document_extension)

        return workbook

    def data_retrieval_option(self):
        data_retrieval_options = str(input(
            'Select 1 if you would like to retrieve data' /
            'from the entire worksheet or 2 to retrieve' /
            'from a specific cell?'))

        if data_retrieval_options == 1:
            return self.worksheet_data_retrieval.find_worksheet(
                self.open_workbook)


class WorksheetDataRetrival(DataRetrieval):

    def __init__(self):
        self.sheet_data = []
        self.sorted_data = []
        self.code_keys = []

    def find_worksheet(self, workbook):
        worksheet_name = str(input('Enter the name of the worksheet: '))
        worksheet = workbook.sheet_by_name(worksheet_name)

        return worksheet

    def read_and_retrieve_worksheet_data(self):
        worksheet = self.find_worksheet(self.open_workbook)
        self.code_keys = [worksheet.cell(
            0, col).value for col in range(worksheet.ncols)]

        for row in range(1, worksheet.nrows):
            sheet_data_dictionary = {self.code_keys[col]: worksheet.cell(
                row, col).value for col in range(worksheet.ncols)}
            self.sheet_data.append(sheet_data_dictionary)

        return self.sheet_data

    def choose_columns(self):
        print('The fields retrieved are as follows', ' '.join(self.code_keys))
        columns = str(input(
            'Enter the names of the two columns you need as they appear' /
            'seperated by a comma: ')).split(',')

        columns_required = list(set(columns).intersection(self.code_keys))

        return columns_required

    def parse_data_to_dictionary_list(self):
        if len(self.code_keys) > 2:
            columns = self.choose_columns()
            final_data = list(map(
                lambda entry: {key: value for key, value in entry.items()
                               if key in columns}, self.read_and_retrieve_worksheet_data()))

            for data in final_data:
                if data.keys() != columns:
                    sorted_dict = dict(sorted(data.items(), reverse=True))
                    self.sorted_data.append(sorted_dict)
                else:
                    self.sorted_data = final_data
        else:
            self.sorted_data = self.read_and_retrieve_worksheet_data()

        return self.sorted_data


generate_json = WorksheetDataRetrival()
generate_json.parse_data_to_dictionary_list()
