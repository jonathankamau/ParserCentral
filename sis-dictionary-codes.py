import json
import logging
from code_lists import WorksheetDataRetrival


class GenerateJSON:

    def __init__(self):
        self.worksheet_data = WorksheetDataRetrival()

    def create_json_file(self):
        sis_data = []
        for sis in self.worksheet_data.sorted_data:
            sis_dictionary = {"from": list(sis.values())[
                0], "to": list(sis.values())[1]}
            sis_data.append(sis_dictionary)

        with open('sis_dictionary.json', 'w') as output_file:
            output_file.write(
                '[' +
                ',\n'.join(json.dumps(i) for i in sis_data) +
                ']\n')
        
        print(sis_data)


generate_json = GenerateJSON()
generate_json.create_json_file()
