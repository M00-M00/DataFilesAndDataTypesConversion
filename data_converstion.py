import json
import csv
import os
import xlrd
from xlsxwriter.workbook import Workbook


# dict_data = {"a": {dict}, "b": {dict} ...}

def create_dict_to_csv(filename: str, dict_data: dict):
    with open(filename, "w") as outfile:
        w = csv.DictWriter(outfile, dict_data[0].keys())
        w.writeheader()
        for a in dict_data:
            w.writerow(dict_data[a])

def append__dict_to_csv(filename: str, data: dict):
    with open(filename, "ab") as outfile:
        w = csv.DictWriter(outfile, data.keys())
        w.writerow(data)


# dict_data = [{dict},{dict} ...}
def create_list_of_dict_to_csv(filename: str, dict_data: list):
    with open(filename, "w") as outfile:
        w = csv.DictWriter(outfile, dict_data[0].keys())
        w.writeheader()
        for d in dict_data:
            w.writerow(d)


def append_list_of_dict_to_csv(filename: str, data: list):
    with open(filename, "ab") as outfile:
        w = csv.DictWriter(outfile, data.keys())
        for d in data: 
            w.writerow(d)

def dictionary_from_dictionary_with_key(old_dict: dict, key) -> dict:
    new_dict = dict( [ (old_dict[a][key] , old_dict[a]) for a in old_dict] )
    return new_dict


def flatten_dict(dd: dict, separator: str ="_" , prefix: str  ="") -> dict: 
    return { str(prefix) + separator + k if prefix else k : v 
             for kk, vv in dd.items() 
             for k, v in flatten_dict(vv, separator, kk).items() 
             } if isinstance(dd, dict) else { prefix : dd } 


def dict_save_json(filename:str, data: dict):
    with open(filename,"w") as outfile:
        json.dump(data, outfile)


def load_dict_from_json(filename:str) -> dict :
    if os.path.exists(filename) == False:
        print("File doesn't exist")
    else:
        with open(filename) as j:
            json_data = json.load(j)
            return(json_data)


def list_to_dict(list:list) -> dict:
    result_dict = {}
    for n in list:
        result_dict[list.index(n)] = n
    return result_dict


def list_to_dict_with_key(list: list, key: str) -> dict:
    result_dict = {}
    for n in list:    
        k = n[key]
        result_dict[k] = n
    return result_dict


def change_dict_key(d: dict, key: str) -> dict:
    result_dict = {}
    for n in d:
        k = d[n][key]
        result_dict[k] = d[n]
    return result_dict


def csv_to_excel(filename: str):
    output_filename = filename.split(".")[0] + ".csv"
    workbook = Workbook(filename[:-4] + ".xlsx")
    worksheet = workbook.add_worksheet()
    with open(output_filename, "rt") as f:
        reader = csv.reader(f)
        for r, row in enumerate(reader):
            for c, col in enumerate(row):
                worksheet.write(r, c, col)
    workbook.close()


def excel_to_json(input_file: str):
    workbook = xlrd.open_workbook(input_file)
    worksheet = workbook.sheet_by_index(0)
    output_filename = input_file.split(".")[0] + ".json"
    data = []
    keys = [v.value for v in worksheet.row(0)]
    for row_number in range(worksheet.nrows):
        if row_number == 0:
            continue
        row_data = {}
        for col_number, cell in enumerate(worksheet.row(row_number)):
            row_data[keys[col_number]] = cell.value
        data.append(row_data)
    with open(output_filename, 'w') as json_file:
        json_file.write(json.dumps({'data': data}))


