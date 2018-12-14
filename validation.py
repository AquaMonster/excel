import columnmanager as cm
from openpyxl.worksheet.datavalidation import DataValidation
import os


cwd = os.getcwd()
file_name = "book1.xlsx"  # change in config file
FILE_PATH = cwd + "\{}".format(file_name)

cm = cm.ColumnManager(FILE_PATH, "Sheet1")

name_string = "TI,Shell & Core,Resi"


# def add_validation(validation_list, location):
#     dv = DataValidation(type="list", formula1=''.format(validation_list), allow_blank=True)
#
#     cell = cm.sheet[location]
#     cell.value = "Choose Job Type"
#     dv.add(cell)


def main():
    cm.sheet.insert_rows(1)
    dv = DataValidation(type="list", formula1='"alex,hannah,frog"', allow_blank=True)
    cm.sheet.add_data_validation(dv)
    c1 = cm.sheet["A1"]
    c1.value = "Dog"
    dv.add(c1)

    cm.save_doc(FILE_PATH)


if __name__ == "__main__":
    main()
