import setup
import os
import json

# dress the excel files
# get summary data
# build summary file

SHEET_TITLE = "Bid Breakdown"  # get from config file
JOB_TYPES = '" || Tele/Data Utility, || Power Devices"'  # get from config file


def get_file_paths(file_extention):
    # should this be in setup?
    cwd = os.getcwd()
    files = os.listdir(cwd)
    excel_files = [file for file in files if file[len(file)-5:] == file_extention]
    excel_file_paths = [cwd + "\{}".format(file_name) for file_name in excel_files]
    return excel_file_paths


def dress_excel_files(sheet_title):
    # should this be in setup?
    excel_file_paths = get_file_paths(".xlsx")
    column_managers = setup.get_column_managers(excel_file_paths, sheet_title)
    setup.dress_excel_file(JOB_TYPES, column_managers)


def get_data_structure(column_manager, title):
    """
    build data structure with systems and corresponding values
    """
    job_type = column_manager.get_jobtype()
    file_data = {job_type: {}}
    systems = setup.get_base_bid_values(column_manager, "System")
    values = setup.get_base_bid_values(column_manager, title)

    if len(systems) == len(values):
        try:
            system_values = list(zip(systems, values))

            for system_value in system_values:
                file_data[job_type][system_value[0]] = system_value[1]
        except:
            print("Values list and systems list are not the same length")

    return file_data


def main():
    # dress_excel_files(SHEET_TITLE)
    paths = excel_file_paths = get_file_paths(".xlsx")
    column_managers = setup.get_column_managers(excel_file_paths, SHEET_TITLE)
    labor_hour = get_data_structure(column_managers[0], "Labor $/Hr")
    print(labor_hour["|| Tele/Data Utility"][" || Distribution"])
    print(labor_hour["|| Tele/Data Utility"][" || Temporary Power"])

    with open("config.json", "r") as data:
        things = json.load(data)

    print(things[0])


if __name__ == "__main__":
    main()
