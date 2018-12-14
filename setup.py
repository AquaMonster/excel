import columnmanager


def get_column_managers(excel_file_paths, SHEET_TITLE):
    """
    return list of manager objects for each excel, given file paths
    """
    try:
        return [columnmanager.ColumnManager(file_path, SHEET_TITLE)
                for file_path in excel_file_paths]
    except:
        print("Problem getting column managers")


def get_base_bid_values(columnmanager, title):
    bid_item_values = columnmanager.get_column_values("Bid Item", row=2)
    system_values = columnmanager.get_column_values(title, row=2)

    results = []

    for value in range(len(bid_item_values)):
        if bid_item_values[value] == " || Base Bid":
            results.append(system_values[value])
    return results


def build_metrics_columns(manager):
    """
    uses excel manager object to build new metrics columns
    """
    manager.gen_labordollar_perhour_column(with_formulas=False)
    manager.gen_laborhours_unitarea(with_formulas=False)
    manager.color_column("Labor $/Hr")
    manager.color_column("Labor Hours/Unit Area")


def dress_excel_file(JOB_TYPES, column_managers):
    """
    adds datavalidation and metrics columns to given excel file manager objects
    """
    for manager in column_managers:
        # this try block prevents traceback errors from being shown in console?
        try:
            build_metrics_columns(manager)
            manager.add_validation(JOB_TYPES)
            manager.save_doc()
        except:
            print("Problem with column manager named {}".format(manager))
            manager.show_error_location()
