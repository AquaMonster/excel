import columnmanager
import os


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
