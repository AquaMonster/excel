import columnmanager as cm
# from openpyxl.styles import PatternFill


file_path = "c:\\users\\atilley\\desktop\\export4.xlsx"
cm = cm.ColumnManager(file_path, "Bid Breakdown")


def main():
    cm.gen_labordollar_perhour_column(with_formulas=True)
    cm.gen_laborhours_unitarea(with_formulas=True)
    cm.color_column("Labor $/Hr")
    cm.color_column("Labor Hours/Unit Area")
    cm.save_doc()

    # print(Color.__attrs__)


if __name__ == "__main__":
    main()
