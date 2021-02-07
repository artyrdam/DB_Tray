import csv
import os
import openpyxl
from tkinter import filedialog as fd


def get_excel_data():
    path = fd.askopenfilename()

    wb_obj = openpyxl.load_workbook(path)
    sheet_obj = wb_obj.active

    data_list = []
    # prints all objects in column 1
    for i in range(1, sheet_obj.max_row + 1):
        cell_obj = sheet_obj.cell(row=i, column=1)
        #print(cell_obj.value)
        data_list.append(cell_obj.value)
    # remove dupes
    data_list = list(dict.fromkeys(data_list))

    x = path.split(".")[0]
    test = x + ".csv"
    #print("path: " + test)

    with open(test, 'w') as myfile:
        wr = csv.writer(myfile, quoting=csv.QUOTE_ALL)
        for val in data_list:
            wr.writerow([val])


if __name__ == "__main__":
    get_excel_data()