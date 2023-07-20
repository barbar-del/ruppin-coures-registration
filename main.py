import xlrd
import functions_file as func
from tkinter import Tk, Listbox, Button, MULTIPLE

def main():
    print("start of prog")
    # Get the file name by using drag and drop
    file_name=func.drag_and_drop()
    file_name=file_name[:-1]
    workbook = xlrd.open_workbook(file_name)

    # Access the first sheet
    sheet = workbook.sheet_by_index(0)

    # Get the header row
    header_row = sheet.row_values(0)
    courselist = sheet.col_values(6)
    courselist.pop(0)
    # Print the headers
    print("header row:", header_row)

    # Select courses and get their indices
    course_names, index = func.display_and_select(courselist)

    # Check for conflicts and write to file
    func.check_conflicts_and_write(course_names, header_row, file_name)

    print("end of prog")   

if __name__ == "__main__":
    main()
