import xlrd
import functions_file as func
from tkinter import Tk, Listbox, Button, MULTIPLE

def main():
    print("start of prog")
    file_name=func.drag_and_drop()
    file_name=file_name[:-1]
    workbook = xlrd.open_workbook(file_name)

    sheet = workbook.sheet_by_index(0)

    header_row = sheet.row_values(0)

    semesterlist = sheet.col_values(0)[1:]
    daylist = sheet.col_values(1)[1:]
    courselist = sheet.col_values(6)[1:]
    startHourlist = sheet.col_values(2)[1:]
    endHourlist = sheet.col_values(3)[1:]

    items = list(zip(courselist, semesterlist, daylist, startHourlist, endHourlist))

    print("header row:", header_row)

    course_names, index = func.display_and_select(items)

    func.check_conflicts_and_write(course_names, header_row, file_name)

    print("end of prog")   

if __name__ == "__main__":
    main()
