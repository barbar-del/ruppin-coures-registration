
import xlrd
import functions_file as func
from tkinter import Tk, Listbox, Button, MULTIPLE

print("start of prog")
#get the file name by using drag and drop
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

# header_row_reversed = []

# for string in header_row:
#     reversed_string = string[::-1]

#     header_row_reversed.append(reversed_string)
# print ("reversed",header_row_reversed)

course_names,index= func.display_and_select(courselist)
# print(chose)
# print(index)
my_indices = [6,15,7]  # Replace with the indices of your selected courses
func.check_conflicts_and_write(course_names,header_row,file_name)

################################################



print("end of prog")    



