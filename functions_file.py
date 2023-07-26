# this file hold the functioms that are used in mains.py
import pandas as pd
from tkinter import Tk, Listbox, Button, MULTIPLE, Scrollbar
from tkinterdnd2 import TkinterDnD
import os
import tkinter as tk


def display_and_select(items):
    root = Tk()
    root.title('choose course')
    root.geometry('500x500')

    scrollbar = Scrollbar(root)
    scrollbar.pack(side='right', fill='y')

    listbox = Listbox(root, selectmode=MULTIPLE, width=0, yscrollcommand=scrollbar.set)
    listbox.pack(fill='both', expand=True)

    scrollbar.config(command=listbox.yview)

    for item in items:
        item_str = f" ||   שם הקורס:  {item[0]} || סמסטר:  {item[1]} || יום:  {item[2]} || שעת התחלה: {item[3]} || שעת סיום:  {item[4]} ||      "
        listbox.insert('end', item_str)

    selected_indices = []
    selected_values = []

    def on_button_click():
        selected_indices.extend(listbox.curselection())
        selected_values.extend([items[int(item)] for item in listbox.curselection()])
        root.quit()

    button = Button(root, text='Select', command=on_button_click)
    button.pack()

    root.mainloop()

    return selected_indices, selected_values


def check_conflicts_and_write(course_names,rows_header,  input_file="wow.xls", output_file="output.xlsx", conflicts_file="conflicts.xlsx"):
    df = pd.read_excel(input_file)
    selected_courses = df.iloc[course_names]

    grouped = selected_courses.groupby([rows_header[0], rows_header[1]])

    conflicts = []

    for name, group in grouped:
        for i in range(len(group)):
            for j in range(i+1, len(group)):
                course_i = group.iloc[i]
                course_j = group.iloc[j]

                if course_i[2] < course_j[3] and course_i[3] > course_j[2]:
                    conflicts.append({"Conflict": f"{course_i[6]} מתנגש עם {course_j[6]} בסמסטר {name[0]} ,ביום {name[1]}."})

    if conflicts:

        conflicts_df = pd.DataFrame(conflicts)
        conflicts_df.to_excel(conflicts_file, index=False, engine='openpyxl')
        print("\nThere are conflicts in your selected courses. Check the file", conflicts_file)
        return False
    else:
        selected_courses.to_excel(output_file, index=False, engine='openpyxl')
        print("You can register without conflicts. The selected courses have been written to", output_file)
        return True
    

def drag_and_drop():
    def drop(event):
        file_path = event.data
        filename_var.set(os.path.basename(file_path))  
        window.destroy()  

    window = TkinterDnD.Tk()
    window.withdraw()  
    window.geometry("800x600")  

    filename_var = tk.StringVar()  

    drop_target = tk.Label(window, text="Drag and drop the .xls file here")
    drop_target.pack(fill=tk.BOTH, expand=True)
    drop_target.drop_target_register('DND_Files')
    drop_target.dnd_bind('<<Drop>>', drop)

    window.deiconify()  
    window.mainloop()

    return filename_var.get() 
