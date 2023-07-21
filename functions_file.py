# this file hold the functioms that are used in mains.py
import pandas as pd
from tkinter import Tk, Listbox, Button, MULTIPLE, Scrollbar
from tkinterdnd2 import TkinterDnD
import os
import tkinter as tk

########################################################3
# this function is used to display the courses in a listbox and select the courses
def display_and_select(items):
    # Create a new tkinter window
    root = Tk()
    root.title('choose course')  # Set the window title
    # Set the window size
    root.geometry('500x500')  # width x height

    # Create a scrollbar
    scrollbar = Scrollbar(root)
    scrollbar.pack(side='right', fill='y')

    # Create a new Listbox widget, set it to allow multiple items to be selected
    # Set the width to 0 to make the Listbox automatically size to fit its contents
    listbox = Listbox(root, selectmode=MULTIPLE, width=0, yscrollcommand=scrollbar.set)
    listbox.pack(fill='both', expand=True)

    # Attach listbox to scrollbar
    scrollbar.config(command=listbox.yview)

    # Add each item to the Listbox
    for item in items:
        listbox.insert('end', item)

    selected_indices = []
    selected_values = []

    # Define a function that will be called when the button is clicked
    def on_button_click():
        # Get a list of indices of the selected items
        selected_indices.extend(listbox.curselection())
        # Get a list of the selected items themselves
        selected_values.extend([items[int(item)] for item in listbox.curselection()])
        # Close the tkinter window
        root.quit()

    # Create a Button that will call on_button_click when clicked
    button = Button(root, text='Select', command=on_button_click)
    button.pack()

    # Run the tkinter event loop
    root.mainloop()

    # After the window is closed, return the selected indices and values
    return selected_indices, selected_values


##############################################################################################################

# this function is used to check the conflicts and write the selected courses to a new file
#if there is a conflict it will return false write the conflicts to a file named conflicts.xlsx
# if there is no conflict it will return true wirte the selected courses to a new file named output.xlsx
#in both cases there will be an otput file

def check_conflicts_and_write(course_names,rows_header,  input_file="wow.xls", output_file="output.xlsx", conflicts_file="conflicts.xlsx"):
    # Load the xls file data into a pandas DataFrame
    df = pd.read_excel(input_file)
    print(input_file)
    # Extract only the rows corresponding to the selected courses
    selected_courses = df.iloc[course_names]

    # Group courses by semester and day
    grouped = selected_courses.groupby([rows_header[0], rows_header[1]]) # Group by semester and day

    conflicts = []  # List to keep track of conflicts

    # For each group (a unique combination of semester and day)
    for name, group in grouped:
        # Check every pair of courses in the group for a time conflict
        for i in range(len(group)):
            for j in range(i+1, len(group)):
                course_i = group.iloc[i]
                course_j = group.iloc[j]

                # If course i starts before course j ends and ends after course j starts, there is a conflict
                if course_i[2] < course_j[3] and course_i[3] > course_j[2]:
                    # Add a description of the conflict to the conflicts list
                    conflicts.append({"Conflict": f"{course_i[6]} מתנגש עם {course_j[6]} בסמסטר {name[0]} ,ביום {name[1]}."})

    # Write the conflicts to a file
    if conflicts:

        conflicts_df = pd.DataFrame(conflicts)
        
        conflicts_df.to_excel(conflicts_file, index=False, engine='openpyxl')
        print("\nThere are conflicts in your selected courses. Check the file", conflicts_file)
        return False
    else:
        # No conflict exists. Write the selected courses to a new .xlsx file
        selected_courses.to_excel(output_file, index=False, engine='openpyxl')
        print("You can register without conflicts. The selected courses have been written to", output_file)
        return True
    
    
    
    
    

##############################################################################################################

def drag_and_drop():
    def drop(event):
        file_path = event.data
        filename_var.set(os.path.basename(file_path))  # Update the filename_var
        window.destroy()  # Close the window after getting the file

    window = TkinterDnD.Tk()
    window.withdraw()  # Hide the main window
    window.geometry("800x600")  # Make the window larger

    filename_var = tk.StringVar()  # Create a StringVar to hold the filename

    drop_target = tk.Label(window, text="Drag and drop the .xls file here")
    drop_target.pack(fill=tk.BOTH, expand=True)
    drop_target.drop_target_register('DND_Files')
    drop_target.dnd_bind('<<Drop>>', drop)

    window.deiconify()  # Show the main window
    window.mainloop()

    return filename_var.get()  # Return the filename

##############################################################################################################