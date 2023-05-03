import openpyxl
import tkinter as tk
from tkinter import filedialog
import tkinter.messagebox as msg

user_select = tk.Tk()
user_select.withdraw()

file_path = filedialog.askopenfilename(title="Select MACdongles file", filetypes=[("Excel Files", "*.xlsx")])

if file_path:
    print("File selected:", file_path)
else:
    print("No file selected")

# Load the specified Excel spreadsheet
workbook = openpyxl.load_workbook(file_path)
worksheet = workbook.active

# Create a tkinter window and frame to hold the buttons
root = tk.Tk()
root.title("MACdongles")
frame = tk.Frame(root, padx=20, pady=10, bg="aquamarine")
frame.pack(fill="both", expand=True)

# Define functions to lock and unlock the window
def lock_window():
    root.attributes("-topmost", True)

def unlock_window():
    root.attributes("-topmost", False)

# Define a function to copy the label of the pressed button to the clipboard
def copy_text(label):
    root.clipboard_clear()
    root.clipboard_append(label)
    msg.showinfo("Copy Successful", f"Label '{label}' copied to clipboard.")



# Iterate over the rows in the spreadsheet and create a button for each cell containing a valid MAC address
for row_idx, row in enumerate(worksheet.iter_rows()):
    for col_idx, cell in enumerate(row):
        if cell.value and isinstance(cell.value, str) and len(cell.value) == 17 and all(c.isalnum() for c in cell.value.replace(":", "")):
            label = cell.value
 
            button = tk.Button(frame, text=label, command=lambda l=label: copy_text(l))
            button.grid(row=row_idx, column=col_idx+1, padx=5, pady=5) # add 1 to the column index to make room for the label
            row_label = tk.Label(frame, text=str(row_idx)) # create a label for the row number
            row_label.grid(row=row_idx, column=0, padx=5, pady=5) # add the label to the same row as the button, in the first column

 #           button = tk.Button(frame, text=label, command=lambda l=label: copy_text(l))
 #           button.grid(row=row_idx, column=col_idx, padx=5, pady=5)



# Create buttons to lock and unlock the window
lock_button = tk.Button(root, text="Lock", command=lock_window)
lock_button.pack(side="left", padx=10, pady=10)

unlock_button = tk.Button(root, text="Unlock", command=unlock_window)
unlock_button.pack(side="left", padx=10, pady=10)



# Use the tkinter main loop to display the window
root.mainloop()
