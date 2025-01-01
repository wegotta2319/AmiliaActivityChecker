'''
import subprocess
import sys

# Function to install required libraries
def install_requirements():
    try:
        import pandas
        import tkinter
        import tkinterdnd2
    except ImportError:
        print("Required libraries not found. Installing now...")
        subprocess.check_call([sys.executable, "-m", "pip", "install", "pandas", "tk", "tkinterdnd2"])
        print("Requirements installed.")


# Run the installation check
install_requirements()
'''

import pandas as pd
import tkinter as tk
import numpy as np
from tkinter import ttk, filedialog
from tkinterdnd2 import TkinterDnD, DND_FILES
from datetime import datetime


def validate_dates_and_cost():
    try:
        sheet_name = entry_sheet_name.get()
        start_date_input = entry_start_date.get()
        end_date_input = entry_end_date.get()
        min_cost_input = entry_min_cost.get()
        max_cost_input = entry_max_cost.get()

        # Validate and process the min and max cost inputs
        try:
            min_cost = float(min_cost_input) if min_cost_input else 0
            max_cost = float(max_cost_input) if max_cost_input else float('inf')
            
            min_cost = round(min_cost, 3)
            max_cost = round(max_cost, 3)
        except ValueError:
            text_output.delete(1.0, tk.END)
            text_output.insert(tk.END, "Error: Please enter valid numeric values for minimum and maximum costs.")
            return
         
        if not file_path.get():
            raise ValueError("No file selected. Please drag and drop a file.")

        df = pd.read_excel(file_path.get(), sheet_name=sheet_name)

        # Default start and end dates if fields are blank
        start_date = datetime.strptime(start_date_input, "%m/%d/%Y").date() if start_date_input else None
        end_date = datetime.strptime(end_date_input, "%m/%d/%Y").date() if end_date_input else None

        # Initialize all variables
        num_of_wrong_start_date = 0
        num_of_correct_start_date = 0
        num_of_wrong_end_date = 0
        num_of_correct_end_date = 0
        num_of_both_out_of_bounds = 0
        num_of_invalid_cost = 0
        num_of_valid_cost = 0
        total_num_of_activities = len(df)

        results_text = "Validation Results:\n\n"

        # Convert the 'Start date' and 'End date' columns to datetime objects in the dataframe
        df['Start date'] = pd.to_datetime(df['Start date'], format='%m/%d/%Y')
        df['End date'] = pd.to_datetime(df['End date'], format='%m/%d/%Y')
        

        # Checks to see that both date fields are filled
        if (start_date_input and not end_date_input) or (end_date_input and not start_date_input):
                text_output.delete(1.0, tk.END)
                text_output.insert(tk.END, "Error: Make sure both Date fields are filled in.")
                return
        # Checks that both cost fields are filled
        elif (min_cost_input and not max_cost_input) or (max_cost_input and not min_cost_input):
            text_output.delete(1.0, tk.END)
            text_output.insert(tk.END, "Error: Make sure both Cost fields are filled in.")
            return
        # iterates through each row in the spread sheet
        for i, row in df.iterrows():
            row_start_date = row['Start date'].date()
            row_end_date = row['End date'].date()
            row_cost = row['Cost']  # Changed to 'Cost' as per your specification

            # Check if both dates are out of bounds, only if start_date and end_date are provided
            if start_date and end_date:
                if row_start_date < start_date and row_end_date > end_date:
                    results_text += (
                        f"{df.loc[i, 'Activity']} has both start and end dates out of bounds "
                        f"(Start: {row_start_date}, End: {row_end_date}).\n"
                    )
                    num_of_both_out_of_bounds += 1
                else:
                    # Validate start date
                    if row_start_date < start_date:
                        results_text += (
                            f"{df.loc[i, 'Activity']} start date is before the desired start date "
                            f"(Start: {row_start_date}).\n"
                        )
                        num_of_wrong_start_date += 1
                    else:
                        num_of_correct_start_date += 1

                    # Validate end date
                    if row_end_date > end_date:
                        results_text += (
                            f"{df.loc[i, 'Activity']} end date is after the desired end date "
                            f"(End: {row_end_date}).\n"
                        )
                        num_of_wrong_end_date += 1
                    else:
                        num_of_correct_end_date += 1
            
            
            # Validate cost
            if min_cost and max_cost:
                if row_cost < min_cost or row_cost > max_cost:
                    results_text += (
                        f"{df.loc[i, 'Activity']} has an invalid cost (Cost: {row_cost}). "
                        f"Expected cost between {min_cost} and {max_cost}.\n"
                    )
                    num_of_invalid_cost += 1
                else:
                    num_of_valid_cost += 1

        # If only costs are given only prints cost info
        if start_date == None and end_date == None:
                results_text += (
                f"\n\nDate Summary:\n\n"
                f"No dates given to validate.\n"
                f"\nCost Summary:\n\n"
                f"Number of valid costs: {num_of_valid_cost}\n"
                f"Number of invalid costs: {num_of_invalid_cost}\n"
            )
        # If only dates are given only prints dates info
        elif min_cost == None and max_cost == None:
            results_text += (
                f"\n\nDate Summary:\n\n"
                f"Number of valid start dates: {num_of_correct_start_date} / {total_num_of_activities}\n"
                f"Number of invalid start dates: {num_of_wrong_start_date} / {total_num_of_activities}\n"
                f"Number of valid end dates: {num_of_correct_end_date} / {total_num_of_activities}\n"
                f"Number of invalid end dates: {num_of_wrong_end_date} / {total_num_of_activities}\n\n"
                f"No price range given to validate.\n"
            )
        # If both are given print all info
        else:    
            results_text += (
                f"\n\nDate Summary:\n\n" 
                f"Number of valid start dates: {num_of_correct_start_date} / {total_num_of_activities}\n"
                f"Number of invalid start dates: {num_of_wrong_start_date} / {total_num_of_activities}\n"
                f"Number of valid end dates: {num_of_correct_end_date} / {total_num_of_activities}\n"
                f"Number of invalid end dates: {num_of_wrong_end_date} / {total_num_of_activities}\n"
                f"\nCost Summary:\n\n"
                f"Number of valid costs: {num_of_valid_cost}\n"
                f"Number of invalid costs: {num_of_invalid_cost}\n"
            )
            

        text_output.delete(1.0, tk.END)
        text_output.insert(tk.END, results_text)

    except Exception as e:
        text_output.delete(1.0, tk.END)
        text_output.insert(tk.END, f"Error: {str(e)}")

# Handles Drapg and Drop Feature
def on_file_drop(event):
    file_path.set(event.data)
    label_file_path.config(text=f"Selected File: {event.data}")

# Handles Manual Upload Feature
def upload_file():
    global file_path

    selected_file = filedialog.askopenfilename(
        title="Select a file",
        filetypes=(("Excel Files", ".*xlsx;*.xls"), ("All Files", "*.*"))
    )
    if selected_file:
        file_path.set(selected_file)
        label_file_path.config(text=f"Selected Files: {selected_file}")
    else:
        label_file_path.config(text="No File Selected. Please Try Again.")

def clear_fields():
    #clears all fields and the file path
    entry_start_date.delete(0,tk.END)
    entry_end_date.delete(0,tk.END)
    entry_sheet_name.delete(0,tk.END)
    file_path.set("")
    label_file_path.config(text = "Drag and Drop a file here")
    text_output.delete(1.0,tk.END)
    entry_min_cost.delete(0,tk.END)
    entry_max_cost.delete(0,tk.END)

def quit_program():
    #Quits the program
    window.quit()


# GUI setup
window = TkinterDnD.Tk()
window.title("Amilia Date and Cost Checker")

# Set ttk theme
style = ttk.Style()
style.theme_use("classic")  # Options: "clam", "alt", "default", "classic", etc.

# Configure Style
style.configure("TButton", font=("Arial", 12))  # Default style for all buttons
style.configure("TButton.validate.TButton", background="lightgreen", font=("Arial", 12, "bold"))
style.configure("TButton.clear.TButton", background="amber", foreground="darkamber", font=("Arial", 12, "bold"))
style.configure("TButton.quit.TButton", background="red", foreground="white", font=("Arial", 12, "bold"))


# Configure for window resizing (set all rows and columns to be resizable)
window.grid_rowconfigure(0, weight=1)  # Row 0 (input fields frame)
window.grid_rowconfigure(1, weight=1)  # Row 1 (output frame)
window.grid_columnconfigure(0, weight=1)  # Column 0 (main column)

# Frame for input fields
frame_inputs = ttk.Frame(window, padding=10)
frame_inputs.grid(row=0, column=0, sticky="nsew", padx=10, pady=10, )

# Frame for output
frame_output = ttk.Frame(window, padding=10)
frame_output.grid(row=1, column=0, sticky="nsew", padx=10, pady=10)

# File path label
file_path = tk.StringVar()
label_file_path = ttk.Label(frame_inputs, text="Drag and drop a file here", relief="solid", font=("Arial", 14))
label_file_path.grid(row=0, column=0, columnspan=2, sticky="ew", pady=5, ipadx=10, ipady=10)

label_file_path.drop_target_register(DND_FILES)
label_file_path.dnd_bind('<<Drop>>', on_file_drop)

# Input labels and entry fields
ttk.Label(frame_inputs, text="Start Date (MM/DD/YYYY):").grid(row=1, column=0, sticky="w", pady=5)
entry_start_date = ttk.Entry(frame_inputs, width=30)
entry_start_date.grid(row=1, column=1, sticky="ew", pady=5)

ttk.Label(frame_inputs, text="End Date (MM/DD/YYYY):").grid(row=2, column=0, sticky="w", pady=5)
entry_end_date = ttk.Entry(frame_inputs, width=30)
entry_end_date.grid(row=2, column=1, sticky="ew", pady=5)

ttk.Label(frame_inputs, text="Minimum Cost:").grid(row=3, column=0, sticky="w", pady=5)
entry_min_cost = ttk.Entry(frame_inputs, width=30)
entry_min_cost.grid(row=3, column=1, sticky="ew", pady=5)

ttk.Label(frame_inputs, text="Maximum Cost:").grid(row=4, column=0, sticky="w", pady=5)
entry_max_cost = ttk.Entry(frame_inputs, width=30)
entry_max_cost.grid(row=4, column=1, sticky="ew", pady=5)

ttk.Label(frame_inputs, text="Sheet Name (Required):").grid(row=5, column=0, sticky="w", pady=5)
entry_sheet_name = ttk.Entry(frame_inputs, width=30)
entry_sheet_name.grid(row=5, column=1, sticky="ew", pady=5)

# Validate Button
button_validate = ttk.Button(frame_inputs, text="Validate", command=validate_dates_and_cost, style="TButton.validate.TButton")
button_validate.grid(row=6, column=0, sticky="ew", pady=5, padx=5)

# Clear Button
button_clear = ttk.Button(frame_inputs, text="Clear", command=clear_fields, style="TButton.clean.TButton")
button_clear.grid(row=6, column=1, sticky="ew", pady=5, padx=5)

# Upload File Button
button_upload = ttk.Button(frame_inputs, text="Upload File", command=upload_file)
button_upload.grid(row=6, column=2 , sticky="ew", pady=5, padx=5)

# Quit Button
button_quit = ttk.Button(frame_inputs, text="Quit", command=quit_program, style="TButton.quit.TButton")
button_quit.grid(row=7, column=0, columnspan=2, sticky="ew", pady=5, padx=5)

# Output Textbox
text_output = tk.Text(frame_output, wrap="word", height=15, width=40, font=("Ariel", 12), bd=2, relief="solid")
text_output.grid(row=0, column=0, padx=10, pady=10, sticky="nsew")

# Configure the output frame to expand with window size
frame_output.grid_rowconfigure(0, weight=1)
frame_output.grid_columnconfigure(0, weight=1)

# Start the Tkinter loop
window.mainloop()