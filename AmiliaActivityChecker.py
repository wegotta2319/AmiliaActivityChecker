import pandas as pd
import tkinter as tk
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
            tree_output.delete(*tree_output.get_children())
            tree_output.insert("", "end", values=("Error", "Invalid Input", "Please enter valid numeric values for minimum and maximum costs."))
            return

        if(start_date_input and not end_date_input) or (end_date_input and not start_date_input):
            tree_output.insert("","end", values=("Date Validation", "Error", "Both Start Date and End Date must be filled."))
            return

        if(min_cost_input and not max_cost_input) or (max_cost_input and not min_cost_input):
            tree_output.insert("", "end", values=("Cost Validation", "Error", "Both Minimun Cost and Maximum Cost must be filled."))
            return    

        if not file_path.get():
            raise ValueError("No file selected. Please drag and drop a file.")

        df = pd.read_excel(file_path.get(), sheet_name=sheet_name)

        # Default start and end dates if fields are blank
        start_date = datetime.strptime(start_date_input, "%m/%d/%Y").date() if start_date_input else None
        end_date = datetime.strptime(end_date_input, "%m/%d/%Y").date() if end_date_input else None

        # Check for missing columns
        required_columns = ['Start date', 'End date', 'Cost', 'Activity']
        missing_columns = [col for col in required_columns if col.lower() not in map(str.lower, df.columns)]

        if missing_columns:
            tree_output.delete(*tree_output.get_children())
            tree_output.insert("", "end", values=("Error", "Missing Columns", f"Required columns: {', '.join(missing_columns)}"))
            return

        # Initialize variables
        num_of_wrong_start_date = 0
        num_of_correct_start_date = 0
        num_of_wrong_end_date = 0
        num_of_correct_end_date = 0
        num_of_both_out_of_bounds = 0
        num_of_invalid_cost = 0
        num_of_valid_cost = 0
        total_num_of_activities = len(df)

        # Convert the 'Start date' and 'End date' columns to datetime objects in the dataframe
        df['Start date'] = pd.to_datetime(df['Start date'], format='%m/%d/%Y')
        df['End date'] = pd.to_datetime(df['End date'], format='%m/%d/%Y')

        # Clear previous results
        tree_output.delete(*tree_output.get_children())

        # Iterate through each row in the spreadsheet
        for i, row in df.iterrows():
            row_start_date = row['Start date'].date()
            row_end_date = row['End date'].date()
            row_cost = row['Cost']

            # Both dates out of bounds
            if start_date and end_date and row_start_date < start_date and row_end_date > end_date:
                tree_output.insert("", "end", values=(
                    df.loc[i, 'Activity'],
                    "Both Dates Out of Bounds",
                    f"Start: {row_start_date}, End: {row_end_date}"
                ))
                num_of_both_out_of_bounds += 1
            else:
                # Start date validation
                if start_date and row_start_date < start_date:
                    tree_output.insert("", "end", values=(
                        df.loc[i, 'Activity'],
                        "Invalid Start Date",
                        f"Start: {row_start_date}, Expected >= {start_date}"
                    ))
                    num_of_wrong_start_date += 1
                else:
                    num_of_correct_start_date += 1

                # End date validation
                if end_date and row_end_date > end_date:
                    tree_output.insert("", "end", values=(
                        df.loc[i, 'Activity'],
                        "Invalid End Date",
                        f"End: {row_end_date}, Expected <= {end_date}"
                    ))
                    num_of_wrong_end_date += 1
                else:
                    num_of_correct_end_date += 1

            # Cost validation
            if min_cost and max_cost and (row_cost < min_cost or row_cost > max_cost):
                tree_output.insert("", "end", values=(
                    df.loc[i, 'Activity'],
                    "Invalid Cost",
                    f"Cost: {row_cost}, Expected between {min_cost} and {max_cost}"
                ))
                num_of_invalid_cost += 1
            else:
                num_of_valid_cost += 1

        # Append summary
        tree_output.insert("", "end", values=(
            "Summary", "Valid Start Dates", f"{num_of_correct_start_date} / {total_num_of_activities}"
        ))
        tree_output.insert("", "end", values=(
            "Summary", "Invalid Start Dates", f"{num_of_wrong_start_date} / {total_num_of_activities}"
        ))
        tree_output.insert("", "end", values=(
            "Summary", "Valid End Dates", f"{num_of_correct_end_date} / {total_num_of_activities}"
        ))
        tree_output.insert("", "end", values=(
            "Summary", "Invalid End Dates", f"{num_of_wrong_end_date} / {total_num_of_activities}"
        ))
        tree_output.insert("", "end", values=(
            "Summary", "Valid Costs", f"{num_of_valid_cost}"
        ))
        tree_output.insert("", "end", values=(
            "Summary", "Invalid Costs", f"{num_of_invalid_cost}"
        ))

    except Exception as e:
        # Error Handling
        tree_output.delete(*tree_output.get_children())
        tree_output.insert("", "end", values=("Error", "Exception", str(e)))


def on_file_drop(event):
    file_path.set(event.data)
    label_file_path.config(text=f"Selected File: {event.data}")


def upload_file():
    global file_path

    selected_file = filedialog.askopenfilename(
        title="Select a file",
        filetypes=(("Excel Files", ".*xlsx;*.xls"), ("All Files", "*.*"))
    )
    if selected_file:
        file_path.set(selected_file)
        label_file_path.config(text=f"Selected File: {selected_file}")
    else:
        label_file_path.config(text="No File Selected. Please Try Again.")


def clear_fields():
    tree_output.delete(*tree_output.get_children())
    entry_start_date.delete(0, tk.END)
    entry_end_date.delete(0, tk.END)
    entry_sheet_name.delete(0, tk.END)
    file_path.set("")
    label_file_path.config(text="Drag and Drop a file here")
    entry_min_cost.delete(0, tk.END)
    entry_max_cost.delete(0, tk.END)


def quit_program():
    window.quit()


# GUI setup
window = TkinterDnD.Tk()
window.title("Amilia Date and Cost Checker")

# Set ttk theme
style = ttk.Style()
style.theme_use("classic")

# Styling for all the buttons
style.configure("TButton", font=("Arial", 12))  # Default style for all buttons
style.configure("TButton.validate.TButton", background="lightgreen", font=("Arial", 12, "bold"))
style.configure("TButton.clear.TButton", background="orange", font=("Arial", 12, "bold"))
style.configure("TButton.quit.TButton", background="red", foreground="white", font=("Arial", 12, "bold"))

# Frame for input fields
frame_inputs = ttk.Frame(window, padding=10)
frame_inputs.grid(row=0, column=0, sticky="nsew", padx=10, pady=10)

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
button_clear = ttk.Button(frame_inputs, text="Clear", command=clear_fields, style = "TButton.clear.TButton")
button_clear.grid(row=6, column=1, sticky="ew", pady=5, padx=5)

# Upload File Button
button_upload = ttk.Button(frame_inputs, text="Upload File", command=upload_file, )
button_upload.grid(row=6, column=2, sticky="ew", pady=5, padx=5)

# Quit Button
button_quit = ttk.Button(frame_inputs, text="Quit", command=quit_program, style="TButton.quit.TButton")
button_quit.grid(row=7, column=0, columnspan=2, sticky="ew", pady=5, padx=5)

# Treeview for output
tree_output = ttk.Treeview(frame_output, columns=("Activity", "Issue", "Details"), show="headings", height=15)
tree_output.grid(row=0, column=0, padx=10, pady=10, sticky="nsew")

# Define column headings
tree_output.heading("Activity", text="Activity")
tree_output.heading("Issue", text="Issue")
tree_output.heading("Details", text="Details")

# Configure column widths
tree_output.column("Activity", width=150, anchor="w")
tree_output.column("Issue", width=150, anchor="w")
tree_output.column("Details", width=400, anchor="w")

# Scrollbar for the Treeview
scrollbar = ttk.Scrollbar(frame_output, orient="vertical", command=tree_output.yview)
tree_output.configure(yscrollcommand=scrollbar.set)
scrollbar.grid(row=0, column=1, sticky="ns")

# Configure grid weights
window.columnconfigure(0, weight=1)
window.rowconfigure(1, weight=1)
frame_output.columnconfigure(0, weight=1)
frame_output.rowconfigure(0, weight=1)

# Key Bindings

window.bind("<Return>", lambda event: validate_dates_and_cost())
window.bind("<Escape>", lambda event: quit_program())

# Start the GUI loop
window.mainloop()