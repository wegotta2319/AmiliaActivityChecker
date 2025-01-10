import pandas as pd
import tkinter as tk
from fuzzywuzzy import fuzz
from tkinter import *
from tkcalendar import DateEntry
from tkinter import ttk, filedialog, Listbox, messagebox
from tkinterdnd2 import TkinterDnD, DND_FILES
from datetime import datetime

# Initialization of variable to take user input
entry_sheet_name = None
entry_start_date = None
entry_end_date = None
entry_min_cost = None
entry_max_cost = None
entry_ledger_code = None

def search_treeview(treeview, search_term):
    # Clean the search term by splitting and lowercasing
    search_terms = [term.strip().lower() for term in search_term.split()]  
    
    if not search_terms:
        messagebox.showinfo("Search", "Please enter a search term.")
        return
    
    found = False
    highest_match = None
    highest_score = 0  # Track the highest fuzzy match score
    
    # Loop through the items and remove any previous highlights
    for item in treeview.get_children():
        treeview.item(item, tags=())  # Remove any previous highlight tag
    
    # Loop through the items again to find matches
    for item in treeview.get_children():
        item_values = treeview.item(item, "values")
        
        # We will now check the similarity score for the "Activity" column only
        activity = str(item_values[0]).lower()  # Assuming activity is the first column
        
        # Use fuzzywuzzy to get the similarity score for each item
        score = fuzz.partial_ratio(search_term.lower(), activity)  # Partial ratio for fuzzy matching
        
        if score > highest_score:
            highest_match = item
            highest_score = score
        
        # If the score is above a certain threshold, we consider it a match
        # You can adjust this threshold as per your needs (e.g., 80 out of 100)
        if score >= 70:  
            treeview.item(item, tags=('highlight',))
            found = True
    
    if highest_match:
        treeview.item(highest_match, tags=('highlight',))  # Highlight the most similar activity
    
    if not found:
        messagebox.showinfo("Search", "No similar matches found.")
    else:
        # Apply the highlight tag with a background color change
        treeview.tag_configure('highlight', background='yellow')

def search_output():
    search_term = search_entry.get()  # Get the search term from the user input
    search_treeview(tree_output, search_term)  # Perform search in the Treeview


def validate_dates_and_cost():
    try:
        # Retreaves all user inputs
        sheet_name = entry_sheet_name.get()
        start_date_input = entry_start_date.get()
        end_date_input = entry_end_date.get()
        min_cost_input = entry_min_cost.get()
        max_cost_input = entry_max_cost.get()
        ledger_code_input = entry_ledger_code.get()

        try:
            min_cost = float(min_cost_input) if min_cost_input else 0
            max_cost = float(max_cost_input) if max_cost_input else float('inf')

            min_cost = round(min_cost, 3)
            max_cost = round(max_cost, 3)
        except ValueError:
            tree_output.delete(*tree_output.get_children())
            tree_output.insert("", "end", values=("Error", "Invalid Input", "Please enter valid numeric values for minimum and maximum costs."))
            return
        
        # If one date box is filled then both boxes have to be filled (can be the same date if looking for a specific date)
        if (start_date_input and not end_date_input) or (end_date_input and not start_date_input):
            tree_output.insert("", "end", values=("Date Validation", "Error", "Both Start Date and End Date must be filled."))
            return

        # If one cost box is filled requires both boxes to be filled (can be the same value)
        if (min_cost_input and not max_cost_input) or (max_cost_input and not min_cost_input):
            tree_output.insert("", "end", values=("Cost Validation", "Error", "Both Minimum Cost and Maximum Cost must be filled."))
            return    

        # Checks to see if there is a file/file path selected
        if not file_path.get():
            raise ValueError("No file selected. Please drag and drop a file.")

        # Uploads excel sheet to a panda dataframe to be analyzed
        df = pd.read_excel(file_path.get(), sheet_name=sheet_name)
        global filtered_df
        filtered_df = df
        
        # Changes user date inputs to 'date' data type in order to be compared to the sheet dates
        start_date = datetime.strptime(start_date_input, "%m/%d/%y").date() if start_date_input else None
        end_date = datetime.strptime(end_date_input, "%m/%d/%y").date() if end_date_input else None

        # Defines the columns required of the sheet
        required_columns = ['Start date', 'End date', 'Cost', 'Activity', 'Ledger code']
        missing_columns = [col for col in required_columns if col.lower() not in map(str.lower, df.columns)]
        
        # Gets rid of the empty columns if there are any
        if missing_columns:
            tree_output.delete(*tree_output.get_children())
            tree_output.insert("", "end", values=("Error", "Missing Columns", f"Required columns: {', '.join(missing_columns)}"))
            return

        # Initializes all variables counters
        num_of_wrong_start_date = 0
        num_of_correct_start_date = 0
        num_of_wrong_end_date = 0
        num_of_correct_end_date = 0
        num_of_both_out_of_bounds = 0
        num_of_invalid_cost = 0
        num_of_valid_cost = 0
        num_of_ledger_code_activities = 0
        total_num_of_activities = len(df)

        df['Start date'] = pd.to_datetime(df['Start date'], format='%m/%d/%y')
        df['End date'] = pd.to_datetime(df['End date'], format='%m/%d/%y')

        tree_output.delete(*tree_output.get_children())
        # Checks to see if the user entered a ledger code to be searched for
        if ledger_code_input:
            ledger_code_input = ledger_code_input.strip() # Deletes the extra spaces before and after the code to make it easier for the user to copy and paste the code from the sheet
            df = df[df['Ledger code'].str.contains(ledger_code_input, na=False, case=False)]
            num_of_ledger_code_activities = len(df)
            if df.empty:
                tree_output.insert("", "end", values=("No Match", "Ledger code", f"No entries found for {ledger_code_input}."))
                return
        else:
            num_of_ledger_code_activities = 0
        # Loop to check every row in the excel sheet
        for i, row in df.iterrows():
            row_start_date = row['Start date'].date()
            row_end_date = row['End date'].date()
            row_cost = row['Cost']

            # Checks start date and end date to see if its within the time frame
            if start_date and end_date and row_start_date < start_date and row_end_date > end_date:
                tree_output.insert("", "end", values=(
                    df.loc[i, 'Activity'],
                    "Both Dates Out of Bounds",
                    f"Start: {row_start_date}, End: {row_end_date}"
                ))
                num_of_both_out_of_bounds += 1
            else:
                if start_date and row_start_date < start_date:
                    tree_output.insert("", "end", values=(
                        df.loc[i, 'Activity'],
                        "Invalid Start Date",
                        f"Start: {row_start_date}, starts before the expected start date of: {start_date}"
                    ))
                    num_of_wrong_start_date += 1
                else:
                    num_of_correct_start_date += 1

                if end_date and row_end_date > end_date:
                    tree_output.insert("", "end", values=(
                        df.loc[i, 'Activity'],
                        "Invalid End Date",
                        f"End: {row_end_date}, ends after the expected end date of: {end_date}"
                    ))
                    num_of_wrong_end_date += 1
                else:
                    num_of_correct_end_date += 1
            # checks all costs with in the sheet to make sure it is in range
            if min_cost and max_cost and (row_cost < min_cost or row_cost > max_cost):
                tree_output.insert("", "end", values=(
                    df.loc[i, 'Activity'],
                    "Invalid Cost",
                    f"Cost: {row_cost}, Expected between {min_cost} and {max_cost}"
                ))
                num_of_invalid_cost += 1
            else:
                num_of_valid_cost += 1
                
        # Itterates through the sheet for the ledger codes   
        if ledger_code_input:
            for i, row in df.iterrows():
                tree_output.insert("", "end", values=(
                    row['Activity'],
                    "Ledger Code Match",
                    f"Ledgercode: {row['Ledger code']}"
                ))
            
            
            # Insert ledger code summary
            tree_output.insert("", "end", values=(
                "Summary", 
                "Ledger Code Matches", 
                f"{num_of_ledger_code_activities} activities match Ledger Code: {ledger_code_input}"
            ))

        # Existing summaries for start dates, end dates, and costs (print the tree)
        tree_output.insert("", "end", values=(
            "Summary", "Valid Start Dates", f"{max(0, num_of_correct_start_date)} / {total_num_of_activities}"
        ))
        tree_output.insert("", "end", values=(
            "Summary", "Invalid Start Dates", f"{max(0, num_of_wrong_start_date)} / {total_num_of_activities}"
        ))
        tree_output.insert("", "end", values=(
            "Summary", "Valid End Dates", f"{max(0, num_of_correct_end_date)} / {total_num_of_activities}"
        ))
        tree_output.insert("", "end", values=(
            "Summary", "Invalid End Dates", f"{max(0, num_of_wrong_end_date)} / {total_num_of_activities}"
        ))
        tree_output.insert("", "end", values=(
            "Summary", "Valid Costs", f"{max(0, num_of_valid_cost)}"
        ))
        tree_output.insert("", "end", values=(
            "Summary", "Invalid Costs", f"{max(0, num_of_invalid_cost)}"
        ))

    except Exception as e:
        tree_output.delete(*tree_output.get_children())
        tree_output.insert("", "end", values=("Error", "Exception", str(e)))

def show_calendar_start(event=None):
    global calendar_start
    calendar_start = DateEntry(window, date_pattern="dd/mm/yyyy")
    calendar_start.place(x=entry_start_date.winfo_x(), y=entry_start_date.winfo_y() + entry_start_date.winfo_height())
    calendar_start.bind("<FocusOut>", lambda e: calendar_start.place_forget())


def show_calendar_end(event=None):
    global calendar_end
    calendar_end = DateEntry(window, date_pattern = "dd/mm/yyyy")
    calendar_end.place(x=entry_start_date.winfo_x(), y=entry_start_date.winfo_y() + entry_start_date.winfo_height())
    calendar_end.bind("<FocusOut>", lambda e: calendar_start.place_forget())

def hide_calendar(event):
    if 'calendar_start' in globals() and 'calendar_end' in globals():
        if event.widget not in (entry_start_date, entry_end_date, calendar_start, calendar_end):
            if calendar_start.winfo_ismapped():
                calendar_start.place_forget()
            if calendar_end.winfo_ismapped():
                calendar_end.place_forget()

# Function to extract ledger codes after file upload:
def extract_ledger_codes():
    if not file_path.get():
        label_file_path.config(text="No file selected.")
        return []
    try:
        df = pd.read_excel(file_path.get())
        if 'Ledger code' not in df.columns:
            label_file_path.config(text="No 'Ledger code' column found in the file.")
            return []
        unique_ledger_codes = df['Ledger code'].dropna().unique()
        return sorted(map(str, unique_ledger_codes))
    except Exception as e:
        label_file_path.config(text=f"Error reading file: {e}")
        return []

# Autocomplete functionality
def autocomplete_ledger_code(event):
    if 'suggestion_box' not in globals():
        setup_suggestion_box()

    # Get the current input
    query = entry_ledger_code.get().lower()

    # Filter ledger codes based on input
    filtered_codes = [code for code in ledger_codes if query in code.lower()]

    # Update the suggestion box
    suggestion_box.delete(0, tk.END)
    for code in filtered_codes:
        suggestion_box.insert(tk.END, code)

    # Adjust height dynamically based on the filtered results
    suggestion_box.config(height=min(10, len(filtered_codes)))

    # Show the suggestion box
    show_suggestion_box()


def update_suggestion_box(suggestions):
    suggestion_box.delete(0, tk.END)
    for suggestion in suggestions:
        suggestion_box.insert(tk.END, suggestion)

# Function to handle suggestion selection
def select_suggestion(event):
    selected_index = suggestion_box.curselection()
    if selected_index:
        selected_code = suggestion_box.get(selected_index)
        entry_ledger_code.delete(0, tk.END)
        entry_ledger_code.insert(0, selected_code)
    suggestion_box.place_forget()  # Hide the suggestion box after selection

# Setup ledger code suggestion box
def setup_suggestion_box():
    global suggestion_box, ledger_codes

    # Extract ledger codes
    ledger_codes = extract_ledger_codes()

    # Initialize suggestion box if not already created
    if 'suggestion_box' not in globals():
        suggestion_box = Listbox(window)

        # Handle selection from the suggestion box
        suggestion_box.bind("<<ListboxSelect>>", select_suggestion)

    # Dynamically adjust suggestion box height to fit the number of ledger codes
    suggestion_box.config(height=min(10, len(ledger_codes)))  # Show up to 10 items

    # Bind <FocusOut> to the root window to hide the suggestion box
    window.bind("<Button-1>", handle_focus_out)

def handle_focus_out(event):
    # Get the widget under the mouse pointer
    widget = event.widget

    # Check if the clicked widget is not the suggestion box or entry field
    if widget not in (entry_ledger_code, suggestion_box):
        suggestion_box.place_forget()


# Trigger suggestion box placement
def show_suggestion_box(event=None):
    if 'suggestion_box' not in globals():
        setup_suggestion_box()

    # Populate the suggestion box with all ledger codes
    suggestion_box.delete(0, tk.END)  # Clear any existing items
    for code in ledger_codes:
        suggestion_box.insert(tk.END, code)

    # Position the suggestion box below the entry field
    x = entry_ledger_code.winfo_rootx() - window.winfo_rootx()
    y = entry_ledger_code.winfo_rooty() + entry_ledger_code.winfo_height() - window.winfo_rooty()
    suggestion_box.place(x=x, y=y, width=entry_ledger_code.winfo_width())

    # Ensure entry retains focus to allow typing
    entry_ledger_code.focus_set()


# Allows user to drag and drop a file to be proccessed
def on_file_drop(event):
    file_path.set(event.data)
    label_file_path.config(text=f"Selected File: {event.data}")

# Allows user to upload a file from system
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


# Clears all fields in the program
def clear_fields():
    tree_output.delete(*tree_output.get_children())
    entry_start_date.delete(0, tk.END)
    entry_end_date.delete(0, tk.END)
    entry_sheet_name.delete(0, tk.END)
    file_path.set("")
    label_file_path.config(text="Drag and Drop a file here")
    entry_min_cost.delete(0, tk.END)
    entry_max_cost.delete(0, tk.END)
    entry_ledger_code.delete(0,tk.END)
    

# Quits the program
def quit_program():
    window.quit()

# GUI setup
window = TkinterDnD.Tk()
window.title("Amilia Date and Cost Checker")
window.geometry("800x600")

# Set ttk theme
style = ttk.Style()
style.theme_use("clam")

# Styling for buttons
style.configure("TButton", font=("Arial", 12))  # Default style for all buttons
style.configure("TButton.validate.TButton", background="lightgreen", font=("Arial", 12, "bold"))
style.configure("TButton.clear.TButton", background="orange", font=("Arial", 12, "bold"))
style.configure("TButton.quit.TButton", background="red", foreground="white", font=("Arial", 12, "bold"))

# Frames
## Input frame
frame_inputs = ttk.Frame(window, padding=10)
frame_inputs.grid(row=0, column=0, sticky="nsew", padx=10, pady=10)

## Output frame
frame_output = ttk.Frame(window, padding=10)
frame_output.grid(row=1, column=0, sticky="nsew", padx=10, pady=10)

# File path label
file_path = tk.StringVar()
label_file_path = ttk.Label(frame_inputs, text="Drag and drop a file here", relief="solid", font=("Arial", 14))
label_file_path.grid(row=0, column=0, columnspan=2, sticky="ew", pady=5, ipadx=10, ipady=10)
label_file_path.drop_target_register(DND_FILES)
label_file_path.dnd_bind('<<Drop>>', on_file_drop)

# Input labels and entry fields
fields = [
    ("Start Date (MM/DD/YYYY):", "entry_start_date"),
    ("End Date (MM/DD/YYYY):", "entry_end_date"),
    ("Minimum Cost:", "entry_min_cost"),
    ("Maximum Cost:", "entry_max_cost"),
    ("Ledger Code:", "entry_ledger_code"),
    ("Sheet Name (Required):", "entry_sheet_name"),
]

for i, (label_text, var_name) in enumerate(fields):
    ttk.Label(frame_inputs, text=label_text).grid(row=i + 1, column=0, sticky="w", pady=5)
    globals()[var_name] = ttk.Entry(frame_inputs, width=40)
    globals()[var_name].grid(row=i + 1, column=1, sticky="ew", pady=5)

# Frame for buttons
frame_buttons = ttk.Frame(frame_inputs)
frame_buttons.grid(row=len(fields) + 1, column=0, columnspan=2, pady=10)

# Buttons
button_validate = ttk.Button(frame_buttons, text="Validate", command=validate_dates_and_cost, style="TButton.validate.TButton")
button_validate.grid(row=0, column=0, padx=5)

button_clear = ttk.Button(frame_buttons, text="Clear", command=clear_fields, style="TButton.clear.TButton")
button_clear.grid(row=0, column=1, padx=5)

button_upload = ttk.Button(frame_buttons, text="Upload File", command=upload_file)
button_upload.grid(row=0, column=2, padx=5)

button_quit = ttk.Button(frame_buttons, text="Quit", command=quit_program, style="TButton.quit.TButton")
button_quit.grid(row=0, column=3, padx=5)

# Treeview for output
tree_output = ttk.Treeview(frame_output, columns=("Activity", "Issue", "Details"), show="headings", height=15)
tree_output.grid(row=0, column=0, padx=10, pady=10, sticky="nsew")

# Define column headings
tree_output.heading("Activity", text="Activity")
tree_output.heading("Issue", text="Issue")
tree_output.heading("Details", text="Details")

# Configure column widths
tree_output.column("Activity", width=200, anchor="w")
tree_output.column("Issue", width=200, anchor="w")
tree_output.column("Details", width=400, anchor="w")

# Scrollbar for the Treeview
scrollbar = ttk.Scrollbar(frame_output, orient="vertical", command=tree_output.yview)
tree_output.configure(yscrollcommand=scrollbar.set)
scrollbar.grid(row=0, column=1, sticky="ns")

# Configure grid weights for resizing
window.columnconfigure(0, weight=1)
window.rowconfigure(1, weight=1)
frame_output.columnconfigure(0, weight=1)
frame_output.rowconfigure(0, weight=1)

# Ledger code drop box
entry_ledger_code.bind("<FocusIn>", show_suggestion_box)
entry_ledger_code.bind("<FocusOut>", lambda _: suggestion_box.place_forget())

# Add ledger code suggestion setup after file upload
button_upload.config(command=lambda: [upload_file(), setup_suggestion_box()])

# Key Bindings
window.bind("<Return>", lambda event: validate_dates_and_cost())
window.bind("<Escape>", lambda event: quit_program())
window.bind("<Button-1>", hide_calendar) # Handles hiding the calendar if the user clicks elsewhere

# Calendar customization for start and end dates
entry_start_date = DateEntry(
    frame_inputs,
    width=12,
    background='blue',
    foreground='white',
    borderwidth=2,
    arrowsize=15,
    font=("Arial", 12),
    headersbackground='white',
    headersforeground='black',
    selectbackground='blue',
    selectforeground='white',
    normalbackground='white',
    normalforeground='blue',
    weekendbackground='lightblue',
    weekendforeground='darkblue',
    othermonthbackground='lightgray',
    othermonthforeground='blue'
)
entry_start_date.grid(row=1, column=1, sticky="ew", pady=5)

entry_end_date = DateEntry(
    frame_inputs,
    width=12,
    background='blue',
    foreground='white',
    borderwidth=2,
    arrowsize=15,
    font=("Arial", 12),
    headersbackground='white',
    headersforeground='black',
    selectbackground='blue',
    selectforeground='white',
    normalbackground='white',
    normalforeground='blue',
    weekendbackground='lightblue',
    weekendforeground='darkblue',
    othermonthbackground='lightgray',
    othermonthforeground='blue'
)
entry_end_date.grid(row=2, column=1, sticky="ew", pady=5)

# Search functionality
ttk.Label(frame_inputs, text="Search Activity:").grid(row=8, column=0, sticky="w", pady=5)
search_entry = ttk.Entry(frame_inputs, width=40)
search_entry.grid(row=8, column=1, pady=5)

btn_search = ttk.Button(frame_inputs, text="Search", command=search_output)
btn_search.grid(row=8, column=2, sticky="e", pady=5)

# Clearing the Calendar fields so it doesnt filter by Calendar of the bat
entry_start_date.delete(0, tk.END)
entry_end_date.delete(0, tk.END)
    
# Start the GUI loop
window.mainloop()