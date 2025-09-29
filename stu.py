import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import pandas as pd
from datetime import datetime
import os
from tkcalendar import Calendar

# --- File & Columns ---
ATTENDANCE_FILE = "attendance.csv"
COLUMNS = ["Date", "Roll Number", "Name", "Subject", "Class", "Section", "Class Type", "Status"]

# --- Load CSV ---
def load_csv():
    if not os.path.exists(ATTENDANCE_FILE):
        df = pd.DataFrame(columns=COLUMNS)
        df.to_csv(ATTENDANCE_FILE, index=False)
        return df
    try:
        df = pd.read_csv(ATTENDANCE_FILE)
        df = df[[col for col in df.columns if col in COLUMNS]]
        for col in COLUMNS:
            if col not in df.columns:
                df[col] = ""
        df = df[COLUMNS]
    except pd.errors.ParserError:
        df = pd.DataFrame(columns=COLUMNS)
    df.to_csv(ATTENDANCE_FILE, index=False)
    return df

# --- Mark Attendance ---
def mark_attendance():
    roll = roll_var.get()
    name = name_var.get()
    subject = subject_var.get()
    class_type = class_type_var.get()
    class_name = class_var.get()
    section = section_var.get()
    status = status_var.get()
    date = calendar.selection_get().strftime("%Y-%m-%d")

    if not (roll and name and subject and class_type and class_name and section and status):
        messagebox.showwarning("Input Error", "Please fill all fields")
        return

    df_existing = load_csv()
    new_data = pd.DataFrame([[date, roll, name, subject, class_name, section, class_type, status]], columns=COLUMNS)
    df_updated = pd.concat([df_existing, new_data], ignore_index=True)
    df_updated.to_csv(ATTENDANCE_FILE, index=False)
    messagebox.showinfo("Success", f"Attendance marked for {name} ({roll}) [{class_type}] in {subject}")
    clear_fields()

# --- Generate Report Window ---
def generate_report_window():
    year = year_var.get()
    month = month_var.get()
    class_name = report_class_var.get()
    section_name = report_section_var.get()
    class_type = report_class_type_var.get()

    df = load_csv()
    if df.empty:
        messagebox.showwarning("No Data", "No attendance records found.")
        return

    df['Date'] = pd.to_datetime(df['Date'], errors='coerce')
    df = df.dropna(subset=['Date'])

    # Filter
    if year:
        df = df[df['Date'].dt.year == int(year)]
    if month:
        df = df[df['Date'].dt.month == int(month)]
    if class_name != "All":
        df = df[df['Class'] == class_name]
    if section_name != "All":
        df = df[df['Section'] == section_name]
    if class_type != "All":
        df = df[df['Class Type'] == class_type]

    if df.empty:
        messagebox.showwarning("No Data", "No records found for selected filters")
        return

    # Attendance summary
    summary = df.groupby(['Roll Number','Name','Subject']).agg(
        Total_Lectures=('Status','count'),
        Present=('Status', lambda x: (x=='Present').sum())
    ).reset_index()
    summary['Attendance %'] = round((summary['Present']/summary['Total_Lectures'])*100,2)

    # Show in a new window
    report_window = tk.Toplevel(root)
    report_window.title("Attendance Report")
    report_window.geometry("900x500")

    ttk.Label(report_window, text=f"Session Report: Year={year or 'All'}, Month={month or 'All'}").pack(pady=5)

    tree = ttk.Treeview(report_window, columns=['Roll','Name','Subject','Total Lectures','Present','Attendance %'], show='headings')
    for col in ['Roll','Name','Subject','Total Lectures','Present','Attendance %']:
        tree.heading(col, text=col)
        tree.column(col, width=130)
    for idx, row in summary.iterrows():
        tree.insert('', tk.END, values=[row['Roll Number'], row['Name'], row['Subject'], row['Total_Lectures'], row['Present'], row['Attendance %']])
    tree.pack(expand=True, fill='both')

    # Download button
    def download_report():
        file_path = filedialog.asksaveasfilename(defaultextension=".xlsx",
                                                 filetypes=[("Excel files","*.xlsx")],
                                                 initialfile=f"Attendance_Report.xlsx")
        if file_path:
            with pd.ExcelWriter(file_path, engine='openpyxl') as writer:
                df.to_excel(writer, index=False, sheet_name='Raw Attendance')
                summary.to_excel(writer, index=False, sheet_name='Attendance Summary')
            messagebox.showinfo("Saved", f"Report saved successfully at:\n{file_path}")

    ttk.Button(report_window, text="Download Report", command=download_report).pack(pady=10)

# --- Search Attendance ---
def search_attendance():
    search_by_val = search_var.get()
    query = search_entry.get().strip()
    if not query:
        messagebox.showwarning("Input Error", "Please enter a value to search")
        return

    df = load_csv()
    df['Roll Number'] = df['Roll Number'].astype(str)
    df['Subject'] = df['Subject'].astype(str)

    if search_by_val == "Roll Number":
        matches = df[df['Roll Number'].str.lower() == query.lower()]
    elif search_by_val == "Subject":
        matches = df[df['Subject'].str.lower() == query.lower()]
    else:
        messagebox.showerror("Error", "Invalid search type")
        return

    if matches.empty:
        messagebox.showinfo("No Records", "No matching records found")
        return

    result_window = tk.Toplevel(root)
    result_window.title("Search Results")
    result_window.geometry("800x400")

    tree = ttk.Treeview(result_window, columns=COLUMNS, show="headings")
    for col in COLUMNS:
        tree.heading(col, text=col)
        tree.column(col, width=100)
    for idx, row in matches.iterrows():
        color = "green" if row['Status']=="Present" else "red"
        tree.insert("", tk.END, values=list(row), tags=("status",))
    tree.tag_configure("status", background=color)
    tree.pack(expand=True, fill="both")

# --- View Old Records ---
def view_old_records():
    df = load_csv()
    if df.empty:
        messagebox.showinfo("No Records", "Attendance file is empty.")
        return

    old_window = tk.Toplevel(root)
    old_window.title("Old Attendance Records")
    old_window.geometry("850x400")

    tree = ttk.Treeview(old_window, columns=COLUMNS, show="headings")
    for col in COLUMNS:
        tree.heading(col, text=col)
        tree.column(col, width=100)
    for idx, row in df.iterrows():
        tree.insert("", tk.END, values=list(row))
    tree.pack(expand=True, fill="both")

# --- Clear Fields ---
def clear_fields():
    roll_var.set("")
    name_var.set("")
    subject_var.set("")
    status_var.set("Present")
    class_type_var.set("Theory")
    class_var.set("")
    section_var.set("")

# --- New Session ---
def new_session():
    if messagebox.askyesno("New Session", "This will delete all previous records. Continue?"):
        df = pd.DataFrame(columns=COLUMNS)
        df.to_csv(ATTENDANCE_FILE, index=False)
        messagebox.showinfo("New Session", "Fresh attendance session started.")

# --- GUI ---
root = tk.Tk()
root.title("Student Attendance System")
root.geometry("900x700")

# --- Variables ---
roll_var = tk.StringVar()
name_var = tk.StringVar()
subject_var = tk.StringVar()
class_type_var = tk.StringVar(value="Theory")
status_var = tk.StringVar(value="Present")
class_var = tk.StringVar()
section_var = tk.StringVar()

year_var = tk.StringVar()
month_var = tk.StringVar()
report_class_var = tk.StringVar(value="All")
report_section_var = tk.StringVar(value="All")
report_class_type_var = tk.StringVar(value="All")

search_var = tk.StringVar(value="Roll Number")
search_entry = tk.StringVar()

# --- Top Buttons ---
ttk.Button(root, text="Start New Session (Clear Old Records)", command=new_session).pack(pady=5)
ttk.Button(root, text="View Old Records", command=view_old_records).pack(pady=5)

# --- Calendar ---
calendar_frame = ttk.LabelFrame(root, text="Select Date")
calendar_frame.pack(fill="x", padx=10, pady=5)
calendar = Calendar(calendar_frame, selectmode='day', date_pattern='yyyy-mm-dd')
calendar.pack(padx=5, pady=5)

# --- Mark Attendance Frame ---
frame1 = ttk.LabelFrame(root, text="Mark Attendance")
frame1.pack(fill="x", padx=10, pady=5)

ttk.Label(frame1, text="Roll Number:").grid(row=0, column=0, padx=5, pady=5)
ttk.Entry(frame1, textvariable=roll_var).grid(row=0, column=1, padx=5, pady=5)

ttk.Label(frame1, text="Name:").grid(row=0, column=2, padx=5, pady=5)
ttk.Entry(frame1, textvariable=name_var).grid(row=0, column=3, padx=5, pady=5)

ttk.Label(frame1, text="Subject:").grid(row=1, column=0, padx=5, pady=5)
ttk.Entry(frame1, textvariable=subject_var).grid(row=1, column=1, padx=5, pady=5)

ttk.Label(frame1, text="Class Type:").grid(row=1, column=2, padx=5, pady=5)
ttk.Combobox(frame1, textvariable=class_type_var, values=["Theory","Practical"]).grid(row=1, column=3, padx=5, pady=5)

ttk.Label(frame1, text="Class:").grid(row=2, column=0, padx=5, pady=5)
ttk.Combobox(frame1, textvariable=class_var, values=["CSE","ECE","MECH","IT"]).grid(row=2, column=1, padx=5, pady=5)

ttk.Label(frame1, text="Section:").grid(row=2, column=2, padx=5, pady=5)
ttk.Combobox(frame1, textvariable=section_var, values=["A","B","C"]).grid(row=2, column=3, padx=5, pady=5)

ttk.Label(frame1, text="Status:").grid(row=3, column=0, padx=5, pady=5)
ttk.Combobox(frame1, textvariable=status_var, values=["Present","Absent"]).grid(row=3, column=1, padx=5, pady=5)

ttk.Button(frame1, text="Mark Attendance", command=mark_attendance).grid(row=3, column=2, padx=5, pady=5)
ttk.Button(frame1, text="Clear Fields", command=clear_fields).grid(row=3, column=3, padx=5, pady=5)

# --- Generate Report Frame ---
frame2 = ttk.LabelFrame(root, text="Generate Report")
frame2.pack(fill="x", padx=10, pady=5)

ttk.Label(frame2, text="Year (YYYY):").grid(row=0, column=0, padx=5, pady=5)
ttk.Entry(frame2, textvariable=year_var).grid(row=0, column=1, padx=5, pady=5)

ttk.Label(frame2, text="Month (MM, optional):").grid(row=0, column=2, padx=5, pady=5)
ttk.Entry(frame2, textvariable=month_var).grid(row=0, column=3, padx=5, pady=5)

ttk.Label(frame2, text="Class:").grid(row=1, column=0, padx=5, pady=5)
ttk.Combobox(frame2, textvariable=report_class_var, values=["All","CSE","ECE","MECH","IT"]).grid(row=1, column=1, padx=5, pady=5)

ttk.Label(frame2, text="Section:").grid(row=1, column=2, padx=5, pady=5)
ttk.Combobox(frame2, textvariable=report_section_var, values=["All","A","B","C"]).grid(row=1, column=3, padx=5, pady=5)

ttk.Label(frame2, text="Class Type:").grid(row=2, column=0, padx=5, pady=5)
ttk.Combobox(frame2, textvariable=report_class_type_var, values=["All","Theory","Practical"]).grid(row=2, column=1, padx=5, pady=5)

ttk.Button(frame2, text="Generate Report", command=generate_report_window).grid(row=2, column=2, padx=5, pady=5)

# --- Search Frame ---
frame3 = ttk.LabelFrame(root, text="Search Attendance")
frame3.pack(fill="x", padx=10, pady=5)

ttk.Label(frame3, text="Search By:").grid(row=0, column=0, padx=5, pady=5)
ttk.Combobox(frame3, textvariable=search_var, values=["Roll Number","Subject"]).grid(row=0, column=1, padx=5, pady=5)

search_entry_box = ttk.Entry(frame3, textvariable=search_entry)
search_entry_box.grid(row=0, column=2, padx=5, pady=5)

ttk.Button(frame3, text="Search", command=search_attendance).grid(row=0, column=3, padx=5, pady=5)

# --- Run App ---
root.mainloop()
