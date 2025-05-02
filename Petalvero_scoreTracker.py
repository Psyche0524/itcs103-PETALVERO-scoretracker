import customtkinter as ctk
from tkinter import messagebox
from openpyxl import Workbook, load_workbook
import os

# === Excel Setup ===
file_name = "student_scores.xlsx"

def create_excel_file():
    if not os.path.exists(file_name):
        wb = Workbook()
        ws = wb.active
        ws.title = "Scores"
        ws.append(["Name", "Score", "Status"])  # Changed Remarks to Status
        wb.save(file_name)

def add_student_record(name, score):
    wb = load_workbook(file_name)
    ws = wb.active
    status = "Pass" if score >= 75 else "Fail"
    ws.append([name, score, status])  # Changed remark to status
    wb.save(file_name)

def get_all_records():
    wb = load_workbook(file_name)
    ws = wb.active
    return [row for row in ws.iter_rows(min_row=2, values_only=True)]

# === GUI Setup ===
ctk.set_appearance_mode("dark")
ctk.set_default_color_theme("blue")

app = ctk.CTk()
app.title("Student Score Tracker")
app.geometry("500x400")

create_excel_file()

title_label = ctk.CTkLabel(app, text="Enter Student Details", font=("Arial", 20))
title_label.pack(pady=10)

name_entry = ctk.CTkEntry(app, placeholder_text="Student Name")
name_entry.pack(pady=10)

score_entry = ctk.CTkEntry(app, placeholder_text="Score")
score_entry.pack(pady=10)

def save_score():
    name = name_entry.get().strip()
    try:
        score = int(score_entry.get())
        if not name:
            messagebox.showwarning("Missing Info", "Please enter a student name.")
            return
        add_student_record(name, score)
        messagebox.showinfo("Success", f"{name}'s score has been saved!")
        name_entry.delete(0, 'end')
        score_entry.delete(0, 'end')
    except ValueError:
        messagebox.showerror("Invalid Input", "Score must be a number.")

def view_records():
    records = get_all_records()
    if not records:
        messagebox.showinfo("No Records", "There are no records yet.")
        return

    # Create styled records window
    records_window = ctk.CTkToplevel(app)
    records_window.title("All Student Records")
    records_window.geometry("450x350")

    title = ctk.CTkLabel(records_window, text="Student Records", font=("Arial", 18))
    title.pack(pady=10)

    textbox = ctk.CTkTextbox(records_window, width=400, height=250, corner_radius=10, font=("Courier", 12))
    textbox.pack(pady=10)

    # Add formatted records to textbox
    textbox.insert("end", f"{'Name':<20}{'Score':<10}{'Status':<10}\n")
    textbox.insert("end", "-" * 40 + "\n")
    for r in records:
        textbox.insert("end", f"{r[0]:<20}{r[1]:<10}{r[2]:<10}\n")

    textbox.configure(state="disabled")  # Make read-only

save_btn = ctk.CTkButton(app, text="Save Score", command=save_score)
save_btn.pack(pady=10)

view_btn = ctk.CTkButton(app, text="View All Records", command=view_records)
view_btn.pack(pady=5)

app.mainloop()
