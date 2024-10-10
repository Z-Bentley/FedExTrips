import tkinter as tk
from tkinter import ttk

def submit_data():
    print("Name:", name_var.get())
    print("Age:", age_var.get())
    print("Gender:", gender_var.get())

root = tk.Tk()
root.title("Editable Form")
root.geometry("500x400")

# Pre-filled data
name_var = tk.StringVar(value="John Doe")
age_var = tk.StringVar(value="30")
gender_var = tk.StringVar(value="Male")

# Input fields with pre-filled data
ttk.Label(root, text="Name:").grid(row=0, column=0)
ttk.Entry(root, textvariable=name_var).grid(row=0, column=1)

ttk.Label(root, text="Age:").grid(row=1, column=0)
ttk.Entry(root, textvariable=age_var).grid(row=1, column=1)

ttk.Label(root, text="Gender:").grid(row=2, column=0)
ttk.Combobox(root, textvariable=gender_var, values=["Male", "Female", "Other"]).grid(row=2, column=1)

# Submit button
ttk.Button(root, text="Submit", command=submit_data).grid(row=3, column=1)

root.mainloop()
