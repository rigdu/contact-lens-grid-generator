import pandas as pd
import tkinter as tk
from tkinter import messagebox, filedialog
import json
import os
import subprocess
import platform

# ---------- Settings ----------
CONFIG_FILE = "last_inputs.json"

default_values = {
    "sph_start": "0.00",
    "sph_step": "0.25",
    "sph_max": "2.00",
    "cyl_start": "0.25",
    "cyl_step": "0.50",
    "cyl_max": "3.00",
    "axis_start": "10",
    "axis_step": "10",
    "axis_max": "180"
}

def load_last_inputs():
    if os.path.exists(CONFIG_FILE):
        try:
            with open(CONFIG_FILE, "r") as f:
                return json.load(f)
        except:
            pass
    return default_values.copy()

def save_inputs(values):
    with open(CONFIG_FILE, "w") as f:
        json.dump(values, f)

def reset_to_defaults():
    for key, var in input_fields.items():
        var.set(default_values[key])
    messagebox.showinfo("Reset", "Fields have been reset to default values.")

def generate_and_export():
    try:
        values = {k: float(v.get()) for k, v in input_fields.items()}
    except ValueError:
        messagebox.showerror("Input Error", "Please enter valid numeric values in all fields.")
        return

    save_inputs({k: str(v.get()) for k, v in input_fields.items()})

    data = []
    sph = values["sph_start"]
    while sph <= values["sph_max"] + 0.0001:
        cyl = values["cyl_start"]
        while cyl <= values["cyl_max"] + 0.0001:
            axis = values["axis_start"]
            while axis <= values["axis_max"]:
                data.append([round(sph, 2), round(cyl, 2), axis])
                axis += values["axis_step"]
            cyl += values["cyl_step"]
        sph += values["sph_step"]

    df = pd.DataFrame(data, columns=["SPH", "CYL", "Axis"])
    file_path = filedialog.asksaveasfilename(defaultextension=".xlsx",
                                             filetypes=[("Excel files", "*.xlsx")],
                                             title="Save Excel File")
    if file_path:
        df.to_excel(file_path, index=False)
        messagebox.showinfo("Success", f"Lens grid saved as:\n{file_path}")
        try:
            if platform.system() == "Windows":
                os.startfile(file_path)
            elif platform.system() == "Darwin":
                subprocess.call(['open', file_path])
            else:
                subprocess.call(['xdg-open', file_path])
        except Exception as e:
            print("Could not open file:", e)

# ---------- GUI Setup ----------
root = tk.Tk()
root.title("Lens Grid Generator")

frame = tk.Frame(root, padx=20, pady=20)
frame.pack()

input_fields = {k: tk.StringVar() for k in default_values}
last_values = load_last_inputs()
for key in input_fields:
    input_fields[key].set(last_values.get(key, default_values[key]))

def create_input(label, var):
    tk.Label(frame, text=label, anchor='w', width=25).pack()
    tk.Entry(frame, textvariable=var).pack()

tk.Label(frame, text="SPH Values", font=('Arial', 10, 'bold')).pack(pady=(10, 0))
create_input("SPH Start", input_fields["sph_start"])
create_input("SPH Step", input_fields["sph_step"])
create_input("SPH Max", input_fields["sph_max"])

tk.Label(frame, text="CYL Values", font=('Arial', 10, 'bold')).pack(pady=(10, 0))
create_input("CYL Start", input_fields["cyl_start"])
create_input("CYL Step", input_fields["cyl_step"])
create_input("CYL Max", input_fields["cyl_max"])

tk.Label(frame, text="Axis Values", font=('Arial', 10, 'bold')).pack(pady=(10, 0))
create_input("Axis Start", input_fields["axis_start"])
create_input("Axis Step", input_fields["axis_step"])
create_input("Axis Max", input_fields["axis_max"])

# Buttons
btn_frame = tk.Frame(frame)
btn_frame.pack(pady=20)

tk.Button(btn_frame, text="Generate & Export", command=generate_and_export, bg='green', fg='white').pack(side=tk.LEFT, padx=10)
tk.Button(btn_frame, text="Reset to Default", command=reset_to_defaults, bg='red', fg='white').pack(side=tk.LEFT, padx=10)

root.mainloop()
