import pandas as pd
import tkinter as tk
from tkinter import messagebox, filedialog
import json
import os
import subprocess
from tkinterdnd2 import DND_FILES, TkinterDnD

# ---------- Main Functions ----------

def load_inputs_from_file(file_path=None):
    if not file_path:
        file_path = filedialog.askopenfilename(
            title="Select Input File",
            filetypes=[("JSON files", "*.json"), ("Excel files", "*.xlsx *.xls")]
        )
        if not file_path:
            return

    try:
        if file_path.endswith('.json'):
            with open(file_path, 'r') as f:
                data = json.load(f)
        elif file_path.endswith(('.xls', '.xlsx')):
            df = pd.read_excel(file_path)
            data = pd.Series(df.Value.values, index=df.Field).to_dict()
        else:
            messagebox.showerror("Invalid File", "Only .json or .xlsx files are supported.")
            return

        for key, var in input_fields.items():
            if key in data:
                var.set(str(data[key]))

    except Exception as e:
        messagebox.showerror("Load Error", f"Error loading file:\n{str(e)}")

def generate_and_export():
    try:
        values = {k: float(v.get()) for k, v in input_fields.items()}
    except ValueError:
        messagebox.showerror("Input Error", "Please enter valid numeric values in all fields.")
        return

    # Generate data
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

    # Export
    df = pd.DataFrame(data, columns=["SPH", "CYL", "Axis"])
    file_path = filedialog.asksaveasfilename(defaultextension=".xlsx",
                                             filetypes=[("Excel files", "*.xlsx")],
                                             title="Save Excel File")
    if file_path:
        df.to_excel(file_path, index=False)
        messagebox.showinfo("Success", f"Lens grid saved as:\n{file_path}")
        try:
            os.startfile(file_path)  # Windows-specific
        except AttributeError:
            subprocess.call(['open', file_path])  # macOS
        except Exception as e:
            print("Could not open file:", e)

# ---------- GUI Setup ----------
root = TkinterDnD.Tk()
root.title("Lens Grid Generator")

frame = tk.Frame(root, padx=20, pady=20)
frame.pack()

input_fields = {
    "sph_start": tk.StringVar(),
    "sph_step": tk.StringVar(),
    "sph_max": tk.StringVar(),
    "cyl_start": tk.StringVar(),
    "cyl_step": tk.StringVar(),
    "cyl_max": tk.StringVar(),
    "axis_start": tk.StringVar(),
    "axis_step": tk.StringVar(),
    "axis_max": tk.StringVar()
}

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

def on_drop(event):
    filepath = event.data.strip().replace("{", "").replace("}", "")
    if os.path.exists(filepath):
        load_inputs_from_file(filepath)

load_btn = tk.Button(btn_frame, text="Drop JSON/Excel Here or Click", command=load_inputs_from_file, bg='orange')
load_btn.pack(side=tk.LEFT, padx=10)
load_btn.drop_target_register(DND_FILES)
load_btn.dnd_bind('<<Drop>>', on_drop)

tk.Button(btn_frame, text="Generate & Export", command=generate_and_export, bg='green', fg='white').pack(side=tk.LEFT, padx=10)

root.mainloop()
