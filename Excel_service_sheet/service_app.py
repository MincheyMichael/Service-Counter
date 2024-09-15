# import panda to turn Excel sheet into df
import pandas as pd
# import tkinter for GUI and Treeview
import tkinter as tk
from tkinter import ttk
# custom Tkinter
import customtkinter

file = "/Users/michaelminchey/Downloads/Service_sheet(Sheet1).csv"
reset_df = pd.read_csv(file)
reset_df["Tech Number"] = reset_df["Tech Number"].astype(str)
df = pd.read_csv("service_data.csv")
df["Tech Number"] = df["Tech Number"].astype(str)

window = tk.Tk()
window.geometry("1350x650")
window.title("Renee's Service Tracker")


# -------------------------------------------------------------------------------

def display_data():
    global treeview

    if treeview:
        treeview.destroy()

    frame = ttk.Frame(master=window)
    frame.grid(row=0, column=0, sticky="nsew")

    treeview = ttk.Treeview(frame, columns=list(df.columns), show="headings")

    for col in df.columns:  # column names for display
        treeview.heading(col, text=col)
        treeview.column(col, width=150)

    for index, row in df.iterrows():  # values for each row
        treeview.insert("", "end", values=list(row))

    treeview.grid(row=0, column=0, padx=3, pady=3, sticky="nsew")

    style = ttk.Style()

    style.configure("Treeview",
                    background="black",
                    forground="white",
                    feildbackground="silver",
                    font=("Arial", 17,))

    style.configure("Treeview.Heading", font=("Arial", 13))

    frame.grid_rowconfigure(0, weight=1)
    frame.grid_columnconfigure(0, weight=0)

    window.grid_rowconfigure(0, weight=1)
    window.grid_columnconfigure(0, weight=1)


def refresh_treeview():
    global treeview
    display_data()


def save_df():
    csv_file_path = "service_data.csv"
    df.to_csv(csv_file_path, index=False)


def reset():
    global df
    df = reset_df
    save_df()
    display_data()


def display_tech(tech_info):
    global df
    tree = ttk.Treeview(window, columns=list(tech_info.columns), show="headings")

    for col in tech_info.columns:
        tree.heading(col, text=col)
        tree.column(col, width=100)

    for index, row in tech_info.iterrows():  # values for each row
        tree.insert("", "end", values=list(row))

    tree.grid(row=6, column=0, padx=3, pady=3, sticky="nsew")


# -----------------------------------------control frame---------------------------------------------------
def control_frame():
    controls_frame = tk.Frame(window)
    controls_frame.grid(row=0, column=1, padx=5, pady=5, sticky="new")

    label1 = tk.Label(controls_frame, text="Service A", width=10)
    label1.grid(row=1, column=0, padx=0, pady=5)
    label2 = tk.Label(controls_frame, text="Service B", width=10)
    label2.grid(row=2, column=0, padx=0, pady=5)
    label3 = tk.Label(controls_frame, text="Customer Pay", width=10)
    label3.grid(row=3, column=0, padx=0, pady=5)
    label4 = tk.Label(controls_frame, text="Pre-Paid", width=10)
    label4.grid(row=4, column=0, padx=0, pady=5)
    label5 = tk.Label(controls_frame, text="Warranty", width=10)
    label5.grid(row=5, column=0, padx=0, pady=5)

    # -------------------ENTRY------------------------------------

    def entry_button():
        global treeview
        tech_number = e1.get()
        tech_info = df.loc[df["Tech Number"] == tech_number]
        display_tech(tech_info)
        if tech_number == "reset":
            reset()

    b1 = tk.Button(controls_frame, text="Enter", command=entry_button, width=5)
    b1.grid(row=0, column=1, columnspan=2, padx=0, pady=5)
    e1 = tk.Entry(controls_frame, width=7)
    e1.grid(row=0, column=0, columnspan=2, padx=0, pady=5)

    # ------------------ Add service buttons -------------------

    def add_service_a():
        tech = e1.get()
        tech_info = df.index[df["Tech Number"] == tech][0]
        df.at[tech_info, "Service A"] += 1
        display_data()
        save_df()

    def add_service_b():
        tech = e1.get()
        tech_info = df.index[df["Tech Number"] == tech][0]
        df.at[tech_info, "Service B"] += 1
        display_data()
        save_df()

    def add_service_cp():
        tech = e1.get()
        tech_info = df.index[df["Tech Number"] == tech][0]
        df.at[tech_info, "Customer Pay"] += 1
        display_data()
        save_df()

    def add_service_ppm():
        tech = e1.get()
        tech_info = df.index[df["Tech Number"] == tech][0]
        df.at[tech_info, "Pre-paid"] += 1
        display_data()
        save_df()

    def add_service_warranty():
        tech = e1.get()
        tech_info = df.index[df["Tech Number"] == tech][0]
        df.at[tech_info, "Warranty"] += 1
        display_data()
        save_df()

    as1 = tk.Button(controls_frame, text="Add", command=add_service_a, width=5)
    as1.grid(row=1, column=1, padx=0, pady=5)
    bs1 = tk.Button(controls_frame, text="Add", command=add_service_b, width=5)
    bs1.grid(row=2, column=1, padx=0, pady=5)
    cp1 = tk.Button(controls_frame, text="Add", command=add_service_cp, width=5)
    cp1.grid(row=3, column=1, padx=0, pady=5)
    pp1 = tk.Button(controls_frame, text="Add", command=add_service_ppm, width=5)
    pp1.grid(row=4, column=1, padx=0, pady=5)
    w1 = tk.Button(controls_frame, text="Add", command=add_service_warranty, width=5)
    w1.grid(row=5, column=1, padx=0, pady=5)

    # -------------------- Subtract buttons -----------------------

    def sub_service_a():
        tech = e1.get()
        tech_info = df.index[df["Tech Number"] == tech][0]
        df.at[tech_info, "Service A"] -= 1
        display_data()
        save_df()

    def sub_service_b():
        tech = e1.get()
        tech_info = df.index[df["Tech Number"] == tech][0]
        df.at[tech_info, "Service B"] -= 1
        display_data()
        save_df()

    def sub_service_cp():
        tech = e1.get()
        tech_info = df.index[df["Tech Number"] == tech][0]
        df.at[tech_info, "Customer Pay"] -= 1
        display_data()
        save_df()

    def sub_service_ppm():
        tech = e1.get()
        tech_info = df.index[df["Tech Number"] == tech][0]
        df.at[tech_info, "Pre-paid"] -= 1
        display_data()
        save_df()

    def sub_service_warranty():
        tech = e1.get()
        tech_info = df.index[df["Tech Number"] == tech][0]
        df.at[tech_info, "Warranty"] -= 1
        display_data()
        save_df()

    as2 = tk.Button(controls_frame, text="Subtract", command=sub_service_a, width=5)
    as2.grid(row=1, column=2, padx=0, pady=5)
    bs2 = tk.Button(controls_frame, text="Subtract", command=sub_service_b, width=5)
    bs2.grid(row=2, column=2, padx=0, pady=5)
    cp2 = tk.Button(controls_frame, text="Subtract", command=sub_service_cp, width=5)
    cp2.grid(row=3, column=2, padx=0, pady=5)
    pp2 = tk.Button(controls_frame, text="Subtract", command=sub_service_ppm, width=5)
    pp2.grid(row=4, column=2, padx=0, pady=5)
    w2 = tk.Button(controls_frame, text="Subtract", command=sub_service_warranty, width=5)
    w2.grid(row=5, column=2, padx=0, pady=5)

    controls_frame.grid_rowconfigure(0, weight=1)
    controls_frame.grid_columnconfigure(0, weight=0)

    window.grid_rowconfigure(0, weight=1)
    window.grid_columnconfigure(0, weight=1)


def app():
    display_data()
    control_frame()


treeview = None
app()

window.mainloop()
