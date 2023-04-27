import pandas as pd
import tkinter as tk
from tkinter import filedialog
from tkinter.filedialog import askopenfile
import numpy as np
import os

root = tk.Tk()

canvas = tk.Canvas(root,width=600,height = 300)
canvas.grid(columnspan=6,rowspan=6)


#Read Excel Files
def read_excel_file():
        file_path = filedialog.askopenfilename(initialdir = "/", title = "Select file for Invoice Register", filetypes = (("Excel files", "*.xlsx"), ("all files", "*.*")))
        global df
        df = pd.read_excel(file_path, skiprows=18)
        ir_file_label.config(text=os.path.basename(file_path))
        df["Payment Date\n"] = pd.to_datetime(df["Payment Date\n"], errors='coerce')
        return df
def read_excel_file1():
        file_path = filedialog.askopenfilename(initialdir = "/", title = "Select file for Invoice Register of PY", filetypes = (("Excel files", "*.xlsx"), ("all files", "*.*")))
        global df2
        df2 = pd.read_excel(file_path, skiprows=18)
        ir1_file_label.config(text=os.path.basename(file_path))
        df2["Payment Date\n"] = pd.to_datetime(df2["Payment Date\n"], errors='coerce')
        return df2

def read_excel_file_SR():
        file_path1 = filedialog.askopenfilename(initialdir = "/", title = "Select file for Supplier Report", filetypes = (("Excel files", "*.xlsx"), ("all files", "*.*")))
        global df1
        df1 = pd.read_excel(file_path1, skiprows=18)
        sr_file_label.config(text=os.path.basename(file_path1))
        df1['Full Address'] = df1[['Address Line1', 'Address Line2', 'Address Line3', 'Address Line4']].apply(lambda x: ' '.join(x.dropna().astype(str)), axis=1)
        df1['Legal Entity'] = df1['Full Legal Entity Name'].str[:4]
    
        df1['VAT number'] = np.where(df1['Tax Registration Number'].isna(), df1['Tax Payer ID'], df1['Tax Registration Number'])
        
        return df1

# function to filter the merged_df data based on user inputs
def filter_data(df, df1,df2, legal_entity, exclude_country,year):
    le = "0" + str(legal_entity)
    df3 = []
    
    for index, row in df2.iterrows():
        if row["Payment Date\n"].year == int(year)and row["Legal g Entity\n"] == int(legal_entity)  and row["Line Type\n"] == "Item" and row["Pay Group\n"] != "Employee":
           df3.append(row)
    for index, row in df.iterrows():
        if row["Payment Date\n"].year ==int(year) and row["Legal g Entity\n"] == int(legal_entity)  and row["Line Type\n"] == "Item" and row["Pay Group\n"] != "Employee":
            df3.append(row)
    df3 = pd.DataFrame(df3)
    df3["Full Description"] = df3["Account Description\n"].astype(str) + '-' + df3["Invoice Distribution Description\n"].astype(str)
    fgrouped = df3.groupby(["Invoice Number\n", "Full Description"], as_index=False).agg({"Invoice Distribution Amount\n": "sum"}).drop_duplicates()
    grouped = fgrouped.merge(df3,left_on ="Invoice Number\n",right_on="Invoice Number\n",how = 'left')

    merged_df = grouped.merge(df1,  left_on='Supplier Name\n', right_on='Supplier Name', how='left')
    result = []
    for index,row in merged_df.iterrows():
        if row["Legal Entity"]==le and row["Country"] != exclude_country:
            result.append(row)
    result = pd.DataFrame(result)
    
    filtered_df = result[["Supplier Name\n", "Full Address","VAT number", "Invoice Number\n",
                          "Full Description_x", "Invoice Distribution Amount\n_x",
                           "Currency\n", "Payment Status\n", "Payment Date\n","Line Type\n","Legal g Entity\n"]]
    filtered_df = filtered_df.drop_duplicates(subset=["Invoice Number\n", "Full Description_x"])
    
    return filtered_df 



#buttons
    # IR file selection button and command
ir_button = tk.Button(root, text="Choose file for Invoice Register", command=read_excel_file)
ir_button.grid(row=0, column=0, padx=10, pady=10)

    # SR file selection button
sr_button = tk.Button(root, text="Choose file for Supplier Report", command=read_excel_file_SR)
sr_button.grid(row=1, column=0, padx=10, pady=10)
    # IRPY file selection button
ir1_button = tk.Button(root, text="Choose file for previous year's Invoice Register", command=read_excel_file1)
ir1_button.grid(row=0, column=2, padx=10, pady=10)
    
    # IR file label to show the selected file
ir_file_label = tk.Label(root, text="")
ir_file_label.grid(row=0, column=1, padx=10, pady=10)

    # SR file label to show the selected file
sr_file_label = tk.Label(root, text="")
sr_file_label.grid(row=1, column=1, padx=10, pady=10)   
    # IRPY file label to show the selected file
ir1_file_label = tk.Label(root, text="")
ir1_file_label.grid(row=0, column=4, padx=10, pady=10)   



   # Year input label
year_label = tk.Label(root, text="Year:")
year_label.grid(row=2, column=0, padx=10, pady=10)

    # Year input entry
year_entry = tk.Entry(root)
year_entry.grid(row=2, column=1, padx=10, pady=10)

    # Country to exclude input label
country_label = tk.Label(root, text="Country to exclude:")
country_label.grid(row=3, column=0, padx=10, pady=10)

    # Country to exclude input entry
country_entry = tk.Entry(root)
country_entry.grid(row=3, column=1, padx=10, pady=10)

    # Legal entity input label
legal_label = tk.Label(root, text="Legal Entity (with 0 in front):")
legal_label.grid(row=4, column=0, padx=10, pady=10)
    
    # Legal entity input entry
legal_entry = tk.Entry(root)
legal_entry.grid(row=4, column=1, padx=10, pady=10)    
    
def write_to_excel(filtered_df):
    file_path = filedialog.asksaveasfilename(defaultextension=".xlsx", initialdir = "/", title = "Save file", filetypes = (("Excel files", "*.xlsx"), ("all files", "*.*")))
    filtered_df.to_excel(file_path, index=False)


#write_to_excel(fd)
def retrieve_inputs():
        year = year_entry.get()
        country_exclude = country_entry.get()
        legal_entity = legal_entry.get()

        print("Year:", year)
        print("Country to Exclude:", country_exclude)
        print("Legal Entity:", legal_entity)
        
        return year, country_exclude, legal_entity
def call_filter_data():
    year, country_exclude, legal_entity = retrieve_inputs()
    if "df" not in globals():
        print("Please choose an IR file first")
        return
    if "df1" not in globals():
        print("Please choose a SR file first")
        return
   
    fd = filter_data(df,df1,df2, legal_entity,country_exclude,year) 
    write_to_excel(fd) 

#execution button
execute_button = tk.Button(root, text="Execute", command=retrieve_inputs)
execute_button.grid(row=5, column=0, padx=10, pady=10)

#execution button
filter_button = tk.Button(root, text="filter", command=call_filter_data)
filter_button.grid(row=6, column=0, padx=10, pady=10)
root.mainloop()
