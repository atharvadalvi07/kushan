import tkinter as tk
from tkinter import ttk
import pandas as pd

# Load the Excel file
excel_file = 'NexTier - Frac Dataset.xlsx'
df = pd.read_excel(excel_file)

# Fleet and region options
fleet_names = ["Frac 08", "Frac 10", "Frac 18", "Frac 21", "Frac 32", "Frac 38", "Frac 39", "Frac 40", "Frac 90"]
regions = ["Northeast", "West", "Central", "West Texas"]

def calculate_efficiencies():
    fleet = fleet_var.get()
    region = region_var.get()
    
    filtered_df = df[(df['Fleet Name'] == fleet) & (df['Region Name'] == region)]
    
    Total_efficiency = filtered_df['Pumping Hours'].sum() / filtered_df['Time Available (hrs)'].sum()
    NPT_Percentage = filtered_df['NPT (All)'].sum() / filtered_df['Duration (hrs)'].sum()
    Pumping_efficiency = filtered_df['Pumping Hours'].sum() / filtered_df['Duration (hrs)'].sum()
    UnavoidableNPT = filtered_df['NPT (Customer + Weather)'].sum()
    AvoidableNPT = filtered_df['NPT (All)'].sum() - UnavoidableNPT
    
    result_text.set(f"Total efficiency: {Total_efficiency:.2f}\n"
                    f"NPT Percentage: {NPT_Percentage:.2f}\n"
                    f"Pumping efficiency: {Pumping_efficiency:.2f}\n"
                    f"Unavoidable NPT: {UnavoidableNPT:.2f} hrs\n"
                    f"Avoidable NPT: {AvoidableNPT:.2f} hrs")

# Create the main window
root = tk.Tk()
root.title("Fleet Efficiency Calculator")

# Create dropdown menus
fleet_var = tk.StringVar()
region_var = tk.StringVar()
result_text = tk.StringVar()

ttk.Label(root, text="Select Fleet:").grid(column=0, row=0, padx=10, pady=10)
fleet_menu = ttk.Combobox(root, textvariable=fleet_var, values=fleet_names)
fleet_menu.grid(column=1, row=0, padx=10, pady=10)

ttk.Label(root, text="Select Region:").grid(column=0, row=1, padx=10, pady=10)
region_menu = ttk.Combobox(root, textvariable=region_var, values=regions)
region_menu.grid(column=1, row=1, padx=10, pady=10)

# Create a button to calculate efficiencies
calculate_button = ttk.Button(root, text="Calculate", command=calculate_efficiencies)
calculate_button.grid(column=0, row=2, columnspan=2, padx=10, pady=10)

# Create a label to display the results
result_label = ttk.Label(root, textvariable=result_text, justify="left")
result_label.grid(column=0, row=3, columnspan=2, padx=10, pady=10)

# Run the application
root.mainloop()
