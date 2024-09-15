import tkinter as tk
from tkinter import ttk
import pandas as pd
import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg

excel_file = 'NexTier - Frac Dataset.xlsx'
df = pd.read_excel(excel_file)
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
    
    labels = ['Pumping Hours', 'Non-Pumping Hours']
    sizes = [filtered_df['Pumping Hours'].sum(), filtered_df['Time Available (hrs)'].sum() - filtered_df['Pumping Hours'].sum()]
    colors = ['#ff9999','#66b3ff']
    explode = (0.1, 0)  
    
    fig, ax = plt.subplots()
    fig.set_size_inches(4, 4)  # Adjust the size here
    ax.pie(sizes, explode=explode, labels=labels, colors=colors, autopct='%1.1f%%',
           shadow=True, startangle=90)
    ax.axis('equal') 
    
    canvas = FigureCanvasTkAgg(fig, master=chart_frame)
    canvas.draw()
    canvas.get_tk_widget().pack(side=tk.TOP, fill=tk.BOTH, expand=True)
    
    labels_npt = ['Unavoidable NPT', 'Avoidable NPT']
    sizes_npt = [UnavoidableNPT, AvoidableNPT]
    colors_npt = ['#ffcc99','#99ff99']
    explode_npt = (0.1, 0) 
    
    fig_npt, ax_npt = plt.subplots()
    fig_npt.set_size_inches(4, 4)  # Adjust the size here
    ax_npt.pie(sizes_npt, explode=explode_npt, labels=labels_npt, colors=colors_npt, autopct='%1.1f%%',
               shadow=True, startangle=90)
    ax_npt.axis('equal') 
    
    canvas_npt = FigureCanvasTkAgg(fig_npt, master=chart_frame)
    canvas_npt.draw()
    canvas_npt.get_tk_widget().pack(side=tk.TOP, fill=tk.BOTH, expand=True)

def reset():
    fleet_var.set('')
    region_var.set('')
    result_text.set('')
    for widget in chart_frame.winfo_children():
        widget.destroy()

root = tk.Tk()
root.title("Fleet Efficiency Calculator")

fleet_var = tk.StringVar()
region_var = tk.StringVar()
result_text = tk.StringVar()

ttk.Label(root, text="Select Fleet:").grid(column=0, row=0, padx=10, pady=10)
fleet_menu = ttk.Combobox(root, textvariable=fleet_var, values=fleet_names)
fleet_menu.grid(column=1, row=0, padx=10, pady=10)

ttk.Label(root, text="Select Region:").grid(column=0, row=1, padx=10, pady=10)
region_menu = ttk.Combobox(root, textvariable=region_var, values=regions)
region_menu.grid(column=1, row=1, padx=10, pady=10)

calculate_button = ttk.Button(root, text="Calculate", command=calculate_efficiencies)
calculate_button.grid(column=0, row=2, columnspan=2, padx=10, pady=10)

reset_button = ttk.Button(root, text="Reset", command=reset)
reset_button.grid(column=0, row=3, columnspan=2, padx=10, pady=10)
result_label = ttk.Label(root, textvariable=result_text, justify="left")
result_label.grid(column=0, row=4, columnspan=2, padx=10, pady=10)

chart_frame = tk.Frame(root)
chart_frame.grid(column=0, row=5, columnspan=2, padx=10, pady=10, sticky="nsew")
text = tk.Text(root, wrap="none")
text.grid(row=0, column=0)
scrollbar = tk.Scrollbar(root, orient="vertical", command=text.yview)
scrollbar.grid(row=0, column=1, sticky="ns")


root.mainloop()