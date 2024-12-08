'''
Benjamin  Chen
MFG 598
Final Project

'''
'''
Problem Definition: 
Currently, the sponsors have no way to accuratly predict which motor and gearbox combination to use based on the door specifications. 
The aim of the program is to create a graphical user interface that allows user to easily identify the best combintaiton as well as givng the abilty to save the output and/or add new motors/gearboxes. 

'''


import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import pandas as pd
import os

EXCEL_FILE_PATH = "Pivot Calculator Formated(86).xlsx"

PASSCODE = "1234"  


# This function calcualtes the torque based on the weight and pivor location of the door
def calculate_torque():
    try:
        weight = float(door_weight_entry.get())  
        width = float(door_width_entry.get())  
        height = float(door_height_entry.get()) 
        pivot_location = float(pivot_location_entry.get())  

        g = 9.81  

        force = weight * g  

        center_of_mass = width / 2
        distance_to_pivot = abs(center_of_mass - pivot_location) 

        torque = force * distance_to_pivot 

        output_text.config(state=tk.NORMAL)
        output_text.delete(1.0, tk.END)
        output_text.insert(tk.END, f"Required Torque: {torque:.2f} Nm\n")
        output_text.config(state=tk.DISABLED)
    except ValueError:
        output_text.config(state=tk.NORMAL)
        output_text.delete(1.0, tk.END)
        output_text.insert(tk.END, "Invalid input. Please enter valid numbers.\n")
        output_text.config(state=tk.DISABLED)

#This function handles creating all of the combinations for the different sets of motor and gearboxes. It should be noted that the thrid gearbox MUST be in every combination
# because it is a right angle gearbox. Once the user enters the information, this function will also filter out the combinations and then display the top outputs based on torque and efficiency.
def find_motor_gearbox():
    excel_data = pd.ExcelFile(EXCEL_FILE_PATH)

    motor_data = excel_data.parse('Motor Data')
    third_gearbox_data = excel_data.parse('Third Gearbox Data')
    first_gearbox_data = excel_data.parse('First Gearbox Data')
    second_gearbox_data = excel_data.parse('Second Gearbox Data')

    torque_text = output_text.get("1.0", tk.END).strip()
    if not torque_text.startswith("Required Torque:"):
        messagebox.showerror("Error", "Please calculate the required torque first.")
        return

    required_torque = float(torque_text.split(":")[1].strip().split()[0])  

    motor_data['Torque (Nm)'] = pd.to_numeric(motor_data['Torque (Nm)'], errors='coerce')
    third_gearbox_data['Rated Output Torque (Nm)'] = pd.to_numeric(third_gearbox_data['Rated Output Torque (Nm)'], errors='coerce')
    third_gearbox_data['Ratio (:1)'] = pd.to_numeric(third_gearbox_data['Ratio (:1)'], errors='coerce')
    third_gearbox_data['Efficiency (%)'] = pd.to_numeric(third_gearbox_data['Efficiency (%)'], errors='coerce')
    first_gearbox_data['Ratio (:1)'] = pd.to_numeric(first_gearbox_data['Ratio (:1)'], errors='coerce')
    first_gearbox_data['Efficiency (%)'] = pd.to_numeric(first_gearbox_data['Efficiency (%)'], errors='coerce')
    second_gearbox_data['Ratio (:1)'] = pd.to_numeric(second_gearbox_data['Ratio (:1)'], errors='coerce')
    second_gearbox_data['Efficiency (%)'] = pd.to_numeric(second_gearbox_data['Efficiency (%)'], errors='coerce')

    motor_data = motor_data.dropna(subset=['Torque (Nm)'])
    third_gearbox_data = third_gearbox_data.dropna(subset=['Rated Output Torque (Nm)', 'Efficiency (%)', 'Ratio (:1)'])

    combinations = []
    for _, motor in motor_data.iterrows():
        motor_torque = motor['Torque (Nm)']
        motor_pn = motor['P/N']

        for _, third_gearbox in third_gearbox_data.iterrows():
            third_ratio = third_gearbox['Ratio (:1)']
            third_efficiency = third_gearbox['Efficiency (%)']
            third_pn = third_gearbox['P/N']

            combined_torque_1 = motor_torque * third_ratio
            if combined_torque_1 >= required_torque:
                combinations.append({
                    'Motor P/N': motor_pn,
                    'Gearbox 1 P/N': None,
                    'Gearbox 2 P/N': None,
                    'Gearbox 3 P/N': third_pn,
                    'Combined Torque (Nm)': combined_torque_1,
                    'Combined Efficiency (%)': third_efficiency
                })

            for _, first_gearbox in first_gearbox_data.iterrows():
                first_ratio = first_gearbox['Ratio (:1)']
                first_efficiency = first_gearbox['Efficiency (%)']
                first_pn = first_gearbox['P/N']

                combined_torque_2 = motor_torque * first_ratio * third_ratio
                combined_efficiency_2 = first_efficiency * third_efficiency
        if combined_torque_2 >= required_torque:
            combinations.append({
                'Motor P/N': motor_pn,
                'Gearbox 1 P/N': first_pn,
                'Gearbox 2 P/N': None,
                'Gearbox 3 P/N': third_pn,
                'Combined Torque (Nm)': combined_torque_2,
                'Combined Efficiency (%)': combined_efficiency_2
            })

    for _, second_gearbox in second_gearbox_data.iterrows():
        second_ratio = second_gearbox['Ratio (:1)']
        second_efficiency = second_gearbox['Efficiency (%)']
        second_pn = second_gearbox['P/N']

        combined_torque_3 = motor_torque * second_ratio * third_ratio
        combined_efficiency_3 = second_efficiency * third_efficiency
        if combined_torque_3 >= required_torque:
            combinations.append({
                'Motor P/N': motor_pn,
                'Gearbox 1 P/N': None,
                'Gearbox 2 P/N': second_pn,
                'Gearbox 3 P/N': third_pn,
                'Combined Torque (Nm)': combined_torque_3,
                'Combined Efficiency (%)': combined_efficiency_3
            })

    for _, first_gearbox in first_gearbox_data.iterrows():
        for _, second_gearbox in second_gearbox_data.iterrows():
            first_ratio = first_gearbox['Ratio (:1)']
            first_efficiency = first_gearbox['Efficiency (%)']
            first_pn = first_gearbox['P/N']
            second_ratio = second_gearbox['Ratio (:1)']
            second_efficiency = second_gearbox['Efficiency (%)']
            second_pn = second_gearbox['P/N']

            combined_torque_4 = motor_torque * first_ratio * second_ratio * third_ratio
            combined_efficiency_4 = first_efficiency * second_efficiency * third_efficiency 
            if combined_torque_4 >= required_torque:
                combinations.append({
                    'Motor P/N': motor_pn,
                    'Gearbox 1 P/N': first_pn,
                    'Gearbox 2 P/N': second_pn,
                    'Gearbox 3 P/N': third_pn,
                    'Combined Torque (Nm)': combined_torque_4,
                    'Combined Efficiency (%)': combined_efficiency_4
                })

    results_df = pd.DataFrame(combinations)

    results_window = tk.Toplevel(root)
    results_window.title("Motor and Gearbox Combinations")

    tree = ttk.Treeview(results_window, columns=('Motor P/N', 'Gearbox 1 P/N', 'Gearbox 2 P/N', 'Gearbox 3 P/N', 'Combined Torque (Nm)', 'Combined Efficiency (%)'), show='headings')
    tree.heading('Motor P/N', text='Motor P/N')
    tree.heading('Gearbox 1 P/N', text='Gearbox 1 P/N')
    tree.heading('Gearbox 2 P/N', text='Gearbox 2 P/N')
    tree.heading('Gearbox 3 P/N', text='Gearbox 3 P/N')
    tree.heading('Combined Torque (Nm)', text='Combined Torque (Nm)')
    tree.heading('Combined Efficiency (%)', text='Combined Efficiency (%)')
    tree.pack(fill=tk.BOTH, expand=True)

    results_df = results_df.sort_values(by=['Combined Efficiency (%)', 'Combined Torque (Nm)'],ascending=[False, False]).head(3)

    for _, row in results_df.iterrows():
        tree.insert("", tk.END, values=tuple(row.values))

    tk.Label(results_window, text="Enter File Name to Save Results:").pack(pady=5)
    file_name_entry = tk.Entry(results_window)
    file_name_entry.pack(pady=5)

    save_btn = tk.Button(
        results_window, 
        text="Save Results", 
        command=lambda: save_results_with_name(results_df, file_name_entry)
    )
    save_btn.pack(pady=10)

# This function allows users to save the outputs into a new excel sheet allowing users to the information later if needed
def save_results_with_name(results_df, file_name_entry):
        if results_df.empty:
            messagebox.showerror("Error", "No results to save.")
            return

        file_name = file_name_entry.get().strip()
        if not file_name:
            messagebox.showerror("Error", "Please enter a file name.")
            return

        if not file_name.endswith(".xlsx"):
            file_name += ".xlsx"

        file_path = filedialog.asksaveasfilename(
            initialfile=file_name,
            defaultextension=".xlsx",
            filetypes=[("Excel Files", "*.xlsx"), ("All Files", "*.*")]
        )

        if file_path:  
            results_df.to_excel(file_path, index=False)
            messagebox.showinfo("Success", f"Results saved to {file_path}")

# This function handels the authentication needed to gain access to add motors or gearboxes to the dataset. The defined password is set at the very top of this code labled "PASSWORD".
def authentication():
    def verify_passcode():
        entered_passcode = passcode_entry.get()
        if entered_passcode == PASSCODE:
            auth_window.destroy()
            show_buttons()
            messagebox.showinfo("Access Granted", "You can now add new motors and/or gearboxes to the data set.")
        else:
            messagebox.showerror("Error", "Invalid passcode. Access denied.")

    auth_window = tk.Toplevel(root)
    auth_window.title("Enter Passcode")

    tk.Label(auth_window, text="Enter Passcode:").grid(row=0, column=0, padx=10, pady=5)
    passcode_entry = tk.Entry(auth_window, show="*") 
    passcode_entry.grid(row=0, column=1, padx=10, pady=5)

    submit_btn = tk.Button(auth_window, text="Submit", command=verify_passcode)
    submit_btn.grid(row=1, column=0, columnspan=2, pady=10)


# This function hides the buttons to allow users to add motor and gearboxes untill after the authentication is verified
def show_buttons():
    global add_motor_btn, add_first_gearbox_btn, add_second_gearbox_btn, add_third_gearbox_btn

    tk.Label(root, text="Add New Items:").grid(row=8, column=0, columnspan=2, pady=10)

    add_motor_btn = tk.Button(root, text="Add Motor", command=lambda: add_item("Motor"))
    add_motor_btn.grid(row=9, column=0, padx=10, pady=5)

    add_first_gearbox_btn = tk.Button(root, text="Add First Gearbox", command=lambda: add_item("First Gearbox"))
    add_first_gearbox_btn.grid(row=9, column=1, padx=10, pady=5)

    add_second_gearbox_btn = tk.Button(root, text="Add Second Gearbox", command=lambda: add_item("Second Gearbox"))
    add_second_gearbox_btn.grid(row=10, column=0, padx=10, pady=5)

    add_third_gearbox_btn = tk.Button(root, text="Add Third Gearbox", command=lambda: add_item("Third Gearbox"))
    add_third_gearbox_btn.grid(row=10, column=1, padx=10, pady=5)

# This function appends the added data into the existing exel sheet used for the program. This ensures that the data is always up to date. 
def append_to_excel(section, data):
    try:
        if not os.path.exists(EXCEL_FILE_PATH):
            messagebox.showerror("Error", f"{EXCEL_FILE_PATH} does not exist.")
            return


        excel_data = pd.ExcelFile(EXCEL_FILE_PATH)
        sheet_data = excel_data.parse(section + " Data")

        new_data = pd.DataFrame([data])

        updated_data = pd.concat([sheet_data, new_data], ignore_index=True)

        with pd.ExcelWriter(EXCEL_FILE_PATH, mode='a', if_sheet_exists='replace') as writer:
            updated_data.to_excel(writer, sheet_name=section + " Data", index=False)

        messagebox.showinfo("Success", f"New item added to {section.lower()} section!")
    except Exception as e:
        messagebox.showerror("Error", f"An error occurred: {str(e)}")

# This function creates the fields to allow users to enter the information of the motor or gearbox that they are trying to add. It ensures that the users input all the required infomration before allowing it to be appended into the exel sheet.
def add_item(section):
    def save_item():
        p_n = p_n_entry.get().strip()
        torque = torque_entry.get().strip()
        ratio = ratio_entry.get().strip()
        efficiency = efficiency_entry.get().strip()

        if section != "Motor":
            if not p_n or not torque or not ratio or not efficiency:
                messagebox.showerror("Input Error", "Please fill out all fields before trying to add a new gearbox to the dataset.")
                return  

            try:
                torque = float(torque)
                ratio = float(ratio)
                efficiency = float(efficiency)
            except ValueError:
                messagebox.showerror("Input Error", "Please ensure torque, ratio, and efficiency are valid numbers.")
                return

        new_row = {
            "P/N": p_n,
            "Torque (Nm)": float(torque) if section == "Motor" else None,
            "Rated Output Torque (Nm)": float(torque) if section != "Motor" else None,
            "Ratio (:1)": float(ratio) if section != "Motor" else None,
            "Efficiency (%)": float(efficiency) if section != "Motor" else None
        }

        append_to_excel(section, new_row)
        add_window.destroy()


    add_window = tk.Toplevel(root)
    add_window.title(f"Add New {section}")

    tk.Label(add_window, text="P/N:").grid(row=0, column=0, padx=10, pady=5)
    p_n_entry = tk.Entry(add_window)
    p_n_entry.grid(row=0, column=1, padx=10, pady=5)

    tk.Label(add_window, text="Torque (Nm):").grid(row=1, column=0, padx=10, pady=5)
    torque_entry = tk.Entry(add_window)
    torque_entry.grid(row=1, column=1, padx=10, pady=5)

    tk.Label(add_window, text="Ratio (:1):").grid(row=2, column=0, padx=10, pady=5)
    ratio_entry = tk.Entry(add_window)
    ratio_entry.grid(row=2, column=1, padx=10, pady=5)

    tk.Label(add_window, text="Efficiency (%):").grid(row=3, column=0, padx=10, pady=5)
    efficiency_entry = tk.Entry(add_window)
    efficiency_entry.grid(row=3, column=1, padx=10, pady=5)

    save_btn = tk.Button(add_window, text="Save", command=save_item)
    save_btn.grid(row=4, column=0, columnspan=2, pady=10)


# This handles the GUI genration and button placements as well as calling the necessary functions to give the users the needed outputs. 
root = tk.Tk()
root.title("Pivot Door Motor and Gearbox Configurator")

tk.Label(root, text="Door Weight (kg):").grid(row=0, column=0, padx=10, pady=5)
door_weight_entry = tk.Entry(root)
door_weight_entry.grid(row=0, column=1, padx=10, pady=5)

tk.Label(root, text="Door Width (m):").grid(row=1, column=0, padx=10, pady=5)
door_width_entry = tk.Entry(root)
door_width_entry.grid(row=1, column=1, padx=10, pady=5)

tk.Label(root, text="Door Height (m):").grid(row=2, column=0, padx=10, pady=5)
door_height_entry = tk.Entry(root)
door_height_entry.grid(row=2, column=1, padx=10, pady=5)

tk.Label(root, text="Pivot Point Location (m):").grid(row=3, column=0, padx=10, pady=5)
pivot_location_entry = tk.Entry(root)
pivot_location_entry.grid(row=3, column=1, padx=10, pady=5)

calculate_btn = tk.Button(root, text="Calculate Torque", command=calculate_torque)
calculate_btn.grid(row=4, column=0, columnspan=2, pady=10)

find_motor_btn = tk.Button(root, text="Find Motor/Gearbox", command=find_motor_gearbox)
find_motor_btn.grid(row=5, column=0, columnspan=2, pady=10)

output_label = tk.Label(root, text="Required Torque:", font=("Arial", 12, "bold"))
output_label.grid(row=6, column=0, columnspan=2, pady=10)

output_text = tk.Text(root, height=10, width=50, state=tk.DISABLED)
output_text.grid(row=7, column=0, columnspan=2, padx=10, pady=5)

tk.Label(root, text="Add New Items:").grid(row=8, column=0, columnspan=2, pady=10)

auth_btn = tk.Button(root, text="Authenticate", command=authentication)
auth_btn.grid(row=11, column=0, columnspan=2, pady=10)

root.mainloop()
