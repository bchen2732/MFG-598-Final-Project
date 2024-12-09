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

# Initiate the program with these required libraries. The command to install them is in the "READ_ME" section in the Github
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import pandas as pd
import os

# File path to the base spreadsheet that is given by the sponsors. Here the reletive path is being used. 
EXCEL_FILE_PATH = "Pivot Calculator Formated(86).xlsx"

# Predifned password that the user is able to change to whatever they want. Here I used "1234" for simplicity
PASSCODE = "1234"  

# This function calcualtes the torque based on the weight and pivort location of the door. It also allows the choice for the user to take into account wind load. 
def calculate_torque():
    try:
        weight = float(door_weight_entry.get())  # Weight of the door in kg
        width = float(door_width_entry.get())  # Width of the door in meters
        height = float(door_height_entry.get())  # Height of the door in meters
        pivot_location = float(pivot_location_entry.get())  # Pivot location in meters

        #Calculate the torque needed to move the door by using a given angular acceration and calculating the moment of inertia. Then multiplying to get the troque needed to move the door. 
        angular_acc = 0.223 #rad/s^2 , this was derived by timming a video that i found of a door opening and closing online on the sponsors website
        moment_of_inertia = ((1/12) * weight * width**2)  + (weight * pivot_location**2)
        distance_to_pivot = pivot_location
        torque_door= angular_acc * moment_of_inertia 

        torque_wind = 0 #Initialized as 0 if the user chooses not to incorporate wind in the calculations

        if include_wind_var.get():
            wind_speed = float(wind_speed_entry.get())  # Wind speed in m/s
            air_density = float(air_density_entry.get())  # Air density in kg/m^3
            
            # Calculate wind pressure
            wind_pressure = 0.5 * air_density * (wind_speed ** 2)

            # We are asuming that the wind is direclty acting on the face of the door. This means that one side of the door from the pivot will have a pushing force assiting the motor, while the other side is acting against the motor. A picture explaining this calculation can be found in the github
            door_area_1 = height * distance_to_pivot
            door_area_2 = height * (abs(distance_to_pivot-width))

            wind_force_1 = wind_pressure * door_area_1 # Wind force acting at center on first section of the door
            wind_force_2 = wind_pressure * door_area_2 # Wind force acting at center on the second section of the door

            torque_wind = ((abs(distance_to_pivot-width)/2) * wind_force_2) - ((distance_to_pivot/2) * wind_force_1)  # Subtracting the torques since they are acting against each other

        # Total torque is calcculated by the sum of torque of the door and wind load
        total_torque = torque_door + torque_wind

        # Apply a safety factor of 10% as requested by sponsors
        safety_factor = 1.10
        total_torque_with_safety = total_torque * safety_factor

        output_text.config(state=tk.NORMAL)
        output_text.delete(1.0, tk.END)
        output_text.insert(tk.END, f"Required Torque:\n")
        output_text.insert(tk.END, f"- For the Door: {torque_door:.2f} Nm\n")
        if include_wind_var.get():
            output_text.insert(tk.END, f"- Due to Wind: {torque_wind:.2f} Nm\n")
        output_text.insert(tk.END, f"- Total (Before Safety Factor): {total_torque:.2f} Nm\n")
        output_text.insert(tk.END, f"- Total (With 10% Safety Factor): {total_torque_with_safety:.2f} Nm\n")
        output_text.config(state=tk.DISABLED)

    except ValueError:  # Handles invalid inputs 
        output_text.config(state=tk.NORMAL)
        output_text.delete(1.0, tk.END)
        output_text.insert(tk.END, "Invalid input. Please enter valid numbers.\n")
        output_text.config(state=tk.DISABLED)

# This function allows the user to toggle if they want to include wind load in the calculations or not
def toggle_wind_load():
    if include_wind_var.get():
        wind_speed_entry.config(state=tk.NORMAL)
        air_density_entry.config(state=tk.NORMAL)
    else:
        wind_speed_entry.delete(0, tk.END)
        air_density_entry.delete(0, tk.END)
        wind_speed_entry.config(state=tk.DISABLED)
        air_density_entry.config(state=tk.DISABLED)

#This function handles creating all of the combinations for the different sets of motor and gearboxes. It should be noted that the thrid gearbox MUST be in every combination
# because it is a right angle gearbox. Once the user enters the information, this function will also filter out the combinations and then display the top outputs based on torque and efficiency.
def find_motor_gearbox():
    excel_data = pd.ExcelFile(EXCEL_FILE_PATH)

    motor_data = excel_data.parse('Motor Data')
    third_gearbox_data = excel_data.parse('Third Gearbox Data')
    first_gearbox_data = excel_data.parse('First Gearbox Data')
    second_gearbox_data = excel_data.parse('Second Gearbox Data')

    torque_text = output_text.get("1.0", tk.END).strip()

    for line in torque_text.split("\n"):
        if "- Total (With 10% Safety Factor):" in line:
            required_torque = float(line.split(":")[1].strip().split()[0])
             

    # Pulling all the data with in the given spreadsheet and converting them in to numerical values to be used for calcuations
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

    # Building all the different possible combinations given the dataset. The motor can be fitted with one gear box, two gear boxes, or three gearboxes.
    # It should be noted that the "Third Gearbox" must be used in all combintations as it is the right angles gearbox and is required due to the way these motor and gearboxes are configured for the doors. 
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
                combined_efficiency_2 = first_efficiency * third_efficiency * 100 #The efficenticy is turned into a decimal form to be used for calculation, here we multiply it by 100 to convert it back to percentage.
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
        combined_efficiency_3 = second_efficiency * third_efficiency * 100
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
            combined_efficiency_4 = first_efficiency * second_efficiency * third_efficiency * 100
            if combined_torque_4 >= required_torque:
                combinations.append({
                    'Motor P/N': motor_pn,
                    'Gearbox 1 P/N': first_pn,
                    'Gearbox 2 P/N': second_pn,
                    'Gearbox 3 P/N': third_pn,
                    'Combined Torque (Nm)': combined_torque_4,
                    'Combined Efficiency (%)': combined_efficiency_4
                })

    # Creates a dataframe for all the combinations
    results_df = pd.DataFrame(combinations)

    # Creates a pop up window to dispay the resulting top 3 combinations for the given input door specifications
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

    # Creates a section to allow users to choose a file name before saving results to the computer
    tk.Label(results_window, text="Enter File Name to Save Results:").pack(pady=5)
    file_name_entry = tk.Entry(results_window)
    file_name_entry.pack(pady=5)

    # Creates a save button 
    save_btn = tk.Button(
        results_window, 
        text="Save Results", 
        command=lambda: save_results_with_name(results_df, file_name_entry)
    )
    save_btn.pack(pady=10)

# This function allows users to save the outputs into a new excel sheet allowing users to the information later if needed
def save_results_with_name(results_df, file_name_entry):
        file_name = file_name_entry.get().strip()

        #Ensures that there is a file name before saving
        if not file_name:
            messagebox.showerror("Error", "Please enter a file name.")
            return

        if not file_name.endswith(".xlsx"):
            file_name += ".xlsx"

        # Saves the file as a xlsx to the operating computer 
        file_path = filedialog.asksaveasfilename(
            initialfile=file_name,
            defaultextension=".xlsx",
            filetypes=[("Excel Files", "*.xlsx"), ("All Files", "*.*")]
        )

        # Creates a popup window allowing users to know that the file was saved succesfuflly as well as display the file path for easy identification in the future
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
            messagebox.showinfo("Access Granted", "You can now add new motors and/or gearboxes to the data set.") # Confirmation the password is correct
        else:
            messagebox.showerror("Error", "Invalid passcode. Access denied.") # Error message if passowrd is wrong

    auth_window = tk.Toplevel(root)
    auth_window.title("Enter Passcode")


    # Creates a pop up window to allow for users to enter the password
    tk.Label(auth_window, text="Enter Passcode:").grid(row=0, column=0, padx=10, pady=5)
    passcode_entry = tk.Entry(auth_window, show="*") 
    passcode_entry.grid(row=0, column=1, padx=10, pady=5)

    submit_btn = tk.Button(auth_window, text="Submit", command=verify_passcode)
    submit_btn.grid(row=1, column=0, columnspan=2, pady=10)


# This function hides the buttons to allow users to add motor and gearboxes untill after the authentication is verified as an added bouns security feature
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

        #Write into the base excel sheet that contains all of the motor and gearbox data
        with pd.ExcelWriter(EXCEL_FILE_PATH, mode='a', if_sheet_exists='replace') as writer: 
            updated_data.to_excel(writer, sheet_name=section + " Data", index=False)

        #Display a confirmation to show users that what they wanted to add has been added
        messagebox.showinfo("Success", f"New item added to {section.lower()} section!")

    # Display an error with the error code to help users understand why it might not have accepted the addition
    except Exception as e:
        messagebox.showerror("Error", f"An error occurred: {str(e)}")

# This function creates the fields to allow users to enter the information of the motor or gearbox that they are trying to add. It ensures that the users input all the required infomration before allowing it to be appended into the exel sheet.
def add_item(section):
    def save_item():
        p_n = p_n_entry.get().strip()
        torque = torque_entry.get().strip()
        ratio = ratio_entry.get().strip()
        efficiency = efficiency_entry.get().strip()

        # The motor section in the given base spreadsheet only contained torque and speed. To ensure that the data set stays consistent, we force the user to only input those sections for the motor only. 
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

        # Creates a new row to be ready to be written into the base excel document
        new_row = {
            "P/N": p_n,
            "Torque (Nm)": float(torque) if section == "Motor" else None,
            "Rated Output Torque (Nm)": float(torque) if section != "Motor" else None,
            "Ratio (:1)": float(ratio) if section != "Motor" else None,
            "Efficiency (%)": float(efficiency) if section != "Motor" else None
        }

        append_to_excel(section, new_row) # Calls the append_to_excel function to write the row into the dataset
        add_window.destroy()

    # This all generates the buttons to allow for the user to choose if they want to add a motor or gearbox into the dataset. Note: These buttons only show up AFTER the user authenticates with the password. 
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

# Wind load option  buttons 
include_wind_var = tk.BooleanVar(value=False)
tk.Checkbutton(root, text="Include Wind Load", variable=include_wind_var, command=toggle_wind_load).grid(row=4, column=0, columnspan=2)

tk.Label(root, text="Wind Speed (m/s):").grid(row=5, column=0)
wind_speed_entry = tk.Entry(root, state=tk.DISABLED)
wind_speed_entry.grid(row=5, column=1)

tk.Label(root, text="Air Density (kg/m³):").grid(row=6, column=0)
air_density_entry = tk.Entry(root, state=tk.DISABLED)
air_density_entry.grid(row=6, column=1)

find_motor_btn = tk.Button(root, text="Find Motor/Gearbox", command=find_motor_gearbox)
find_motor_btn.grid(row=8, column=0, columnspan=2, pady=10)

output_label = tk.Label(root, text="Output:", font=("Arial", 12, "bold"))
output_label.grid(row=9, column=0, columnspan=2, pady=10)

output_text = tk.Text(root, height=10, width=50, state=tk.DISABLED)
output_text.grid(row=10, column=0, columnspan=2, padx=10, pady=5)

# Calculate Button
calculate_btn = tk.Button(root, text="Calculate Torque", command=calculate_torque)
calculate_btn.grid(row=7, column=0, columnspan=2, pady=10)

tk.Label(root, text="Add New Motors/Gearboxes to the Dataset:").grid(row=11, column=0, columnspan=2, pady=10)

# Authenticate Button 
auth_btn = tk.Button(root, text="Authenticate", command=authentication)
auth_btn.grid(row=12, column=0, columnspan=2, pady=10)

root.mainloop()
