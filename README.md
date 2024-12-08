# Configurator for Motor and Gerabox Combinations for Pivot Doors

- Benjamin Chen
- MFG 598 
- Final Project

## **Overview**
This project completed as a final project for MFG 598 is built using python. 
This program allows users to use a graphical user interface to calculate the required torque for a piovt door based on the user-provided door specifications. 
The program then outputs the three best motor/gearbox combinations to provide that torque. 
The option to also add in new iterations of motors or gearboxes into the existing data base is provided as well. 
---

## **Features**
### 1. **Torque Calculation**
- Accepts user input for:
  - Door weight (in kilograms).
  - Door dimensions (width and height in meters).
  - Pivot location relative to the door's center.
- Computes the required torque

---

### 2. **Motor and Gearbox Selection**
- Stores motor and gearbox specifications in an Excel dataset.
- Provides the best three combinations of motors and gearboxes that satisfy the required torque.
- Sorts results by efficiency and torque

---

### 3. **Data Management**
- Provides functionality to update the dataset if new motor and/or gearboxes are needed to be added to the data set:
  - Add new motors or gearboxes.

---

### 4. **Results Management**
- Displays results within the GUI.
- Gives the user the option to save the output as a CSV file for later use:
  - Users can specify a file name and save location.

---

### 5. **Security and Access Control**
- A set password is used to grant access to adding iterations to the dataset, ensureing only people with the right authority are able to do so. 

---

## **Requirements**
- Python 3
- Required Libraries:
  - `pandas`
  - `tkinter`

Install these libraries using:
```bash
pip install pandas tkinter
