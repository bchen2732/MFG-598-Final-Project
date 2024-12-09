# Configurator for Motor and Gearbox Combinations for Pivot Doors

- Benjamin Chen
- MFG 598 
- Final Project

## **Overview**
This project completed as a final project for MFG 598 is built using python. 
This program allows users to use a graphical user interface to calculate the required torque for a pivot door based on the user-provided door specifications. 
The program then outputs the three best motor/gearbox combinations to provide that torque. 
The option to also add in new iterations of motors or gearboxes into the existing data base is provided as well as saving the outputted combinations for later use. 
---

## **Methodology**
The code generates a GUI that allows for the user to input information regarding specifications of a door. These specifications are used to calculate the torque needed to initiate movement of the door, as the door will continue to move after due to inertia. The torque is calculated by using the angular acceleration, obtained by timing a video of a pivot door opening from a video on the sponsors website, and multiplying it to the moment of inertia. 

The wind loading is calculated by dividing the door at the pivot into two sections. Assuming the wind is acting on the face of the door, if the door is opening the wind would assist the motor on one side while acting against the door on the other. Knowing this, the torques for both of the sections were calculated, then subtracted from each other to obtain the total torque acting at the pivot location.  In order to calculate the torques, the area of each section is multiplied by the wind pressure. This is the force acting at the middle of the sections, so to obtain the torque acting at the pivot, this force is multiplied by the distance between the center of each section and the pivot location. The wind torque is then added to the required torque for the door outputting the total required torque.

After obtaining the total torque required for the given door, each combination of motor and gearboxes can be generated and filtered to find the best possible combinations. A motor can be fitted with just one gearbox, two gearboxes, or three gearboxes. A note that was given by the sponsor stated that the combinations must use a gearbox from the section “Third Gearbox Data” as these are the right angle gearbox that they use for every build. The best combinations are filtered by torque output as well as overall efficiency. There is an added feature that allows the user to save the output as a .xlsx file in case they need to use the information again. This process prompts the user to include a file name.  

A feature that was requested to be added from the sponsors was the ability to add in data to the original dataset when needed. In order to protect this portion from unauthorized users, a password protected section was added into the GUI to allow those with authority to be the only ones to be able to append data. Once authenticated, the user is able to add data into the forth different sections. Looking at the given data, the motor section only uses Torque and Rated Speed as their numerical inputs, so to ensure the data stays consistent, when inputting information for the motors, if the other fields are filled out the program will prompt you with an error. 


## **Features**
### 1. **Torque Calculation**
- Accepts user input for:
  - Door weight (in kilograms).
  - Door dimensions (width and height in meters).
  - Pivot location from the edge of the door.
  - Wind Speed
  - Wind Pressure
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
- Gives the user the option to save the output as a .xlsx file for later use:
  - Users can specify a file name and save location.

---

### 5. **Security and Access Control**
- A set password is used to grant access to adding iterations to the dataset, ensureing only people with the right authority are able to do so. 

---

## **Requirements to Run This Code**
- Python 3
- Required Libraries:
  - `pandas`
  - `tkinter`

Install these libraries using:
```bash
pip install pandas tkinter
