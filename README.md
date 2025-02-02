# **Automation in Data Entry**  
A Python application that automates data entry by storing user input in both an **Excel file** and a **SQLite database**. The stored data can later be visualized in **Power BI**.

## **Features**
- **Graphical User Interface (GUI)** using **Tkinter**  
- **User Input Form** for collecting personal and product details  
- **Automatic Data Storage**  
  - Saves data to an **Excel file**  
  - Stores data in a **SQLite database**  
- **Power BI Compatibility**  
- **Terms & Conditions Acceptance Checkbox**  

## **Installation**  

### **Prerequisites**  
Make sure you have Python installed (version 3.x recommended). You also need to install the following dependencies:

```bash
pip install openpyxl
```

## **How to Use**
1. Run the Python script:  
   ```bash
   python script_name.py
   ```
2. Enter user details in the GUI.
3. Click **"Enter data"** to store the information.  
4. Data will be saved in both:
   - An **Excel file** (`path-direction.xlsx` – update the path in the script)
   - A **SQLite database** (`path-direction.db`– update the path in the script)
5. The stored data can later be visualized in **Power BI**.

## **Project Structure**
```
Automation-in-Data-Entry/
│-- script.py           # Main Python script
│-- path-direction.xlsx # Excel file (auto-generated)
│-- dath-direction.db   # SQLite database (auto-generated)
```

## **Technologies Used**
- **Python** (Tkinter, openpyxl, sqlite3)  
- **Excel** (for data storage)  
- **SQLite** (for database management)  
- **Power BI** (for visualization)  

## **Future Improvements**
- Add data validation for better error handling  
- Implement CSV export functionality  
- Enhance UI with modern styling 
