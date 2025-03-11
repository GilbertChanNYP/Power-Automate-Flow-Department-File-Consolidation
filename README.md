# Power Automate Flow: Department File Consolidation

##  Overview
This Power Automate Desktop flow consolidates financial data from multiple department files into a **Master Excel file**. It extracts relevant data from predefined tables in the data files and sums the values into the corresponding cells in the Master file.

---
##  **Key Features**
 - **Automated Data Collection** – Eliminates manual data entry and ensures accuracy.  
 - **Dynamic Table Selection** – Users can choose which table to consolidate.  
 - **Error Handling** – Ensures that empty cells are skipped to prevent errors.  
 - **Scalability** – Works for any number of department files. 

---

##  **Flow Breakdown**
### 1️. **Select Department Files Folder**
- The flow starts with a **folder selection dialog**.
- The user selects a folder containing all department files.

### 2️. **Retrieve Files in the Selected Folder**
- The flow retrieves all files from the selected folder.

### 3️. **Select the Master File**
- A file selection dialog prompts the user to choose the **Master Excel File**.

### 4️. **Open Master Excel File**
- The selected Master Excel file is opened in an existing Excel process.

### 5️. **Define Table Ranges**
The flow categorizes financial data into the following groups:

| **Category**                        | **Start Row** | **End Row** |
|--------------------------------------|--------------|--------------|
| **Variable Expense**                 | 69           | 75           |
| **Income**                            | 8            | 11           |
| **Salaries & Social Charges**         | 15           | 65           |
| **External Contract**                 | 79           | 83           |
| **Infrastructure**                    | 87           | 106          |
| **Travel, Training & Meeting**        | 110          | 117          |

These categories are stored in a **list (`TableList`)**.

### 6️. **User Selects a Table for Consolidation**
- A list of table names is presented to the user.
- The user selects the table to process.
- The corresponding row range (`Value1`, `Value2`) is retrieved.

### 7️. **Process Each Department File**
- The flow iterates through each department file.
- Each file is opened in **Excel**.

### 8️. **Extract and Consolidate Data**
- The flow loops through **columns 58 to 72** (DB2025).
- For each column:
  - Loops through rows within the selected category’s range.
  - Reads the corresponding cell values from:
    - **The department file (`ExcelData`)**
    - **The Master file (`MasterData`)**
  - If values exist in both files, they are summed.
  - The summed values are stored in a temporary list.

### 9️. **Update the Master File**
- The summed values are written to the **Master Excel File** in the corresponding column and starting row (`Value1`).
- The temporary list is cleared before processing the next column.

---

##  **How to Use This Flow**
1. Open **Power Automate Desktop**.
2. Open the provided text file: **`FYP Power Automate Flow.txt`**.
3. Copy the entire content of the text file.
4. In Power Automate Desktop, create a **New Flow** and make sure the Power Fx option is **NOT** enabled.
5. Paste the copied content into the new flow.
6. Save the flow.
7. Run the flow and follow the prompts to select folders and files.
8. Choose the financial category to consolidate.
9. The automation will process the data and update the Master Excel file accordingly.
