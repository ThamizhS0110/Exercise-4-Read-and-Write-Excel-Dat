# Exercise-4-Read-and-Write-Excel-Data

~~~
Name : W Allen Johnston Ozario  
Reg.No : 21222411004  
~~~

## Aim:
To create a UiPath workflow that reads data from one Excel file and writes it to another Excel file using Excel activities.

## Materials Required:
UiPath Studio (Community or Enterprise Edition)
Microsoft Excel installed

Two Excel files:
Input.xlsx (source file with data)
Output.xlsx (destination file to store copied data)

## Procedure:
 Step 1: Create a New Process
Open UiPath Studio and create a new process named ExcelReadWriteExample.

 Step 2: Prepare Excel Files
Create an Excel file named Input.xlsx with sample data, e.g.:

![image](https://github.com/user-attachments/assets/75d24509-308e-472b-b40f-38e4c2f98ce0)
Save this file in your project directory.

 Step 3: Add Excel Application Scope
Drag and drop Excel Application Scope activity.

Set the path to "Input.xlsx".

 Step 4: Read Range Activity
Inside the Excel Scope, drag a Read Range activity.

Properties:
  SheetName: "Sheet1"
  Range: Leave empty to read the whole sheet
  Output: inputDataTable (Create a variable of type DataTable)

 Step 5: Add Second Excel Application Scope
Below the first scope, drag another Excel Application Scope.

Set the path to "Output.xlsx" (can be a new empty file).

 Step 6: Write Range Activity
Inside the second scope, drag a Write Range activity.

Properties:
  DataTable: inputDataTable
  SheetName: "Sheet1"
  Starting Cell: "A1"
  Add Headers: ✔️ (checked)

## OUTPUT:
The contents of Input.xlsx will be duplicated into Output.xlsx after the process runs.
![image](https://github.com/user-attachments/assets/75d24509-308e-472b-b40f-38e4c2f98ce0)

![image](https://github.com/user-attachments/assets/eb11afca-2d45-47c9-bbb9-a7ea9a7f1522)

## Result:
The UiPath workflow successfully reads data from Input.xlsx and writes it to Output.xlsx using Excel activities.
