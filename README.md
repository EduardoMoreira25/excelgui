Usage
To use this program, follow these steps:

Open a terminal or command prompt and navigate to the directory where you saved the data_entry_form.py file.
Run the command: python data_entry_form.py
The program will display a window with several fields for the user to fill out. The user can enter their name, city, favorite color, languages spoken, and number of children. When the user clicks the "Submit" button, the data is saved to an Excel file.
The user can also click the "Clear" button to reset all the input fields.
Saving Data
By default, the program saves data to an Excel file named gui-excel.xlsx. You can change the name of the file by modifying the EXCEL_FILE variable in the code.

Creating an Executable
To create an executable for this program, you can use a tool like PyInstaller. Follow these steps:

Install PyInstaller using pip: pip install pyinstaller
Navigate to the directory containing the data_entry_form.py file.
Run the command: pyinstaller data_entry_form.py --onefile
The executable will be created in the dist directory.
