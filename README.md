# KISDI_BudgetReportManager

## Overview
This project focuses on managing and Excel files through an interactive user interface. The primary functionality includes dynamically adding new sheets to existing Excel files, as well as providing drag-and-drop capabilities in a Treeview for easier data manipulation. The aim is to simplify data management without the need to open the Excel file manually each time. Simultaneously, it is helpful to write your budget reports.

## Features
- **Dynamic Sheet Addition**: Add new sheets to an existing Excel file based on the file's original location, making it easier to expand and update data without overwriting or creating new files.
- **Drag-and-Drop Treeview**: Interactively update data and manage files using a drag-and-drop feature integrated into the Treeview UI, enhancing user experience and reducing manual steps.
- **Automatic Data Updates**: Modify and add data to Excel files automatically without the need to open the files through Excel, saving time and improving workflow efficiency.
- **Error Handling**: Handles scenarios such as attempting to modify a file that is already open in Excel, preventing accidental overwrites.

## Getting Started
### Prerequisites
- Python 3.x
- Required libraries:
  - ```openpyxl``` for working with Excel files
  - ```tkinter``` for the graphical user interface (GUI)
  - ```pandas``` for advanced data manipulation (if needed)
  - ```pyinstaller``` for making the ```.exe``` files (if needed)

You can install the required libraries using ```pip```:
```
pip install openpyxl pandas tkinter pyinstaller
```

## Installation
1. Clone this repository:
   ```
   git clone https://github.com/dolphin1404/KISDI_BudgetReportManager.git
   ```
2. Navigate to the project directory:
   ```
   cd KISDI_BudgetReportManager
   ```
3. Run the main application:
   ```
   python KISDI_Budget.py  
   ```

### Usage
1. Open the application.
2. Select your excel file to insert in budget report.
3. The Treeview interface will display the structure of the existing Excel files.
4. Push the button, then there are new sheets that makes it easy to write your budget report on your excel file. Probably, this button will save the file without opening Excel manually.

## Contributing
If you'd like to contribute to this project, feel free to fork the repository, make improvements, and submit pull requests. Please ensure that your code is well-tested and adheres to the project’s coding style.

### Steps for Contributing:
1. Fork the repository.
2. Create a new branch ```(git checkout -b feature-branch)```.
3. Make your changes.
4. Commit your changes ```(git commit -am 'Add new feature')```.
5. Push to the branch ```(git push origin feature-branch)```.
6. Create a new Pull Request.
