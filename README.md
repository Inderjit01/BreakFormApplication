## Overview

The Rest Period Compliance Automation Program was developed for a company I work for to streamline the process of ensuring all employees comply with legal requirements regarding rest periods. This program automates the completion of a legal form that every employee is required to sign each pay period, confirming their compliance with rest period regulations. By automating this process, the program saves time and ensures consistency and accuracy across the company.

## Features

- **Automation of Legal Compliance**: The program automatically generates and fills out the necessary forms that employees must sign, confirming whether they have taken their legally required rest periods.
- **User-Friendly Interface**: A graphical user interface (GUI) makes it easy for employees to interact with the program and complete the form with minimal effort.
- **Batch Processing**: The program supports batch processing, allowing multiple employees to complete their forms in one session.
- **Error Handling**: Built-in error handling ensures that any issues during form completion are flagged and addressed promptly.

## How It Works

### Main Components

1. **`main.py`**: The central script that drives the program. It handles the overall workflow, including calling other modules, processing data, and interacting with the user through the GUI.

2. **`Application_GUI.py`**: This module is responsible for creating the user interface, allowing employees to easily navigate and complete their required forms.

3. **`FileHandling.py`**: Manages file operations, such as reading and writing the form data, saving completed forms, and ensuring data integrity.

4. **`information.txt` & `information2.txt`**: These files contain the legal text that informs employees of their rights and responsibilities regarding rest periods. The text is presented to the employee as part of the form completion process.

5. **Batch File (`launch.bat`)**: A batch script that checks for necessary software dependencies, installs them if needed, and then launches the program. It ensures that Python and the required libraries (PyQt5, docx2pdf, python-docx, beautifulsoup4) are installed and up-to-date.

### Usage Instructions

1. **Installation**:
   - Ensure Python is installed on your system. The batch file (`launch.bat`) will check for Python and required libraries, and install them if they are not already present.
   - Run `launch.bat` to start the program.

2. **Running the Program**:
   - The program will search for the `BreakFormApplication` directory on your C: or D: drive.
   - Once the program is located, it will open a user-friendly interface where employees can complete and submit their rest period compliance forms.
   - Employees will be prompted to review the information contained in `information.txt` and `information2.txt` to ensure they understand their obligations and rights regarding rest periods.

3. **Form Completion**:
   - Each employee must initial the form indicating whether they have taken their required rest periods.
   - Completed forms are saved automatically, and the data is processed to determine if a Rest Period Premium is required.
