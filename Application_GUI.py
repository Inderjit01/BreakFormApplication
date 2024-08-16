from FileHandling import FileManipulation
from PyQt5.QtWidgets import (QMainWindow, QWidget, QVBoxLayout, QTextEdit, QHBoxLayout, QMessageBox, 
                QTableWidget, QTableWidgetItem, QLineEdit, QLabel, QFormLayout, QPushButton, QDialog)
from PyQt5.QtCore import Qt
from PyQt5 import QtGui, QtMultimedia, QtCore
import datetime
from datetime import date
import calendar
from docx2pdf import convert
import os
import random


class MyWindow(QMainWindow):
    def __init__(self):
        # Initialize the main window and set up UI components
        super(MyWindow, self).__init__()
        self.setup_window() # configure window properties - size and title
        self.setup_ui() # Initialize and set up the user interface
        # Create an instance of FileManipulation
        self.file_manipulation = FileManipulation(self.table_widget, self.two_week_table, self.name_input)

    def setup_window(self):
        # Set the size and position of the main window
        self.setGeometry(400, 200, 1050, 700) 
        # Set the window title
        self.setWindowTitle("Rest Break Form")

    def setup_ui(self):
        # create and set up the central widget that everything gets added to
        self.central_widget = QWidget()
        self.setCentralWidget(self.central_widget)
        
        # Create a vertical layout to organize widgets vertically
        self.layout = QVBoxLayout()
        self.central_widget.setLayout(self.layout)

        # Initialize and set up the table widget for the two-week period
        self.two_week_table = self.create_two_week_table()
        
        # Initialize and set up the table widget to display predefined data
        self.table_widget = self.create_table_widget()
        
        # Create the form layout with input fields and submit button
        self.form_layout = self.create_form_layout()
        
        # Create the QTextEdit widget to display formatted information
        self.restInformation1 = self.create_rest_information_display()

        # Add the QTextEdit and form layout to the main layout
        self.layout.addWidget(self.restInformation1)
        self.layout.addLayout(self.form_layout)

    def create_form_layout(self):
        # Create a form layout to organize labels and input fields
        form_layout = QFormLayout()
        
        # Create the fields for user input and buttons
        self.name_input = self.create_input_field("Enter your name")
        self.start_date_input = self.changeDateInputFormat()
        self.changeDate = QPushButton("Change Date")
        self.submit_button = QPushButton("Submit")
        
        # When the button is pressed it will go to these method calls
        self.changeDate.clicked.connect(self.on_changeDate_clicked)
        self.submit_button.clicked.connect(self.on_submit_clicked)
        
        # Added the buttons to the form layout
        form_layout.addRow(QLabel("Pay Period:"), self.two_week_table)                
        form_layout.addRow(QLabel("Start Date:"), self.start_date_input)
        form_layout.addRow(self.changeDate)
        form_layout.addRow(QLabel("Name:"), self.name_input)
        form_layout.addRow(self.submit_button)

        return form_layout

    def changeDateInputFormat(self):
        #This makes the button appear horizontal for start_date_input
        self.start_year_input = self.create_input_field("Year")
        self.start_year_input.setMaxLength(4) # Limit input to 4 characters
        
        self.start_month_input = self.create_input_field("Month")
        self.start_month_input.setMaxLength(2)
        
        self.start_day_input = self.create_input_field("Day")
        self.start_day_input.setMaxLength(2)

        date_layout = QHBoxLayout()
        date_layout.addWidget(self.start_year_input)
        date_layout.addWidget(self.start_month_input)
        date_layout.addWidget(self.start_day_input)
        
        return date_layout
        
    def create_input_field(self, placeholder_text):
        # Create a QLineEdit input field with the specified placeholder text
        input_field = QLineEdit()
        input_field.setPlaceholderText(placeholder_text)
        return input_field

    def create_rest_information_display(self):
    # Create a QTextEdit widget to display formatted information
        text_edit = QTextEdit()
        text_edit.setReadOnly(True)  # Make the QTextEdit read-only
    
        # Read and format the files
        content1 = self.read_and_format_file('information1.txt')
        content2 = self.read_and_format_file('information2.txt')
    
        # Convert the data table to HTML
        table_html = self.convert_table_to_html()
    
        # Combine all content into a single HTML string
        combined_html = f"{content1}<br><br>{table_html}<br><br>{content2}<br><br>"
        text_edit.setHtml(combined_html)

        return text_edit
    
    def read_and_format_file(self, filename):
        try:
            with open(filename, 'r', encoding='utf-8') as file:
                content = file.read()
                # Directly return content if it's in proper HTML format
                return content
        except FileNotFoundError:
            return f"Error: The file '{filename}' was not found."
        except Exception as e:
            return f"An error occurred: {e}"

    def create_table_widget(self):
        # Create a QTableWidget to diaplay predefined data
        table = QTableWidget()
        data = [["Workdays of less than 3.5 hours", "Rest period may be waived by employee"],
                ["Workdays of 3.5 - 6 hours", "1 rest period"],
                ["Workdays of 6 - 10 hours", "2 rest periods"],
                ["Workdays of 10 - 14 hours", "3 rest periods"]]
        table.setRowCount(len(data))
        table.setColumnCount(len(data[0]))

        # Populate the tabel with data and make cells non-editable
        for row in range(len(data)):
            for col in range(len(data[0])):
                item = QTableWidgetItem(data[row][col])
                item.setFlags(item.flags() & ~Qt.ItemIsEditable)
                table.setItem(row, col, item)

        return table

    def create_two_week_table(self):
        # Create a QTableWidget for a two-week period
        table = QTableWidget()
        table.setRowCount(3)  # Two weeks + one row for comments
        table.setColumnCount(7)  # One column for each day of the week
        table.setHorizontalHeaderLabels(['Sunday', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday'])

        # Calculate the start and end dates for the two-week period        
        today = datetime.date.today()
        weekday = today.weekday()
        days_since_sunday = (weekday + 1) % 7
        start_date = today - datetime.timedelta(days=days_since_sunday + 14)

        # Fill the table with the dates
        for week in range(2):
            # week is the row
            current_date = start_date + datetime.timedelta(weeks=week)
            for day in range(7):
                # Day is the column
                date = current_date + datetime.timedelta(days=day)
                item = QTableWidgetItem(date.strftime("%Y-%m-%d"))
                item.setFlags(item.flags() & ~Qt.ItemIsEditable)
                table.setItem(week, day, item)
                

        # Prefilling comments in two_week_table       
        sunday = QTableWidgetItem("I did not work on either Sundays.")
        table.setItem(2, 0, sunday)
        
        monday = QTableWidgetItem("I took my breaks at the appropriate time on both Mondays.")
        table.setItem(2, 1, monday)
        
        tuesday = QTableWidgetItem("I took my breaks at the appropriate time on both Tuesdays.")
        table.setItem(2, 2, tuesday)
        
        wednesday = QTableWidgetItem("I took my breaks at the appropriate time on both Wednesdays.")
        table.setItem(2, 3, wednesday)
        
        thursday = QTableWidgetItem("I took my breaks at the appropriate time on both Thursdays.")
        table.setItem(2, 4, thursday)
        
        friday = QTableWidgetItem("I took my breaks at the appropriate time on both Fridays.")
        table.setItem(2, 5, friday)
        
        saturday = QTableWidgetItem("I did not work on either Saturdays.")
        table.setItem(2, 6, saturday)

        # Resize rows to fit content and connect cell changes to a handler
        table.resizeRowsToContents()
        table.cellChanged.connect(self.on_cell_changed)

        return table
    
    def on_changeDate_clicked(self):
        #This gets the inputs from user
        year = self.start_year_input.text()
        month = self.start_month_input.text()
        day = self.start_day_input.text()
        
        #This calcuates how many days there are in this month
        today = datetime.date.today()
        today_year = today.year
        today_month = today.month
        daysInMonth = calendar.monthrange(today_year, today_month)[1]
              
        # if everything is empty highlight all parts red
        if year == "" and month == "" and day == "":
            self.incorrectDataInput(self.start_year_input)
            self.incorrectDataInput(self.start_month_input)
            self.incorrectDataInput(self.start_day_input)
            return
        
        # This makes sure that the year input from user is correct
        if len(year) == 4 and year.isdigit() and year.startswith('20'):
            self.correctDataInput(self.start_year_input)
        else:
            self.incorrectDataInput(self.start_year_input)
            return

        # This makes sure that the month input from user is correct
        if (len(month) == 1 or len(month) == 2) and month.isdigit() and 1 <= int(month) <= 12:
            self.correctDataInput(self.start_month_input)
        else:
            self.incorrectDataInput(self.start_month_input)
            return
            
        # This makes sure that the day input from user is correct 
        if (len(day) == 1 or len(day) == 2) and day.isdigit() and 1 <= int(day) <= daysInMonth:
            self.correctDataInput(self.start_day_input)

        else:
            self.incorrectDataInput(self.start_day_input)
            return
        
        # add one zero to the front of month or day if it is less that 1 digit
        month = month.zfill(2)
        day = day.zfill(2)
        
        # Checks if user entered a date in the future
        user_date = date(int(year), int(month), int(day))
        if user_date > today:
            self.incorrectDataInput(self.start_year_input)
            self.incorrectDataInput(self.start_month_input)
            self.incorrectDataInput(self.start_day_input)
            return
        
        # Use user date to update two week table
        self.update_two_week_table(year, month, day)
        
    def update_two_week_table(self, year, month, day):
        # Convert year, month, day to integers
        try:
            year = int(year)
            month = int(month)
            day = int(day)
        except ValueError:
            print("Invalid input: year, month, or day is not an integer.")
            return
    
        # Validate the input date
        try:
            start_date = datetime.date(year, month, day)
        except ValueError:
            print("Invalid date.")
            return
        
        # Calculate the closest Sunday to the left of the calendar
        weekday = start_date.weekday()  # Monday = 0, ..., Sunday = 6
        if weekday == 6:
            start_of_week = start_date
        else:
        # Find the most recent Sunday before or on the start_date
            start_of_week = start_date - datetime.timedelta(days=(weekday + 1))    
        
        # Update the table with the new dates
        for week in range(2):
            current_date = start_of_week + datetime.timedelta(weeks=week)
            for day in range(7):
                date = current_date + datetime.timedelta(days=day)
                item = QTableWidgetItem(date.strftime("%Y-%m-%d"))
                item.setFlags(item.flags() & ~Qt.ItemIsEditable)
                self.two_week_table.setItem(week, day, item)
    
        # Resize rows to fit content
        self.two_week_table.resizeRowsToContents()
        
    def incorrectDataInput(self, userInput):
        # If user input for the date is bad highlight it red
        userInput.setStyleSheet("background-color: lightcoral;")

    def correctDataInput(self, userInput):
        # This rehighlights userInput to white. THis is sepeificly for when userInput was wrong the first time.
        userInput.setStyleSheet("background-color: white;")

    def convert_table_to_html(self):
        # Convert the QTableWidget data to HTML format
        html = '<table border="1" cellspacing="2" cellpadding="5">'
        for row in range(self.table_widget.rowCount()):
            html += '<tr>'
            for col in range(self.table_widget.columnCount()):
                item = self.table_widget.item(row, col)
                text = item.text() if item else ''
                html += f'<td>{text}</td>'
            html += '</tr>'
        html += '</table>'
        return html

    def on_cell_changed(self, row, column):
        # Resize the row to fit content if a cell in the comment row changes
        if row == 2:  # If the changed cell is in the comment row
            self.two_week_table.resizeRowToContents(row)

    def on_submit_clicked(self):
        #Checks if name is at least two characters otherwise highlight red
        name = self.name_input.text()

        if len(name) <= 2 or self.file_manipulation.notOnlyLetters(name):
            self.incorrectDataInput(self.name_input)
            return
        else:
            self.correctDataInput(self.name_input)
            
        #Creates confirmation popup window
        self.submit_confirmation_window()
        
    def submit_confirmation_window(self):
        #Creates a pop window for the user to confirm the dates and name
        msg = QMessageBox()
        msg.setWindowTitle("Confirmation")
        
        start_date = self.two_week_table.item(0, 0).text() # grab the start date from the table
        end_date = self.two_week_table.item(1, 6).text() # grab the end date from the table
        name = self.name_input.text() # grab the name the user inputed

        msg.setText(f"The payperiod is for {start_date} to {end_date}. \nThis form is being completed by {name}.")
        
        #Adds the i icon to the left
        msg.setIcon(QMessageBox.Information)
        
        # Adding the cancel and ok buttons
        msg.setStandardButtons(QMessageBox.Cancel | QMessageBox.Ok)
        msg.setDefaultButton(QMessageBox.Cancel)
        
        # Connect buttonClicked signal to a handler method
        button_result = msg.exec_() # the button that the user presses is button_result
        
        # Passes the button presses to the method
        self.handle_button_result(button_result)

    def handle_button_result(self, result):
        if result == QMessageBox.Ok:
           # Perform the action associated with the Ok button
           self.document_generation_feedback()
       # No action needed for Cancel button, so simply return
        
    def document_generation_feedback(self):
        name = self.name_input.text()
        name = name.replace(" ", "_")
        
        #List of people who can see the memes
        name_list = ["Inderjit_Singh", "Michal_Fabian", "Jose_Ramirez", 
                     "Andres_Aranda", "Corey_Van_Oostende", "Rene_Lopez",
                     "Ulises_Romero"]
        
        image = ["ssg_logo.jpg", "Fortnite.jpg", "Eagle.jpg", "chun.jpg", "Corey.jpg", "Smash.jpg", 
                 "Crazy_Frog.jpg", "pizza.jpg", "Moist.jpg", "Andres.jpg"]
        audio = ["Notification_Sound.mp3", "Fortnite.mp3", "Eagle.mp3", "Notification_Sound.mp3", 
                 "Notification_Sound.mp3", "Smash.mp3", "Crazy_Frog.mp3", "Notification_Sound.mp3",
                 "Moist.mp3", "Notification_Sound.mp3"]
        
        #This creates the third popup when the document generation is complete
        dialog = QDialog(self)
        dialog.setWindowTitle("Document Generation")
        
        layout = QVBoxLayout(dialog)
        
        if name in name_list:
            # Select a random number
            if name == "Andres_Aranda":
                meme = random.randint(1, 10) - 1
            else:
                meme = random.randint(1, 9) - 1
            
            # Selects the meme
            image_label = QLabel(dialog)
            image_label.setPixmap(QtGui.QPixmap(f"memes/{image[meme]}"))
            image_label.setScaledContents(True)
            layout.addWidget(image_label)
            
            # Set maximum size for the QLabel
            image_label.setMaximumSize(500, 400)
            
            # Selects the audio for the meme
            player = QtMultimedia.QMediaPlayer()
            audio_file = QtCore.QUrl.fromLocalFile(f"memes/{audio[meme]}")
            content = QtMultimedia.QMediaContent(audio_file)
            player.setMedia(content)
            
        else:
            # Default option for image and audio
            
            image_label = QLabel(dialog)
            image_label.setPixmap(QtGui.QPixmap("ssg_logo.jpg"))
            image_label.setScaledContents(True)
            layout.addWidget(image_label)
            
            player = QtMultimedia.QMediaPlayer()
            audio_file = QtCore.QUrl.fromLocalFile("Notification_Sound.mp3")
            content = QtMultimedia.QMediaContent(audio_file)
            player.setMedia(content)
        
        message_label = QLabel("Your document has been generated. Please go to Babisha's folder in SSG Customer Access and sign the document.", dialog)
        message_label.setWordWrap(True)  # Allow text to wrap within the label
        layout.addWidget(message_label)

        # This calls the method the generated the documents
        self.file_manipulation.generate_docx(f'{name}.docx')
        print(f"DOCX generated: {name}.docx")
        
        # This turns the doc file to pdf
        convert(f"{name}.docx", f"{name}.pdf")
        print(f"PDF generated: {name}.pdf")
        
        #This calls the method that removes the docx file from directory
        self.file_manipulation.delete_doc(f'{name}.docx')
        
        #This calls the method that moves the pdf to the correct location
        self.file_manipulation.move_pdf(f'{name}.pdf')
        
        
        close_button = QPushButton("Close", dialog)
        layout.addWidget(close_button)
        
        # Plays the audio clip
        player.play()
        
        close_button.clicked.connect(dialog.accept)
        
        dialog.exec_()
        
        
        
        
        
        
        
        
    
        
    
        
        