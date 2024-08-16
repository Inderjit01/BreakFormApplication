import sys
from PyQt5.QtWidgets import QApplication
from Application_GUI import MyWindow

def window():
    # Create a QApplication instance, that allows compatiablity with all type of devices
    app = QApplication(sys.argv)
    
    # Create an instance of MyWindow that creates the window and then displays it
    win = MyWindow()
    win.show()
    
    # Creates an event loop that will constantly check if x on the window is pressed to allow the program to exit
    sys.exit(app.exec_())

# The application will only run on this class
if __name__ == "__main__":
    window()
