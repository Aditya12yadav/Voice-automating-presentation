# Voice-automating-presentation
This project is a Python GUI application that allows users to open and control a PowerPoint presentation using voice commands. The application provides a simple interface for browsing and selecting a presentation file, and it uses speech recognition to listen for voice commands to navigate through the slides.

It is a graphical user interface (GUI) application written in Python using the PySide6 library. The application allows users to open and control a PowerPoint presentation using voice commands. Here is a breakdown of the code:

1. Importing necessary modules:
   - PySide6 modules for GUI components
   - socket for networking
   - ntpath and os for file path manipulation
   - win32com.client for interacting with PowerPoint
   - ctypes.wintypes for retrieving system folders
   - speech_recognition for speech recognition capabilities

2. Defining global variables and constants:
   - CSIDL_PERSONAL and SHGFP_TYPE_CURRENT for retrieving the user's personal folder path
   - `pathProjects` variable to store the current working directory
   - `r` object for speech recognition

3. Defining utility functions:
   - `path_leaf()` function to extract the filename from a given path

4. Defining the UI class `Ui_Dialog`:
   - The `setupUi()` method sets up the GUI components, such as labels, line edits, buttons, and layouts.
   - Signal connections and event handlers are defined to perform actions when buttons are clicked or text is edited.

5. The `checkOK()` method is called when the OK button is clicked. It opens a PowerPoint presentation using the `win32com.client` module and starts a slide show. It continuously listens for voice commands using the `speech_recognition` module and performs actions accordingly (e.g., going to the next or previous slide).

6. The `getProjectName()` method is called when the Browse button is clicked. It opens a file dialog to select a PowerPoint presentation file and sets the selected file's path in the line edit.

7. The `textEdited()` method updates the window title based on the edited text in the line edit.

8. The `retranslateUi()` method sets the text and titles of GUI components.

9. The main block initializes the GUI, creates an instance of the `Ui_Dialog` class, and displays the dialog.
