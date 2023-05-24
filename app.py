from PySide6 import QtCore
import socket
from PySide6.QtWidgets import *
import ntpath
import os
import win32com.client
import ctypes.wintypes
CSIDL_PERSONAL = 5
SHGFP_TYPE_CURRENT = 0
import speech_recognition as sr

r = sr.Recognizer()

buf = ctypes.create_unicode_buffer(ctypes.wintypes.MAX_PATH)
ctypes.windll.shell32.SHGetFolderPathW(None, CSIDL_PERSONAL, None, SHGFP_TYPE_CURRENT, buf)
pathProjects = os.getcwd()
if not os.path.exists((buf.value)+'\\'):
    os.makedirs((buf.value)+'\\')
    pathProjects = (buf.value)+'\\'

def path_leaf(path):
    head, tail = ntpath.split(path)
    return tail or ntpath.basename(head)

class Ui_Dialog(object):
    def setupUi(self, Dialog):
        Dialog.setObjectName("Dialog")
        Dialog.resize(1086, 206)
        self.verticalLayout_3 = QVBoxLayout(Dialog)
        self.verticalLayout_3.setObjectName("verticalLayout_3")
        self.horizontalLayout_2 = QHBoxLayout(Dialog)
        self.horizontalLayout_2.setObjectName("horizontalLayout_2")
        self.horizontalLayout = QHBoxLayout()
        self.horizontalLayout.setObjectName("horizontalLayout")
        self.verticalLayout_2 = QVBoxLayout()
        self.verticalLayout_2.setObjectName("verticalLayout_2")

        self.label_2 = QLabel(Dialog)
        self.label_2.setObjectName("label_2")
        self.verticalLayout_2.addWidget(self.label_2)
        self.horizontalLayout.addLayout(self.verticalLayout_2)
        self.verticalLayout = QVBoxLayout()
        self.verticalLayout.setObjectName("verticalLayout")
        self.lineEdit_4 = ""
        self.FlagAccept = False

        self.lineEdit_2 = QLineEdit(Dialog)
        self.lineEdit_2.setObjectName("lineEdit_2")
        self.verticalLayout.addWidget(self.lineEdit_2)
        self.horizontalLayout.addLayout(self.verticalLayout)
        self.horizontalLayout_2.addLayout(self.horizontalLayout)
        self.pushButton = QPushButton(Dialog)
        self.pushButton.setObjectName("pushButton")
        self.horizontalLayout_2.addWidget(self.pushButton)
        self.verticalLayout_3.addLayout(self.horizontalLayout_2)
        self.buttonBox = QDialogButtonBox(Dialog)
        self.buttonBox.setOrientation(QtCore.Qt.Horizontal)
        self.buttonBox.setStandardButtons(QDialogButtonBox.Cancel|QDialogButtonBox.Ok)
        self.buttonBox.setObjectName("buttonBox")
        self.verticalLayout_3.addWidget(self.buttonBox)

        self.retranslateUi(Dialog)
        self.buttonBox.accepted.connect(Dialog.accept)
        self.buttonBox.accepted.connect(self.checkOK)
        self.buttonBox.rejected.connect(Dialog.reject)
        QtCore.QMetaObject.connectSlotsByName(Dialog)

        self.pushButton.clicked.connect(self.getProjectName)
        self.projectName = pathProjects
        self.lineEdit_2.setText(self.projectName)
        self.lineEdit_2.textEdited.connect(self.textEdited)

        self.Dialog =Dialog
    
    def checkOK(self):
        self.FlagAccept = True
        app = win32com.client.Dispatch("PowerPoint.Application")
        presentation = app.Presentations.Open(self.projectName, ReadOnly=1)
        # app.WindowState = 3
        # app.Activate()
        presentation.SlideShowSettings.Run()			 
        while(1):
            with sr.Microphone() as source2:
                print("Recognising...\n")
                r.adjust_for_ambient_noise(source2, duration=0.2)
                audio = r.listen(source2)
                MyText = r.recognize_google(audio)
                MyText = MyText.lower()
                if('next' in MyText):
                    presentation.SlideShowWindow.View.Next()
                elif('previous' in MyText):
                    presentation.SlideShowWindow.View.Previous()
                else:
                    pass
            # if(presentation.SlideShowWindow.View.Exit()):
            #     app.Quit()
        
        # s = socket.socket()		  
        # port = 12345			 
        # s.connect(('127.0.0.1', port))
        # while(True): 
        #     com = s.recv(1024).decode()
        #     print(com)
        #     if "next" in com:
        #         presentation.SlideShowWindow.View.Next()
        #     elif "previous" in com:
        #         presentation.SlideShowWindow.View.Previous()
        #     else:
        #         pass
        # s.close()	 

        # while(1):
        #     print('rr')
        #     r = sr.Recognizer()
        #     with sr.Microphone()as source:
        #         print('Say Something')
        #         audio = r.listen(source)
                
        #         print('Done')
        #         try:
        #             text = r.recognize_sphinx(audio)
        #             print(text)
        #             if text.find('next')!=-1:
        #                 presentation.SlideShowWindow.View.Next()
        #                 print('sliding....')
        #             elif text.find('bingo')!=-1:
        #                 presentation.SlideShowWindow.View.Previous()
        #             else:
        #                 pass
                    
        #         except:
        #             print('error')
        #     presentation.SlideShowWindow.View.Exit()
        #     app.Quit()
    
    def getProjectName(self):
        projectName = QFileDialog.getOpenFileName(filter="Data (*.pptx)")
        print("project name: ",projectName[0])
        self.projectName = projectName[0]
        self.lineEdit_2.setText(self.projectName)
        self.Dialog.setWindowTitle((str(path_leaf(projectName[0]))))
        textIs=(self.lineEdit_2.text())
        self.Dialog.setWindowTitle(path_leaf(projectName[0]))
        if textIs !="":
            self.buttonBox.setEnabled(True)   
        else:
            self.buttonBox.setEnabled(False)
          
    def retranslateUi(self, Dialog):
        _translate = QtCore.QCoreApplication.translate
        Dialog.setWindowTitle(_translate("Dialog", "Open Presentation"))
        self.pushButton.setText(_translate("Dialog", "Browse..."))
        self.label_2.setText(_translate("Dialog", "Location: "))

    def textEdited(self):
        self.Dialog.setWindowTitle(path_leaf(self.lineEdit_2.text()))

if __name__ == "__main__":
    import sys
    app = QApplication(sys.argv)
    Dialog = QDialog()
    ui = Ui_Dialog()
    ui.setupUi(Dialog)
    Dialog.show()
    sys.exit(app.exec())
