from PyQt5.QtWidgets import QApplication, QWidget, QPushButton, QVBoxLayout, QFileDialog, QLabel, QLineEdit

class MyApp(QWidget):
    def __init__(self):
        super().__init__()
        self.initUI()
        # Inside the initUI() method
        self.resize(500, 100)  # Set width to 500 and height to 100

    def initUI(self):
        # Create a QLabel for the output file name
        outputFileNameLabel = QLabel('Output File Name: (optional)', self)

        # Create a QLineEdit for the output file name
        self.outputFileNameEdit = QLineEdit(self)

        # Create "Load Folder" button and connect it to a method
        loadFolderBtn = QPushButton('Load Folder', self)
        loadFolderBtn.clicked.connect(self.load_folder)

        # Create "Generate report" button and connect it to a method
        generateReportBtn = QPushButton('Generate report', self)
        generateReportBtn.clicked.connect(self.generate_report)

        # Create a vertical box layout and add the widgets
        vbox = QVBoxLayout()
        vbox.addWidget(outputFileNameLabel)
        vbox.addWidget(self.outputFileNameEdit)
        vbox.addWidget(loadFolderBtn)
        vbox.addWidget(generateReportBtn)

        # Set the layout of the window
        self.setLayout(vbox)

        self.setWindowTitle('My App')
        self.show()

    def load_folder(self):
        # Open a QFileDialog when the "Load Folder" button is clicked
        folder_path = QFileDialog.getExistingDirectory(self, 'Select Folder')
        print(f'Folder path: {folder_path}')

    def generate_report(self):
        # Placeholder method for "Generate report" button
        print('Generate report button clicked')

if __name__ == '__main__':
    import sys
    app = QApplication(sys.argv)
    ex = MyApp()
    sys.exit(app.exec_())
