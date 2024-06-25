from PyQt5.QtWidgets import QApplication, QWidget, QPushButton, QVBoxLayout, QFileDialog, QLabel, QLineEdit, QMessageBox
import os
import glob
import sys

sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), '..')))
from report_generation.interaction import Interaction
from report_generation.docx_editor import DocxEditor

class MyApp(QWidget):
    def __init__(self):
        super().__init__()
        self.initUI()
        self.resize(500, 100)
        self.folder_path = None

    def initUI(self):
        outputFileNameLabel = QLabel('Output File Name: (optional)', self)
        self.outputFileNameEdit = QLineEdit(self)
        loadFolderBtn = QPushButton('Load Folder', self)
        loadFolderBtn.clicked.connect(self.load_folder)
        generateReportBtn = QPushButton('Generate report', self)
        generateReportBtn.clicked.connect(self.generate_report)
        vbox = QVBoxLayout()
        vbox.addWidget(outputFileNameLabel)
        vbox.addWidget(self.outputFileNameEdit)
        vbox.addWidget(loadFolderBtn)
        vbox.addWidget(generateReportBtn)
        self.setLayout(vbox)
        self.setWindowTitle('DGNB Report Generator')
        self.show()

    def load_folder(self):
        self.folder_path = QFileDialog.getExistingDirectory(self, 'Select Folder')
        self._check_folder_structure(self.folder_path)

    def _check_folder_structure(self, folder_path):
        files = os.listdir(folder_path)
        if any(file.endswith('.xlsx') for file in files) and any(file.endswith('.docx') for file in files):
            pass
        else:
            QMessageBox.critical(self, 'Error', 'Sorry folder structure doesn\'t comply with requirements.😞')

    def generate_report(self):
        docx_file = glob.glob(os.path.join(self.folder_path, '*.docx'))[0]
        xlsx_file = glob.glob(os.path.join(self.folder_path, '*.xlsx'))[0]
        data = Interaction(xlsx_file)
        report = DocxEditor(docx_file)
        sheet_name = "SQ_Auditoreingaben "

        report.replace_key_words(data , sheet_name)
        images_folder = self.search_folder(self.folder_path, 'Images')
        report.replace_term_with_image("##Image##",images_folder)
        if self.outputFileNameEdit.text() != "":
            report.save_changes(self.outputFileNameEdit.text())
        else:
            report.save_changes("Output")

        # Show a success message
        QMessageBox.information(self, "Success", "Report successfully generated!🥳\nPlease check the folder.")

    def search_folder(self,parent_folder, folder_name):
        for root, dirs, files in os.walk(parent_folder):
            if folder_name in dirs:
                return os.path.join(root, folder_name)
        return None        

if __name__ == '__main__':
    import sys
    app = QApplication(sys.argv)
    ex = MyApp()
    sys.exit(app.exec_())

