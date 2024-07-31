from PyQt5.QtWidgets import QApplication, QWidget, QPushButton, QVBoxLayout, QFileDialog, QLabel, QLineEdit, QMessageBox
from pathlib import Path
import os
import glob
import sys
import argparse
# Insert the parent directory into the system path
sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), '..')))

# Import necessary modules from the report_generation package
from report_generation.interaction import Interaction
from report_generation.docx_editor import DocxEditor

class MyApp(QWidget):
    def __init__(self, file_path):
        super().__init__()
        self.excel_path=file_path
        self.folder_path=Path(file_path).parent
        self.generate_report(self.excel_path)
        

    def initUI(self):
        # Create a QVBoxLayout instance
        layout = QVBoxLayout()

        # Create a QLabel widget and add it to the layout
        label = QLabel(f"Generating Report from: {self.excel_path}", self)
        layout.addWidget(label)

        # Set the layout to the QWidget
        self.setLayout(layout)

        # Set the window title and size
        self.setWindowTitle('DGNB Report Generator')
        self.resize(300, 100)


    def _check_folder_structure(self, folder_path):
        # Check if the folder contains the necessary files
        files = os.listdir(folder_path)
        if any(file.endswith('.xlsm') for file in files) and any(file.endswith('.docx') for file in files):
            pass
        else:
            # Show an error message if the folder structure is not correct
            QMessageBox.critical(self, 'Error', 'Sorry folder structure doesn\'t comply with requirements.😞')

    def generate_report(self, excel_path):
        # Generate a report based on the files in the selected folder
        self.folder_path = os.path.dirname(self.excel_path)
        self._check_folder_structure(self.folder_path)
        docx_file = glob.glob(os.path.join(self.folder_path, '**', '*Pre-Check.docx'), recursive=True)[0]
        if not docx_file:
            raise FileNotFoundError(f"No file ending with 'Pre-Check.docx' found in {self.folder_path}")
        xlsx_file = self.excel_path
        data = Interaction(xlsx_file)
        report = DocxEditor(docx_file)
        sheet_name = "SQ_Auditoreingaben "

        report.replace_key_words(data , sheet_name)
        images_folder = self.search_folder(self.folder_path, 'Images')
        report.replace_term_with_image("##Image##",images_folder)
        report.save_changes("Output")

        # Show a success message
        QMessageBox.information(self, "Success", "Report successfully generated!🥳\nPlease check the folder.")

    def search_folder(self,parent_folder, folder_name):
        # Search for a specific folder within the parent folder
        for root, dirs, files in os.walk(parent_folder):
            if folder_name in dirs:
                return os.path.join(root, folder_name)
        return None        

if __name__ == '__main__':
    # Start the application
    import sys

    parser = argparse.ArgumentParser(description='Process an Excel file.')
    parser.add_argument('file_path', type=str, help='Path to the Excel file')
    args = parser.parse_args()



    app = QApplication(sys.argv)
    ex = MyApp(args.file_path)
    sys.exit(app.exec_())
