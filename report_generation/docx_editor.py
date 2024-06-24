from docx import Document
import os
from docx.oxml import parse_xml
import re
from report_generation.interaction import Interaction
from docx.shared import Inches

class DocxEditor:
    """
    A class to edit a .docx file.
    """

    def __init__(self, file_path : str):
        """
        Initialize the DocxEditor with a .docx file path.
        
        :param file_path: The path to the .docx file to be edited.
        """
        self.file_path = file_path
        self.doc = Document(file_path)

    def add_text(self, term : str, text_to_add : str):
        """
        Add a new paragraph with the specified text after the paragraph containing the term.
        
        :param term: The term to search for in the document.
        :param text_to_add: The text to add in a new paragraph after the paragraph containing the term.
        """
        for para in self.doc.paragraphs:
            if term in para.text:
                new_para_xml = parse_xml(
                    r'<w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"></w:p>')
                para._p.addnext(new_para_xml)
                new_run = parse_xml(
                    r'<w:r xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:t xml:space="preserve">%s</w:t></w:r>' % text_to_add)
                new_para_xml.append(new_run)
                break

    def add_bullet_list(self, term : str, bullet_list : list):
        """
        Add a bullet list after the paragraph containing the term.
        
        :param term: The term to search for in the document.
        :param bullet_list: The list of items to add as bullet points after the paragraph containing the term.
        """
        bullet_list.reverse()
        for para in self.doc.paragraphs:
            if term in para.text:
                for item in bullet_list:
                    new_para_xml = parse_xml(
                        r'<w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"></w:p>')
                    para._p.addnext(new_para_xml)
                    new_run = parse_xml(
                        r'<w:r xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:t xml:space="preserve">• %s</w:t></w:r>' % item)
                    new_para_xml.append(new_run)
                break

    def replace_text_block(self, term : str, replacement_text: str):
        """
        Replace the paragraph containing the term with the replacement text.
        
        :param term: The term to search for in the document.
        :param replacement_text: The text to replace the paragraph containing the term.
        """
        for para in self.doc.paragraphs:
            if term in para.text:
                for run in para.runs[::-1]:
                    run.clear()
                para.add_run(replacement_text)
                break

    def save_changes(self, new_file_name=None):
        """
        Save the changes to a new file.
        If a new file name is provided, use it. Otherwise, use the default name 'new_' + old file name.
        
        :param new_file_name: The name of the new file where the changes will be saved. If not provided, the default name 'new_' + old file name will be used.
        """
        dir_name = os.path.dirname(self.file_path)
        if new_file_name is None:
            base_name = 'new_' + os.path.basename(self.file_path)
        else:
            base_name = new_file_name + os.path.splitext(self.file_path)[1]
        self.new_file_path = os.path.join(dir_name, base_name)
        self.doc.save(self.new_file_path)

    def replace_specific_text(self, term: str, replacement_text: str):
        """
        Replace a specific text within the paragraph containing the term with the replacement text.

        :param term: The term to search for in the document.
        :param replacement_text: The text to replace the specific term.
        """
        for para in self.doc.paragraphs:
            if term in para.text:
                new_text = para.text.replace(term, replacement_text)
                para.clear()
                para.add_run(new_text)

    def replace_key_words(self, data : Interaction , sheet_name : str, key_word = None):
        """
        This function replaces key words in the document with data from an Excel sheet. It searches for patterns of the form "Excel (word)" in the document. For each match, it retrieves the corresponding data from the specified Excel sheet and replaces the match with this data.

        :param data: An Interaction object which contains the data from the Excel sheet.
        :param sheet_name: The name of the Excel sheet from which to retrieve data.
        :param key_word: The key word to search for in the document. If not provided, the function will search for the pattern "Excel (word)".
        """
        pattern = r"Excel \(\w+\)"
        for para in self.doc.paragraphs:
            matches = re.findall(pattern, para.text)
            for match in matches:
                cell_reference = str(match.split("(")[1].split(")")[0])
                print(cell_reference)              
                replacement_text = str(data.get_data_at_cell(sheet_name, cell_reference))
                self.replace_specific_text(match, replacement_text)

    def replace_term_with_image(self, term: str):
        """
        Search for the term in the document and replace it with an image from the specified file path.
        The image is scaled to fit the width of the page and a caption is added below the image.
        The caption is formatted with the "Caption" style.

        :param term: The term to search for in the document.
        """
        for i, para in enumerate(self.doc.paragraphs):
            if term in para.text:
                print(para.text)

                file_name, img_cap = self.get_image_data(para.text)
                img_path = self.find_image(file_name, r"C:\\Users\\vmylavarapu\\Pictures")
                if img_path != None:

                    para.clear()
                    run = para.add_run()
                    run.add_picture(img_path, width=Inches(6))
                    # Add the caption to a new paragraph with the "Caption" style
                    new_para_xml = parse_xml(r'<w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"></w:p>')
                    para._p.addnext(new_para_xml)
                    new_run = parse_xml(r'<w:r xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:t xml:space="preserve">%s</w:t></w:r>' % img_cap)
                    new_para_xml.append(new_run)
                    self.doc.paragraphs[i+1].style = self.doc.styles['Caption']
                    break


    @staticmethod
    def get_image_data(info : str):
        """
        This function extracts the image name and caption from a string. The string is expected to be in the format "##Image##image_name##image_caption##".

        :param info: The string containing the image name and caption.
        :return: A tuple containing the image name and caption.
        """
        parts = info.split("##")

        img_name = parts[2]
        img_caption = parts[3]

        return img_name, img_caption

    @staticmethod
    def find_image(image_name :str, folder_path : str):
        """
        This function searches for an image file in a specified folder. The image file is expected to have a specific name and be in one of the
        following formats: .jpeg, .jpg, .png, .tiff.

        :param image_name: The name of the image file to search for.
        :param folder_path: The path of the folder in which to search for the image file.
        :return: The path of the image file if found, otherwise None.
        """
        image_formats = [".jpeg", ".jpg", ".png",".tiff"]

        for filename in os.listdir(folder_path):
            if filename.startswith(image_name) and any(filename.endswith(format) for format in image_formats):
                return os.path.join(folder_path, filename)

        return None
