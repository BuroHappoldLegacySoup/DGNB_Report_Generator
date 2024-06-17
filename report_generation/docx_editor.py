from docx import Document
import os
from docx.oxml.ns import nsdecls
from docx.oxml import parse_xml


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

    def replace_text(self, term : str, replacement_text: str):
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
