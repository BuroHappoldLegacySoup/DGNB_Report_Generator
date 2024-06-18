from report_generation.interaction import Interaction
from dependencies import check_packages
from report_generation.docx_editor import DocxEditor
import re

def main():
    check_packages()


if __name__ == "__main__":
    main()
    excel_fp = r"C:\Users\vmylavarapu\Buro Happold\Germany Computational Team - General\3 Development\3 Cities\DGNB Precheck Report Generator\Q20_Gew.Tab_Mehrsprachig_230524_gesch.xlsx"
    wrd_fp = r"C:\Users\vmylavarapu\Desktop\240410_Elbhafen_DGNB-SQ20-Pre-Check.docx"
    
    data = Interaction(excel_fp)
    report = DocxEditor(wrd_fp)
    sheet_name = "SQ_Auditoreingaben "

    report.replace_key_words(data , sheet_name) 
    report.save_changes("OP_TEST_run")