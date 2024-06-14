from report_generation.interaction import Interaction
from report_generation.report import Report

if __name__ == "__main__":
    excel_fp = r"C:\Users\vmylavarapu\Buro Happold\Germany Computational Team - General\3 Development\3 Cities\DGNB Precheck Report Generator\Q20_Gew.Tab_Mehrsprachig_230524_gesch.xlsx"
    intrctn = Interaction(excel_fp)

    print(intrctn.get_data_at_cell('SQ_Auditoreingaben ','G21')) # output should be 80.0