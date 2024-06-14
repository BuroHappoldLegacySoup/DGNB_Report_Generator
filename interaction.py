from python_calamine import CalamineWorkbook

class Interaction:

    def __init__(self, excel_file : str):
        self.excel_file = excel_file

    def get_data_at_cell(self, sheet_name : str, cell : str):
        all_data = CalamineWorkbook.from_path(self.excel_file).get_sheet_by_name(sheet_name).to_python(skip_empty_area=False)
        row, col = self._cell_to_indices(cell)
        return all_data[row][col]

    def _cell_to_indices(self,cell):
        col_idx = 0

        for i in range(len(cell)):
            if cell[i].isdigit():
                for j in range(i):
                    col_idx = col_idx * 26 + (ord(cell[j]) - ord('A') + 1)
                row_idx = int(cell[i:])
                return row_idx -1 , col_idx -1
