from python_calamine import CalamineWorkbook


class Interaction:
    """
    This class represents an interaction with an Excel file.
    """

    def __init__(self, excel_file: str):
        """
        Initializes the Interaction with the given Excel file.

        :param excel_file: The path to the Excel file.
        """
        self.excel_file = excel_file

    def get_data_at_cell(self, sheet_name: str, cell: str):
        """
        Retrieves the data at a specific cell in a specific sheet of the Excel file.

        :param sheet_name: The name of the sheet.
        :param cell: The cell to retrieve data from.
        :return: The data at the specified cell.
        """
        all_data = CalamineWorkbook.from_path(
            self.excel_file).get_sheet_by_name(sheet_name).to_python(skip_empty_area=False)
        row, col = self._cell_to_indices(cell)
        return all_data[row][col]

    def _cell_to_indices(self, cell):
        """
        Converts a cell name (like 'A1') to row and column indices.

        :param cell: The cell name.
        :return: A tuple of (row index, column index).
        """
        col_idx = 0

        for i in range(len(cell)):
            if cell[i].isdigit():
                for j in range(i):
                    col_idx = col_idx * 26 + (ord(cell[j]) - ord('A') + 1)
                row_idx = int(cell[i:])
                return row_idx - 1, col_idx - 1