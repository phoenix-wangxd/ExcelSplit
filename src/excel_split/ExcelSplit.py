import time
import logging
from typing import Optional
from pathlib import Path
from dataclasses import dataclass, field
from openpyxl import load_workbook, workbook, worksheet

# type hints
WorkBook = workbook.workbook.Workbook
WorkSheet = worksheet


def check_path_is_file(path):
    if not Path(path).is_file():
        raise FileExistsError(f"{path = } not a file!!")


class MyLog:
    def __init__(self, folder_name: Optional[str] = 'run_logs'):
        """
        Initialize a new instance
        :param folder_name:Name of the folder where logs are stored
        """
        logger = logging.getLogger()
        logger.setLevel(logging.INFO)

        log_folder = Path('.').joinpath(folder_name)
        self.__creat_log_folder(log_folder)

        _log_time = time.strftime('%Y%m%d%H%M', time.localtime(time.time()))
        log_name = f"{_log_time}.log"
        log_file = str(log_folder.joinpath(log_name))
        fh = logging.FileHandler(log_file, mode='w')
        fh.setLevel(logging.DEBUG)

        fmt = "%(asctime)s - %(filename)s[line:%(lineno)d] " \
              "- %(levelname)s: %(message)s"
        formatter = logging.Formatter(fmt)
        fh.setFormatter(formatter)

        logger.addHandler(fh)
        self.logger = logger

    @staticmethod
    def __creat_log_folder(log_folder: Path):
        if not log_folder.is_dir():
            log_folder.mkdir()


@dataclass
class ExcelObj:
    """Excel Object """
    orig_file_path: str
    new_file_name: str = field(init=False)

    def __post_init__(self):
        check_path_is_file(self.orig_file_path)
        _parent_path = Path(self.orig_file_path).parent
        _new_file_name = self.orig_file_path.replace('.xls', '_new.xls')

        self.wb: Optional[WorkBook] = None  # Excel WorkBook
        self.orig_sheet_name_list: list = list()
        self.orig_first_sheet_name: Optional[str] = None
        self.orig_first_sheet: WorkSheet = None  # Excel First Work Sheet Object

        self.new_file_path = str(_parent_path.joinpath(_new_file_name))
        self.new_sheet_names: list = list()  # All new sheet name List
        self.new_sheets: list = list()  # All new sheet object List


class ExcelSplit:
    def __init__(self, orig_file_path: str, split_numb: int = 8000):
        """
        Initialize a new instance
        :param orig_file_path:  Path of the original Excel file
        :param split_numb: Every split_numb records are cut into a new sheet
        """
        self.logger = MyLog().logger
        self.excel = ExcelObj(orig_file_path=orig_file_path)
        _parent_path = Path(orig_file_path).parent
        _file_name = self.excel.orig_file_path.replace('.xls', '_new.xls')
        self.new_file_path = _parent_path.joinpath(_file_name)
        self.new_sheet_prefix: str = 'data_'
        self.split_num = split_numb
        self.__open_file()

    def __open_file(self):
        wb: WorkBook = load_workbook(filename=self.excel.orig_file_path)
        orig_sheet_names = wb.sheetnames
        orig_first_sheet_name = orig_sheet_names[0]
        self.excel.wb = wb
        self.excel.orig_sheet_name_list = orig_sheet_names
        self.excel.orig_first_sheet_name = orig_first_sheet_name
        self.excel.orig_first_sheet = wb[orig_first_sheet_name]

        orig_max_row = self.excel.orig_first_sheet.max_row
        max_sheet_numb = int(orig_max_row / self.split_num) + 1
        self.logger.info(f"{self.excel.orig_first_sheet.dimensions = }")
        self.logger.info(f"{self.excel.orig_first_sheet.max_row = }")

        _new_sheet_list = [f"{self.new_sheet_prefix}{i + 1}" for i in
                           range(max_sheet_numb)]
        self.excel.new_sheet_names = _new_sheet_list
        self.logger.info(f"{self.excel.new_sheet_names = }")

    def creat_all_new_sheets(self):
        for i in self.excel.new_sheet_names:
            sheet = self.excel.wb.create_sheet(i)
            self.excel.new_sheets.append(sheet)

        self.logger.info(f"now all {self.excel.wb.sheetnames = }")
        self.logger.info(f"all new sheets {self.excel.new_sheets = }")

    def write_all_new_sheet_record(self):
        for ws_name in self.excel.new_sheet_names:
            ws_index = int(ws_name.replace(self.new_sheet_prefix, ''))
            start_row_numb = self.split_num * (ws_index - 1) + 1

            _records = self.get_orig_sheet_mult_rows(
                start_row_numb=start_row_numb, count=self.split_num)

            if not _records:
                self.__write_one_new_sheet(ws_name=ws_name, start_row_numb=1,
                                           src_data=_records)

    def get_orig_sheet_mult_rows(self, start_row_numb: Optional[int] = 1,
                                 count: Optional[int] = 8000):
        """
        Get multiple records from the original sheet
        :param start_row_numb: Start reading row number, Must be greater than 0
        :param count: Number of records to be read
        :return: Results of records read
        """
        ws: WorkSheet = self.excel.orig_first_sheet
        if (not isinstance(start_row_numb, int)) or start_row_numb <= 0:
            self.logger.error(f"{start_row_numb = } must be greater than 0 ")
            raise ValueError(f"{start_row_numb = } must be greater than 0 ")

        if start_row_numb > ws.max_row:
            self.logger.warning(f"{start_row_numb = } > {ws.max_row = }")
            return tuple()

        if (not isinstance(count, int)) or count <= 0:
            self.logger.error(f"{count = } must be greater than 0 ")
            raise ValueError(f"{count = } must be greater than 0 ")

        max_row = start_row_numb + count - 1
        if max_row > ws.max_row:
            max_row = ws.max_row

        return ws[start_row_numb: max_row]

    def __write_one_new_sheet(self, ws_name, start_row_numb, src_data) -> None:
        """
        Write the records read from the original table in a new sheet
        :param ws_name: The name of the sheet to be written
        :param start_row_numb: Start write row number, Must be greater than 0
        :param src_data: Original record
        """
        ws = self.excel.wb[ws_name]
        self.logger.info(f"start write {len(src_data) = }")
        for i, v in enumerate(src_data):
            row = i + start_row_numb
            value = v[0].value
            self.logger.debug(f'{ws = }, {row}, {value = }')
            ws[row] = value

    def save_to_disk(self, new_file_path: Optional[str] = None):
        _file_path = new_file_path if new_file_path else self.new_file_path
        self.excel.wb.save(_file_path)


if __name__ == '__main__':
    file_path = f'file_001.xlsx'
    a = ExcelSplit(file_path)
    a.creat_all_new_sheets()
    a.write_all_new_sheet_record()
    a.save_to_disk()
