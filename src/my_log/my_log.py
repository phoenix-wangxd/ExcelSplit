import logging
import sys
import time
from pathlib import Path
from typing import Optional


class MyLog:
    def __init__(self, folder_name: Optional[str] = 'run_logs',
                 file_level: Optional[str] = 'info'):
        """
        Initialize a new instance
        :param folder_name:Name of the folder where logs are stored
        """
        logger = logging.getLogger()
        logger.setLevel(logging.DEBUG)

        # log format config
        fmt = "%(asctime)s - %(filename)s[line:%(lineno)d] " \
              "- %(levelname)s: %(message)s"
        formatter = logging.Formatter(fmt)

        # log file config
        log_folder = Path('.').joinpath(folder_name)
        self.__creat_log_folder(log_folder)
        _log_time = time.strftime('%Y%m%d%H%M', time.localtime(time.time()))
        log_file_name = f"{_log_time}.log"
        log_file_path = str(log_folder.joinpath(log_file_name))

        file_handler = logging.FileHandler(log_file_path, mode='w')
        file_handler.setLevel(logging.DEBUG)
        file_handler.setFormatter(formatter)
        stream_handler = logging.StreamHandler(sys.stdout)
        stream_handler.setLevel(logging.INFO)
        stream_handler.setFormatter(formatter)

        logger.addHandler(stream_handler)
        logger.addHandler(file_handler)
        self.logger = logger

    @staticmethod
    def __creat_log_folder(log_folder: Path):
        if not log_folder.is_dir():
            log_folder.mkdir()

