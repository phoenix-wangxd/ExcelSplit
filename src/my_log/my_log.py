import logging
import sys
import time
from pathlib import Path
from typing import Optional


class MyLog:
    """My customized log class"""

    def __init__(self, app_name: Optional[str] = 'app',
                 folder_name: Optional[str] = 'run_logs',
                 stream_level: Optional[str] = 'info',
                 file_level: Optional[str] = 'debug'):
        """
        Initialize a new log instance
        :param app_name:APP name
        :param folder_name:Name of the folder where logs are stored
        :param stream_level:Log level when write to stream
        :param file_level:Log level when write to log files
        """
        # Log format rule
        __fmt_tmp = "%(asctime)s - %(filename)s[line:%(lineno)d] "
        fmt = f"{__fmt_tmp}- %(levelname)s: %(message)s"
        formatter = logging.Formatter(fmt)

        # Log file path config
        log_folder = Path('.').joinpath(folder_name)
        self.__creat_log_folder(log_folder)
        _log_time = time.strftime('%Y%m%d%H%M', time.localtime(time.time()))
        log_file_name = f"{_log_time}.log"
        log_file_path = str(log_folder.joinpath(log_file_name))

        # Set root logger
        logger = logging.getLogger(app_name)
        logger.setLevel(logging.INFO)
        logger.propagate = False

        file_handler = logging.FileHandler(log_file_path, mode='w')
        file_handler.setLevel(self.get_log_level(file_level))
        file_handler.setFormatter(formatter)

        stream_handler = logging.StreamHandler(sys.stdout)
        stream_handler.setLevel(self.get_log_level(stream_level))
        stream_handler.setFormatter(formatter)

        logger.addHandler(stream_handler)
        logger.addHandler(file_handler)
        for index, log_handler in enumerate(logger.handlers):
            logger.info(f"log_handler_{index}--> {log_handler}")
        self.logger = logger

    @staticmethod
    def get_log_level(level_str: Optional[str]):
        """
        Get log level from string
        :param level_str:Log level string
        """
        _name = ('CRITICAL', 'ERROR', 'WARNING', 'INFO', 'DEBUG', 'NOTSET')
        level_str_upper = level_str.upper()
        if level_str_upper not in _name:
            raise ValueError(f"input {level_str=} not in {_name}!")

        if level_str_upper == _name[0]:
            return logging.CRITICAL
        if level_str_upper == _name[1]:
            return logging.ERROR
        if level_str_upper == _name[2]:
            return logging.WARNING
        if level_str_upper == _name[3]:
            return logging.INFO
        if level_str_upper == _name[4]:
            return logging.DEBUG
        if level_str_upper == _name[5]:
            return logging.NOTSET

    @staticmethod
    def __creat_log_folder(log_folder: Path):
        if not log_folder.is_dir():
            log_folder.mkdir()
