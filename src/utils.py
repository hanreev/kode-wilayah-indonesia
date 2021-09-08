# -*- coding: utf-8 -*-

import logging
import os
import re
import sys
from typing import Any, Tuple, Union

BASE_DIR = os.path.abspath(os.path.dirname(os.path.dirname(__file__)))
DATA_DIR = os.path.join(BASE_DIR, 'src', 'data')
OUT_DIR = os.path.join(BASE_DIR, 'output')
LOGS_DIR = os.path.join(BASE_DIR, 'logs')

os.makedirs(OUT_DIR, exist_ok=True)
os.makedirs(LOGS_DIR, exist_ok=True)


def create_logger(filename: str, level: Union[int, str, None] = logging.INFO, console: bool = True) -> logging.Logger:
    log_formatter = logging.Formatter('%(asctime)s [%(levelname)s]: %(message)s')
    logger = logging.getLogger()
    file_handler = logging.FileHandler(filename=filename, mode='w')
    file_handler.setFormatter(log_formatter)
    logger.addHandler(file_handler)

    if console:
        console_handler = logging.StreamHandler()
        console_handler.setFormatter(log_formatter)
        logger.addHandler(console_handler)

    logger.setLevel(level)

    if '-v' in sys.argv or '--verbose' in sys.argv:
        logger.setLevel(logging.DEBUG)

    return logger


def log_filename(fname: str) -> str:
    basename, ext = os.path.splitext(os.path.basename(fname))
    return os.path.join(LOGS_DIR, '{}.log'.format(basename))


def touch(fname: str, times: Union[Tuple[int, int], Tuple[float, float], None] = None):
    with open(fname, 'a'):
        os.utime(fname, times)


def normalize_value(value: Any) -> str:
    if value is None:
        return ''
    value = str(value).strip().replace('\n', ' ')
    value = re.sub(r'^\d+\s+', '', value)
    return re.sub(r'\s+', ' ', value)
