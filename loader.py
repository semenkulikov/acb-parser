import logging
import os
from logging.handlers import RotatingFileHandler


BASE_DIR = os.path.dirname(os.path.abspath(__file__))

log_formatter = logging.Formatter('%(asctime)s | %(levelname)s | %(funcName)s - %(message)s')
my_handler = RotatingFileHandler(os.path.join(BASE_DIR, "parser.log"), mode='a', maxBytes=2 * 1024 * 1024,
                                 backupCount=1, encoding="utf8", delay=0)
my_handler.setFormatter(log_formatter)
my_handler.setLevel(logging.INFO)
