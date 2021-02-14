import logging
import config


class Logger:
    def __init__(self, name):
        # "[%(filename)s:%(lineno)s - %(funcName)20s() ] %(message)s"
        formatter = logging.Formatter("%(asctime)s:%(levelname)s:%(name)s:%(funcName)s:%(message)s")
        # stream_formatter = logging.Formatter("%(name)s: %(message)s")

        file_handler = logging.FileHandler(config.logs_path)
        file_handler.setFormatter(formatter)

        stream_handler = logging.StreamHandler()
        stream_handler.setFormatter(formatter)

        self.logger = logging.getLogger(name)
        self.logger.setLevel(logging.INFO)

        self.logger.addHandler(file_handler)
        self.logger.addHandler(stream_handler)
