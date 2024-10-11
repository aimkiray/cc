import logging

def setup_logging():
    logging.basicConfig(filename='cc.log', filemode='a', format='%(asctime)s - %(levelname)s - %(message)s', level=logging.INFO, encoding='utf-8')

