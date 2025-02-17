# mod_Logging.py

import logging


# =========================================================================
# 設定日誌
# =========================================================================
def init_logging():
    logging.basicConfig(
        filename='process_log.txt',
        level=logging.INFO,
        format='%(asctime)s - %(levelname)s - %(message)s'
    )

def logging_process_step(msg):
    print(msg)
    logging.info(msg)

def logging_exc_error(msg, error):
    print(f'{msg}，錯誤訊息: {error}')
    logging.error(f"{msg}，錯誤訊息: {error}", exc_info=True)

def logging_exception(msg, error):
    print(f'{msg}，錯誤訊息: {error}')
    logging.exception(f"{msg}，錯誤訊息: {error}", exc_info=True)

def logging_warning(msg):
    print(msg)
    logging.info(msg)
