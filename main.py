from config import logger
from dispatcher import dispatch
from performer import perform

if __name__ == '__main__':
    try:
        dispatch()
    except Exception as e:
        logger.exception(str(e))
    try:
        perform()
    except Exception as e:
        logger.exception(str(e))
