from config import logger
from dispatcher import dispatch
from performer import parking, operations, sales

if __name__ == '__main__':
    try:
        dispatch()
    except Exception as e:
        logger.exception(str(e))
    try:
        parking()
    except Exception as e:
        logger.exception(str(e))

    try:
        operations()
    except Exception as e:
        logger.exception(str(e))

    try:
        sales()
    except Exception as e:
        logger.exception(str(e))
