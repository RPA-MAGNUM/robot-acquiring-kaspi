import traceback

from config import logger
from dispatcher import dispatch
from performer import parking, operations, sales, prepare

if __name__ == '__main__':
    delta = 0
    result = False

    try:
        logger.warning('> разбивка')
        result = prepare(delta)
        logger.warning('+ разбивка')
    except Exception as e:
        logger.warning(f'- разбивка {str(e)}')

    if result:
        try:
            logger.warning('> sales')
            sales(delta)
            logger.warning('+ sales')
        except Exception as e:
            logger.warning(f'- sales {str(e)}')