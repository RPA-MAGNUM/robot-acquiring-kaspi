from config import logger
from dispatcher import dispatch
from performer import parking, operations, sales, prepare

if __name__ == '__main__':
    delta = 1
    try:
        logger.warning('> подготовка')
        dispatch(delta)
        logger.warning('+ подготовка')
    except Exception as e:
        logger.warning(f'- подготовка {str(e)}')

    try:
        logger.warning('> парковка')
        parking(delta)
        logger.warning('+ парковка')
    except Exception as e:
        logger.warning(f'- парковка {str(e)}')