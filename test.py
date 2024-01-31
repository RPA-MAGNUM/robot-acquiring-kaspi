from config import logger
from dispatcher import dispatch
from performer import operations, parking

try:
    logger.warning('> подготовка')
    dispatch()
    logger.warning('+ подготовка')
except Exception as e:
    logger.warning(f'- подготовка {str(e)}')

try:
    logger.warning('> операции')
    operations(1)
    logger.warning('+ операции')
except Exception as e:
    logger.warning(f'- операции {str(e)}')
    logger.exception()

try:
    logger.warning('> парковка')
    parking()
    logger.warning('+ парковка')
except Exception as e:
    logger.warning(f'- парковка {str(e)}')