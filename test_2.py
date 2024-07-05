from config import logger
from dispatcher import dispatch
from performer import parking, operations, sales, prepare

if __name__ == '__main__':
    delta = 0
    try:
        logger.warning('> операции')
        operations(delta)
        logger.warning('+ операции')
    except Exception as e:
        logger.warning(f'- операции {str(e)}')
        logger.exception()