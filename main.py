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

    try:
        logger.warning('> операции')
        operations(delta)
        logger.warning('+ операции')
    except Exception as e:
        logger.warning(f'- операции {str(e)}')
        logger.exception()

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
