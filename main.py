from gui import main_window
from loguru import logger

logger.add('debug.log')

@logger.catch()
def main():
    main_window()


if __name__ == '__main__':
    main()
