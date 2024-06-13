import logging

# Define custom colors using ANSI escape codes
COLORS = {
    'DEBUG': '\033[94m',   # Blue
    'INFO': '\033[92m',    # Green
    'WARNING': '\033[93m', # Yellow
    'ERROR': '\033[91m',   # Red
    'CRITICAL': '\033[95m' # Magenta
}
RESET = '\033[0m'  # Reset to default color

# Custom formatter to add color to log messages
class ColoredFormatter(logging.Formatter):
    def format(self, record):
        levelname = record.levelname
        message = super().format(record)
        color = COLORS.get(levelname, RESET)
        return f'{color}{message}{RESET}'

def configure_logging():
    # Create the root logger
    logger = logging.getLogger()
    logger.setLevel(logging.DEBUG)  # Set the root logger level to DEBUG

    # Remove all previous handlers to avoid duplicate logs
    for handler in logger.handlers[:]:
        logger.removeHandler(handler)

    # Configure logging to file
    file_handler = logging.FileHandler('DummyDeviceBuilder.log')
    file_handler.setLevel(logging.DEBUG)

    # Configure logging to console with colored formatter
    console_handler = logging.StreamHandler()
    console_handler.setLevel(logging.DEBUG)
    console_handler.setFormatter(ColoredFormatter('%(asctime)s - %(name)s - %(levelname)s - %(message)s'))

    # Add handlers to the logger
    logger.addHandler(file_handler)
    logger.addHandler(console_handler)

if __name__ == '__main__':
    configure_logging()
    logging.info('Logger configuration test message')
