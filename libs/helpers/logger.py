import logging
import time

logging.basicConfig(level='DEBUG', filemode='w', format='%(process)d-%(levelname)s-%(message)s')
timestr = time.strftime("%Y%m%d-%H%M%S")

log = logging.getLogger()

c_handler = logging.FileHandler('WARNING.log')
f_handler = logging.FileHandler('INFO.log')
c_handler.setLevel(logging.CRITICAL)
f_handler.setLevel(logging.INFO)

# Create formatters and add it to handlers
c_format = logging.Formatter('%(asctime)s-%(name)s-%(levelname)s-%(message)s')
f_format = logging.Formatter('%(asctime)s-%(name)s-%(levelname)s-%(message)s')

c_handler.setFormatter(c_format)
f_handler.setFormatter(f_format)

# Add handlers to the logger
log.addHandler(c_handler)
log.addHandler(f_handler)
