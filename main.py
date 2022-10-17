import sys, getopt
import logging
import json
from timi import timemanager

logger = logging.getLogger(__name__)

# logging reference https://basicincome30.com/python-log-output
LOG_LEVEL_FILE = 'DEBUG'
LOG_LEVEL_CONSOLE = 'INFO'

_detail_formatting = '%(asctime)s %(levelname)-8s [%(module)s#%(funcName)s %(lineno)d] %(message)s'

try:
  # Above LOG_LEVEL_FILE_CONSOLE will be output in log file
  logging.basicConfig(
    level=getattr(logging, LOG_LEVEL_FILE),
    format=_detail_formatting,
    filename='./logs/test.log'
  )
except FileNotFoundError:
  logger.info("logs directory not found")

console = logging.StreamHandler()
console.setLevel(getattr(logging, LOG_LEVEL_CONSOLE))
console_formatter = logging.Formatter(_detail_formatting)
console.setFormatter(console_formatter)


# add console handler to main logger
logger.addHandler(console)

# do the same to timi logger
logging.getLogger("timi").addHandler(console)

def main(argv):
  logger.info("main starts")

  # read_input
  inputdata = ''

  # read_temporary_excel
  outputfile = ''

  # change action by the options
  try:
    opts, args = getopt.getopt(argv,"hi:o:", ["inputdata=","outputfile="])
  except getopt.GetoptError:
    print('main.py -i <inputdata> -o <outputfile>')
    sys.exit(2)

  for opt, arg in opts:
    if opt == '-h':
      print('main.py -i <inputdata> -o <outputfile>')
      sys.exit()
    elif opt in ("-i", "--inputdata"):
      inputdata = arg
    elif opt in ("-o", "--outputfile"):
      outputfile = arg

  data = json.loads(inputdata)

  timi_instance = timemanager(data, outputfile)
  timi_instance.edit_timecard()

  logger.info("main ends")

if __name__ == "__main__":
  main(sys.argv[1:])
