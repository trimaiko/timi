import sys, getopt
from openpyxl import load_workbook

def write_cell(data, cell):
  cell = data

def main(argv):
  print("main starts")
  # read_input
  inputfile = ''
  # read_template_excel
  outputfile = ''
  
  try:
    opts, args = getopt.getopt(argv,"hi:o:", ["inputfile=","outputfile="])
  except getopt.GetoptError:      
    print('main.py -i <inputfile> -o <outputfile>')
    sys.exit(2)
  
  for opt,arg in opts:
    if opt == '-h':
      print('main.py -i <inputfile> -o <outputfile>')
      sys.exit()
    elif opt in ("-i", "--inputfile"):
      inputfile = arg
    elif opt in ("-o", "--ofile"):
      outputfile = arg
  
  wb = load_workbook(filename = outputfile)

  target_sheet = wb['target_sheet']

  target_sheet['A1'] = inputfile

  # write_to_outputfile
  wb.save(filename = outputfile)

  print("main ends")

if __name__ == "__main__":
   main(sys.argv[1:])
