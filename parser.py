# Author
# Daniel Varab
# djam@itu.dk
# 26/03 - 2014

import xlrd, json, os, re, sys
from collections import OrderedDict

def find_xls_files(folder_path):
  files = [f for f in os.listdir(folder_path) if re.match(r'[A-Za-z]*.*.xls', f)]
  return files

def xls_to_json(path):
  # Open the workbook and select the first worksheet
  wb = xlrd.open_workbook(path)
  sh = wb.sheet_by_index(0)
   
  # List to hold dictionaries
  transaction_list = []
   
  # Iterate through each row in worksheet and fetch values into dict
  # Notice range is hardcoded to 6 where the spreadsheet for iZettle starts
  for rownum in range(6, sh.nrows):

      transaction = OrderedDict()
      row_values = sh.row_values(rownum)

      #Skip digital fees
      if "Digital Fee" in row_values[4]: continue
      if "iZettle Fee" in row_values[4]: continue

      transaction['date']     = row_values[0]
      transaction['time']     = row_values[1]
      transaction['product']  = row_values[4]
      transaction['variant']  = row_values[5]
      transaction['price']    = row_values[6]

      transaction_list.append(transaction)

  parsed_xls = json.dumps(transaction_list)

  return parsed_xls

""" MAIN EXECUTION SEGMENT BELOW """

# Find all xls files in dir to be parsed
xls = find_xls_files(".")

# Serialize and print to terminal
for filepath in xls:
  if(len(sys.argv) == 2):
    # If a second argument is given to the script, it dumps the output in a file by that name
    with open(sys.argv[1], 'w') as f:
      f.write(xls_to_json(filepath))
  else:
    # Writes output to terminal
    print(xls_to_json(filepath))
