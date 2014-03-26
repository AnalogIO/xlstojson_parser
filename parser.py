# Author
# Daniel Varab
# djam@itu.dk
# 26/03 - 2014

import xlrd
from collections import OrderedDict
import json, os, re

def find_xls_files(folder_path):
  files = [f for f in os.listdir('.') if re.match(r'[A-Za-z]*.*.xls', f)]
  return files

def xls_to_json(path):
  # Open the workbook and select the first worksheet
  wb = xlrd.open_workbook('test_file.xls')
  sh = wb.sheet_by_index(0)
   
  # List to hold dictionaries
  transaction_list = []
   
  # Iterate through each row in worksheet and fetch values into dict
  # Notice range is hardcoded to 6 where the spreadsheet starts
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

      """
      print("date " + row_values[0])
      print("time " + row_values[1])
      print("product " + row_values[4])
      print("variant " + row_values[5])
      
      print(" ")

      """

      transaction_list.append(transaction)
   
  # Serialize the list of dicts to JSON
  #j = json.dumps(cars_list)

  parsed_xls = json.dumps(transaction_list)

  # Write to file
  """
  with open('data.json', 'w') as f:
      f.write(parsed_xls)
  """

  return parsed_xls

xls = find_xls_files(".")

for filepath in xls:
  print(xls_to_json(filepath))