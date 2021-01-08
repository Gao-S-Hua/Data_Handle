import os
import re
from os.path import isfile, join
from common import TestResult

def gen_data(dir):
  result_list = []
  # Get valid file list
  all_files = os.listdir(dir)
  files = [f for f in all_files if isfile(join(dir, f))]
  for file in files:
    [SN, temperature] = file.split('_')
    temperature = temperature.replace('.txt', '')
    f = open(dir + '/' + file, "r")
    data = handle_data(f.read())
    new_result = TestResult(SN, temperature)
    new_result.data = data
    result_list.append(new_result)
  print("| Raw Data has been capture")
  return result_list

def handle_data(text):
  pattern = r'PCIE FMAX (.*) MHz'
  data = re.findall(pattern, text)
  num_data = []
  for single_number in data:
    num_data.append(float(single_number))
  return num_data

