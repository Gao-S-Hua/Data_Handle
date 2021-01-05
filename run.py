import os
from utils.data_fetch import gen_data
from utils.common import TestResult
from utils.gen_excel import gen_excel

# Fetch test data
result_list = gen_data('./data/')

# Generate the excel
gen_excel(result_list)