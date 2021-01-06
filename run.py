import os
from utils.data_fetch import gen_data
from utils.common import TestResult
from utils.excel_charts import store_excel_charts
from utils.gen_excel import gen_excel
from utils.gen_ppt import gen_ppt
# from utils.gen_ppt import gen_ppt
# Fetch test data
result_list = gen_data('./data/')

# Generate the excel
summary = gen_excel(result_list)

# Save Charts
store_excel_charts(summary)

# Generate the PPT Report
gen_ppt(summary)