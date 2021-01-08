import os
from utils.data_fetch import gen_data
from utils.common import TestResult
from utils.store_excel_charts import store_excel_charts
from utils.gen_excel import gen_excel
from utils.gen_ppt import gen_ppt
from utils.prepare import prepare

# STEP 1. Program Initialization
prepare()

# STEP 2. Fetch test data
result_list = gen_data('./data/')

# STEP 3. Generate the excel
summary = gen_excel(result_list)

# STEP 4. Save Charts
store_excel_charts(summary)

# STEP 5. Generate the PPT Report
gen_ppt(summary)