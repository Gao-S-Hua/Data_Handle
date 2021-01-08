import win32com.client
import sys
from PIL import ImageGrab
import os
from common import get_path


SRC_FILE = get_path("\\outputs\\Result.xlsx")
IMAGE_PATH = get_path("\\outputs\\img\\")

def store_excel_charts(summary):
  # os.mkdir(IMAGE_PATH)
  o = win32com.client.Dispatch("Excel.Application")
  o.Visible = 0
  o.DisplayAlerts = 0
  wb = o.Workbooks.Open(SRC_FILE)
  try:
    page_number = wb.Sheets.Count
    for i in range(1, page_number + 1):
      sheet = o.Sheets(i)
      for n, chart in enumerate(sheet.Shapes):
        chart.Copy()
        image = ImageGrab.grabclipboard()
        image.save(IMAGE_PATH + str(i) +'.png', 'png')
        pass
      pass
  except:
    print("Unexpected error:", sys.exc_info()[0])
    print('** Error: Chart to Image Failed')
  wb.Close(True)
  o.Quit()
  print("| Excel Charts have been exported")

if __name__ == "__main__":
  store_excel_charts()