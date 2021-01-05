from openpyxl import Workbook
from openpyxl.chart import BarChart, Series, Reference, LineChart

def add_summarize(test_result, wb):
  ws = wb.create_sheet("Summary")
  title = ["Pass", "Fail"]
  ws.append(title)
  pass_count = 0
  fail_count = 0
  for result in test_result:
    for data in result.data:
      if data > 100 :
        pass_count = pass_count + 1
      else:
        fail_count = fail_count + 1
  ws.append([pass_count, fail_count])
  add_counter(ws)
  return wb


def gen_excel(test_result):
  wb = Workbook()
  add_summarize(test_result, wb)
  for result in test_result:
    ws = wb.create_sheet(result.SN + "_" + result.temperature)
    title = ["SN", "Temperature", "Fmax(MHz)"]
    ws.append(title)
    for data in result.data:
      ws.append([result.SN, result.temperature, data])
    add_fmax_plot(ws)
  page_one=wb.get_sheet_by_name('Sheet')
  wb.remove_sheet(page_one)
  wb.save("result.xlsx")


def add_fmax_plot(ws):
  c1 = LineChart()
  c1.title = "PCIE Fmax"
  c1.style = 13
  c1.y_axis.title = 'Frequency (Fmax)'
  c1.x_axis.title = 'Test count'
  data = Reference(ws, min_col=3, min_row=2, max_col=3, max_row=8)
  c1.add_data(data, titles_from_data=False)

  s1 = c1.series[0]
  s1.marker.symbol = "star"
  s1.marker.graphicalProperties.solidFill = "#000000" # Marker filling
  s1.marker.graphicalProperties.line.solidFill = "#000000" # Marker outline
  s1.graphicalProperties.line.noFill = False
  s1.smooth = True
  ws.add_chart(c1, "A10")

def add_counter(ws):
  chart1 = BarChart()
  chart1.type = "col"
  chart1.style = 10
  chart1.title = "PCIE Test Summary"
  chart1.y_axis.title = 'Number'
  chart1.x_axis.title = ''
  data1 = Reference(ws, min_col=1, min_row=1, max_row=2, max_col=1)
  chart1.add_data(data1, titles_from_data=True)
  data2 = Reference(ws, min_col=2, min_row=1, max_row=2, max_col=2)
  chart1.add_data(data2, titles_from_data=True)
  # cats = Reference(ws, min_col=1, min_row=1, max_row=1, max_col=2)
  # chart1.set_categories(cats)
  ws.add_chart(chart1, "A10")
