import os
import xlwings as xw

from typing import List, Tuple
from datetime import datetime
from string import ascii_uppercase

COLOR_RED = (255, 0, 0,)
COLOR_WHITE = (255, 255, 255, )
COLOR_BLACK = (0, 0, 0, )
COLOR_BLUE = (0, 0, 255, )

def format_header(cells: List[str], values: List[str], 
  foreground_color: Tuple[(int, int, int)] = COLOR_WHITE, 
  background_color: Tuple[(int, int, int)] = COLOR_BLACK): 
  header_info = zip(cells, values)
  for info in header_info:
    current_sheet[info[0]].value = info[1]
    current_sheet[info[0]].font.bold = True
    current_sheet[info[0]].color = background_color
    current_sheet[info[0]].font.color = foreground_color

def fill_row(description: str, budget: float, row: int):
    values = list([datetime.now(), description, budget])
    for column in range(len(columns)):
      col = ascii_uppercase[column]
      current_sheet[f'{col}{row}'].value = values[column]
    if budget < 0:
      current_sheet[f"C{row}"].font.color = COLOR_RED


with xw.App(visible=False) as xl:
  workbooks = xl.books

  current_wb = workbooks[0]

  sheets = current_wb.sheets
  current_sheet = sheets[0]

  columns = ["A1", "B1", "C1"]
  values = ["Data", "Descrição", "Valor"]
  format_header(columns, values)

  rows_values = [
    ("Compra de carne para o churrasco", -150),
    ("Bonificação hora extra", 70),
    ("Gratificação de aniversário", 10),
    ("Compra de barra de chocolate", -7.5)
  ]

  for idx, value in enumerate(rows_values):
    fill_row(*value, idx + 2)

  current_wb.save(os.path.join('xlsx', 'caixa.xlsx'))

  