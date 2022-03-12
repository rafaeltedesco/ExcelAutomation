import xlwings as xw
from threading import Thread
import random
import pythoncom
from datetime import datetime
import os

threads = []
cpus = os.cpu_count()

def record_xlsx():
  pythoncom.CoInitialize()
  with xw.App(visible=False) as xl:
    workbooks = xl.books
    current_wb = workbooks[0]
    sheets = current_wb.sheets
    current_sheet = sheets[0]
    current_sheet["A1"].value = random.randint(1, 100)
    current_wb.save(os.path.join(
      'xlsx', f'{datetime.now().strftime("%d%m%Y_%H%M%S")}_2022.xlsx'))

if __name__ == '__main__':

  for cpu in range(cpus):
    threads.append(Thread(target=record_xlsx))

  for thread in threads:
    thread.start()

  for thread in threads:
    thread.join()

  print('Finalizado')



  