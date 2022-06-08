from openpyxl import Workbook
import os

wb = Workbook()
dest_filename = 'Piphagor.xlsx'
ws = wb.active
ws.title = 'Piphagor'
sise = 10  # Table size

piphagor = [[0 for j in range(sise)] for i in range(sise)]


for i in range(sise):
    for j in range(sise):
        piphagor[i][j] = (i + 1) * (j + 1)


for i in range(sise):
    for j in range(sise):
        ws.cell(i + 1, j + 1).value = piphagor[i][j]


wb.save(os.path.join(os.getcwd(), dest_filename))
wb.close()
