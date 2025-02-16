import xlwings as xw

from mod_excel_access import reset_han_ji_cells

wb = xw.apps.active.books.active
reset_han_ji_cells(wb)