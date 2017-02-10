from random import gauss
from statistics import mean, stdev
import xlrd

sheet_modify = xlrd.open_workbook("jiaoduandmian.xls").sheet_by_index(0)
sheet_carbon = xlrd.open_workbook("tanhuaqiangdu.xls").sheet_by_index(0)




