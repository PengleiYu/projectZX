from random import gauss
from statistics import mean, stdev
import xlrd

arr = [43, 48, 43, 42, 47, 41, 46, 43, 45, 45, 41, 40, 42, 53, 41, 39]
arr2 = [39, 41, 41, 42, 40, 41, 36, 37, 39, 37, 41, 39, 41, 39, 43, 41]
avg = mean(arr2)
sigma = stdev(arr2)
print(avg, sigma)
