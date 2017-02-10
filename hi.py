from random import gauss, randrange
from statistics import mean, stdev

import xlrd

sheet_carbon = xlrd.open_workbook("tanhuaqiangdu.xls").sheet_by_index(0)
sheet_modify = xlrd.open_workbook("jiaoduandmian.xls").sheet_by_index(0)


class Forward(object):
    def __init__(self):
        self.index_avg = 5
        self.index_face = 2
        self.index_carbon = 13

    def compute_strength(self, arr):
        print("Forward.compute_strength:arr=", arr)
        avg = mean(arr)
        col0 = sheet_modify.col_values(0)
        n_row_arg = col0.index(int(avg + 0.5))
        n_col_arg = self.index_avg

        modify_arg = sheet_modify.cell_value(n_row_arg, n_col_arg)

        avg += modify_arg

        col10 = sheet_modify.col_values(10)
        n_row_face = col10.index(int(avg + 0.5))
        n_col_face = 10 + self.index_face
        modify_face = sheet_modify.cell_value(n_row_face, n_col_face)

        avg += modify_face

        n_col_carbon = self.index_carbon
        n_row_carbon = sheet_carbon.col_values(0).index(int(avg + 0.5))

        avg = sheet_carbon.cell_value(n_row_carbon, n_col_carbon)
        # print("result avg=", avg)
        return avg

    @staticmethod
    def print_computed_arr(arr):
        val_avg = mean(arr)
        val_stddev = stdev(arr)
        val_min = min(arr)
        return val_avg, val_stddev, val_min


class Utils(object):
    def __init__(self, arr1=None, arr2=None):
        self.index_arr_source = arr1
        self.value_arr_source = arr2
        self.s_ave = 43.3
        self.s_sigma = 4.69
        self.s_min = 35.5

    def index_of_arr(self, val):
        # arr_source = col13
        # arr_target = col0
        # arr_source = list()
        # arr_target = list()

        index = -1
        for key in range(len(self.index_arr_source)):
            try:
                _v = float(self.index_arr_source[key])
            except ValueError:
                continue
            if val == _v:
                index = key
            elif val < _v:
                index = key if key >= (self.index_arr_source[key] + self.index_arr_source[key - 1]) / 2 else (key - 1)
                break
        return index

    def value_of_arr(self, index):
        return self.value_arr_source[index]

    @staticmethod
    def add(arr1, arr2):
        arr = list()
        for key in range(len(arr1)):
            # print(type(arr1[key]), type(arr2[key]))
            # print(arr1[key] + arr2[key])
            arr.append(arr1[key] + arr2[key])
        return arr

    def backward(self, arr):
        col0 = sheet_carbon.col_values(0)
        col13 = sheet_carbon.col_values(13)

        arr_to_search = self.add(sheet_modify.col_values(10), sheet_modify.col_values(12))

        self.index_arr_source = col13
        indexes_carbon = list(map(self.index_of_arr, arr))
        print("indexes_carbon:", indexes_carbon)

        self.value_arr_source = col0
        values_carbon = list(map(self.value_of_arr, indexes_carbon))

        self.index_arr_source = arr_to_search
        indexes_face = list(map(self.index_of_arr, values_carbon))

        self.value_arr_source = sheet_modify.col_values(10)
        values_face = list(map(self.value_of_arr, indexes_face))
        return values_face

    def generate_origin(self, need_input):
        if need_input:
            t = input("请输入均值,若不输入则默认值为%s:\t" % self.s_ave)
            self.s_ave = float(t) if t else self.s_ave
            if self.s_ave < 10.1 or self.s_ave > 56.4:
                input("均值输入范围错误,按任意键退出!")
                exit()

            t = input("请输入标准差,若不输入则默认值为%s:\t" % self.s_sigma)
            self.s_sigma = float(t) if t else self.s_sigma

            t = input("请输入最小值,若不输入则默认值为%s:\t" % self.s_min)
            self.s_min = float(t) if t else self.s_min

        s_count = 10

        array = list()
        # while len(array) == 0 or min(array) < 10.1 or max(array) > 56.4:
        while True:
            array.clear()
            for key in range(s_count):
                array.append(gauss(self.s_ave, self.s_sigma))

            if min(array) < 10.1 or max(array) > 56.4:
                print("数据范围不正确, 重新生成...")
            else:
                break

        avers = Utils().backward(array)
        arr = list()
        # while True:
        for key in avers:
            #         while True:
            temp_arr = list()
            for j in range(10):
                temp_arr.append(gauss(key, 1))
                #             if min(temp_arr) >= 20 and max(temp_arr) <= 50:
                arr.append(temp_arr)
        # break
        #             else:
        #                 print(array)
        #                 print(avers)
        #                 print("生成数据范围有误...")

        return arr


utils = Utils()
forward = Forward()
need = True

ll = set()
while True:
    while True:
        arr_list = utils.generate_origin(need)
        avg_arr = list(map(forward.compute_strength, arr_list))
        result = forward.print_computed_arr(avg_arr)
        if abs(utils.s_min - result[2]) > 0.1:
            print("最小值偏差太大, 重新生成...")
            v = utils.s_min - result[2]
            ll.add(v)
            print(len(ll))
            need = False
        elif abs(utils.s_ave - result[0]) > 0.3:
            print("均值偏差太大, 重新生成...")
            need = False
        else:
            break

    print("预期:均值=%s, 标准差=%s, 最小值=%s" % (utils.s_ave, utils.s_sigma, utils.s_min))
    print("结果:均值=%s, 标准差=%s, 最小值=%s" % (result[0], result[1], result[2]))
    s = input("任意键继续,退出请按'n':\t")
    if s == 'n':
        break
    else:
        need = True
