from random import gauss, randrange, randint, shuffle
from random import random
from statistics import mean, stdev
from time import sleep

import xlrd, xlwt
# from xlutils3 import copy
from xlutils import copy

sheet_carbon = xlrd.open_workbook("tanhuaqiangdu.xls").sheet_by_index(0)
sheet_modify = xlrd.open_workbook("jiaoduandmian.xls").sheet_by_index(0)
book = xlrd.open_workbook("output.xlsx")
# sheet_output=book.sheet_by_index(0)
wb = copy.copy(book)
write_sheet = wb.get_sheet(0)


class Forward(object):
    def __init__(self):
        self.index_avg = 5
        self.index_face = 2
        self.carbon = 6
        self.index_carbon = self.carbon * 2 + 1

    def write_to_excel(self, arr_arr):
        for i in range(len(arr_arr)):
            self.compute_strength(arr_arr[i], index=i)

    def compute_strength(self, arr, index=None):
        able = index is not None
        ss = "%.1f"
        if able:
            for i in range(len(arr)):
                write_sheet.write(5 + index, i + 1, arr[i])

        # print("Forward.compute_strength:arr=", arr)
        avg = mean(arr)  # 该行原始数据的平均值
        if able:
            write_sheet.write(5 + index, 18, ss % avg)
            write_sheet.write(20 + index, 1, ss % avg)
        # sheet_output.cell(5 + index, 18)
        #     sheet_output.cell(20 + index, 1)
        col0 = sheet_modify.col_values(0)
        n_row_arg = col0.index(int(avg + 0.5))
        n_col_arg = self.index_avg

        modify_arg = sheet_modify.cell_value(n_row_arg, n_col_arg)
        if able:
            write_sheet.write(20 + index, 2, ss % modify_arg)

        avg += modify_arg  # 加上角度修正值
        if able:
            write_sheet.write(20 + index, 3, ss % avg)

        col10 = sheet_modify.col_values(10)
        n_row_face = col10.index(int(avg + 0.5))
        n_col_face = 10 + self.index_face
        modify_face = sheet_modify.cell_value(n_row_face, n_col_face)
        if able:
            write_sheet.write(20 + index, 4, ss % modify_face)

        avg += modify_face  # 加上浇筑面修正值
        if able:
            write_sheet.write(20 + index, 5, ss % avg)

        n_col_carbon = self.index_carbon
        n_row_carbon = sheet_carbon.col_values(0).index(int(avg + 0.5))
        if able:
            write_sheet.write(20 + index, 6, ss % self.carbon)

        avg = sheet_carbon.cell_value(n_row_carbon, n_col_carbon)  # 经过碳化深度换算
        if able:
            write_sheet.write(20 + index, 7, ss % avg)
            write_sheet.write(20 + index, 8, ss % (avg + 6))
            write_sheet.write(5 + index, 19, ss % (avg + 6))
        # print("result avg=", avg)
        return avg

    @staticmethod
    def print_computed_arr(arr):
        # print("arr which to computed = ", arr)
        val_avg = mean(arr)
        val_stddev = stdev(arr)
        val_min = min(arr)
        return val_avg, val_stddev, val_min


class Utils(object):
    def __init__(self, arr1=None, arr2=None):
        self.index_arr_source = arr1
        self.value_arr_source = arr2
        self.s_ave = 33.5
        self.s_sigma = 2.59
        self.s_min = 28.5

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
                index = key if key >= (self.index_arr_source[key] + self.index_arr_source[
                    key - 1]) / 2 else (key - 1)
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
        # print("indexes_carbon:", indexes_carbon)

        self.value_arr_source = col0
        values_carbon = list(map(self.value_of_arr, indexes_carbon))
        # print("values_carbon = ", values_carbon)

        self.index_arr_source = arr_to_search
        indexes_face = list(map(self.index_of_arr, values_carbon))

        self.value_arr_source = sheet_modify.col_values(10)
        values_face = list(map(self.value_of_arr, indexes_face))
        # print("values_face = ", values_face)
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

        # 最终结果数组
        array = list()
        # while len(array) == 0 or min(array) < 10.1 or max(array) > 56.4:
        while True:
            array.clear()
            for key in range(s_count):
                array.append(gauss(self.s_ave, self.s_sigma))

            # 限制碳化列表最终值范围
            if min(array) < 10.1 or max(array) > 39.1:
                print("数据范围不正确, 重新生成...")
            else:
                break
        # print("final data = ", array)
        # 原始数据的平均值列表
        avers = Utils().backward(array)  # todo
        # print("original avg data = ", avers)
        # 原始数据，list每个元素是一行数据
        arr = list()
        # while True:
        for i in range(len(avers)):
            #         while True:
            while True:
                # sleep(0.5)
                temp_arr = list()
                for j in range(10):
                    temp_arr.append(gauss(avers[i], 3))
                # print("arr[i]=", avers[i], ",min=", min(temp_arr), ",max=", max(temp_arr))
                temp_arr = list(map(int, temp_arr))
                if min(temp_arr) >= 20 and max(temp_arr) <= 50:
                    arr.append(temp_arr)
                    break
                else:
                    # print(array)
                    # print(avers)
                    print("第", i, "行原始数据", temp_arr, "有误，重新生成")
        return arr

    @staticmethod
    def shuffle_arr(arr_to_add):
        # sheet_write = xlrd.open_workbook("output.xlsx").sheet_by_index(0)
        for i in arr_to_add:
            v_min = min(i)
            v_max = max(i)
            for j in range(3):
                t1 = randint(27, v_min - 1)
                i.append(t1)
                t2 = randint(v_max + 1, 53)
                i.append(t2)
            shuffle(i)


utils = Utils()
forward = Forward()
need = True

# ll = set()
while True:
    arr_list = None
    result = None
    while True:
        arr_list = utils.generate_origin(need)
        #        print("原始数据：", arr_list)
        avg_arr = list(map(forward.compute_strength, arr_list))  # 正向计算结果数组
        result = forward.print_computed_arr(avg_arr)  # 结果数组的均值和方差
        # print("result = ", result)
        if abs(utils.s_min - result[2]) > 0.4:
            print("最小值偏差=", abs(utils.s_min - result[2]), ",太大, 重新生成...")
            v = utils.s_min - result[2]
            # ll.add(v)
            # print(len(ll))
            need = False
        elif abs(utils.s_ave - result[0]) > 1:
            print("均值偏差", abs(utils.s_ave - result[0]), "太大, 重新生成...")
            need = False
        else:
            break
        sleep(1)

    print("预期:均值=%s, 标准差=%s, 最小值=%s" % (utils.s_ave, utils.s_sigma, utils.s_min))
    print("结果:均值=%s, 标准差=%s, 最小值=%s" % (result[0], result[1], result[2]))
    # print(arr_list)
    Utils.shuffle_arr(arr_list)
    # print(arr_list)
    # s = input("任意键继续,退出请按'n':\t")
    # if s == 'n':
    forward.write_to_excel(arr_list)
    wb.save("hello.xls")
    break
    # else:
    #     need = True
