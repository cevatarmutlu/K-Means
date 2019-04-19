import math
import xlrd
import random
import numpy
from xlwt import Workbook


def euclidean(process_row, center):
    return math.sqrt(sum([pow(x[0] - x[1], 2) for x in zip(process_row, center)]))


class KMeans:

    def __init__(self, num_k: int, excel_path, excel_sheet_index: int):
        self.k = num_k
        self.path = excel_path
        self.sheet_index = excel_sheet_index
        self.excel = self.__open_excel()
        self.cluster_array = self.__divide()
        self.center_array = self.__create_array()
        self.cluster_check_array = []

    def __open_excel(self):
        wb = xlrd.open_workbook(self.path)
        sheet = wb.sheet_by_index(self.sheet_index)
        return sheet

    def __create_array(self):
        array = []
        for i in range(self.k):
            array.append([])
        return array

    def __divide(self):
        index_array = self.__create_array()

        for row in range(self.excel.nrows):
            random_val = random.randint(0, self.k - 1)
            index_array[random_val].append(row)
        return index_array

    def print_cluster_array(self):
        for i in range(len(self.cluster_array)):
            print(self.cluster_array[i])

    def calculate_centers(self):
        center_array = []
        for cluster in range(len(self.cluster_array)):
            center_calculate_array = [0] * self.excel.ncols
            for cluster_index in range(len(self.cluster_array[cluster])):
                str_to_float_array = list(map(float, self.excel.row_values(self.cluster_array[cluster][cluster_index])))
                center_calculate_array = [sum(x) for x in zip(center_calculate_array, str_to_float_array)]
            center_calculate_array = [x / len(self.cluster_array[cluster]) for x in center_calculate_array]  # or use numpy
            center_array.append(center_calculate_array)
        self.center_array = center_array

    def cluster_index_change(self, cluster, cluster_index):
            process_row = list(map(float, self.excel.row_values(self.cluster_array[cluster][cluster_index])))
            center_distance_array = []
            for center_index in range(len(self.center_array)):
                center_distance_array.append(euclidean(process_row, self.center_array[center_index]))
            min_center_index = numpy.argmin(center_distance_array)
            if min_center_index != cluster:
                self.cluster_array[min_center_index].append(self.cluster_array[cluster].pop(cluster_index))
                return True
            else:
                return False

    def print_cluster(self):
        for i in range(len(self.cluster_array)):
            print(f'Cluster {i + 1}')
            for j in range(len(self.cluster_array[i])):
                print(f'Row: {self.excel.row_values(self.cluster_array[i][j])}, Row İndex {self.cluster_array[i][j]}')

    def calculate(self):
        for cluster in range(len(self.cluster_array)):
            row = 0
            while True:
                while True:
                    self.calculate_centers()
                    if row == len(self.cluster_array[cluster]):
                        break
                    if self.cluster_index_change(cluster, row):
                        pass
                    else:
                        break

                if row == len(self.cluster_array[cluster]):
                    break
                row += 1

        print("Completed", end="\n\n\n")
        for i in range(len(self.cluster_array)):
            self.cluster_array[i] = sorted(self.cluster_array[i])

    def excel_write(self):
        wb = Workbook()
        sheet = wb.add_sheet('result')
        a = 0
        for i in range(len(self.cluster_array)):
            for j in range(len(self.cluster_array[i])):
                data = self.excel.row_values(self.cluster_array[i][j]) + [i + 1, self.cluster_array[i][j]]
                for cell in range(len(data)):
                    sheet.write(a, cell, data[cell])
                a += 1
        wb.save('result.xlsx')


k = KMeans(3, "../iris.xlsx", 0)
k.calculate()
k.print_cluster()
k.excel_write()



