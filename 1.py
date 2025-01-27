import openpyxl
wb = openpyxl.reader.excel.load_workbook(filename = "Книга3.xlsx", data_only = True)
wb.active = 0
sheet = wb.active

from statistics import mean
from numpy import sqrt


stop_param, i = 1, 1
while sheet['A' + str(i)].value != None:
    stop_param += 1
    i += 1
print(stop_param)

column_A, column_B = [], []                     #column_A - values of the Y,    column_B - X
for i in range(3, stop_param):
    column_A.append(sheet['A' + str(i)].value)
    column_B.append(sheet['B' + str(i)].value)

values_YX, values_XX, values_YY = [], [], []
for i in range(stop_param - 3):
    values_YX.append(column_A[i] * column_B[i])
    values_XX.append(column_B[i] * column_B[i])
    values_YY.append(column_A[i] * column_A[i])

#==========================================================
# МНК для y = kx

# print('The average value for column_A = ', mean(column_A))
# print('The average value for column_B = ', mean(column_B))
# print('The average value for values_YX = ', mean(values_YX))
# print('The average value for values_XX = ', mean(values_XX))
k = mean(values_YX) / mean(values_XX)                                                                  # k = <YX> / <X^2>
print('k = ', k)

#==========================================================
# #МНК для y = kx + b
#
# k = (mean(values_YX) - mean(column_A) * mean(column_B)) / (mean(values_XX) - mean(column_B)**2)          # k = (<XY> - <Y><X>) / (<X^2> - <X>^2)
# b = mean(column_A) - k * mean(column_B)                                                                  # b = <Y> - k * <X>
# print('k = ', k)
# print('b = ', b)

#==========================================================
# Случайная погрешность МНК для y = kx + b

N = stop_param                                                                                          # N - количество измерений
print((mean(values_YY) / mean(values_XX) - k**2) / (N - 2))
σk = sqrt((mean(values_YY) / mean(values_XX) - k**2) / (N - 2))                                         # σk = sqrt((<Y^2> / <X^2> - k^2) / (N - 2))
σb = σk * sqrt(mean(values_XX))                                                                         # σb = σk * sqrt(<X^2>)
print('σk = ', σk, 'σb = ', σb)

#===========================================================
# График с МНК, линейная регрессия

import numpy as np
import matplotlib.pyplot as plt

# Данные
x = np.array(column_B)
y = np.array(column_A)

# Подбор коэффициентов линейной регрессии
coefficients = np.polyfit(x, y, 1)

slope = coefficients[0]
intercept = coefficients[1]
print(f"Наклон (slope): {slope}")
print(f"Смещение (intercept): {intercept}")

# Построение линии регрессии
y_pred = slope * x + intercept

# Визуализация
plt.scatter(x, y, color = 'blue', label = 'Данные')
plt.plot(x, y_pred, color = 'red', label = 'Линейная регрессия')
plt.xlabel('Размерность по x')
plt.ylabel('Размерность по y')
plt.title('График методом наименьших квадратов')
plt.legend()
plt.grid()                                                                  # отвечает за сетку на графике
plt.show()

for i in range(1, stop_param):
    print(sheet['A' + str(i)].value, sheet['B' + str(i)].value, sheet['C' + str(i)].value, sheet['D' + str(i)].value, sep = '          ')

