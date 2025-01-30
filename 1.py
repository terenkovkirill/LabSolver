import openpyxl
wb = openpyxl.reader.excel.load_workbook(filename = "Книга3.xlsx", data_only = True)
wb.active = 0
sheet = wb.active

import numpy as np
import matplotlib.pyplot as plt

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
k = np.mean(values_YX) / np.mean(values_XX)                                                                  # k = <YX> / <X^2>
print(f'k = {k}     b = 0')

#==========================================================
# #МНК для y = kx + b
#
# k = (mean(values_YX) - mean(column_A) * mean(column_B)) / (mean(values_XX) - mean(column_B)**2)          # k = (<XY> - <Y><X>) / (<X^2> - <X>^2)
# b = mean(column_A) - k * mean(column_B)                                                                  # b = <Y> - k * <X>
# print(f'k = {k}        b = {b}')

#==========================================================
# Случайная погрешность (σ_k, σ_b) МНК для y = kx + b

N = stop_param - 3    #TODO: изменить определение N                                                         # N - количество измерений

σ_k = np.sqrt((np.mean(values_YY) / np.mean(values_XX) - k**2) / (N - 2))                                         # σk = sqrt((<Y^2> / <X^2> - k^2) / (N - 2))
σ_b = σ_k * np.sqrt(np.mean(values_XX))                                                                        # σb = σk * sqrt(<X^2>)
print('σ_k = ', σ_k, 'σ_b = ', σ_b)


#==========================================================
##Погрешности
#Систематическая погрешность
σx_syst = 0                  #TODO: добавить σx_syst

#Случайная погрешность для точек графика

std_dev = np.std(column_B, ddof = 1)                                                                     #стандартное отклонение
σx_ind = std_dev                                                                                         #погрешность тодельных точек по X (то, что на графике)
σx_rand = σx_ind / np.sqrt(N)                                                                            #погрешность средняя (для расчётов)
σ_x = np.sqrt(σx_rand**2 + σx_syst**2)                                                                   #суммарная погрешность

print(f'σx_rand = {σx_rand}     σx_syst = {σx_syst}       σ_x = {σ_x}')


θ = 0.001                   #TODO: добавить множитель производной
σy_syst = 0                 #TODO: добавить σy_syst
σy_ind = θ * σx_ind                                                                                      #погрешность тодельных точек по Y (то, что на графике)
σy_rand = σy_ind / np.sqrt(N)                                                                            #погрешность средняя (для расчётов)
σ_y = np.sqrt(σy_rand**2 + σy_syst**2)
print(f'σy_rand = {σy_rand}     σy_syst = {σy_syst}       σ_y = {σ_y}')


#Для y = Ax^n
#σ_y = n * y * σ_x / x
#ε_y = n * ε_x

#Для z = x^n * y^n
#σ_z = sqrt((n*z*σ_x / x)**2 + (m*z*σ_y / y)**2)
#ε_z = sqrt(n**2 * ε_x**2 + m**2 * ε_y**2)

#===========================================================
## График с МНК, линейная регрессия

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
plt.errorbar(x, y, xerr = σx_ind, yerr = σy_ind, fmt = 'o', color = 'blue', label = 'Данные с погрешностями X')
plt.plot(x, y_pred, color = 'red', label = 'Линейная регрессия')
plt.xlabel('Размерность по x')                                  #TODO: сменить размерность
plt.ylabel('Размерность по y')
plt.title('График y(x) по МНК')                                 #TODO: заменить y(x)
plt.legend()
plt.grid()                                                                              # отвечает за сетку на графике
plt.show()

for i in range(1, stop_param):
    print(sheet['A' + str(i)].value, sheet['B' + str(i)].value, sheet['C' + str(i)].value, sheet['D' + str(i)].value, sep = '          ')

#TODO: улучшить график (разобраться в функциях, подписи, несколько прямых, цвет и толщина шрифта)