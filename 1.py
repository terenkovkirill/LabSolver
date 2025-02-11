# import openpyxl
# wb = openpyxl.reader.excel.load_workbook(filename = "Книга3.xlsx", data_only = True)
# wb.active = 0
# sheet = wb.active


import numpy as np
import matplotlib.pyplot as plt
from scipy.stats import linregress

# stop_param, i = 1, 1
# while sheet['A' + str(i)].value != None:
#     stop_param += 1
#     i += 1
# print(stop_param)

# Y, X = [], []                     #column_A - values of the Y,    column_B - X
# for i in range(2, stop_param):
#     Y.append(sheet['A' + str(i)].value)
#     X.append(sheet['B' + str(i)].value)

stop_param = 7
Y = np.array([125, 102, 79, 61]) / 39.8
Y2 = np.array([113, 90, 70, 51]) / 40.7
Y3 = np.array([107, 85, 67, 49]) / 41.6
Y4 = np.array([101, 81, 64, 48]) / 42.5
Y5 = np.array([99, 83, 67, 49]) / 43.3
X = np.array([4, 3.5, 3, 2.5])

values_YX, values_XX, values_YY = [], [], []
for i in range(stop_param - 3):
    values_YX.append(Y[i] * X[i])
    values_XX.append(Y[i] * X[i])
    values_YY.append(Y[i] * Y[i])

#==========================================================
# МНК для y = kx

# print('The average value for Y = ', mean(Y))
# print('The average value for X = ', mean(X))
# print('The average value for values_YX = ', mean(values_YX))
# print('The average value for values_XX = ', mean(values_XX))
# k = np.mean(values_YX) / np.mean(values_XX)                                                                  # k = <YX> / <X^2>
# print(f'k = {k}     b = 0')

#==========================================================
#МНК для y = kx + b

k = (np.mean(values_YX) - np.mean(Y) * np.mean(X)) / (np.mean(values_XX) - np.mean(X)**2)          # k = (<XY> - <Y><X>) / (<X^2> - <X>^2)
b = np.mean(Y) - k * np.mean(X)                                                                  # b = <Y> - k * <X>
print(f'k = {k}        b = {b}')

#==========================================================
# Случайная погрешность (σ_k, σ_b) МНК для y = kx + b

N = stop_param - 3    #TODO: изменить определение N                                                         # N - количество измерений

σ_k = np.sqrt((np.mean(values_YY) / np.mean(values_XX) - k**2) / (N - 2))                                      # σk = sqrt((<Y^2> / <X^2> - k^2) / (N - 2))
σ_b = σ_k * np.sqrt(np.mean(values_XX))                                                                        # σb = σk * sqrt(<X^2>)
print('σ_k = ', σ_k, 'σ_b = ', σ_b)


#================================================================================
##Погрешности для лаб с точками, удовлетворяющими нормальному распределению
#Систематическая погрешность
σx_syst = 0                  #TODO: добавить σx_syst

#Случайная погрешность для точек графика

# std_dev = np.std(X, ddof = 1)                                                                            #стандартное отклонение
# σx_ind = std_dev                                                                                         #погрешность отдельных точек по X (то, что на графике)
# σx_rand = σx_ind / np.sqrt(N)                                                                            #погрешность средняя (для расчётов)
# σ_x = np.sqrt(σx_rand**2 + σx_syst**2)                                                                   #суммарная погрешность
#print(f'σx_rand = {σx_rand}     σx_syst = {σx_syst}       σ_x = {σ_x}')

# θ = 0.001                   #TODO: добавить множитель производной
# σy_syst = 0                 #TODO: добавить σy_syst
# σy_ind = θ * σx_ind                                                                                      #погрешность тодельных точек по Y (то, что на графике)
# σy_rand = σy_ind / np.sqrt(N)                                                                            #погрешность средняя (для расчётов)
# σ_y = np.sqrt(σy_rand**2 + σy_syst**2)
# print(f'σy_rand = {σy_rand}     σy_syst = {σy_syst}       σ_y = {σ_y}')

#================================================================================

σx_ind = 0.1
σy_ind = 3

#Для y = Ax^n
#σ_y = n * y * σ_x / x
#ε_y = n * ε_x

#Для z = x^n * y^n
#σ_z = sqrt((n*z*σ_x / x)**2 + (m*z*σ_y / y)**2)
#ε_z = sqrt(n**2 * ε_x**2 + m**2 * ε_y**2)

#===========================================================
## График с МНК, линейная регрессия

fig = plt.figure(figsize = (12*0.9,10*0.9))                     #TODO: настроить figsize
ax = fig.add_subplot(1, 1, 1)
ax.set_title('Зависимость ΔP(ΔT) по МНК')
ax.set_xlabel('ΔP, атм')
ax.set_ylabel('ΔT, °C')
ax.errorbar(X, Y, σx_ind, σy_ind / 39.8, fmt = 'o', label = '15,5 °C', color = 'black')
ax.errorbar(X, Y2, σx_ind, σy_ind / 40.7, fmt = 'o', label = '25,5 °C', color = 'red')
ax.errorbar(X, Y3, σx_ind, σy_ind / 41.6, fmt = 'o', label = '35 °C', color = 'blue')
ax.errorbar(X, Y4, σx_ind, σy_ind / 42.5, fmt = 'o', label = '45 °C', color = 'green')
ax.errorbar(X, Y5, σx_ind, σy_ind / 43.3, fmt = 'o', label = '55 °C', color = 'orange')
ax.legend()
ax.grid(which = 'both', linewidth = 0.5)

x_line = np.linspace(np.min(X), np.max(X))

slope, intercept, _, _, _ = linregress(X, Y)
y_line = slope * x_line + intercept
ax.plot(x_line, y_line, linestyle = '--', color = 'black')

slope2, intercept2, _, _, _ = linregress(X, Y2)
y_line2 = slope2 * x_line + intercept2
ax.plot(x_line, y_line2, linestyle = '--', color = 'red')

slope3, intercept3, _, _, _ = linregress(X, Y3)
y_line3 = slope3 * x_line + intercept3
ax.plot(x_line, y_line3, linestyle = '--', color = 'blue')

slope4, intercept4, _, _, _ = linregress(X, Y4)
y_line4 = slope4 * x_line + intercept4
ax.plot(x_line, y_line4, linestyle = '--', color = 'green')

slope5, intercept5, _, _, _ = linregress(X, Y5)
y_line5 = slope5 * x_line + intercept5
ax.plot(x_line, y_line5, linestyle = '--', color = 'orange')

plt.savefig("plot.png", dpi = 300)
plt.show()

# Данные
# x = np.array(X)
# y = np.array(Y) / 39.8
#
# # Подбор коэффициентов линейной регрессии
# coefficients = np.polyfit(x, y, 1)
#
# slope = coefficients[0]
# intercept = coefficients[1]
# print(f"Наклон (slope): {slope}")
# print(f"Смещение (intercept): {intercept}")
#
# # Построение линии регрессии
# y_pred = slope * x + intercept
#
# # Визуализация
# plt.errorbar(x, y, xerr = σx_ind, yerr = σy_ind, fmt = 'o', color = 'blue', label = 'Данные с погрешностями X')
# plt.plot(x, y_pred, color = 'red', label = 'Линейная регрессия')
# plt.xlabel('ΔP, атм')                                  #TODO: сменить размерность
# plt.ylabel('ΔT, °C')
# plt.title('Зависимость ΔP(ΔT) по МНК')                                 #TODO: заменить y(x)
# plt.legend()
# plt.grid()                                                                              # отвечает за сетку на графике
# plt.show()

# for i in range(1, stop_param):
#     print(sheet['A' + str(i)].value, sheet['B' + str(i)].value, sheet['C' + str(i)].value, sheet['D' + str(i)].value, sep = '          ')

