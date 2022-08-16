import openpyxl, os

workbook = openpyxl.load_workbook('D:\Тексты для Жукова.xlsx')
sheet = workbook.active
for i in range(181, 202):
    file = open("D:/Text for DM/" + str(i - 1) + '.txt', 'w')
    for j in range(1, 8):
        file.write(str(sheet[i][j].value))
        if j == 4:
            file.write(', ')
            continue 
        file.write('\n')
    file.close()

