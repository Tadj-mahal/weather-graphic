import pandas as pd
import matplotlib.pyplot as plt
import xlsxwriter

# Чтение данных из файла Excel в объект DataFrame
df = pd.read_excel('file.xlsx')

# Создание диаграммы столбцов с метками ячеек Excel
ax = df.plot.bar(x='Категории', y='Данные')
for i, v in enumerate(df['Данные']):
    ax.text(i, v + 1, str(v), ha='center')

# Создание нового листа Excel с таблицей данных и диаграммой
writer = pd.ExcelWriter('file.xlsx', engine='xlsxwriter')
df.to_excel(writer, sheet_name='Sheet1', index=False)
workbook  = writer.book
worksheet = writer.sheets['Sheet1']

# Добавление данных диаграммы в объект chart
chart = workbook.add_chart({'type': 'column'})
chart.add_series({
    'name': 'Данные',
    'categories': '=Sheet1!$A$2:$A$6',
    'values': '=Sheet1!$B$2:$B$6',
})

# Вставка диаграммы в лист Excel
worksheet.insert_chart('D2', chart)
writer.save()