from pygismeteo import Gismeteo
import pandas as pd
import tkinter as tk
import openpyxl
import xlsxwriter
import datetime


def get_weather(city):
    gismeteo = Gismeteo()
    search_results = gismeteo.search.by_query(city)
    city_id = search_results[0].id
    current = gismeteo.current.by_id(city_id)
    temperature = current.temperature.air.c
    condition = current.description.full
    return temperature, condition


def update_weather():
    city = city_entry.get()
    temperature, condition = get_weather(city)
    current_datetime = datetime.datetime.now()
    formatted_datetime = current_datetime.strftime("%Y-%m-%d %H:%M:%S")
    temperature_label.config(text=f"Temperature: {temperature}°C")
    condition_label.config(text=f"Condition: {condition}")
    current_data.config(text=f"Current date and time: {formatted_datetime}")

# Add weather data in Excel table
def add_in_excel():
    city = city_entry.get()
    temperature, condition = get_weather(city)
    current_datetime = datetime.datetime.now()
    formatted_datetime = current_datetime.strftime("%Y-%m-%d %H:%M:%S")

    filename = 'weather.xlsx'
    sheetname = 'weather1'
    df = pd.read_excel(filename, sheet_name=sheetname)
    new_row = pd.DataFrame([[city, temperature, condition, current_datetime]],
                  index=[city], columns=['city', 'temperature', 'condition', 'data'])
    df = df.append(new_row, ignore_index=False)
    df.to_excel(filename, sheet_name=sheetname, index=False)
    add_chart(filename, sheetname)
    save_excel.config(text=f"Successfull saved Excel-file")

def add_chart(filename, sheetname):
    # Чтение данных из файла Excel в объект DataFrame
    df = pd.read_excel(filename)

    # Создание диаграммы столбцов с метками ячеек Excel
    ax = df.plot.bar(x='city', y='temperature')
    for i, v in enumerate(df['temperature']):
        ax.text(i, v + 1, str(v), ha='center')

    # Создание нового листа Excel с таблицей данных и диаграммой
    writer = pd.ExcelWriter(filename, engine='xlsxwriter')
    df.to_excel(writer, sheet_name=sheetname, index=False)
    workbook  = writer.book
    worksheet = writer.sheets[sheetname]

    # Добавление данных диаграммы в объект chart
    chart = workbook.add_chart({'type': 'column'})
    chart.add_series({
        'name': 'Temperature',
        'categories': '=weather1!$A$2:$A$10',
        'values': '=weather1!$B$2:$B$10',
    })

    # Вставка диаграммы в лист Excel
    worksheet.insert_chart('G2', chart)
    writer.save()


# Create the main window
window = tk.Tk()
window.title("Weather App")

# Create a label for the city entry field
city_label = tk.Label(window, text="Enter city name:")
city_label.pack()

# Create an entry field for the city name
city_entry = tk.Entry(window)
city_entry.pack()

# Create a button to update the weather
update_button = tk.Button(
    window, text="Update Weather", command=update_weather)
update_button.pack()

excel_button = tk.Button(window, text="Add in Excel", command=add_in_excel)
excel_button.pack()

# Create a label to display the temperature
temperature_label = tk.Label(window, text="Temperature: ")
temperature_label.pack()

# Create a label to display the weather condition
condition_label = tk.Label(window, text="Condition: ")
condition_label.pack()

current_data = tk.Label(window, text="Current date and time: ")
current_data.pack()


save_excel = tk.Label(window, text="Excel-file is not saved")
save_excel.pack()

# Start the main event loop
window.mainloop()
