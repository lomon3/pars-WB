import requests
import os
import pandas as pd
from openpyxl import load_workbook
from time import sleep

def get_data(url, search_strings):
    """ Получение и фильтрация данных по заданным строкам. """
    response = requests.get(url)
    response.raise_for_status()
    data = response.json()

    # Фильтрация данных
    filtered_data = [item for item in data if any(s in str(item) for s in search_strings)]
    print (filtered_data)
    return filtered_data

def update_excel(file_name, new_data):
    """ Обновление Excel файла новыми данными, без дублирования. """
    if not os.path.exists(file_name) or not file_name.endswith('.xlsx'):
        # Если файл не существует или не является Excel-файлом, создаём новый
        pd.DataFrame(new_data).to_excel(file_name, index=False)
        return

    try:
        # Попытка открыть существующий файл
        book = load_workbook(file_name)
        writer = pd.ExcelWriter(file_name, engine='openpyxl')
        writer.book = book
        writer.sheets = {ws.title: ws for ws in book.worksheets}
        existing_data = pd.read_excel(file_name)
    except Exception as e:
        print(f"Ошибка при открытии файла: {e}")
        return

    # Преобразование списков в строковые представления
    new_df = pd.DataFrame(new_data)
    new_df = new_df.applymap(lambda x: ','.join(x) if isinstance(x, list) else x)

    # Добавление новых данных, избегая дублирования
    updated_data = pd.concat([existing_data, new_df]).drop_duplicates().reset_index(drop=True)
    updated_data.to_excel(writer, index=False)

    # Сохранение файла
    writer.save()


def main():
    url = "https://banners-website.wildberries.ru/public/v1/banners?urltype=1024&apptype=2&displaytype=3&longitude=37.60805&latitude=55.775565&country=1&culture=ru"
    file_name = "BrendsInMediaTracker.xlsx"
    
    # Ввод частоты запросов и поисковых строк
    interval = int(input("Введите интервал запросов в секундах: "))
    search_strings = input("Введите строки для поиска, разделенные запятой: ").split(",")

    try:
        while True:
            print("Отправка запроса...")
            data = get_data(url, search_strings)
            if data:
                update_excel(file_name, data)
                print("Файл обновлен.")
            else:
                print("Новых данных для добавления нет.")
            sleep(interval)
    except KeyboardInterrupt:
        print("Выполнение остановлено пользователем.")

if __name__ == "__main__":
    main()