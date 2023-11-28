import requests
import pandas as pd
import time
import logging
import os

# Настройка логирования
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

# Функция для проверки, открыт ли файл Excel
def is_excel_open(filename):
    try:
        os.rename(filename, filename)
        return False
    except PermissionError:
        return True
    
# Функция для получения данных с API
def get_data():
    try:
        response = requests.get("https://banners-website.wildberries.ru/public/v1/banners?urltype=1024&apptype=2&displaytype=3&longitude=37.60805&latitude=55.775565&country=1&culture=ru")
        response.raise_for_status()
        return response.json()
    except requests.RequestException as e:
        logging.error(f"Ошибка при запросе: {e}")
        return []

# Функция для обработки данных
def process_data(data, keywords):
    processed_data = []
    for item in data:
        if any(keyword in str(item) for keyword in keywords):
            item['Href'] = 'https://www.wildberries.ru' + item.get('Href', '')
            item['Src'] = 'https://static-basket-01.wb.ru/vol1/crm-bnrs' + item.get('Src', '')
            processed_data.append(item)
    return processed_data

# Функция для обновления существующих данных
def update_data(existing_data, new_data):
    for new_item in new_data:
        # Поиск совпадения
        match = None
        for existing_item in existing_data:
            if all(new_item[k] == existing_item[k] for k in new_item if k not in ['Rv', 'PlacementOptions']):
                match = existing_item
                break

        if match:
            # Обновление LocationType и PlacementOptions с новыми уникальными значениями
            if str(new_item['Rv']) not in str(match['Rv']):
                match['Rv'] = f"{match['Rv']}\n{new_item['Rv']}"
#match['Rv'] = f"{match['Rv']}\n{new_item['Rv']}"
            if str(new_item['PlacementOptions']) not in str(match['PlacementOptions']):
                match['PlacementOptions'] = f"{match['PlacementOptions']}\n{new_item['PlacementOptions']}"
        else:
            existing_data.append(new_item)

# Загрузка существующих данных из Excel
try:
    existing_data_df = pd.read_excel('BrendsInMediaTrecker.xlsx')
    existing_data = existing_data_df.to_dict('records')
except FileNotFoundError:
    existing_data_df = pd.DataFrame()
    existing_data = []

# Запрос интервала и ключевых слов у пользователя
interval = int(input("Введите интервал в секундах: "))
keywords = input("Введите ключевые слова через запятую: ").split(',')

try:
    while True:
        new_data = get_data()
        processed_new_data = process_data(new_data, keywords)

        # Обновление существующих данных
        update_data(existing_data, processed_new_data)

        try:
            # Попытка сохранения данных в Excel
            new_df = pd.DataFrame(existing_data)
            new_df.to_excel('BrendsInMediaTrecker.xlsx', index=False)
            logging.info("Файл Excel обновлен.")
        except PermissionError:
            logging.error("Ошибка доступа к файлу Excel. Пожалуйста, закройте файл и нажмите Enter для продолжения...")
            input()
            # Попытка сохранения данных в Excel
            new_df = pd.DataFrame(existing_data)
            new_df.to_excel('BrendsInMediaTrecker.xlsx', index=False)
            logging.info("Файл Excel обновлен.")
        time.sleep(interval)
except KeyboardInterrupt:
    logging.info("Скрипт остановлен пользователем.")

input("byebye")
