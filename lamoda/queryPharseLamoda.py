import asyncio
import aiohttp
import pandas as pd
from bs4 import BeautifulSoup
import re
import json

# Функция для очистки и преобразования JavaScript-объекта в валидный JSON
def clean_json(js_object_str):
    # Добавляем кавычки вокруг ключей без кавычек
    fixed_json = re.sub(r'([{,])(\s*)([A-Za-z0-9_]+)(\s*):', r'\1"\3":', js_object_str)
    # Заменяем одинарные кавычки на двойные
    fixed_json = fixed_json.replace("'", '"')
    # Удаляем лишние запятые перед закрывающими скобками
    fixed_json = re.sub(r',\s*}', '}', fixed_json)
    # Удаляем лишние запятые перед закрывающими квадратными скобками
    fixed_json = re.sub(r',\s*\]', ']', fixed_json)
    return fixed_json

# Асинхронная функция для выполнения GET-запроса
async def fetch(session, url):
    async with session.get(url) as response:
        return await response.text()

# Функция для извлечения данных о продуктах
def extract_data(html):
    soup = BeautifulSoup(html, 'html.parser')
    script_text = soup.find('script', string=re.compile('__NUXT__'))
    if not script_text:
        print("Скрипт с данными не найден.")
        return []

    json_text = re.search(r'var __NUXT__ = (\{.*?\});', script_text.string, re.DOTALL).group(1)
    cleaned_json_text = clean_json(json_text)

    try:
        data = json.loads(cleaned_json_text)
    except json.JSONDecodeError as e:
        print(f"Ошибка декодирования JSON: {e}")
        return []

    products_data = data['payload']['state']['payload']['products']
    if not products_data:
        print("Продукты не найдены.")
        return []

    extracted_data = []
    for product in products_data:

        product_info = {
            'sku': product.get('sku', ''),
            'rating': product.get('rating', {}).get('average_rating', ''),
            'reviews_count': product.get('rating', {}).get('reviews_count', ''),
            'old_price_amount': product.get('old_price_amount', ''),
            'price_amount': product.get('price_amount', ''),
            'name': product.get('name', ''),
            'seasons': ', '.join(product.get('seasons', {}).values()),
            'image_url': f"https://a.lmcdn.ru/img600x866{product.get('gallery', [''])[0]}"
        }
        extracted_data.append(product_info)

    return extracted_data

# Основная асинхронная функция для сбора данных
async def scrape_data(query):
    base_url = "https://www.lamoda.ru/catalogsearch/result/"
    all_products = []

    async with aiohttp.ClientSession() as session:
        page = 1
        while True:
            url = f"{base_url}?q={query}&page={page}"
            print(f"Обрабатывается URL: {url}")
            html = await fetch(session, url)
            products = extract_data(html)
            
            if not products:
                break

            all_products.extend(products)
            print(f"Найдено продуктов на странице {page}: {len(products)}")
            page += 1

    return all_products

# Запуск скрипта
query = input("Введите поисковый запрос: ")
data = asyncio.run(scrape_data(query))
print("Made by https://t.me/ArChernushevich")
input("Нажмите Enter, чтобы выйти...")

# Проверка наличия данных перед сохранением
if data:
    # Сохранение результатов в Excel
    df = pd.DataFrame(data)
    df.to_excel('lamoda_products.xlsx', index=False)
    print("Данные сохранены в файл 'lamoda_products.xlsx'")
else:
    print("Данные не получены.")