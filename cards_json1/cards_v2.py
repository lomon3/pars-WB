import pandas as pd
import asyncio
import aiohttp
import math
from tqdm.asyncio import tqdm

def get_url(article):
    b = math.floor(article / 100000)
    c = math.floor(article / 1000)
    
    if b <= 143:
        url = "https://basket-01.wb.ru/vol"
    elif 144 <= b <= 287:
        url = "https://basket-02.wb.ru/vol"
    elif b <= 431:
        url = "https://basket-03.wb.ru/vol"
    elif b <= 719:
        url = "https://basket-04.wb.ru/vol"
    elif b <= 1007:
        url = "https://basket-05.wb.ru/vol"
    elif b <= 1061:
        url = "https://basket-06.wb.ru/vol"
    elif b <= 1115:
        url = "https://basket-07.wb.ru/vol"
    elif b <= 1169:
        url = "https://basket-08.wb.ru/vol"
    elif b <= 1313:
        url = "https://basket-09.wb.ru/vol"
    elif b <= 1601:
        url = "https://basket-10.wb.ru/vol"
    elif b <= 1655:
        url = "https://basket-11.wb.ru/vol"
    elif b <= 1919:
        url = "https://basket-12.wb.ru/vol"
    else:
        url = "https://basket-13.wb.ru/vol"

    url += str(b) + "/part" + str(c) + "/" + str(article) + "/info/ru/card.json"
    return url

def flatten_data(data, prefix=''):
    """
    Рекурсивно разбирает данные и преобразует их в плоскую структуру.
    """
    flattened = {}

    if isinstance(data, dict):
        for key, value in data.items():
            new_key = f"{prefix}_{key}" if prefix else key
            flattened.update(flatten_data(value, new_key))
    elif isinstance(data, list):
        for i, value in enumerate(data):
            new_key = f"{prefix}_{i}" if prefix else str(i)
            flattened.update(flatten_data(value, new_key))
    else:
        flattened[prefix] = data

    return flattened

async def fetch_and_process(session, url):
    async with session.get(url) as response:
        if response.status == 200:
            data = await response.json()
            flattened_data = flatten_data(data)
            # Convert all values to strings to avoid dtype issues
            return {k: str(v) for k, v in flattened_data.items()}
        else:
            print(f"Error {response.status} for URL: {url}")
            return {}

async def main(df):
    async with aiohttp.ClientSession() as session:
        tasks = [fetch_and_process(session, url) for url in df['url']]
        
        # Используем tqdm для отображения прогресса выполнения асинхронных запросов
        results = []
        for task in tqdm(asyncio.as_completed(tasks), total=len(tasks)):
            results.append(await task)

        # Convert results to DataFrame and concatenate with original df
        results_df = pd.DataFrame(results)
        final_df = pd.concat([df, results_df], axis=1)

        # Parse SKUs, if available
        if 'sizes_table_values_skus' in final_df.columns:
            for i in range(10):
                final_df[f'SKU{i+1}'] = final_df['sizes_table_values_skus'].apply(lambda x: x.split(',')[i] if isinstance(x, str) and len(x.split(',')) > i else None)

        # Удаление столбца "Артикул"
        final_df = final_df.drop(columns=['Артикул'])

        final_df.to_excel('cards_updated.xlsx', index=False)

df = pd.read_excel('cards.xlsx')
df['url'] = df['Артикул'].apply(get_url)

# Run the main coroutine
asyncio.run(main(df))

print("\nMade by https://t.me/ArChernushevich")
input("Нажмите Enter, чтобы выйти...")