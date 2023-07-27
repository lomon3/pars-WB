import pandas as pd
import requests
import math

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

def get_data(url):
    response = requests.get(url)
    if response.status_code == 200:
        return response.json()
    else:
        return None

def parse_data(data):
    if data:
        flattened_data = {}
        for key, value in data.items():
            if key == 'grouped_options':
                for group in value:
                    for option in group['options']:
                        flattened_data[option['name']] = option['value']
            elif key == 'options':
                for option in value:
                    flattened_data[option['name']] = option['value']
            elif isinstance(value, list):
                flattened_data[key] = ';'.join(map(str, value))
            elif isinstance(value, dict):
                for sub_key, sub_value in value.items():
                    flattened_data[f'{key}_{sub_key}'] = sub_value
            else:
                flattened_data[key] = value
        return pd.Series(flattened_data)
    else:
        return pd.Series()

df = pd.read_excel('cards.xlsx')

df['url'] = df['Артикул'].apply(get_url)
df['data'] = df['url'].apply(get_data)
df = df.join(df['data'].apply(parse_data))

# Parse data_skus into individual SKU columns
for i in range(10):
    df[f'SKU{i+1}'] = df['data_skus'].apply(lambda x: x[i] if len(x) > i else None)

df.to_excel('cards_updated.xlsx', index=False)