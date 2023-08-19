
# pip install pandas
# pip install requests
# pip install pyinstaller

import requests
import openpyxl
import requests

# Получаем значение spp
def get_spp():
    headers = {
        'Cookie': '_wbauid=3828519991678728635; ___wbu=da44508c-eccc-4fd5-be34-e832e30ae01c.1678728636; __wba_s=1; __bsa=basket-ru-10; ncache=CardSize%3AC516x688%26Sort%3APopular; external-locale=ru; ab.storage.deviceId.a9882122-ac6c-486a-bc3b-fab39ef624c5=%7B%22g%22%3A%22835111c7-634f-4e32-c69f-984dba8398e7%22%2C%22c%22%3A1692345058940%2C%22l%22%3A1692345058940%7D; x-supplier-id-external=d99a1f06-4d18-5d70-8ee4-d0fefc8a193f; BasketUID=c0b53287-84da-4dca-86f0-aa6837fb9e8f; ___wbs=6114f5b9-e827-42f8-bcd2-198b27dab89d.1692443225; _wbSes=CfDJ8K3v9xtc6wlIoUObeqvnAyPFOUtkGmctvZmb6di48tWhcv%2BxsjvL4TzX%2Bb5WxEVnhQyyh8559C1Hv9Uk8s52qVBO3SRSrT1hkb%2F%2FenkCxzaoyj2vejd57vVvrbwHVKDK4spq%2BJH37W7N2LDGQuukMTbOYs%2BXECuSHZteKSYU7K1D; WILDAUTHNEW_V3=567A117E732A8A589124C6B29AC386D4813ABF537D03FD74F64153E8F4760E154A3FE5DB325ACD354F1D4B6240233897C548A1A3DB3120797A7D2020F9EA442DE16A20A8A55ADA4345BE37464B61FBC1150A6B5FF8521AD0FCDB71B989FB99D87CAF7232E100E17A2A2BB9AD91043D8CFD05CC2B12B1356C4DDE7FE801B907B2C597875F67E8AFEAD7A057EB35DD88DCE3501AEECD6253DD4ACF53B61F4390BCEFDB7239C4188CE880FFFC643F67EAA81A8162360B19486C658483B23E5DD06857DCD7615A3D95DB1805923E0AC7B75475D1A97895557D7C80ABA69FE2EFCF1BA799139D31B4EBEF6C8C0177A94528503A9FDDA890FF11F9AA52D8476028F93D7F7A847B850BE83DEEBE00C3B50F96A72B20B6DD231883409DBF35D42D1FA0A568529C76F25CB46C0FA5A5C646F8E5E58F37C26D; um=uid%3Dw7TDssOkw7PCu8KwwrPCs8K4wrXCtcK5wrY%253d%3Aproc%3D100%3Aehash%3D52e858bb1161c9fdc2855f8739496599',
    }

    url = 'https://www.wildberries.ru/webapi/personalinfo/extrainfo'

    while True:
        response = requests.post(url, headers=headers)
        if response.status_code == 200:
            personal_discount = response.json()['value']['discountInfo']['personalDiscount']
            return personal_discount
        else:
            print(f"Ошибка запроса: {response.status_code}. Повторная попытка...")


spp = input("Введите СПП или оставьте пустым для автоматической загрузки СПП: ") or get_spp()
print(f'Сегодня СПП = {spp}%')

# Открыть книгу Excel и получить листы
wb = openpyxl.load_workbook('Отчет по скидкам.xlsx')
old_sheet = wb['Общий отчет']
sheet = wb.create_sheet('pars')

# Пройти по каждой строке на исходном листе
for row in old_sheet.iter_rows(min_row=1, max_col=5, max_row=old_sheet.max_row):
    # Проверить значение ячейки в первом столбце
    if row[0].value in ["WiMi", "BASEUS", "Бренд"]:
        # Создать новый объект строки и заполнить его значениями из соответствующих ячеек
        new_row = []
        for cell in row:
            new_row.append(cell.value)
        # Добавить строку на новый лист
        sheet.append(new_row)

# Сохранить книгу Excel
wb.save('Отчет по скидкам.xlsx')
print("Filtered brands to new sheet")

# Добавляем новые столбцы
sheet.cell(row=1, column=1, value="Арт ВБ")
sheet.cell(row=1, column=2, value="Название")
sheet.cell(row=1, column=3, value="Цена до скидок")
sheet.cell(row=1, column=4, value="Скидка поставщика")
sheet.cell(row=1, column=5, value="Цена после скидки поставщика")
sheet.cell(row=1, column=6, value="СПП")
sheet.cell(row=1, column=7, value="Цена после СПП")
sheet.cell(row=1, column=8, value="сток")
sheet.cell(row=1, column=9, value="рейт")
sheet.cell(row=1, column=10, value="отзывы")

# группируем список
chunk_size = 300
art_list = []
for row in range(2, sheet.max_row + 1):
    art = str(sheet.cell(row=row, column=5).value)
    art_list.append(art)
    article_chunks = [art_list[i:i+chunk_size]
                      for i in range(0, len(art_list), chunk_size)]

row = 2
for chunk in article_chunks:
    url = f"https://card.wb.ru/cards/detail?spp={spp}&reg=1&locale=ru&dest=-1216601,-115136,-421732,123585595&nm={';'.join(chunk)}"
    # print(url)

    # Отправляем запрос и получаем ответ в формате JSON
    try:
        response = requests.get(url)
        data = response.json()
    except Exception as e:
        print(f"Failed to get data for art= {chunk}: {e}")
        continue
    products = data["data"]["products"]
    for product in products:
        if product:

            sku = product["id"]
            name = product["name"]

            price_u = product.get("priceU")
            if price_u is not None:
                price_u = price_u / 100

            sizes = product.get("sizes", [])
            qty = sum(stock["qty"]
                      for size in sizes for stock in size.get("stocks", []))

            rating = product["rating"]
            feedbacks = product["feedbacks"]

            extended = product.get("extended")
            if extended is not None:
                basic_sale = extended.get("basicSale", 0)

                basic_price_u = extended.get("basicPriceU")
                if basic_price_u is not None:
                    basic_price_u = basic_price_u / 100
                else:
                    basic_price_u = price_u

                client_sale = extended["clientSale"]
                client_price_u = extended["clientPriceU"]/100
            else:
                basic_sale = ""
                basic_price_u = ""
                client_sale = ""
                client_price_u = ""

        else:
            sku = ""
            name = ""
            price_u = ""
            basic_sale = ""
            basic_price_u = ""
            client_sale = ""
            client_price_u = ""
            qty = ""
            rating = ""
            feedbacks = ""

        # Записываем результаты в excel-файл
        row_values = [sku, name, price_u, basic_sale, basic_price_u,
                      client_sale, client_price_u, qty, rating, feedbacks]
        for i, value in enumerate(row_values):
            sheet.cell(row=row, column=i+1, value=value)

        print(f"\rSKU #{row-1} {sku} ok", end="")
        row += 1

# Сохраняем excel-файл
wb.save("Отчет по скидкам.xlsx")
print()
print("Made by https://t.me/ArChernushevich")
input("Нажмите Enter, чтобы выйти...")
