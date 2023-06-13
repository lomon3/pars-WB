
# pip install pandas
# pip install requests
# pip install pyinstaller

import requests
import openpyxl

# Получаем значение spp
def get_spp():
    headers = {
        'Cookie': '_wbauid=3828519991678728635; ___wbu=da44508c-eccc-4fd5-be34-e832e30ae01c.1678728636; BasketUID=2b74e6fa-a40f-4902-a99f-bb65a1e430aa; __wba_s=1; _wbSes=CfDJ8J5iFOpolQ9DoidgCbzUNaVSPf4XLXJesHzUfCaQsw5ixoqpE4lqSq4rY8OGN01Z7S23htQlX8lzA8FgYLB8u7TzthDHvi6TNPNKtQUdX%2Fxarx1HAbaJ%2FYv%2B19ZwhtfZF9qgUc5vFSq7PLc1b1YizpYzeNgj7Q1RMA45PIT%2FLjXV; WILDAUTHNEW_V3=04114DE34D74E7AD8E6C60E63D7F37AE2480E4E200A421173DCB230E500097918BE2869B7DA4FD51006926DE3B2B95EE578D16B9394E5CD811D860ED1AF819321D6501ABDA58E26DA50E680D6442F0147B361293B58C98F49EF5B857E2F7CC153E0CDEC85AC4CAD954C46D39BAE73539F8A49C11511A60B4DE7D7A7A0B3B2CEB3EC68A39C5760E93702F266A5773957DA01C20692E53E446310143720F0B88EF799FFCA9E18F619A3E70670143569C4DD5BB1557EC0DA636B7B0C3F2A5574B03E35BA62AD03E6417FB05CE1442DA783680EA832D1F9BCC93CABB4C91290695371A16CE6F5E67EB26C3116FCB4794C05F0945224B994544AD4DCF49B3F7F055F0F0743636F73AD28B1A553B5AE2D0D3EFC656A07FA6AB089A8FB370F51AB6058CDC82497912BA07ED33D47CE9ABF3BD7E20835225; __bsa=basket-ru-10; um=uid%3Dw7TDssOkw7PCu8KwwrPCs8K4wrXCtcK5wrY%253d%3Aproc%3D100%3Aehash%3D52e858bb1161c9fdc2855f8739496599; ___wbs=978839bb-bdc3-4171-8643-45731c261d4f.1684819252; x-supplier-id-external=d99a1f06-4d18-5d70-8ee4-d0fefc8a193f; __wbl=cityId%3D0%26regionId%3D0%26city%3D%D0%9C%D0%BE%D1%81%D0%BA%D0%B2%D0%B0%2C%20%D0%91%D0%B5%D1%80%D0%B5%D0%B3%D0%BE%D0%B2%D0%BE%D0%B9%20%D0%9F%D1%80%D0%BE%D0%B5%D0%B7%D0%B4%205%D0%BA2%26phone%3D84957755505%26latitude%3D55%2C756381000%26longitude%3D37%2C508975000%26src%3D1; __store=117673_122258_122259_125238_125239_125240_6159_507_3158_117501_120602_6158_120762_121709_124731_159402_2737_130744_117986_1733_686_132043_161812_1193_206968_206348_205228_172430_117442_117866; __region=80_115_38_4_64_83_33_68_70_69_30_86_75_40_1_66_110_22_48_31_71_114; __dst=-1029256_-102269_-226149_-446112; ncache=117673_122258_122259_125238_125239_125240_6159_507_3158_117501_120602_6158_120762_121709_124731_159402_2737_130744_117986_1733_686_132043_161812_1193_206968_206348_205228_172430_117442_117866%3B80_115_38_4_64_83_33_68_70_69_30_86_75_40_1_66_110_22_48_31_71_114%3B-1029256_-102269_-226149_-446112; __tm=1684833395',
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
sheet = wb['Общий отчет']

# Добавляем новые столбцы
sheet.cell(row=1, column=2, value="Арт ВБ")
sheet.cell(row=1, column=3, value="Название")
sheet.cell(row=1, column=4, value="Цена до скидок")
sheet.cell(row=1, column=5, value="Скидка поставщика")
sheet.cell(row=1, column=6, value="Цена после скидки поставщика")
sheet.cell(row=1, column=7, value="СПП")
sheet.cell(row=1, column=8, value="Цена после СПП")
sheet.cell(row=1, column=9, value="сток")
sheet.cell(row=1, column=10, value="рейт")
sheet.cell(row=1, column=11, value="отзывы")

# группируем список
chunk_size = 300
art_list = []
for row in range(2, sheet.max_row + 1):
    art = str(sheet.cell(row=row, column=1).value)
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
            sheet.cell(row=row, column=i+2, value=value)

        print(f"\rSKU #{row-1} {sku} ok", end="")
        row += 1

# Сохраняем excel-файл
wb.save("Отчет по скидкам.xlsx")
print()
print("Made by https://t.me/ArChernushevich")
input("Нажмите Enter, чтобы выйти...")
