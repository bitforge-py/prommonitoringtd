<img width="1538" height="720" alt="image" src="https://github.com/user-attachments/assets/a447b3df-7f6a-4f99-8fd6-7a08e06cb622" />

# Мониторинг цен конкурентов

Набор скриптов для e-commerce на Prom.ua: сравнение цен магазина с конкурентами и сбор данных по продавцам. Результат выгружается в Excel.

`demo_competitors.py` открывает витрину клиента, берет товары с ценой и для каждого ищет на Prom такую же позицию. Из выдачи убираются собственный магазин, товары не в наличии и явные аксессуары (чехол или кабель не сравнивается с самим устройством). В отчет попадает минимальная цена живого конкурента. Разница считается как «цена клиента минус цена конкурента»: если клиент дороже, ячейка светло-красная, если дешевле, светло-зеленая.

`scraper.py` обходит категорию маркетплейса и собирает контакты продавцов: телефон, Telegram, Viber, email. Настройки этой части лежат в `config.py`.

Локальные таблицы (`*.xlsx`) в репозиторий не входят.

## Стек

- Python 3.10+
- Playwright
- OpenPyXL
- Rich (прогресс в консоли)

## Быстрый старт

```bash
python -m venv .venv
```

Windows:

```bash
.venv\Scripts\activate
```

macOS и Linux:

```bash
source .venv/bin/activate
```

```bash
pip install -r requirements.txt
playwright install chromium
```

Сравнение цен. В начале `demo_competitors.py` укажите витрину клиента и размер выборки:

```python
TARGET_SHOP_URL = "https://example-shop.prom.ua"
MAX_PRODUCTS = 15
```

`CATALOG_URL` можно оставить пустым (весь каталог) или указать раздел того же сайта.

```bash
python demo_competitors.py
```

Файл отчета: `demo_price_comparison.xlsx`.

Сбор контактов продавцов. В `config.py` задайте `CATEGORY_URL`, затем:

```bash
python scraper.py
```

## Пример таблицы

 В файле Excel положительная разница (клиент дороже) залита светло-красным, отрицательная (клиент дешевле) светло-зеленым. Товары ниже вымышленные.

| Товар клиента | Цена клиента (грн) | Мин. цена конкурента (грн) | Разница (грн) | Конкурент | Ссылка на конкурента | Ссылка на товар клиента |
| --- | ---: | ---: | ---: | --- | --- | --- |
| Портативная станция Nova 1000 | 24990 | 22150 | 2840 | EnergyHub | https://prom.ua/ua/p1001-nova-1000.html | https://example-shop.prom.ua/ua/p900-nova-1000.html |
| Наушники Wave Mini | 1890 | 2050 | -160 | SoundLine | https://prom.ua/ua/p1002-wave-mini.html | https://example-shop.prom.ua/ua/p901-wave-mini.html |
| Роутер Aero AX1800 | 3290 | 3290 | 0 | NetBox | https://prom.ua/ua/p1003-aero-ax1800.html | https://example-shop.prom.ua/ua/p902-aero-ax1800.html |
