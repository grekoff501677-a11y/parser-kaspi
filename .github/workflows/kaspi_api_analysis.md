
# Анализ API запросов Kaspi.kz

## 1. REQUEST HEADERS для /yml/offer-view/offers

### Endpoint (одиночный товар)
```
POST https://kaspi.kz/yml/offer-view/offers/{productId}
```

### Обязательные headers:
```
Accept: application/json, text/*
Accept-Encoding: gzip, deflate, br, zstd
Accept-Language: ru,en;q=0.9
Connection: keep-alive
Content-Type: application/json; charset=UTF-8
Host: kaspi.kz
Origin: https://kaspi.kz
Referer: https://kaspi.kz/shop/p/...
Sec-Fetch-Dest: empty
Sec-Fetch-Mode: cors
Sec-Fetch-Site: same-origin
User-Agent: Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36...
X-Description-Enabled: true
X-KS-City: 551010000
```

### Ключевые headers:
- **X-KS-City**: Код города (551010000 = Павлодар)
- **X-Description-Enabled**: true (включить описания)
- **Content-Type**: application/json; charset=UTF-8

### CURL пример:
```bash
curl -X POST 'https://kaspi.kz/yml/offer-view/offers/145169090' \
  -H 'Accept: application/json, text/*' \
  -H 'Content-Type: application/json; charset=UTF-8' \
  -H 'Origin: https://kaspi.kz' \
  -H 'Referer: https://kaspi.kz/shop/p/...' \
  -H 'User-Agent: Mozilla/5.0...' \
  -H 'X-KS-City: 551010000' \
  -H 'X-Description-Enabled: true' \
  -d '{
    "cityId": "551010000",
    "id": "145169090",
    "merchantUID": [],
    "limit": 5,
    "page": 0,
    "product": {
      "brand": "Kipardo",
      "categoryCodes": ["Rims", "Car goods", "Categories"],
      "baseProductCodes": ["[Master - Rims][kipardo]{литые}{H3317F}{7}{5x114.3}{49}{67.10}{4 диска}"],
      "groups": null,
      "productSeries": []
    },
    "sortOption": "PRICE",
    "highRating": null,
    "searchText": null,
    "isExcellentMerchant": false,
    "zoneId": ["551010000"],
    "installationId": "-1"
  }'
```

---

## 2. ENDPOINT для листинга категории

### URL страницы категории:
```
GET https://kaspi.kz/shop/pavlodar/c/rims/?q=<filter>&sort=<sort>&sc=
```

### Query-параметры:
- **q**: Фильтр в формате `:field1:value1:field2:value2...`
- **sort**: Сортировка (`relevance`, `price_asc`, `price_desc`, и т.д.)
- **sc**: (пустой в примерах)

### Пример фильтра (decoded):
```
:availableInZones:551010000
:category:Rims
:Rims*Equipment:4 диска
:Rims*Type:литые
:allMerchants:Kama
```

### Структура фильтра:
- `availableInZones:551010000` — доступно в зоне (город)
- `category:Rims` — категория
- `Rims*Equipment:4 диска` — подкатегория (4 диска)
- `Rims*Type:литые` — тип (литые)
- `allMerchants:Kama` — **фильтр по продавцу "Kama"**

---

## 3. ENDPOINT для листинга магазина "Kama"

### Метод 1: Через параметр в URL категории
```
GET https://kaspi.kz/shop/pavlodar/c/rims/?q=...:allMerchants:Kama&sort=...
```

**Параметр**: `:allMerchants:Kama` в query-параметре `q`

### Метод 2: Batch-запрос для нескольких товаров магазина
```
POST https://kaspi.kz/yml/offer-view/offers
```

**Body**:
```json
{
  "cityId": "551010000",
  "entries": [
    {"sku": "145169090", "merchantId": "Kama"},
    {"sku": "101402843", "merchantId": "Kama"}
  ],
  "options": ["PRICE"],
  "zoneId": ["551010000"]
}
```

**Параметры**:
- `entries` — массив товаров с указанием `merchantId: "Kama"`
- `options` — какие данные возвращать (`PRICE`, и т.д.)
- `cityId` / `zoneId` — код города

---

## 4. ПАГИНАЦИЯ

### Параметры пагинации (POST /yml/offer-view/offers/{id}):
```json
{
  "limit": 5,
  "page": 0,
  ...
}
```

- **limit**: Количество товаров на страницу (5 по умолчанию)
- **page**: Номер страницы (начинается с 0)

### Пример следующей страницы:
```json
{
  "limit": 5,
  "page": 1,
  ...
}
```

### Ответ включает:
- Массив товаров/офферов
- Информацию о пагинации (предположительно `totalCount`, `hasMore`, и т.д.)

---

## ПРИМЕРЫ ИСПОЛЬЗОВАНИЯ

### 1. Получить офферы товара (с пагинацией):
```bash
# Страница 1 (первые 5)
page=0, limit=5

# Страница 2 (следующие 5)  
page=1, limit=5
```

### 2. Фильтр по магазину в категории:
```
https://kaspi.kz/shop/pavlodar/c/rims/?q=:category:Rims:allMerchants:Kama&sort=relevance
```

### 3. Batch-запрос офферов от конкретного магазина:
```json
POST /yml/offer-view/offers
{
  "entries": [
    {"sku": "145169090", "merchantId": "Kama"}
  ],
  ...
}
```

---

## ВАЖНЫЕ ДЕТАЛИ

1. **Город обязателен**: `cityId` / `X-KS-City` = `551010000` (Павлодар)
2. **merchantId**: Используется `"Kama"` для фильтра по магазину
3. **Пагинация**: `page` (0-based) + `limit` (обычно 5)
4. **Формат фильтров**: Двоеточие-разделенная строка `:field:value:field:value`
5. **Сортировка**: `PRICE`, `POPULARITY`, `RATING`, `RELEVANCE`

