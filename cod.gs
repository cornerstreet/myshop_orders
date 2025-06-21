function fetchShopifyOrders() {
  // --- КОНФІГУРАЦІЯ ---
  // ВСТАВТЕ СВОЇ ДАНІ З SHOPIFY API СЮДИ
  // Store name should be the part before .myshopify.com (e.g., "shop_domain" for shop_domain.myshopify.com)
  const SHOPIFY_STORE_NAME = "shop_domain"; 
  const SHOPIFY_API_KEY = " ";      // Ваш Admin API Key
  const SHOPIFY_ACCESS_TOKEN = "shpat_"; // Ваш Admin API Access Token (починається з 'shpat_')

  const SPREADSHEET_NAME = "lamamarka orders"; // Назва вашої таблиці Google Sheets
  const SHEET_NAME = "Замовлення";             // Назва вкладки для замовлень

  // --- НЕ РЕДАГУЙТЕ НИЖЧЕ, ЯКЩО НЕ ВПЕВНЕНІ ---

  // Використовуємо актуальну версію Shopify Admin API.
  const shopifyBaseUrl = `https://${SHOPIFY_STORE_NAME}.myshopify.com/admin/api/2025-04`; 
  const headers = {
    "X-Shopify-Access-Token": SHOPIFY_ACCESS_TOKEN,
    "Content-Type": "application/json"
  };

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAME);

  // Перевірка наявності аркуша "Замовлення"
  if (!sheet) {
    Logger.log(`Помилка: Аркуш '${SHEET_NAME}' не знайдено.`);
    Browser.msgBox(`Помилка: Аркуш '${SHEET_NAME}' не знайдено. Будь ласка, переконайтеся, що аркуш існує.`);
    return;
  }

  // Отримання існуючих ID замовлень для уникнення дублювання
  const existingOrderIds = getExistingOrderIds(sheet);

  let nextLink = ''; // Змінна для зберігання URL наступної сторінки (курсорна пагінація)
  let hasNextPage = true;
  const ordersToImport = [];

  // Цикл для отримання всіх сторінок замовлень
  while (hasNextPage) {
    let ordersUrl;
    if (nextLink === '') {
      // Перший запит: отримуємо замовлення з певними параметрами
      // Видалено параметр 'page', оскільки він більше не підтримується Shopify [cite: 1.18]
      ordersUrl = `${shopifyBaseUrl}/orders.json?status=any&limit=250`; // status=any отримує всі замовлення, limit до 250 для оптимального запиту
    } else {
      // Наступні запити: використовуємо URL, отриманий з заголовка 'Link' попередньої відповіді
      ordersUrl = nextLink;
    }

    const options = {
      method: "GET",
      headers: headers,
      muteHttpExceptions: true // Дозволяє отримувати деталі відповіді навіть при помилках HTTP (наприклад, 404)
    };

    // Виконання HTTP-запиту до Shopify API
    const response = UrlFetchApp.fetch(ordersUrl, options);
    const responseCode = response.getResponseCode();
    const responseBody = response.getContentText();
    const responseHeaders = response.getHeaders(); // Отримуємо всі заголовки відповіді, включаючи "Link"

    // Обробка успішної відповіді (код 200)
    if (responseCode === 200) {
      const data = JSON.parse(responseBody);
      const orders = data.orders;

      if (orders && orders.length > 0) {
        // Обробка кожного замовлення
        orders.forEach(order => {
          // Перевірка на дублювання замовлення
          if (!existingOrderIds.has(String(order.id))) {
            const customer = order.customer;
            const lineItems = order.line_items;
            let productsDescription = "";

            // Формування опису товарів у замовленні
            if (lineItems && lineItems.length > 0) {
              productsDescription = lineItems.map(item => `${item.title} (${item.quantity} шт.)`).join("; ");
            }

            // Формування рядка даних для вставки в Google Sheets
            const row = [
              String(order.id),
              order.name,
              new Date(order.created_at).toLocaleString(), // Форматування дати
              order.financial_status,
              // ПОЛІ PII (Персональні дані):
              // Зверніть увагу, що ці поля (ім'я, email, телефон) будуть порожніми
              // на тарифному плані "Basic", оскільки Shopify обмежує доступ до PII
              // на цьому плані. Для отримання цих даних потрібно оновити план
              // до "Shopify", "Advanced" або "Plus".
              customer ? `${customer.first_name || ''} ${customer.last_name || ''}`.trim() : '',
              customer ? customer.email || '' : '',
              customer ? customer.phone || '' : '',
              parseFloat(order.total_price),
              order.currency,
              `https://${SHOPIFY_STORE_NAME}.myshopify.com/admin/orders/${order.id}`, // URL замовлення в адмін-панелі Shopify
              productsDescription,
              "Нове", // Початковий статус виготовлення
              "", // Вартість Товарів (для ручного заповнення або майбутньої автоматизації)
              "", // Вартість Доставки (для ручного заповнення або майбутньої автоматизації)
              ""  // Чистий Прибуток (для ручного заповнення або майбутньої автоматизації)
            ];
            ordersToImport.push(row);
          }
        });

        // ЛОГІКА КУРСОРНОЇ ПАГІНАЦІЇ:
        // Перевірка заголовка "Link" для визначення наступної сторінки
        if (responseHeaders["Link"] !== undefined) {
          const links = parseLinkHeader(responseHeaders['Link']); // Використовуємо допоміжну функцію для парсингу заголовка
          if (links['next'] !== undefined) {
            nextLink = links['next']['href']; // Отримуємо повний URL для наступної сторінки
            Logger.log(`Знайдено наступну сторінку для пагінації: ${nextLink}`);
          } else {
            hasNextPage = false; // Більше немає сторінок
          }
        } else {
          hasNextPage = false; // Немає заголовка Link, отже, більше немає сторінок
        }

      } else {
        hasNextPage = false; // У відповіді немає замовлень, завершуємо цикл
      }
    } else {
      // Обробка помилок відповіді API
      Logger.log(`Помилка отримання замовлень Shopify. Код: ${responseCode}, Відповідь: ${responseBody}`);
      Browser.msgBox(`Помилка отримання замовлень Shopify. Перевірте API ключі та дозволи. Код: ${responseCode}`);
      hasNextPage = false; // Зупиняємо цикл при помилці
    }
  }

  // Вставка нових замовлень у таблицю Google Sheets
  if (ordersToImport.length > 0) {
    sheet.getRange(sheet.getLastRow() + 1, 1, ordersToImport.length, ordersToImport[0].length).setValues(ordersToImport);
    Logger.log(`Успішно імпортовано ${ordersToImport.length} нових замовлень.`);
    Browser.msgBox(`Успішно імпортовано ${ordersToImport.length} нових замовлень.`);
  } else {
    Logger.log("Немає нових замовлень для імпорту.");
    Browser.msgBox("Немає нових замовлень для імпорту.");
  }
}

// Допоміжна функція для отримання існуючих ID замовлень з таблиці
function getExistingOrderIds(sheet) {
  const data = sheet.getDataRange().getValues();
  const orderIds = new Set();
  // Починаємо з 1, щоб пропустити рядок заголовків
  for (let i = 1; i < data.length; i++) {
    const orderId = data[i][0]; // ID Замовлення Shopify знаходиться в першій колонці (індекс 0)
    if (orderId) {
      orderIds.add(String(orderId)); // Додаємо ID як рядок для коректного порівняння
    }
  }
  return orderIds;
}

// *** ЦІ ДВІ ФУНКЦІЇ АБСОЛЮТНО НЕОБХІДНІ ДЛЯ КУРСОРНОЇ ПАГІНАЦІЇ ***
// Вони парсять заголовок "Link" HTTP-відповіді, щоб знайти URL для наступної сторінки даних.

function parseLinkHeader(header) {
  var linkexp = /<[^>]*>\s*(\s*;\s*[^\(\)<>@,;:"\/\[\]\?={} \t]+=(([^\(\)<>@,;:"\/\[\]\?={} \t]+)|("[^"]*")))*(,|$)/g;
  var paramexp = /[^\(\)<>@,;:"\/\[\]\?={} \t]+=(([^\(\)<>@,;:"\/\[\]\?={} \t]+)|("[^"]*"))/g;

  var matches = header.match(linkexp);
  var rels = {};
  for (let i = 0; i < matches.length; i++) {
    var split = matches[i].split('>');
    var href = split[0].substring(1);
    var ps = split[1];
    var link = {};
    link.href = href;
    var s = ps.match(paramexp);
    for (let j = 0; j < s.length; j++) {
      var p = s[j];
      var paramsplit = p.split('=');
      var name = paramsplit[0];
      link[name] = unquote(paramsplit[1]);
    }

    if (link.rel !== undefined) {
      rels[link.rel] = link;
    }
  }
  return rels;
}

function unquote(value) {
  if (value.charAt(0) == '"' && value.charAt(value.length - 1) == '"') return value.substring(1, value.length - 1);
  return value;
}


// Функція для налаштування автоматичного запуску скрипту (тригер)
function setupOrderSyncTrigger() {
  // Видаляє всі існуючі тригери для цієї функції, щоб уникнути дублювання
  const triggers = ScriptApp.getProjectTriggers();
  for (let i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === "fetchShopifyOrders") {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }
  // Створює новий тригер: запускати функцію 'fetchShopifyOrders' кожні 15 хвилин
  ScriptApp.newTrigger("fetchShopifyOrders")
    .timeBased()
    .everyMinutes(15) // Можна змінити інтервал: everyHours(1), everyDays(1), etc.
    .create();
  Logger.log("Тригер для синхронізації замовлень встановлено на кожні 15 хвилин.");
  Browser.msgBox("Тригер для синхронізації замовлень встановлено на кожні 15 хвилин.");
}

// Допоміжна функція для ручного запуску скрипту (для тестування)
function runOnce() {
  fetchShopifyOrders();
}
