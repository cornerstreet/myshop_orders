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

  // Отримання існуючих "Номерів Замовлень" для уникнення дублювання
  const existingOrderIdentifiers = getExistingOrderIdentifiers(sheet); // Змінено назву функції та змінної

  let nextLink = ''; // Змінна для зберігання URL наступної сторінки (курсорна пагінація)
  let hasNextPage = true;
  const ordersToImport = [];

  // Цикл для отримання всіх сторінок замовлень
  while (hasNextPage) {
    let ordersUrl;
    if (nextLink === '') {
      // Перший запит: отримуємо замовлення з певними параметрами
      ordersUrl = `${shopifyBaseUrl}/orders.json?status=any&limit=250`;
    } else {
      // Наступні запити: використовуємо URL, отриманий з заголовка 'Link' попередньої відповіді
      ordersUrl = nextLink;
    }

    const options = {
      method: "GET",
      headers: headers,
      muteHttpExceptions: true
    };

    const response = UrlFetchApp.fetch(ordersUrl, options);
    const responseCode = response.getResponseCode();
    const responseBody = response.getContentText();
    const responseHeaders = response.getHeaders();

    if (responseCode === 200) {
      const data = JSON.parse(responseBody);
      const orders = data.orders;

      if (orders && orders.length > 0) {
        orders.forEach(order => {
          // Перевірка на дублювання замовлення за "Номер Замовлення" (order.name)
          // Важливо: Якщо "Номер Замовлення" не є унікальним, цей механізм
          // дедуплікації може працювати некоректно.
          // Оригінальний order.id є гарантовано унікальним, але його прибрали з виводу.
          if (!existingOrderIdentifiers.has(String(order.name))) {
            const customer = order.customer;
            const lineItems = order.line_items;
            let productsDescription = "";

            if (lineItems && lineItems.length > 0) {
              productsDescription = lineItems.map(item => `${item.title} (${item.quantity} шт.)`).join("; ");
            }

            // Нова послідовність стовпчиків:
            // Номер Замовлення, Дата Замовлення, Статус Замовлення Shopify, Ім'я Клієнта, Email Клієнта, Телефон Клієнта,
            // Загальна Сума, Товари Замовлення (Опис), Статус Виготовлення, Вартість Товарів,
            // Вартість Доставки, Чистий Прибуток, URL Замовлення Shopify
            const row = [
              order.name, // Номер Замовлення
              new Date(order.created_at).toLocaleString(), // Дата Замовлення
              order.financial_status, // Статус Замовлення Shopify
              customer ? `${customer.first_name || ''} ${customer.last_name || ''}`.trim() : '', // Ім'я Клієнта
              customer ? customer.email || '' : '', // Email Клієнта
              customer ? customer.phone || '' : '', // Телефон Клієнта
              `${parseFloat(order.total_price)} ${order.currency}`, // Загальна Сума (з валютою)
              productsDescription, // Товари Замовлення (Опис)
              "Нове", // Статус Виготовлення (початковий)
              "", // Вартість Товарів
              "", // Вартість Доставки
              "",  // Чистий Прибуток
              `https://${SHOPIFY_STORE_NAME}.myshopify.com/admin/orders/${order.id}` // URL Замовлення Shopify
            ];
            ordersToImport.push(row);
          }
        });

        if (responseHeaders["Link"] !== undefined) {
          const links = parseLinkHeader(responseHeaders['Link']);
          if (links['next'] !== undefined) {
            nextLink = links['next']['href'];
            Logger.log(`Знайдено наступну сторінку для пагінації: ${nextLink}`);
          } else {
            hasNextPage = false;
          }
        } else {
          hasNextPage = false;
        }

      } else {
        hasNextPage = false;
      }
    } else {
      Logger.log(`Помилка отримання замовлень Shopify. Код: ${responseCode}, Відповідь: ${responseBody}`);
      Browser.msgBox(`Помилка отримання замовлень Shopify. Перевірте API ключі та дозволи. Код: ${responseCode}`);
      hasNextPage = false;
    }
  }

  if (ordersToImport.length > 0) {
    sheet.getRange(sheet.getLastRow() + 1, 1, ordersToImport.length, ordersToImport[0].length).setValues(ordersToImport);
    Logger.log(`Успішно імпортовано ${ordersToImport.length} нових замовлень.`);
    Browser.msgBox(`Успішно імпортовано ${ordersToImport.length} нових замовлень.`);
  } else {
    Logger.log("Немає нових замовлень для імпорту.");
    Browser.msgBox("Немає нових замовлень для імпорту.");
  }
}

// Допоміжна функція для отримання існуючих "Номерів Замовлень" з таблиці
function getExistingOrderIdentifiers(sheet) { // Змінено назву функції
  const data = sheet.getDataRange().getValues();
  const orderIdentifiers = new Set();
  // Починаємо з 1, щоб пропустити рядок заголовків
  for (let i = 1; i < data.length; i++) {
    // "Номер Замовлення" (order.name) тепер перший стовпець (індекс 0)
    // і буде використовуватися для уникнення дублікатів.
    const orderIdentifier = data[i][0];
    if (orderIdentifier) {
      orderIdentifiers.add(String(orderIdentifier));
    }
  }
  return orderIdentifiers;
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
  const triggers = ScriptApp.getProjectTriggers();
  for (let i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === "fetchShopifyOrders") {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }
  ScriptApp.newTrigger("fetchShopifyOrders")
    .timeBased()
    .everyMinutes(15)
    .create();
  Logger.log("Тригер для синхронізації замовлень встановлено на кожні 15 хвилин.");
  Browser.msgBox("Тригер для синхронізації замовлень встановлено на кожні 15 хвилин.");
}

// Допоміжна функція для ручного запуску скрипту (для тестування)
function runOnce() {
  fetchShopifyOrders();
}
