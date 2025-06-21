function fetchShopifyOrders() {
  // --- КОНФІГУРАЦІЯ ---
  const SHOPIFY_STORE_NAME = "shop_domain";
  const SHOPIFY_API_KEY = " ";      // Ваш Admin API Key
  const SHOPIFY_ACCESS_TOKEN = "shpat_"; // Ваш Admin API Access Token

  const SHEET_NAME = "Замовлення";

  // --- НЕ РЕДАГУЙТЕ НИЖЧЕ, ЯКЩО НЕ ВПЕВНЕНІ ---
  const shopifyBaseUrl = `https://${SHOPIFY_STORE_NAME}.myshopify.com/admin/api/2025-04`;
  const headers = {
    "X-Shopify-Access-Token": SHOPIFY_ACCESS_TOKEN,
    "Content-Type": "application/json"
  };

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAME);

  if (!sheet) {
    Logger.log(`Помилка: Аркуш '${SHEET_NAME}' не знайдено.`);
    Browser.msgBox(`Помилка: Аркуш '${SHEET_NAME}' не знайдено. Будь ласка, переконайтеся, що аркуш існує.`);
    return;
  }

  const existingOrderIdentifiers = getExistingOrderIdentifiers(sheet);
  let nextLink = '';
  let hasNextPage = true;
  const ordersToImport = [];

  while (hasNextPage) {
    let ordersUrl = nextLink === '' ? `${shopifyBaseUrl}/orders.json?status=any&limit=250` : nextLink;
    const options = { method: "GET", headers: headers, muteHttpExceptions: true };
    const response = UrlFetchApp.fetch(ordersUrl, options);
    const responseCode = response.getResponseCode();
    const responseBody = response.getContentText();
    const responseHeaders = response.getHeaders();

    if (responseCode === 200) {
      const data = JSON.parse(responseBody);
      const orders = data.orders;

      if (orders && orders.length > 0) {
        orders.forEach(order => {
          if (!existingOrderIdentifiers.has(String(order.name))) {
            const customer = order.customer;
            const lineItems = order.line_items;

            let productTitles = "";
            let productQuantities = "";
            if (lineItems && lineItems.length > 0) {
              productTitles = lineItems.map(item => item.title).join("; ");
              productQuantities = lineItems.map(item => item.quantity).join("; ");
            }

            let shippingCost = 0;
            if (order.shipping_lines && order.shipping_lines.length > 0) {
              order.shipping_lines.forEach(line => {
                shippingCost += parseFloat(line.price);
              });
            }

            const row = [
              order.name, // Номер Замовлення
              new Date(order.created_at).toLocaleString(), // Дата Замовлення
              order.financial_status, // Статус Замовлення Shopify
              customer ? `${customer.first_name || ''} ${customer.last_name || ''}`.trim() : '', // Ім'я Клієнта
              customer ? customer.email || '' : '', // Email Клієнта
              customer ? customer.phone || '' : '', // Телефон Клієнта
              parseFloat(order.total_price), // Загальна Сума (без валюти)
              productTitles, // Товари Замовлення (Назва)
              productQuantities, // Товари Замовлення (Кількість)
              "Нове", // Статус Виготовлення
              parseFloat(order.total_line_items_price), // Вартість Товарів
              shippingCost, // Вартість Доставки
              "",  // Чистий Прибуток
              `https://${SHOPIFY_STORE_NAME}.myshopify.com/admin/orders/${order.id}` // URL Замовлення Shopify
            ];
            ordersToImport.push(row);
          }
        });

        if (responseHeaders["Link"]) {
          const links = parseLinkHeader(responseHeaders['Link']);
          if (links['next']) {
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

function getExistingOrderIdentifiers(sheet) {
  const data = sheet.getDataRange().getValues();
  const orderIdentifiers = new Set();
  for (let i = 1; i < data.length; i++) {
    const orderIdentifier = data[i][0]; // "Номер Замовлення" (order.name) - перший стовпець
    if (orderIdentifier) {
      orderIdentifiers.add(String(orderIdentifier));
    }
  }
  return orderIdentifiers;
}

function parseLinkHeader(header) {
  var linkexp = /<[^>]*>\s*(\s*;\s*[^\(\)<>@,;:"\/\[\]\?={} \t]+=(([^\(\)<>@,;:"\/\[\]\?={} \t]+)|("[^"]*")))*(,|$)/g;
  var paramexp = /[^\(\)<>@,;:"\/\[\]\?={} \t]+=(([^\(\)<>@,;:"\/\[\]\?={} \t]+)|("[^"]*"))/g;
  var matches = header.match(linkexp);
  var rels = {};
  for (let i = 0; matches && i < matches.length; i++) {
    var split = matches[i].split('>');
    var href = split[0].substring(1);
    var ps = split[1];
    var link = { href: href };
    var s = ps.match(paramexp);
    for (let j = 0; s && j < s.length; j++) {
      var p = s[j];
      var paramsplit = p.split('=');
      var name = paramsplit[0];
      link[name] = unquote(paramsplit[1]);
    }
    if (link.rel) {
      rels[link.rel] = link;
    }
  }
  return rels;
}

function unquote(value) {
  if (value.charAt(0) == '"' && value.charAt(value.length - 1) == '"') return value.substring(1, value.length - 1);
  return value;
}

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

function runOnce() {
  fetchShopifyOrders();
}
