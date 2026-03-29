/* ============================================================
 *  Confluence → Google Sheets Integration
 *  Забирает статьи из Confluence и переносит в таблицу
 * ============================================================ */

// ─── Настройки Confluence ───────────────────────────────────
// Заполните эти значения или используйте PropertiesService
const CONFLUENCE_BASE_URL = "";  // например: "https://your-domain.atlassian.net/wiki"
const CONFLUENCE_EMAIL    = "";  // email аккаунта Atlassian
const CONFLUENCE_API_TOKEN = ""; // API token: https://id.atlassian.com/manage-profile/security/api-tokens

// ─── Лист для импорта ───────────────────────────────────────
const CONFLUENCE_SHEET_NAME = "Confluence Data";


/* ============================================================
 *  НАСТРОЙКА CREDENTIALS (безопасное хранение)
 * ============================================================ */

/**
 * Сохраняет Confluence credentials в Script Properties.
 * Запустите один раз вручную, затем очистите константы выше.
 */
function setupConfluenceCredentials() {
  var ui = SpreadsheetApp.getUi();

  var urlResp = ui.prompt("Confluence URL", "Введите базовый URL (например: https://your-domain.atlassian.net/wiki):", ui.ButtonSet.OK_CANCEL);
  if (urlResp.getSelectedButton() !== ui.Button.OK) return;

  var emailResp = ui.prompt("Email", "Введите email аккаунта Atlassian:", ui.ButtonSet.OK_CANCEL);
  if (emailResp.getSelectedButton() !== ui.Button.OK) return;

  var tokenResp = ui.prompt("API Token", "Введите API Token (создайте на id.atlassian.com):", ui.ButtonSet.OK_CANCEL);
  if (tokenResp.getSelectedButton() !== ui.Button.OK) return;

  var props = PropertiesService.getScriptProperties();
  props.setProperty("CONFLUENCE_URL", urlResp.getResponseText().trim().replace(/\/+$/, ""));
  props.setProperty("CONFLUENCE_EMAIL", emailResp.getResponseText().trim());
  props.setProperty("CONFLUENCE_TOKEN", tokenResp.getResponseText().trim());

  ui.alert("Credentials сохранены в Script Properties.");
}

/**
 * Возвращает текущие настройки подключения.
 */
function getConfluenceConfig_() {
  var props = PropertiesService.getScriptProperties();
  var url   = props.getProperty("CONFLUENCE_URL")   || CONFLUENCE_BASE_URL;
  var email = props.getProperty("CONFLUENCE_EMAIL") || CONFLUENCE_EMAIL;
  var token = props.getProperty("CONFLUENCE_TOKEN") || CONFLUENCE_API_TOKEN;

  if (!url || !email || !token) {
    throw new Error("Confluence credentials не настроены. Запустите setupConfluenceCredentials().");
  }
  return { url: url, email: email, token: token };
}


/* ============================================================
 *  HTTP-ЗАПРОСЫ К CONFLUENCE REST API
 * ============================================================ */

/**
 * Выполняет GET-запрос к Confluence REST API.
 */
function confluenceFetch_(endpoint, queryParams) {
  var cfg = getConfluenceConfig_();
  var url = cfg.url + "/rest/api" + endpoint;

  if (queryParams) {
    var parts = [];
    for (var key in queryParams) {
      parts.push(encodeURIComponent(key) + "=" + encodeURIComponent(queryParams[key]));
    }
    url += "?" + parts.join("&");
  }

  var options = {
    method: "get",
    headers: {
      "Authorization": "Basic " + Utilities.base64Encode(cfg.email + ":" + cfg.token),
      "Accept": "application/json"
    },
    muteHttpExceptions: true
  };

  var response = UrlFetchApp.fetch(url, options);
  var code = response.getResponseCode();

  if (code !== 200) {
    throw new Error("Confluence API error " + code + ": " + response.getContentText());
  }

  return JSON.parse(response.getContentText());
}


/* ============================================================
 *  ПОЛУЧЕНИЕ СТРАНИЦ
 * ============================================================ */

/**
 * Получает все страницы из пространства Confluence.
 * @param {string} spaceKey — ключ пространства (например "DEV", "HR")
 * @param {number} [limit] — максимум страниц (по умолчанию 100)
 * @return {Array} массив объектов {id, title, url, created, updated, author, body}
 */
function getConfluencePages(spaceKey, limit) {
  limit = limit || 100;
  var allPages = [];
  var start = 0;
  var batchSize = 25; // Confluence макс. 25 за запрос по умолчанию

  while (allPages.length < limit) {
    var currentLimit = Math.min(batchSize, limit - allPages.length);
    var data = confluenceFetch_("/content", {
      spaceKey: spaceKey,
      type: "page",
      status: "current",
      expand: "body.storage,version,history.createdBy",
      start: start,
      limit: currentLimit
    });

    if (!data.results || data.results.length === 0) break;

    data.results.forEach(function(page) {
      allPages.push(parsePage_(page));
    });

    if (data.size < currentLimit) break;
    start += currentLimit;
  }

  return allPages;
}

/**
 * Получает страницы по метке (label).
 * @param {string} label — метка
 * @param {number} [limit] — максимум страниц
 * @return {Array}
 */
function getConfluencePagesByLabel(label, limit) {
  limit = limit || 100;
  var allPages = [];
  var start = 0;
  var batchSize = 25;

  while (allPages.length < limit) {
    var currentLimit = Math.min(batchSize, limit - allPages.length);
    var data = confluenceFetch_("/content/search", {
      cql: 'label = "' + label + '" AND type = page',
      expand: "body.storage,version,history.createdBy",
      start: start,
      limit: currentLimit
    });

    if (!data.results || data.results.length === 0) break;

    data.results.forEach(function(page) {
      allPages.push(parsePage_(page));
    });

    if (data.size < currentLimit) break;
    start += currentLimit;
  }

  return allPages;
}

/**
 * Получает одну страницу по ID.
 * @param {string} pageId
 * @return {Object}
 */
function getConfluencePageById(pageId) {
  var data = confluenceFetch_("/content/" + pageId, {
    expand: "body.storage,version,history.createdBy"
  });
  return parsePage_(data);
}

/**
 * Поиск страниц по тексту (CQL).
 * @param {string} searchText — текст для поиска
 * @param {number} [limit]
 * @return {Array}
 */
function searchConfluencePages(searchText, limit) {
  limit = limit || 50;
  var data = confluenceFetch_("/content/search", {
    cql: 'text ~ "' + searchText.replace(/"/g, '\\"') + '" AND type = page',
    expand: "body.storage,version,history.createdBy",
    limit: limit
  });

  return (data.results || []).map(parsePage_);
}


/* ============================================================
 *  ПАРСИНГ СТРАНИЦЫ
 * ============================================================ */

function parsePage_(page) {
  var cfg = getConfluenceConfig_();
  var author = "";
  if (page.history && page.history.createdBy) {
    author = page.history.createdBy.displayName || page.history.createdBy.username || "";
  }

  var body = "";
  if (page.body && page.body.storage) {
    body = page.body.storage.value || "";
  }

  return {
    id:       page.id,
    title:    page.title || "",
    url:      cfg.url + page._links.webui,
    created:  page.history ? page.history.createdDate : "",
    updated:  page.version ? page.version.when : "",
    version:  page.version ? page.version.number : 0,
    author:   author,
    bodyHtml: body,
    bodyText: htmlToText_(body)
  };
}

/**
 * Простое преобразование HTML → текст.
 */
function htmlToText_(html) {
  if (!html) return "";
  return html
    .replace(/<br\s*\/?>/gi, "\n")
    .replace(/<\/p>/gi, "\n")
    .replace(/<\/div>/gi, "\n")
    .replace(/<\/tr>/gi, "\n")
    .replace(/<\/li>/gi, "\n")
    .replace(/<[^>]+>/g, "")
    .replace(/&nbsp;/g, " ")
    .replace(/&amp;/g, "&")
    .replace(/&lt;/g, "<")
    .replace(/&gt;/g, ">")
    .replace(/&quot;/g, '"')
    .replace(/\n{3,}/g, "\n\n")
    .trim();
}


/* ============================================================
 *  ЗАПИСЬ В GOOGLE SHEETS
 * ============================================================ */

/**
 * Импортирует страницы из пространства Confluence в лист таблицы.
 * Вызывается из меню.
 */
function importConfluenceBySpace() {
  var ui = SpreadsheetApp.getUi();
  var resp = ui.prompt("Импорт из Confluence", "Введите ключ пространства (Space Key):", ui.ButtonSet.OK_CANCEL);
  if (resp.getSelectedButton() !== ui.Button.OK) return;

  var spaceKey = resp.getResponseText().trim();
  if (!spaceKey) { ui.alert("Ключ пространства не указан."); return; }

  var pages = getConfluencePages(spaceKey);
  writePagesToSheet_(pages);
  ui.alert("Импортировано " + pages.length + " страниц из пространства " + spaceKey + ".");
}

/**
 * Импортирует страницы по метке.
 */
function importConfluenceByLabel() {
  var ui = SpreadsheetApp.getUi();
  var resp = ui.prompt("Импорт по метке", "Введите метку (label):", ui.ButtonSet.OK_CANCEL);
  if (resp.getSelectedButton() !== ui.Button.OK) return;

  var label = resp.getResponseText().trim();
  if (!label) { ui.alert("Метка не указана."); return; }

  var pages = getConfluencePagesByLabel(label);
  writePagesToSheet_(pages);
  ui.alert("Импортировано " + pages.length + " страниц с меткой '" + label + "'.");
}

/**
 * Поиск и импорт по тексту.
 */
function importConfluenceBySearch() {
  var ui = SpreadsheetApp.getUi();
  var resp = ui.prompt("Поиск в Confluence", "Введите текст для поиска:", ui.ButtonSet.OK_CANCEL);
  if (resp.getSelectedButton() !== ui.Button.OK) return;

  var text = resp.getResponseText().trim();
  if (!text) { ui.alert("Текст поиска не указан."); return; }

  var pages = searchConfluencePages(text);
  writePagesToSheet_(pages);
  ui.alert("Найдено и импортировано " + pages.length + " страниц.");
}

/**
 * Записывает массив страниц в лист Google Sheets.
 */
function writePagesToSheet_(pages) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(CONFLUENCE_SHEET_NAME);

  if (!sheet) {
    sheet = ss.insertSheet(CONFLUENCE_SHEET_NAME);
  }

  sheet.clear();

  // Заголовки
  var headers = ["ID", "Заголовок", "Автор", "Создано", "Обновлено", "Версия", "URL", "Содержимое (текст)"];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  sheet.getRange(1, 1, 1, headers.length)
    .setFontWeight("bold")
    .setBackground("#4A86C8")
    .setFontColor("#FFFFFF");

  if (pages.length === 0) return;

  // Данные
  var rows = pages.map(function(p) {
    return [
      p.id,
      p.title,
      p.author,
      p.created ? new Date(p.created) : "",
      p.updated ? new Date(p.updated) : "",
      p.version,
      p.url,
      p.bodyText.substring(0, 50000) // ограничение ячейки
    ];
  });

  sheet.getRange(2, 1, rows.length, headers.length).setValues(rows);

  // Форматирование
  sheet.setFrozenRows(1);
  sheet.autoResizeColumns(1, headers.length);
  sheet.setColumnWidth(8, 400); // содержимое шире
}


/* ============================================================
 *  АВТОМАТИЧЕСКИЙ ИМПОРТ ПО РАСПИСАНИЮ
 * ============================================================ */

/**
 * Устанавливает триггер автоимпорта (раз в час).
 * Требует заполнить SPACE_KEY в Script Properties.
 */
function setupConfluenceAutoSync() {
  var ui = SpreadsheetApp.getUi();
  var resp = ui.prompt(
    "Автосинхронизация",
    "Введите ключ пространства для автоматического импорта каждый час:",
    ui.ButtonSet.OK_CANCEL
  );
  if (resp.getSelectedButton() !== ui.Button.OK) return;

  var spaceKey = resp.getResponseText().trim();
  if (!spaceKey) { ui.alert("Ключ не указан."); return; }

  PropertiesService.getScriptProperties().setProperty("AUTO_SYNC_SPACE", spaceKey);

  // Удаляем старый триггер
  ScriptApp.getProjectTriggers().forEach(function(t) {
    if (t.getHandlerFunction() === "autoSyncConfluence_") {
      ScriptApp.deleteTrigger(t);
    }
  });

  ScriptApp.newTrigger("autoSyncConfluence_")
    .timeBased()
    .everyHours(1)
    .create();

  ui.alert("Автосинхронизация настроена. Пространство: " + spaceKey + ", интервал: каждый час.");
}

/**
 * Автоматический импорт (вызывается триггером).
 */
function autoSyncConfluence_() {
  var spaceKey = PropertiesService.getScriptProperties().getProperty("AUTO_SYNC_SPACE");
  if (!spaceKey) {
    Logger.log("AUTO_SYNC_SPACE не задан. Пропускаем.");
    return;
  }

  var pages = getConfluencePages(spaceKey, 200);
  writePagesToSheet_(pages);
  Logger.log("Автосинхронизация: импортировано " + pages.length + " страниц из " + spaceKey);
}


/* ============================================================
 *  РАСШИРЕНИЕ МЕНЮ
 * ============================================================ */

// Добавляем пункты Confluence в меню onOpen (дополняет основное меню)
function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu("Нагрузка T&A")
    .addItem("Пересчитать всё",           "высчитываниеНагрузки")
    .addItem("Только пересчёт данных",    "recalculateMain_")
    .addItem("Только диаграммы",          "buildCharts_")
    .addItem("Установить триггеры",       "setupTriggers")
    .addSeparator()
    .addItem("Отладка: ошибки по столам", "TA_DebugDeskErrorCounts")
    .addItem("Отладка: стажёры по столам","TA_DebugDeskCounts")
    .addToUi();

  ui.createMenu("Confluence")
    .addItem("Настроить подключение",       "setupConfluenceCredentials")
    .addSeparator()
    .addItem("Импорт по пространству",     "importConfluenceBySpace")
    .addItem("Импорт по метке",            "importConfluenceByLabel")
    .addItem("Поиск и импорт",             "importConfluenceBySearch")
    .addSeparator()
    .addItem("Настроить автосинхронизацию", "setupConfluenceAutoSync")
    .addToUi();
}
