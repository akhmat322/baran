/* ============================================================
 *  Нагрузка T&A — автоматизация Google Sheets
 *  v3.0
 * ============================================================ */

// ─── Листы ───────────────────────────────────────────────────
const MAIN_SHEET_NAME        = "Нагрузка T&A";
const CHILD_SHEET_NAME       = "ОТЧЁТ (СОТРУДНИКИ)";
const STAFF_LIST_SHEET_NAME  = "СПИСОК СОТРУДНИКОВ";

// ─── Стартовая строка данных на главном листе ────────────────
const MAIN_START_ROW = 3;

// ─── Колонки листа ОТЧЁТ (СОТРУДНИКИ) ───────────────────────
const CHILD_COL_TABLE      = 1;   // A — стол / GEO
const CHILD_COL_STATUS_DD  = 10;  // J — dropdown статус
const CHILD_COL_DATE       = 15;  // O — дата
const CHILD_COL_STATUS_TXT = 18;  // R — текстовый дубль статуса

// ─── Колонки листа СПИСОК СОТРУДНИКОВ ────────────────────────
const STAFF_COL_A      = 1;  // A — заголовок стола / имя
const STAFF_COL_STATUS = 2;  // B — статус

// ─── Колонки главного листа ──────────────────────────────────
const MAIN_COL_NAME       = 1;  // A — Сотрудник
const MAIN_COL_ZONE       = 2;  // B — Зона
const MAIN_COL_ERRORS     = 3;  // C — Ошибки (авто)
const MAIN_COL_TRAINEES   = 4;  // D — Стажёры (авто)
const MAIN_COL_HAS_TRAIN  = 5;  // E — Есть обучение (чекбокс)
const MAIN_COL_HAS_PROJ   = 6;  // F — Есть проект (чекбокс)
const MAIN_COL_POINTS     = 7;  // G — Баллы (авто)
const MAIN_COL_PERCENT    = 8;  // H — %Нагрузки (авто)

// ─── Ячейки с настройками баллов (на главном листе, колонки J-K) ──
const CFG_POINTS_PER_TRAINEE  = "K2";  // Баллы за стажёра     (10)
const CFG_PCT_PER_TRAINEE     = "K3";  // % за стажёра         (5%)
const CFG_POINTS_PER_ERROR    = "K4";  // Баллы за ошибку      (2)
const CFG_PCT_PER_ERROR       = "K5";  // % за ошибку          (1%)
const CFG_POINTS_PER_PROJECT  = "K6";  // Баллы за проект      (20)
const CFG_PCT_PER_PROJECT     = "K7";  // % за проект          (20)
const CFG_POINTS_PER_TRAINING = "K8";  // Баллы за обучение    (40)
const CFG_PCT_PER_TRAINING    = "K9";  // % за обучение        (40)

// ─── Фильтрация ─────────────────────────────────────────────
const STATUS_TARGET = "не обработано";
const DAYS_WINDOW   = 30;

// ─── Максимум баллов (200 баллов = 100%) ─────────────────────
const MAX_POINTS = 200;


/* ============================================================
 *  ТРИГГЕРЫ
 * ============================================================ */

function setupTriggers() {
  ScriptApp.getProjectTriggers().forEach(function (trigger) {
    var name = trigger.getHandlerFunction();
    if (name === "onEditHandler" || name === "onChangeHandler") {
      ScriptApp.deleteTrigger(trigger);
    }
  });

  var ssId = SpreadsheetApp.getActiveSpreadsheet().getId();

  ScriptApp.newTrigger("onEditHandler")
    .forSpreadsheet(ssId)
    .onEdit()
    .create();

  ScriptApp.newTrigger("onChangeHandler")
    .forSpreadsheet(ssId)
    .onChange()
    .create();

  Logger.log("Триггеры установлены");
}


/* ============================================================
 *  ОБРАБОТЧИКИ СОБЫТИЙ
 * ============================================================ */

function onEditHandler(e) {
  if (!e || !e.range) return;

  var sheet     = e.range.getSheet();
  var sheetName = sheet.getName();
  var col       = e.range.getColumn();

  // Лист отчёта
  if (sheetName === CHILD_SHEET_NAME) {
    if (col === CHILD_COL_STATUS_DD) {
      mirrorStatusToR_(sheet, e.range.getRow());
    }
    recalculateMain_();
    return;
  }

  // Лист сотрудников
  if (sheetName === STAFF_LIST_SHEET_NAME) {
    recalculateMain_();
    return;
  }

  // Главный лист — зона, чекбоксы обучение/проект
  if (sheetName === MAIN_SHEET_NAME &&
      (col === MAIN_COL_ZONE ||
       col === MAIN_COL_HAS_TRAIN || col === MAIN_COL_HAS_PROJ)) {
    recalculateMain_();
  }
}

function onChangeHandler() {
  try {
    fillStatusMirrorColumnR_();
  } catch (err) {
    Logger.log("onChange: ошибка зеркалирования статусов: " + err.message);
  }
  recalculateMain_();
}


/* ============================================================
 *  ЗЕРКАЛИРОВАНИЕ СТАТУСОВ (J → R)
 * ============================================================ */

function mirrorStatusToR_(sheet, row) {
  if (row < 2) return;
  var value = sheet.getRange(row, CHILD_COL_STATUS_DD).getValue();
  sheet.getRange(row, CHILD_COL_STATUS_TXT).setValue(String(value || "").trim());
}

function fillStatusMirrorColumnR_() {
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CHILD_SHEET_NAME);
  if (!sh) return;

  var lastRow = sh.getLastRow();
  if (lastRow < 2) return;

  var values = sh.getRange(2, CHILD_COL_STATUS_DD, lastRow - 1, 1).getValues();
  var out = values.map(function (row) {
    return [String(row[0] || "").trim()];
  });
  sh.getRange(2, CHILD_COL_STATUS_TXT, lastRow - 1, 1).setValues(out);
}


/* ============================================================
 *  ГЛАВНЫЙ ПЕРЕСЧЁТ — ошибки, стажёры, баллы, %
 * ============================================================ */

function recalculateMain_() {
  var ss         = SpreadsheetApp.getActiveSpreadsheet();
  var mainSheet  = ss.getSheetByName(MAIN_SHEET_NAME);
  var childSheet = ss.getSheetByName(CHILD_SHEET_NAME);
  var staffSheet = ss.getSheetByName(STAFF_LIST_SHEET_NAME);

  if (!mainSheet)  { Logger.log("Лист не найден: " + MAIN_SHEET_NAME);  return; }
  if (!childSheet) { Logger.log("Лист не найден: " + CHILD_SHEET_NAME); return; }
  if (!staffSheet) { Logger.log("Лист не найден: " + STAFF_LIST_SHEET_NAME); return; }

  // ── Читаем настройки баллов ──
  var cfg = readConfig_(mainSheet);

  // ── Считаем ошибки и стажёров ──
  var errorCounts   = countNotProcessedByDesk_(childSheet);
  var traineeCounts = countTraineesByDeskStructured_(staffSheet);

  // ── Определяем диапазон данных ──
  var mainLast = getMainLastDataRow_(mainSheet, MAIN_START_ROW);
  if (mainLast < MAIN_START_ROW) return;
  var numRows = mainLast - MAIN_START_ROW + 1;

  // ── Batch-чтение: зоны + чекбоксы (колонки A-F) ──
  var mainData = mainSheet.getRange(MAIN_START_ROW, 1, numRows, MAIN_COL_HAS_PROJ).getValues();

  // ── Формируем выходные колонки C, D, G, H ──
  var outErrors   = [];
  var outTrainees = [];
  var outPoints   = [];
  var outPercent  = [];

  for (var i = 0; i < numRows; i++) {
    var zoneKey     = normDeskKey_(mainData[i][MAIN_COL_ZONE - 1]);
    var hasTraining = isChecked_(mainData[i][MAIN_COL_HAS_TRAIN - 1]);
    var hasProject  = isChecked_(mainData[i][MAIN_COL_HAS_PROJ - 1]);

    var errors   = zoneKey ? getCountForDeskOrGroup_(errorCounts, zoneKey) : 0;
    var trainees = zoneKey ? getCountForDeskOrGroup_(traineeCounts, zoneKey) : 0;

    // ── Баллы ──
    var points = errors   * cfg.pointsPerError
               + trainees * cfg.pointsPerTrainee
               + (hasProject  ? cfg.pointsPerProject  : 0)
               + (hasTraining ? cfg.pointsPerTraining : 0);

    // ── % нагрузки ──
    var pct = errors   * cfg.pctPerError
            + trainees * cfg.pctPerTrainee
            + (hasProject  ? cfg.pctPerProject  : 0)
            + (hasTraining ? cfg.pctPerTraining : 0);

    // Ограничиваем 100%
    if (pct > 1) pct = 1;

    outErrors.push([errors]);
    outTrainees.push([trainees]);
    outPoints.push([points]);
    outPercent.push([pct]);
  }

  // ── Batch-запись ──
  mainSheet.getRange(MAIN_START_ROW, MAIN_COL_ERRORS,   numRows, 1).setValues(outErrors);
  mainSheet.getRange(MAIN_START_ROW, MAIN_COL_TRAINEES, numRows, 1).setValues(outTrainees);
  mainSheet.getRange(MAIN_START_ROW, MAIN_COL_POINTS,   numRows, 1).setValues(outPoints);
  mainSheet.getRange(MAIN_START_ROW, MAIN_COL_PERCENT,  numRows, 1).setNumberFormat("0%").setValues(outPercent);
}


/* ============================================================
 *  ЧТЕНИЕ КОНФИГУРАЦИИ ИЗ ЯЧЕЕК ЛИСТА
 * ============================================================ */

function readConfig_(mainSheet) {
  var pointsPerTrainee  = parseNumber_(mainSheet.getRange(CFG_POINTS_PER_TRAINEE).getValue())  || 10;
  var pctPerTrainee     = ensureDecimal_(mainSheet.getRange(CFG_PCT_PER_TRAINEE).getValue(),     0.05);
  var pointsPerError    = parseNumber_(mainSheet.getRange(CFG_POINTS_PER_ERROR).getValue())    || 2;
  var pctPerError       = ensureDecimal_(mainSheet.getRange(CFG_PCT_PER_ERROR).getValue(),       0.01);
  var pointsPerProject  = parseNumber_(mainSheet.getRange(CFG_POINTS_PER_PROJECT).getValue())  || 20;
  var pctPerProject     = ensureDecimal_(mainSheet.getRange(CFG_PCT_PER_PROJECT).getValue(),     0.20);
  var pointsPerTraining = parseNumber_(mainSheet.getRange(CFG_POINTS_PER_TRAINING).getValue()) || 40;
  var pctPerTraining    = ensureDecimal_(mainSheet.getRange(CFG_PCT_PER_TRAINING).getValue(),    0.40);

  return {
    pointsPerTrainee:  pointsPerTrainee,
    pctPerTrainee:     pctPerTrainee,
    pointsPerError:    pointsPerError,
    pctPerError:       pctPerError,
    pointsPerProject:  pointsPerProject,
    pctPerProject:     pctPerProject,
    pointsPerTraining: pointsPerTraining,
    pctPerTraining:    pctPerTraining
  };
}

/**
 * Если значение > 1, значит введено как целое число (20 = 20%).
 * Если <= 1, значит уже десятичное (0.05 = 5%).
 * Если пусто или 0, возвращает fallback.
 */
function ensureDecimal_(value, fallback) {
  var n = parseNumber_(value);
  if (!n) return fallback;
  return n > 1 ? n / 100 : n;
}


/* ============================================================
 *  ЧЕКБОКСЫ — установка в колонки E и F
 * ============================================================ */

/**
 * Вставляет чекбоксы (Да/Нет) в колонки E и F для всех строк с данными.
 * Запустить один раз из меню «Нагрузка T&A → Вставить чекбоксы».
 */
function insertCheckboxes() {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(MAIN_SHEET_NAME);
  if (!sheet) return;

  var lastRow = getMainLastDataRow_(sheet, MAIN_START_ROW);
  if (lastRow < MAIN_START_ROW) return;
  var numRows = lastRow - MAIN_START_ROW + 1;

  // Валидация: чекбокс с Да/Нет
  var rule = SpreadsheetApp.newDataValidation()
    .requireCheckbox("Да", "Нет")
    .setAllowInvalid(false)
    .build();

  // Колонка E — Есть обучение
  var rangeE = sheet.getRange(MAIN_START_ROW, MAIN_COL_HAS_TRAIN, numRows, 1);
  rangeE.setDataValidation(rule);
  // Устанавливаем "Нет" только в пустые ячейки
  var valuesE = rangeE.getValues();
  for (var i = 0; i < valuesE.length; i++) {
    if (!valuesE[i][0] || valuesE[i][0] === "" || valuesE[i][0] === false) {
      valuesE[i][0] = "Нет";
    }
  }
  rangeE.setValues(valuesE);

  // Колонка F — Есть проект
  var rangeF = sheet.getRange(MAIN_START_ROW, MAIN_COL_HAS_PROJ, numRows, 1);
  rangeF.setDataValidation(rule);
  var valuesF = rangeF.getValues();
  for (var i = 0; i < valuesF.length; i++) {
    if (!valuesF[i][0] || valuesF[i][0] === "" || valuesF[i][0] === false) {
      valuesF[i][0] = "Нет";
    }
  }
  rangeF.setValues(valuesF);

  // Выравнивание по центру
  rangeE.setHorizontalAlignment("center");
  rangeF.setHorizontalAlignment("center");

  SpreadsheetApp.getUi().alert(
    "Чекбоксы установлены в колонках E и F (строки " +
    MAIN_START_ROW + "–" + lastRow + ").\n" +
    "Нажмите на чекбокс, чтобы переключить Да/Нет."
  );
}


/* ============================================================
 *  ПОДСЧЁТ ОШИБОК ПО СТОЛАМ
 * ============================================================ */

function countNotProcessedByDesk_(childSheet) {
  var lastRow = childSheet.getLastRow();
  var map = new Map();
  if (lastRow < 2) return map;

  var width = Math.max(CHILD_COL_TABLE, CHILD_COL_DATE, CHILD_COL_STATUS_TXT);
  var data  = childSheet.getRange(2, 1, lastRow - 1, width).getValues();

  var now  = new Date();
  var from = new Date(now.getTime() - DAYS_WINDOW * 24 * 60 * 60 * 1000);

  for (var i = 0; i < data.length; i++) {
    var deskKey = normDeskKey_(data[i][CHILD_COL_TABLE - 1]);
    var status  = normStatus_(data[i][CHILD_COL_STATUS_TXT - 1]);
    var d       = toDate_(data[i][CHILD_COL_DATE - 1]);

    if (!deskKey) continue;
    if (status !== STATUS_TARGET) continue;
    if (!d || d < from || d > now) continue;

    map.set(deskKey, (map.get(deskKey) || 0) + 1);
  }
  return map;
}


/* ============================================================
 *  ПОДСЧЁТ СТАЖЁРОВ ПО СТОЛАМ
 * ============================================================ */

function countTraineesByDeskStructured_(staffSheet) {
  var lastRow = staffSheet.getLastRow();
  var map = new Map();
  if (lastRow < 2) return map;

  var data = staffSheet.getRange(1, 1, lastRow, 2).getDisplayValues();
  var currentDesk = "";

  for (var i = 0; i < data.length; i++) {
    var colA = String(data[i][0] || "").trim();
    var colB = String(data[i][1] || "").trim();

    if (isDeskHeader_(colA, colB)) {
      currentDesk = normDeskKey_(colA);
      continue;
    }

    if (normStatus_(colB) !== "стажер") continue;
    if (!currentDesk) continue;

    map.set(currentDesk, (map.get(currentDesk) || 0) + 1);
  }
  return map;
}


/* ============================================================
 *  ГРУППИРОВКА: ТОЧНЫЙ СТОЛ ИЛИ ГРУППА
 * ============================================================ */

function getCountForDeskOrGroup_(countsMap, key) {
  if (/\d/.test(key)) {
    return countsMap.get(key) || 0;
  }
  var sum = 0;
  countsMap.forEach(function (count, deskKey) {
    if (deskKey === key || deskKey.indexOf(key + " ") === 0) {
      sum += count;
    }
  });
  return sum;
}


/* ============================================================
 *  ОПРЕДЕЛЕНИЕ ЗАГОЛОВКА СТОЛА
 * ============================================================ */

function isDeskHeader_(colA, colB) {
  var a = normDeskKey_(colA);
  var b = norm_(colB);
  if (!a || b) return false;

  return /^(египет|марокко|алжир|осн\.стол|турция)(\s+\d+)?$/.test(a)
      || /^интернал$/.test(a)
      || /^бт\s+акк(\s+\d+)?$/.test(a)
      || /^tur-azn/.test(a);
}


/* ============================================================
 *  ФОРМАТИРОВАНИЕ И ДИАГРАММЫ
 * ============================================================ */

function applyConditionalFormatting_() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(MAIN_SHEET_NAME);
  if (!sheet) return;

  sheet.setFrozenRows(2);

  var percentCol = "H";
  var percentRange = sheet.getRange(percentCol + "3:" + percentCol + "1000");
  sheet.clearConditionalFormatRules();

  var rules = [
    SpreadsheetApp.newConditionalFormatRule()
      .whenNumberGreaterThan(0.8)
      .setBackground("#FF5252").setFontColor("#FFFFFF")
      .setRanges([percentRange]).build(),
    SpreadsheetApp.newConditionalFormatRule()
      .whenNumberBetween(0.4, 0.8)
      .setBackground("#FFD740")
      .setRanges([percentRange]).build(),
    SpreadsheetApp.newConditionalFormatRule()
      .whenNumberLessThan(0.4)
      .setBackground("#69F0AE")
      .setRanges([percentRange]).build()
  ];
  sheet.setConditionalFormatRules(rules);
  sheet.autoResizeColumns(1, MAIN_COL_PERCENT);
}

function buildCharts_() {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(MAIN_SHEET_NAME);
  if (!sheet) return;

  var lastRow = getMainLastDataRow_(sheet, MAIN_START_ROW);
  if (lastRow < MAIN_START_ROW) return;

  sheet.getCharts().forEach(function (c) { sheet.removeChart(c); });

  var percentCol = "H";

  // 1. Нагрузка сотрудников
  var chart1 = sheet.newChart()
    .setChartType(Charts.ChartType.PIE)
    .addRange(sheet.getRange("A" + MAIN_START_ROW + ":A" + lastRow))
    .addRange(sheet.getRange(percentCol + MAIN_START_ROW + ":" + percentCol + lastRow))
    .setOption("title", "Нагрузка сотрудников")
    .setOption("pieSliceText", "percentage")
    .setOption("legend.position", "right")
    .setOption("width", 500).setOption("height", 350)
    .setPosition(MAIN_START_ROW, 14, 0, 0)
    .build();
  sheet.insertChart(chart1);

  // 2. Распределение по зонам
  var zoneData = buildZoneData_(sheet, lastRow);
  var tempSheet = getOrCreateHiddenSheet_(ss, "_данные_зон");

  tempSheet.getRange(1, 1).setValue("Зона");
  tempSheet.getRange(1, 2).setValue("Кол-во");
  zoneData.forEach(function (entry, i) {
    tempSheet.getRange(2 + i, 1).setValue(entry[0]);
    tempSheet.getRange(2 + i, 2).setValue(entry[1]);
  });

  var chart2 = sheet.newChart()
    .setChartType(Charts.ChartType.PIE)
    .addRange(tempSheet.getRange(1, 1, zoneData.length + 1, 1))
    .addRange(tempSheet.getRange(1, 2, zoneData.length + 1, 1))
    .setOption("title", "Распределение по зонам")
    .setOption("pieSliceText", "percentage")
    .setOption("legend.position", "right")
    .setOption("width", 500).setOption("height", 350)
    .setPosition(20, 14, 0, 0)
    .build();
  sheet.insertChart(chart2);
}

function buildZoneData_(sheet, lastRow) {
  var data = sheet.getRange("B" + MAIN_START_ROW + ":B" + lastRow).getValues();
  var counts = {};
  data.forEach(function (row) {
    var zone = row[0] || "Без зоны";
    counts[zone] = (counts[zone] || 0) + 1;
  });
  return Object.entries(counts);
}

function getOrCreateHiddenSheet_(ss, name) {
  var existing = ss.getSheetByName(name);
  if (existing) ss.deleteSheet(existing);
  var sh = ss.insertSheet(name);
  sh.hideSheet();
  return sh;
}


/* ============================================================
 *  ПУБЛИЧНАЯ ФУНКЦИЯ: полная настройка листа
 * ============================================================ */

function высчитываниеНагрузки() {
  recalculateMain_();
  applyConditionalFormatting_();
  buildCharts_();
  SpreadsheetApp.getUi().alert("Готово! Данные пересчитаны, диаграммы обновлены.");
}


/* ============================================================
 *  МЕНЮ В ИНТЕРФЕЙСЕ
 * ============================================================ */

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("Нагрузка T&A")
    .addItem("Пересчитать всё",           "высчитываниеНагрузки")
    .addItem("Только пересчёт данных",    "recalculateMain_")
    .addItem("Только диаграммы",          "buildCharts_")
    .addSeparator()
    .addItem("Вставить чекбоксы",         "insertCheckboxes")
    .addItem("Установить триггеры",       "setupTriggers")
    .addSeparator()
    .addItem("Отладка: ошибки по столам", "TA_DebugDeskErrorCounts")
    .addItem("Отладка: стажёры по столам","TA_DebugDeskCounts")
    .addToUi();
}


/* ============================================================
 *  УТИЛИТЫ
 * ============================================================ */

function getMainLastDataRow_(sheet, startRow) {
  var lastRow = sheet.getLastRow();
  if (lastRow < startRow) return startRow - 1;

  var data = sheet.getRange(startRow, 1, lastRow - startRow + 1, 2).getValues();
  var lastFilled = -1;
  for (var i = 0; i < data.length; i++) {
    if (String(data[i][0] || "").trim() || String(data[i][1] || "").trim()) {
      lastFilled = i;
    }
  }
  return lastFilled === -1 ? startRow - 1 : startRow + lastFilled;
}

function norm_(value) {
  return String(value || "")
    .replace(/\u00A0/g, " ")
    .replace(/\s+/g, " ")
    .trim()
    .toLowerCase();
}

function normStatus_(value) {
  return norm_(value).replace(/ё/g, "е");
}

function normDeskKey_(value) {
  var s = norm_(value).replace(/ё/g, "е");
  if (!s) return "";
  s = s.replace(/осн\.?\s*стол/g, "осн.стол");
  return s.replace(/\s+/g, " ").trim();
}

function parseNumber_(value) {
  var n = Number(value);
  return isNaN(n) ? 0 : n;
}

/**
 * Проверяет значение чекбокса.
 * Поддерживает: TRUE/FALSE (стандартный чекбокс),
 *               "Да"/"Нет" (кастомный чекбокс),
 *               "yes"/"no" (английский вариант).
 */
function isChecked_(value) {
  if (value === true) return true;
  if (value === false) return false;
  var s = norm_(value);
  return s === "да" || s === "yes";
}

function toDate_(value) {
  if (value instanceof Date) return value;
  var s = String(value || "").trim();
  if (!s) return null;
  var d = new Date(s);
  return isNaN(d.getTime()) ? null : d;
}


/* ============================================================
 *  ОТЛАДКА
 * ============================================================ */

function TA_DebugDeskErrorCounts() {
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CHILD_SHEET_NAME);
  if (!sh) { Logger.log("Лист не найден: " + CHILD_SHEET_NAME); return; }

  var counts = countNotProcessedByDesk_(sh);
  Logger.log("=== НЕОБРАБОТАННЫЕ ОШИБКИ ПО СТОЛАМ ===");
  counts.forEach(function (count, desk) {
    Logger.log(desk + ": " + count);
  });
}

function TA_DebugDeskCounts() {
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(STAFF_LIST_SHEET_NAME);
  if (!sh) { Logger.log("Лист не найден: " + STAFF_LIST_SHEET_NAME); return; }

  var counts = countTraineesByDeskStructured_(sh);
  Logger.log("=== СТАЖЁРЫ ПО СТОЛАМ ===");
  counts.forEach(function (count, desk) {
    Logger.log(desk + ": " + count);
  });
}
