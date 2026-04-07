/* ============================================================
 *  Нагрузка T&A — автоматизация Google Sheets
 *  v3.0
 *  - Колонка C: "Не обработано" (кол-во со статусом "не обработано")
 *  - Колонка D: "В обработке" (кол-во со статусом "в обработке")
 *  - Исправлен баг: onChange больше не перезаписывает статусы
 *  - Добавлена блокировка LockService против race condition
 *  - Статус читается из колонки J (dropdown), а не R
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
const MAIN_COL_NAME        = 1;  // A — Сотрудник
const MAIN_COL_ZONE        = 2;  // B — Зона
const MAIN_COL_NOT_PROC    = 3;  // C — Не обработано (авто)
const MAIN_COL_IN_PROGRESS = 4;  // D — В обработке (авто)
const MAIN_COL_TRAINEES    = 5;  // E — Стажёры (авто)
const MAIN_COL_PROJECTS    = 6;  // F — Проекты (вручную)
const MAIN_COL_POINTS      = 7;  // G — Баллы (авто)
const MAIN_COL_PERCENT     = 8;  // H — %Нагрузки (авто)

// ─── Ячейки с настройками баллов (на главном листе) ──────────
const CFG_POINTS_PER_TRAINEE = "J2";  // Баллы за стажёра  (10)
const CFG_PCT_PER_TRAINEE    = "J3";  // % за стажёра      (5%)
const CFG_POINTS_PER_ERROR   = "J4";  // Баллы за ошибку   (2)
const CFG_PCT_PER_ERROR      = "J5";  // % за ошибку       (1%)

// ─── Фильтрация ─────────────────────────────────────────────
const STATUS_NOT_PROCESSED = "не обработано";
const STATUS_IN_PROGRESS   = "в обработке";
const DAYS_WINDOW          = 30;

// ─── Максимум баллов (200 баллов = 100%) ─────────────────────
const MAX_POINTS = 200;


/* ============================================================
 *  ТРИГГЕРЫ
 * ============================================================ */

/**
 * Устанавливает onEdit и onChange триггеры.
 * Запускать один раз вручную.
 */
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
    safeRecalculate_();
    return;
  }

  // Лист сотрудников
  if (sheetName === STAFF_LIST_SHEET_NAME) {
    safeRecalculate_();
    return;
  }

  // Главный лист — зона или проекты
  if (sheetName === MAIN_SHEET_NAME && (col === MAIN_COL_ZONE || col === MAIN_COL_PROJECTS)) {
    safeRecalculate_();
  }
}

function onChangeHandler(e) {
  // Не обрабатываем EDIT — это уже делает onEditHandler
  if (e && e.changeType === "EDIT") return;

  // Структурные изменения (вставка/удаление строк и т.д.) — только пересчёт
  safeRecalculate_();
}


/* ============================================================
 *  БЛОКИРОВКА ОТ ПАРАЛЛЕЛЬНОГО ВЫПОЛНЕНИЯ
 * ============================================================ */

function safeRecalculate_() {
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(5000)) {
    Logger.log("safeRecalculate_: не удалось получить блокировку, пропуск");
    return;
  }
  try {
    recalculateMain_();
  } finally {
    lock.releaseLock();
  }
}


/* ============================================================
 *  ЗЕРКАЛИРОВАНИЕ СТАТУСОВ (J → R)
 *  Только при прямом редактировании ячейки в колонке J
 * ============================================================ */

function mirrorStatusToR_(sheet, row) {
  if (row < 2) return;
  var value = sheet.getRange(row, CHILD_COL_STATUS_DD).getValue();
  sheet.getRange(row, CHILD_COL_STATUS_TXT).setValue(String(value || "").trim());
}


/* ============================================================
 *  ГЛАВНЫЙ ПЕРЕСЧЁТ — статусы, стажёры, баллы, %
 * ============================================================ */

/**
 * Единая точка пересчёта: читает все данные один раз,
 * записывает колонки C, D, E, G, H за один проход.
 */
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

  // ── Считаем статусы и стажёров ──
  var statusCounts  = countStatusesByDesk_(childSheet);
  var traineeCounts = countTraineesByDeskStructured_(staffSheet);

  // ── Определяем диапазон данных ──
  var mainLast = getMainLastDataRow_(mainSheet, MAIN_START_ROW);
  if (mainLast < MAIN_START_ROW) return;
  var numRows = mainLast - MAIN_START_ROW + 1;

  // ── Batch-чтение зон и проектов ──
  var mainData = mainSheet.getRange(MAIN_START_ROW, 1, numRows, MAIN_COL_PROJECTS).getDisplayValues();

  // ── Формируем выходные колонки ──
  var outNotProc    = [];
  var outInProgress = [];
  var outTrainees   = [];
  var outPoints     = [];
  var outPercent    = [];

  for (var i = 0; i < numRows; i++) {
    var zoneKey  = normDeskKey_(mainData[i][MAIN_COL_ZONE - 1]);
    var projects = parseNumber_(mainData[i][MAIN_COL_PROJECTS - 1]);

    var notProcessed = zoneKey ? getCountForDeskOrGroup_(statusCounts.notProcessed, zoneKey) : 0;
    var inProgress   = zoneKey ? getCountForDeskOrGroup_(statusCounts.inProgress, zoneKey) : 0;
    var trainees     = zoneKey ? getCountForDeskOrGroup_(traineeCounts, zoneKey) : 0;

    // Баллы считаются по "не обработано" (как раньше по ошибкам)
    var points = notProcessed * cfg.pointsPerError
               + trainees    * cfg.pointsPerTrainee
               + projects;

    var pct = notProcessed * cfg.pctPerError
            + trainees     * cfg.pctPerTrainee;

    if (MAX_POINTS > 0) {
      pct += projects / MAX_POINTS;
    }

    if (pct > 1) pct = 1;

    outNotProc.push([notProcessed]);
    outInProgress.push([inProgress]);
    outTrainees.push([trainees]);
    outPoints.push([points]);
    outPercent.push([pct]);
  }

  // ── Batch-запись ──
  mainSheet.getRange(MAIN_START_ROW, MAIN_COL_NOT_PROC,    numRows, 1).setValues(outNotProc);
  mainSheet.getRange(MAIN_START_ROW, MAIN_COL_IN_PROGRESS, numRows, 1).setValues(outInProgress);
  mainSheet.getRange(MAIN_START_ROW, MAIN_COL_TRAINEES,    numRows, 1).setValues(outTrainees);
  mainSheet.getRange(MAIN_START_ROW, MAIN_COL_POINTS,      numRows, 1).setValues(outPoints);
  mainSheet.getRange(MAIN_START_ROW, MAIN_COL_PERCENT,     numRows, 1).setNumberFormat("0%").setValues(outPercent);
}


/* ============================================================
 *  ЧТЕНИЕ КОНФИГУРАЦИИ ИЗ ЯЧЕЕК ЛИСТА
 * ============================================================ */

function readConfig_(mainSheet) {
  var pointsPerTrainee = parseNumber_(mainSheet.getRange(CFG_POINTS_PER_TRAINEE).getValue()) || 10;
  var pctPerTrainee    = parseNumber_(mainSheet.getRange(CFG_PCT_PER_TRAINEE).getValue())    || 0.05;
  var pointsPerError   = parseNumber_(mainSheet.getRange(CFG_POINTS_PER_ERROR).getValue())   || 2;
  var pctPerError      = parseNumber_(mainSheet.getRange(CFG_PCT_PER_ERROR).getValue())      || 0.01;

  return {
    pointsPerTrainee: pointsPerTrainee,
    pctPerTrainee:    pctPerTrainee,
    pointsPerError:   pointsPerError,
    pctPerError:      pctPerError
  };
}


/* ============================================================
 *  ПОДСЧЁТ СТАТУСОВ ПО СТОЛАМ
 *  Читает из колонки J (dropdown), НЕ из R
 *  Возвращает { notProcessed: Map, inProgress: Map }
 * ============================================================ */

function countStatusesByDesk_(childSheet) {
  var lastRow = childSheet.getLastRow();
  var notProcessed = new Map();
  var inProgress   = new Map();

  if (lastRow < 2) return { notProcessed: notProcessed, inProgress: inProgress };

  var width = Math.max(CHILD_COL_TABLE, CHILD_COL_DATE, CHILD_COL_STATUS_DD);
  var data  = childSheet.getRange(2, 1, lastRow - 1, width).getValues();

  var now  = new Date();
  var from = new Date(now.getTime() - DAYS_WINDOW * 24 * 60 * 60 * 1000);

  for (var i = 0; i < data.length; i++) {
    var deskKey = normDeskKey_(data[i][CHILD_COL_TABLE - 1]);
    var status  = normStatus_(data[i][CHILD_COL_STATUS_DD - 1]);  // J — dropdown
    var d       = toDate_(data[i][CHILD_COL_DATE - 1]);           // O — дата

    if (!deskKey) continue;
    if (!d || d < from || d > now) continue;

    if (status === STATUS_NOT_PROCESSED) {
      notProcessed.set(deskKey, (notProcessed.get(deskKey) || 0) + 1);
    } else if (status === STATUS_IN_PROGRESS) {
      inProgress.set(deskKey, (inProgress.get(deskKey) || 0) + 1);
    }
  }

  return { notProcessed: notProcessed, inProgress: inProgress };
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

  var percentRange = sheet.getRange("H3:H1000");
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

  var chart1 = sheet.newChart()
    .setChartType(Charts.ChartType.PIE)
    .addRange(sheet.getRange("A" + MAIN_START_ROW + ":A" + lastRow))
    .addRange(sheet.getRange("H" + MAIN_START_ROW + ":H" + lastRow))
    .setOption("title", "Нагрузка сотрудников")
    .setOption("pieSliceText", "percentage")
    .setOption("legend.position", "right")
    .setOption("width", 500).setOption("height", 350)
    .setPosition(MAIN_START_ROW, 13, 0, 0)
    .build();
  sheet.insertChart(chart1);

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
    .setPosition(20, 13, 0, 0)
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
    .addItem("Установить триггеры",       "setupTriggers")
    .addSeparator()
    .addItem("Отладка: статусы по столам",  "TA_DebugStatusCounts")
    .addItem("Отладка: стажёры по столам",  "TA_DebugDeskCounts")
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

function TA_DebugStatusCounts() {
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CHILD_SHEET_NAME);
  if (!sh) { Logger.log("Лист не найден: " + CHILD_SHEET_NAME); return; }

  var counts = countStatusesByDesk_(sh);

  Logger.log("=== НЕ ОБРАБОТАНО ПО СТОЛАМ ===");
  counts.notProcessed.forEach(function (count, desk) {
    Logger.log(desk + ": " + count);
  });

  Logger.log("=== В ОБРАБОТКЕ ПО СТОЛАМ ===");
  counts.inProgress.forEach(function (count, desk) {
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
