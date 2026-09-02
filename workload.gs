/* ============================================================
 *  Нагрузка T&A — автоматизация Google Sheets
 *  v5.0
 *  Основано на v4.8 + исправления багов:
 *  - Колонка C: "Не обработано" (вместо "Ошибки")
 *  - Колонка D: "В обработке" (новая)
 *  - Колонки сдвинуты: E-J (было D-I)
 *  - Исправлен баг onChange (не перезаписывает статусы)
 *  - LockService против race condition
 *  - Статус читается из колонки J (dropdown)
 *  - Интернал приравнивается к Осн.Стол при подсчёте
 *  - Поддержка старого и нового нейминга столов
 *  - Сброс смен, часовой пересчёт, dropdowns
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
const MAIN_COL_NAME        = 1;   // A — Сотрудник
const MAIN_COL_ZONE        = 2;   // B — Зона
const MAIN_COL_NOT_PROC    = 3;   // C — Не обработано (авто)
const MAIN_COL_IN_PROGRESS = 4;   // D — В обработке (авто)
const MAIN_COL_TRAINEES    = 5;   // E — Кол-во стажёров (авто)
const MAIN_COL_TRAINING    = 6;   // F — Есть обучение (Да/Нет)
const MAIN_COL_SHIFT       = 7;   // G — Есть ли смена (Да/Нет)
const MAIN_COL_PROJECTS    = 8;   // H — Кол-во проектов
const MAIN_COL_POINTS      = 9;   // I — Нагрузка (ед.)
const MAIN_COL_PERCENT     = 10;  // J — %Нагрузки

// ─── Ячейки с настройками (колонка L) ────────────────────────
const CFG_POINTS_PER_TRAINEE = "L2";
const CFG_PCT_PER_TRAINEE    = "L3";
const CFG_POINTS_PER_ERROR   = "L4";
const CFG_PCT_PER_ERROR      = "L5";
const CFG_POINTS_PER_PROJECT = "L6";
const CFG_PCT_PER_PROJECT    = "L7";
const CFG_PCT_TRAINING       = "L8";

// ─── Фильтрация ─────────────────────────────────────────────
const STATUS_NOT_PROCESSED = "не обработано";
const STATUS_IN_PROGRESS   = "в обработке";
const DAYS_WINDOW          = 30;

// ─── Служебный скрытый лист для диаграмм ─────────────────────
const HIDDEN_CHART_DATA_SHEET = "_chart_data_ta_";

/* ============================================================
 *  ТРИГГЕРЫ
 * ============================================================ */

function setupTriggers() {
  ScriptApp.getProjectTriggers().forEach(function (trigger) {
    var name = trigger.getHandlerFunction();
    if (name === "onEditHandler" || name === "onChangeHandler" || name === "hourlyRecalculate_") {
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

  ScriptApp.newTrigger("hourlyRecalculate_")
    .timeBased()
    .everyHours(1)
    .create();

  Logger.log("Триггеры установлены (onEdit + onChange + каждый час)");
}

function hourlyRecalculate_() {
  safeRecalculate_();
  Logger.log("Часовой пересчёт выполнен: " + new Date());
}

/* ============================================================
 *  ОБРАБОТЧИКИ СОБЫТИЙ
 * ============================================================ */

function onEditHandler(e) {
  if (!e || !e.range) return;

  var sheet     = e.range.getSheet();
  var sheetName = sheet.getName();
  var col       = e.range.getColumn();

  if (sheetName === CHILD_SHEET_NAME) {
    if (col === CHILD_COL_STATUS_DD) {
      mirrorStatusToR_(sheet, e.range.getRow());
    }
    safeRecalculate_();
    return;
  }

  if (sheetName === STAFF_LIST_SHEET_NAME) {
    safeRecalculate_();
    return;
  }

  if (sheetName === MAIN_SHEET_NAME &&
      (col === MAIN_COL_ZONE ||
       col === MAIN_COL_TRAINING ||
       col === MAIN_COL_SHIFT ||
       col === MAIN_COL_PROJECTS)) {
    safeRecalculate_();
  }
}

function onChangeHandler(e) {
  if (e && e.changeType === "EDIT") return;
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
 *  ГЛАВНЫЙ ПЕРЕСЧЁТ
 * ============================================================ */

function recalculateMain_() {
  var ss         = SpreadsheetApp.getActiveSpreadsheet();
  var mainSheet  = ss.getSheetByName(MAIN_SHEET_NAME);
  var childSheet = ss.getSheetByName(CHILD_SHEET_NAME);
  var staffSheet = ss.getSheetByName(STAFF_LIST_SHEET_NAME);

  if (!mainSheet)  { Logger.log("Лист не найден: " + MAIN_SHEET_NAME);  return; }
  if (!childSheet) { Logger.log("Лист не найден: " + CHILD_SHEET_NAME); return; }
  if (!staffSheet) { Logger.log("Лист не найден: " + STAFF_LIST_SHEET_NAME); return; }

  var cfg = readConfig_(mainSheet);

  var statusCounts  = countStatusesByDesk_(childSheet);
  var traineeCounts = countTraineesByDeskStructured_(staffSheet);

  var mainLast = getMainLastDataRow_(mainSheet, MAIN_START_ROW);
  if (mainLast < MAIN_START_ROW) return;
  var numRows = mainLast - MAIN_START_ROW + 1;

  var mainData = mainSheet.getRange(MAIN_START_ROW, 1, numRows, MAIN_COL_PROJECTS).getValues();

  var outNotProc    = [];
  var outInProgress = [];
  var outTrainees   = [];
  var outPoints     = [];
  var outPercent    = [];

  for (var i = 0; i < numRows; i++) {
    var zoneKey     = normDeskKey_(mainData[i][MAIN_COL_ZONE - 1]);
    var isTraining  = isYes_(mainData[i][MAIN_COL_TRAINING - 1]);
    var hasShift    = isYes_(mainData[i][MAIN_COL_SHIFT - 1]);
    var projectsCnt = parseNumber_(mainData[i][MAIN_COL_PROJECTS - 1]);

    var notProcessed = zoneKey ? getCountForDeskOrGroup_(statusCounts.notProcessed, zoneKey) : 0;
    var inProgress   = zoneKey ? getCountForDeskOrGroup_(statusCounts.inProgress, zoneKey) : 0;
    var trainees     = zoneKey ? getCountForDeskOrGroup_(traineeCounts, zoneKey) : 0;

    var points = 0;
    var pct = 0;

    if (hasShift) {
      points = notProcessed * cfg.pointsPerError
             + trainees     * cfg.pointsPerTrainee
             + projectsCnt  * cfg.pointsPerProject;

      if (isTraining) {
        pct = cfg.pctTraining;
      } else {
        pct = notProcessed * cfg.pctPerError
            + trainees     * cfg.pctPerTrainee
            + projectsCnt  * cfg.pctPerProject;

        pct = Math.round(pct * 1000) / 1000;

        if (pct > 1) pct = 1;
        if (pct < 0) pct = 0;
      }
    }

    outNotProc.push([notProcessed]);
    outInProgress.push([inProgress]);
    outTrainees.push([trainees]);
    outPoints.push([points]);
    outPercent.push([pct]);
  }

  mainSheet.getRange(MAIN_START_ROW, MAIN_COL_NOT_PROC, numRows, 1)
    .setNumberFormat("0")
    .setValues(outNotProc);
  mainSheet.getRange(MAIN_START_ROW, MAIN_COL_IN_PROGRESS, numRows, 1)
    .setNumberFormat("0")
    .setValues(outInProgress);
  mainSheet.getRange(MAIN_START_ROW, MAIN_COL_TRAINEES, numRows, 1)
    .setNumberFormat("0")
    .setValues(outTrainees);
  mainSheet.getRange(MAIN_START_ROW, MAIN_COL_POINTS, numRows, 1)
    .setNumberFormat("0")
    .setValues(outPoints);
  mainSheet.getRange(MAIN_START_ROW, MAIN_COL_PERCENT, numRows, 1)
    .setNumberFormat("0.0%")
    .setValues(outPercent);
}

/* ============================================================
 *  ЧТЕНИЕ КОНФИГУРАЦИИ
 * ============================================================ */

function readConfig_(mainSheet) {
  var pointsPerTrainee = parseNumber_(mainSheet.getRange(CFG_POINTS_PER_TRAINEE).getValue()) || 10;
  var pctPerTrainee    = readPercentCell_(mainSheet, CFG_PCT_PER_TRAINEE, 0.05);

  var pointsPerError   = parseNumber_(mainSheet.getRange(CFG_POINTS_PER_ERROR).getValue()) || 2;
  var pctPerError      = readPercentCell_(mainSheet, CFG_PCT_PER_ERROR, 0.01);

  var pointsPerProject = parseNumber_(mainSheet.getRange(CFG_POINTS_PER_PROJECT).getValue()) || 20;
  var pctPerProject    = readPercentCell_(mainSheet, CFG_PCT_PER_PROJECT, 0.20);

  var pctTraining      = readPercentCell_(mainSheet, CFG_PCT_TRAINING, 1.00);

  return {
    pointsPerTrainee: pointsPerTrainee,
    pctPerTrainee: pctPerTrainee,
    pointsPerError: pointsPerError,
    pctPerError: pctPerError,
    pointsPerProject: pointsPerProject,
    pctPerProject: pctPerProject,
    pctTraining: pctTraining
  };
}

function readPercentCell_(sheet, a1, fallback) {
  var range = sheet.getRange(a1);

  var rawValue = range.getValue();
  var display  = String(range.getDisplayValue() || "").trim().replace(",", ".");

  if (!display) return fallback;

  if (display.indexOf("%") !== -1) {
    var n1 = parseFloat(display.replace("%", "").trim());
    return isNaN(n1) ? fallback : n1 / 100;
  }

  var n2 = parseFloat(String(rawValue).replace(",", "."));
  if (isNaN(n2)) return fallback;

  return n2 / 100;
}

/* ============================================================
 *  ВЫПАДАЮЩИЕ СПИСКИ — F и G
 * ============================================================ */

function insertTrainingDropdown() {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(MAIN_SHEET_NAME);
  if (!sheet) return;

  var lastRow = getMainLastDataRow_(sheet, MAIN_START_ROW);
  if (lastRow < MAIN_START_ROW) return;
  var numRows = lastRow - MAIN_START_ROW + 1;

  var rule = SpreadsheetApp.newDataValidation()
    .requireValueInList(["Да", "Нет"], true)
    .setAllowInvalid(false)
    .build();

  var range = sheet.getRange(MAIN_START_ROW, MAIN_COL_TRAINING, numRows, 1);
  range.setDataValidation(rule);

  var values = range.getValues();
  for (var i = 0; i < values.length; i++) {
    var v = String(values[i][0] || "").trim().toLowerCase();
    if (v !== "да" && v !== "нет") {
      values[i][0] = "Нет";
    }
  }
  range.setValues(values);
  range.setHorizontalAlignment("center");

  SpreadsheetApp.getUi().alert(
    "Выпадающий список «Да/Нет» установлен в колонке F (Есть обучение)."
  );
}

function insertShiftDropdown() {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(MAIN_SHEET_NAME);
  if (!sheet) return;

  var lastRow = getMainLastDataRow_(sheet, MAIN_START_ROW);
  if (lastRow < MAIN_START_ROW) return;
  var numRows = lastRow - MAIN_START_ROW + 1;

  var rule = SpreadsheetApp.newDataValidation()
    .requireValueInList(["Да", "Нет"], true)
    .setAllowInvalid(false)
    .build();

  var range = sheet.getRange(MAIN_START_ROW, MAIN_COL_SHIFT, numRows, 1);
  range.setDataValidation(rule);

  var values = range.getValues();
  for (var i = 0; i < values.length; i++) {
    var v = String(values[i][0] || "").trim().toLowerCase();
    if (v !== "да" && v !== "нет") {
      values[i][0] = "Нет";
    }
  }

  range.setValues(values);
  range.setHorizontalAlignment("center");

  SpreadsheetApp.getUi().alert(
    "Выпадающий список «Да/Нет» установлен в колонке G (Есть ли смена)."
  );
}

/* ============================================================
 *  ПОДСЧЁТ СТАТУСОВ ПО СТОЛАМ
 *  Статус читается из колонки J (dropdown)
 *  Интернал приравнивается к Осн.Стол
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
    var row = data[i];

    var deskKey = normErrorDeskKey_(row[CHILD_COL_TABLE - 1]);
    var status  = normStatus_(row[CHILD_COL_STATUS_DD - 1]); // J
    var d       = toDate_(row[CHILD_COL_DATE - 1]);           // O

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

function normErrorDeskKey_(value) {
  var s = normDeskKey_(value);

  // Интернал считаем как основной стол
  if (s === "internal") return "mt";

  return s;
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
  var normalizedKey = normDeskKey_(key);

  if (normalizedKey === "egp") return sumByPrefix_(countsMap, "egp");
  if (normalizedKey === "mar") return sumByPrefix_(countsMap, "mar");
  if (normalizedKey === "alg") return sumByPrefix_(countsMap, "alg");
  if (normalizedKey === "mt")  return sumByPrefix_(countsMap, "mt");
  if (normalizedKey === "tur") return sumByPrefix_(countsMap, "tur");
  if (normalizedKey === "internal") return countsMap.get("internal") || 0;

  if (countsMap.has(normalizedKey)) {
    return countsMap.get(normalizedKey) || 0;
  }

  return sumByPrefix_(countsMap, normalizedKey);
}

function sumByPrefix_(countsMap, prefix) {
  var sum = 0;
  countsMap.forEach(function(count, deskKey) {
    if (deskKey === prefix || deskKey.indexOf(prefix) === 0) {
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

  return /^egp/.test(a)
      || /^mar/.test(a)
      || /^alg/.test(a)
      || /^mt/.test(a)
      || /^tur/.test(a)
      || /^internal$/.test(a)
      || /^бт\s+акк(\s+\d+)?$/.test(a);
}

/* ============================================================
 *  ФОРМАТИРОВАНИЕ И ДИАГРАММЫ
 * ============================================================ */

function applyConditionalFormatting_() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(MAIN_SHEET_NAME);
  if (!sheet) return;

  sheet.setFrozenRows(2);

  var percentRange = sheet.getRange("J3:J1000");
  sheet.clearConditionalFormatRules();

  var rules = [
    SpreadsheetApp.newConditionalFormatRule()
      .whenNumberGreaterThan(0.8)
      .setBackground("#FF5252")
      .setFontColor("#FFFFFF")
      .setRanges([percentRange])
      .build(),

    SpreadsheetApp.newConditionalFormatRule()
      .whenNumberBetween(0.4, 0.8)
      .setBackground("#FFD740")
      .setRanges([percentRange])
      .build(),

    SpreadsheetApp.newConditionalFormatRule()
      .whenNumberLessThan(0.4)
      .setBackground("#69F0AE")
      .setRanges([percentRange])
      .build()
  ];

  sheet.setConditionalFormatRules(rules);
}

function buildCharts_() {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(MAIN_SHEET_NAME);
  if (!sheet) return;

  var lastRow = getMainLastDataRow_(sheet, MAIN_START_ROW);
  if (lastRow < MAIN_START_ROW) return;

  sheet.getCharts().forEach(function(chart) {
    sheet.removeChart(chart);
  });

  var tempSheet = getOrCreateHiddenSheet_(ss, HIDDEN_CHART_DATA_SHEET);

  var employeeNames = sheet.getRange("A" + MAIN_START_ROW + ":A" + lastRow).getDisplayValues();
  var employeeLoads = sheet.getRange("J" + MAIN_START_ROW + ":J" + lastRow).getValues();

  tempSheet.getRange(1, 1).setValue("Сотрудник");
  tempSheet.getRange(1, 2).setValue("%Нагрузки");

  if (employeeNames.length > 0) {
    tempSheet.getRange(2, 1, employeeNames.length, 1).setValues(employeeNames);
    tempSheet.getRange(2, 2, employeeLoads.length, 1).setValues(employeeLoads);
    tempSheet.getRange(2, 2, employeeLoads.length, 1).setNumberFormat("0.0%");
  }

  var zoneData = buildZoneData_(sheet, lastRow);
  var zoneStartRow = 1;
  var zoneStartCol = 4;

  tempSheet.getRange(zoneStartRow, zoneStartCol).setValue("Зона");
  tempSheet.getRange(zoneStartRow, zoneStartCol + 1).setValue("Кол-во");

  if (zoneData.length > 0) {
    tempSheet.getRange(zoneStartRow + 1, zoneStartCol, zoneData.length, 2).setValues(zoneData);
  }

  var chartAnchorCol = getChartAnchorColumn_(sheet);

  var chart1 = sheet.newChart()
    .setChartType(Charts.ChartType.PIE)
    .addRange(tempSheet.getRange(1, 1, employeeNames.length + 1, 1))
    .addRange(tempSheet.getRange(1, 2, employeeLoads.length + 1, 1))
    .setOption("title", "Нагрузка сотрудников")
    .setOption("pieSliceText", "percentage")
    .setOption("legend.position", "right")
    .setOption("width", 500)
    .setOption("height", 330)
    .setPosition(3, chartAnchorCol, 0, 0)
    .build();
  sheet.insertChart(chart1);

  var chart2 = sheet.newChart()
    .setChartType(Charts.ChartType.PIE)
    .addRange(tempSheet.getRange(zoneStartRow, zoneStartCol, zoneData.length + 1, 1))
    .addRange(tempSheet.getRange(zoneStartRow, zoneStartCol + 1, zoneData.length + 1, 1))
    .setOption("title", "Распределение по зонам")
    .setOption("pieSliceText", "percentage")
    .setOption("legend.position", "right")
    .setOption("width", 500)
    .setOption("height", 330)
    .setPosition(21, chartAnchorCol, 0, 0)
    .build();
  sheet.insertChart(chart2);
}

function getChartAnchorColumn_(sheet) {
  var baseRightCol = 12;
  var actualLastCol = sheet.getLastColumn();
  return Math.max(baseRightCol, actualLastCol) + 2;
}

function buildZoneData_(sheet, lastRow) {
  var data = sheet.getRange("B" + MAIN_START_ROW + ":B" + lastRow).getDisplayValues();
  var counts = {};

  data.forEach(function(row) {
    var zone = String(row[0] || "").trim();
    if (!zone || zone === "-") zone = "Без зоны";
    counts[zone] = (counts[zone] || 0) + 1;
  });

  return Object.keys(counts).map(function(zone) {
    return [zone, counts[zone]];
  });
}

function getOrCreateHiddenSheet_(ss, name) {
  var sh = ss.getSheetByName(name);

  if (sh) {
    sh.clear();
  } else {
    sh = ss.insertSheet(name);
  }

  sh.hideSheet();
  return sh;
}

/* ============================================================
 *  ПУБЛИЧНАЯ ФУНКЦИЯ
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
    .addItem("Пересчитать всё",             "высчитываниеНагрузки")
    .addItem("Только пересчёт данных",      "recalculateMain_")
    .addItem("Только диаграммы",            "buildCharts_")
    .addItem("Установить триггеры",         "setupTriggers")
    .addItem("Установить сброс смен (23:00)", "setupShiftResetTrigger")
    .addSeparator()
    .addItem("Dropdown: Обучение (кол. F)",  "insertTrainingDropdown")
    .addItem("Dropdown: Смена (кол. G)",     "insertShiftDropdown")
    .addSeparator()
    .addItem("Отладка: статусы по столам",   "TA_DebugStatusCounts")
    .addItem("Отладка: стажёры по столам",   "TA_DebugDeskCounts")
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
    .replace(/ /g, " ")
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

  if (/^египет/.test(s) || /^egp/.test(s)) return "egp";
  if (/^марокко/.test(s) || /^mar/.test(s)) return "mar";
  if (/^алжир/.test(s) || /^alg/.test(s)) return "alg";
  if (/^осн\.?\s*стол/.test(s) || /^mt/.test(s)) return "mt";
  if (/tur/.test(s)) return "tur";
  if (/^интернал/.test(s) || /^internal/.test(s)) return "internal";

  return s;
}

function parseNumber_(value) {
  var n = Number(value);
  return isNaN(n) ? 0 : n;
}

function isYes_(value) {
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
 *  СБРОС СМЕН И ТРИГГЕР
 * ============================================================ */

function resetAllShiftsToNo() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(MAIN_SHEET_NAME);
  if (!sheet) return;

  var lastRow = getMainLastDataRow_(sheet, MAIN_START_ROW);
  if (lastRow < MAIN_START_ROW) return;

  var numRows = lastRow - MAIN_START_ROW + 1;
  var range = sheet.getRange(MAIN_START_ROW, MAIN_COL_SHIFT, numRows, 1);

  var values = Array.from({ length: numRows }, function () {
    return ["Нет"];
  });

  range.setValues(values);

  safeRecalculate_();

  Logger.log("Все значения в колонке 'Есть ли смена' сброшены на 'Нет'");
}

function setupShiftResetTrigger() {
  ScriptApp.getProjectTriggers().forEach(function(trigger) {
    if (trigger.getHandlerFunction() === "resetAllShiftsToNo") {
      ScriptApp.deleteTrigger(trigger);
    }
  });

  ScriptApp.newTrigger("resetAllShiftsToNo")
    .timeBased()
    .everyDays(1)
    .atHour(23)
    .create();

  Logger.log("Триггер ежедневного сброса смен установлен на 23:00");
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
