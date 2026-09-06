// Глобальная переменная для запоминания выбора (на время сессии)
var rememberChoice = {
  enabled: false,
  choice: null
};

// ОСНОВНАЯ ФУНКЦИЯ - Обновление матрицы трассировки с динамическими тестами
function updateTraceabilityMatrixWithBlocked() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var regressionSheet = spreadsheet.getSheetByName("Регрессионное тестирование");
  var matrixSheet = spreadsheet.getSheetByName("Матрица трассировки");

    // АВТОМАТИЧЕСКОЕ ОКРАШИВАНИЕ перед обновлением
  autoColorColumnA();
  
  var regressionData = regressionSheet.getDataRange().getValues();
  
  // ДИНАМИЧЕСКОЕ ОПРЕДЕЛЕНИЕ ТЕСТОВ: собираем все уникальные testId из регрессии
  var allTestIds = new Set();
  var traceabilityMap = {};
  var testStatusMap = {};
  
  // Собираем данные из регрессионного листа
  for (var i = 1; i < regressionData.length; i++) {
    var row = regressionData[i];
    var rowNum = i + 1;
    
    // Левый блок
    var leftTestId = parseInt(row[0]);
    var leftReqIds = parseRequirementIds(row[11]);
    var leftBlocked = isTestBlocked(regressionSheet, rowNum, 6);
    
    // Правый блок  
    var rightTestId = parseInt(row[12]);
    var rightReqIds = parseRequirementIds(row[23]);
    var rightBlocked = isTestBlocked(regressionSheet, rowNum, 23);
    
    // Добавляем testId в общий набор
    if (leftTestId && !isNaN(leftTestId)) {
      allTestIds.add(leftTestId);
      testStatusMap[leftTestId] = leftBlocked;
      
      if (!traceabilityMap[leftTestId]) {
        traceabilityMap[leftTestId] = [];
      }
      traceabilityMap[leftTestId] = traceabilityMap[leftTestId].concat(leftReqIds);
    }
    
    if (rightTestId && !isNaN(rightTestId)) {
      allTestIds.add(rightTestId);
      testStatusMap[rightTestId] = rightBlocked;
      
      if (!traceabilityMap[rightTestId]) {
        traceabilityMap[rightTestId] = [];
      }
      traceabilityMap[rightTestId] = traceabilityMap[rightTestId].concat(rightReqIds);
    }
  }
  
  // Преобразуем Set в массив и сортируем
  var sortedTestIds = Array.from(allTestIds).sort((a, b) => a - b);
  
  // ОБНОВЛЯЕМ ЗАГОЛОВКИ МАТРИЦЫ
  updateMatrixHeaders(matrixSheet, sortedTestIds);
  
  // Получаем обновленные заголовки после изменения
  var lastMatrixColumn = matrixSheet.getLastColumn();
  var matrixHeaders = matrixSheet.getRange(4, 5, 1, lastMatrixColumn - 4).getValues()[0];
  
  // Получаем требования
  var requirementRange = matrixSheet.getRange(8, 1, matrixSheet.getLastRow() - 7, 1);
  var requirementData = requirementRange.getValues();
  var requirementIds = [];
  var requirementRowMap = {};
  
  for (var i = 0; i < requirementData.length; i++) {
    var reqId = requirementData[i][0];
    var actualRow = i + 8;
    
    if (typeof reqId === 'number' && !isNaN(reqId)) {
      requirementIds.push(reqId);
      requirementRowMap[reqId] = actualRow;
    }
  }
  
  // Убираем дубликаты требований для каждого теста
  for (var testId in traceabilityMap) {
    traceabilityMap[testId] = [...new Set(traceabilityMap[testId])];
  }
  
  // Очищаем матрицу перед обновлением
  var clearRange = matrixSheet.getRange(8, 5, requirementData.length, matrixHeaders.length);
  clearRange.clearContent();
  
  // Обновляем матрицу с новой логикой
  var requirementCoverage = {};
  
  for (var col = 0; col < matrixHeaders.length; col++) {
    var testId = parseInt(matrixHeaders[col]);
    
    if (isNaN(testId) || !traceabilityMap[testId]) continue;
    
    var isBlockedTest = testStatusMap[testId] === true;
    
    for (var r = 0; r < traceabilityMap[testId].length; r++) {
      var requirementId = traceabilityMap[testId][r];
      var targetRow = requirementRowMap[requirementId];
      
      if (targetRow) {
        // Инициализируем данные для требования
        if (!requirementCoverage[requirementId]) {
          requirementCoverage[requirementId] = {
            activeTests: 0,
            blockedTests: 0,
            totalTests: 0
          };
        }
        
        // Увеличиваем счетчики
        requirementCoverage[requirementId].totalTests++;
        if (isBlockedTest) {
          requirementCoverage[requirementId].blockedTests++;
        } else {
          requirementCoverage[requirementId].activeTests++;
        }
        
        // Ставим символ в матрицу
        var cell = matrixSheet.getRange(targetRow, col + 5);
        if (isBlockedTest) {
          cell.setValue("⏸️");
          cell.setBackground("#FFF9C4");
          cell.setFontColor("#7B6D00");
        } else {
          cell.setValue("✓");
          cell.setBackground("#E6F4EA");
          cell.setFontColor("#137333");
        }
      }
    }
  }
  // Добавляем в функцию updateTraceabilityMatrixWithBlocked после основного заполнения:

// Добавляем в функцию updateTraceabilityMatrixWithBlocked после основного заполнения:

function addCoverageCounters(matrixSheet, requirementIds, requirementCoverage) {
  var lastColumn = matrixSheet.getLastColumn();
  var coverageColumn = lastColumn + 1;
  
  // Заголовок для столбца с количеством покрытий
  matrixSheet.getRange(4, coverageColumn).setValue("Кол-во покрытий");
  matrixSheet.getRange(4, coverageColumn).setBackground("#6A0DAD");
  matrixSheet.getRange(4, coverageColumn).setFontColor("white");
  matrixSheet.setColumnWidth(coverageColumn, 120);
  
  // Для каждого требования считаем общее количество покрывающих тестов
  for (var i = 0; i < requirementIds.length; i++) {
    var reqId = requirementIds[i];
    var targetRow = requirementRowMap[reqId];
    
    if (targetRow && requirementCoverage[reqId]) {
      var totalTests = requirementCoverage[reqId].totalTests;
      var cell = matrixSheet.getRange(targetRow, coverageColumn);
      cell.setValue(totalTests);
      
      // Цветовое кодирование
      if (totalTests === 0) {
        cell.setBackground("#FCE8E6"); // красный - нет покрытия
        cell.setFontColor("#C5221F");
      } else if (totalTests === 1) {
        cell.setBackground("#FFF9C4"); // желтый - минимальное покрытие
        cell.setFontColor("#7B6D00");
      } else {
        cell.setBackground("#E6F4EA"); // зеленый - хорошее покрытие
        cell.setFontColor("#137333");
      }
    }
  }
}

  // Применяем цветовое кодирование к требованиям
  applyRequirementColorsFixed(matrixSheet, requirementIds, requirementCoverage);
  applySimpleVisuals(matrixSheet);
  // ВАЖНО: передаем requirementRowMap в addCoverageCounters!
  addCoverageCounters(matrixSheet, requirementIds, requirementCoverage)
    // ДОБАВЬ ЭТУ СТРОЧКУ:
  freezeCoverageColumn();
  addNavigationHelp();
  showSmartCoverageStats();
}



// ФУНКЦИЯ ОБНОВЛЕНИЯ ЗАГОЛОВКОВ МАТРИЦЫ
function updateMatrixHeaders(matrixSheet, testIds) {
  var headerRow = 4;
  var startCol = 5; // Столбец E
  
  // Очищаем старые заголовки
  var lastCol = matrixSheet.getLastColumn();
  if (lastCol >= startCol) {
    matrixSheet.getRange(headerRow, startCol, 1, lastCol - startCol + 1).clearContent();
  }
  
    // Записываем новые заголовки
  if (testIds.length > 0) {
    var headerValues = [testIds];
    matrixSheet.getRange(headerRow, startCol, 1, testIds.length).setValues(headerValues);
  }
  
  // ДОБАВЬТЕ ЭТОТ КОД ДЛЯ АВТО-ПОДБОРА ШИРИНЫ:
  // Авто-подбор ширины для столбцов с тестами
  for (var i = 0; i < testIds.length; i++) {
    var column = startCol + i;
    matrixSheet.autoResizeColumn(column);
 }
}

 // ФУНКЦИЯ ДЛЯ ОПРЕДЕЛЕНИЯ ЗАБЛОКИРОВАННЫХ ТЕСТОВ
 function isTestBlocked(sheet, row, col) {
  try {
    var cell = sheet.getRange(row, col);
    var backgroundColor = cell.getBackground();
    
    // ТОЧНЫЙ ЦВЕТ ДЛЯ ЗАБЛОКИРОВАННЫХ ТЕСТОВ: #e69138
    return backgroundColor === '#e69138';
  } catch (e) {
    Logger.log('Ошибка при проверке цвета: ' + e.toString());
    return false;
  }
 }


// ФУНКЦИЯ ЦВЕТОВОГО КОДИРОВАНИЯ ТРЕБОВАНИЙ
function applyRequirementColorsFixed(matrixSheet, requirementIds, requirementCoverage) {
  var requirementData = matrixSheet.getRange(8, 1, matrixSheet.getLastRow() - 7, 1).getValues();
  var headerColor = matrixSheet.getRange("B6").getBackground();
  
  for (var i = 0; i < requirementData.length; i++) {
    var requirementCell = matrixSheet.getRange(i + 8, 1);
    var requirementId = requirementData[i][0];
    var cellColor = matrixSheet.getRange(i + 8, 1).getBackground();
    
    // Пропускаем заголовки
    if (cellColor === headerColor) continue;
    
    if (typeof requirementId === 'number' && !isNaN(requirementId)) {
      var coverage = requirementCoverage[requirementId];
      
      if (coverage) {
        if (coverage.activeTests > 0) {
          // Есть активные тесты - ЗЕЛЕНЫЙ
          requirementCell.setBackground("#E6F4EA");
          requirementCell.setFontColor("#137333");
          requirementCell.setFontWeight("bold");
        } else if (coverage.blockedTests > 0) {
          // Только заблокированные тесты - ЖЕЛТЫЙ
          requirementCell.setBackground("#FFF9C4");
          requirementCell.setFontColor("#7B6D00");
          requirementCell.setFontWeight("bold");
        } else {
          // Нет тестов - КРАСНЫЙ
          requirementCell.setBackground("#FCE8E6");
          requirementCell.setFontColor("#C5221F");
          requirementCell.setFontWeight("bold");
        }
      } else {
        // Нет покрытия - КРАСНЫЙ
        requirementCell.setBackground("#FCE8E6");
        requirementCell.setFontColor("#C5221F");
        requirementCell.setFontWeight("bold");
      }
    }
  }
}

// ОБЪЕДИНЕННАЯ ФУНКЦИЯ СТАТИСТИКИ
function showSmartCoverageStats() {
  var matrixSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Матрица трассировки");
  
  var requirementData = matrixSheet.getRange(8, 1, matrixSheet.getLastRow() - 7, 4).getValues();
  var testData = matrixSheet.getRange(8, 5, matrixSheet.getLastRow() - 7, matrixSheet.getLastColumn() - 4).getValues();
  
  var coverageStats = {
    total: { requirements: 0, covered: 0, wellCovered: 0 },
    byType: {
      'UI': { requirements: 0, covered: 0, wellCovered: 0, minTests: 1 },
      'A11Y': { requirements: 0, covered: 0, wellCovered: 0, minTests: 1 },
      'NAV': { requirements: 0, covered: 0, wellCovered: 0, minTests: 1 },
      'ST': { requirements: 0, covered: 0, wellCovered: 0, minTests: 1 },
      'FUNC': { requirements: 0, covered: 0, wellCovered: 0, minTests: 2 },
      'DAT': { requirements: 0, covered: 0, wellCovered: 0, minTests: 2 },
      'NOT': { requirements: 0, covered: 0, wellCovered: 0, minTests: 2 },
      'APP': { requirements: 0, covered: 0, wellCovered: 0, minTests: 1 },
      'FIL': { requirements: 0, covered: 0, wellCovered: 0, minTests: 2 }
    }
  };

  // Счетчики для заблокированных требований
  var blockedDetails = [];
  var activeCovered = 0;
  var blockedCovered = 0;
  var notCovered = 0;
  
  for (var i = 0; i < requirementData.length; i++) {
    var reqId = requirementData[i][0];
    var reqName = requirementData[i][1];
    var reqType = requirementData[i][2];
    
    if (typeof reqId !== 'number' || isNaN(reqId)) continue;
    
    coverageStats.total.requirements++;
    
    var detectedType = detectRequirementType(reqType, reqName);
    if (!detectedType) detectedType = 'FUNC';
    
    if (!coverageStats.byType[detectedType]) {
      coverageStats.byType[detectedType] = { requirements: 0, covered: 0, wellCovered: 0, minTests: 2 };
    }
    
    coverageStats.byType[detectedType].requirements++;
    
    var testCount = 0;
    var hasActiveTests = false;
    var hasBlockedTests = false;
    
    for (var j = 0; j < testData[i].length; j++) {
      if (testData[i][j] === "✓") {
        testCount++;
        hasActiveTests = true;
      } else if (testData[i][j] === "⏸️") {
        testCount++;
        hasBlockedTests = true;
      }
    }
    
    if (testCount > 0) {
      coverageStats.total.covered++;
      coverageStats.byType[detectedType].covered++;
      
      var minTestsForType = coverageStats.byType[detectedType].minTests;
      if (testCount >= minTestsForType) {
        coverageStats.total.wellCovered++;
        coverageStats.byType[detectedType].wellCovered++;
      }

      // Учитываем активные и заблокированные
      if (hasActiveTests) {
        activeCovered++;
      } else if (hasBlockedTests) {
        blockedCovered++;
        blockedDetails.push("Требование " + reqId);
      }
    } else {
      notCovered++;
    }
  }
  
  // Показываем объединенную статистику
  showUnifiedStatsDialog(coverageStats, activeCovered, blockedCovered, notCovered, blockedDetails);
}

// ФУНКЦИЯ ДЛЯ ОТОБРАЖЕНИЯ ОБЪЕДИНЕННОЙ СТАТИСТИКИ
function showUnifiedStatsDialog(coverageStats, activeCovered, blockedCovered, notCovered, blockedDetails) {
  var totalRequirements = coverageStats.total.requirements;
  var totalCovered = coverageStats.total.covered;
  
  var totalCoveragePercent = totalRequirements > 0 ? (totalCovered / totalRequirements * 100).toFixed(2) : "0.00";
  var activeCoveragePercent = totalRequirements > 0 ? (activeCovered / totalRequirements * 100).toFixed(2) : "0.00";
  var wellCoveredPercent = totalRequirements > 0 ? (coverageStats.total.wellCovered / totalRequirements * 100).toFixed(2) : "0.00";
  
  var blockedList = blockedDetails.length > 0 ? blockedDetails.slice(0, 10).join(', ') : 'нет';
  if (blockedDetails.length > 10) {
    blockedList += '... (всего: ' + blockedDetails.length + ')';
  }
  
  var typeDetails = '';
  for (var type in coverageStats.byType) {
    var typeData = coverageStats.byType[type];
    if (typeData.requirements > 0) {
      var coveredPercent = (typeData.covered / typeData.requirements * 100).toFixed(1);
      var wellCoveredPercentType = (typeData.wellCovered / typeData.requirements * 100).toFixed(1);
      typeDetails += `
        <div style="margin: 8px 0; padding: 8px; background: #f8f9fa; border-radius: 5px;">
          <strong>${getTypeDisplayName(type)}</strong><br>
          <span style="font-size: 0.9em;">
            Требований: ${typeData.requirements} | 
            Покрыто: ${typeData.covered} (${coveredPercent}%) |
            Хорошо: ${typeData.wellCovered} (${wellCoveredPercentType}%)
          </span>
        </div>`;
    }
  }
  
  var htmlOutput = HtmlService.createHtmlOutput(
    '<div style="font-family: Arial; width: 650px; padding: 20px; background: #f8f9fa; color: #333; border-radius: 10px; border: 1px solid #ddd;">' +
    '<h2 style="margin: 0 0 20px 0; text-align: center; color: #4285F4;">📊 Умная статистика покрытия</h2>' +
    
    // ОБЩАЯ СТАТИСТИКА
    '<div style="background: white; padding: 15px; border-radius: 8px; margin-bottom: 15px; box-shadow: 0 2px 4px rgba(0,0,0,0.1);">' +
    '<div style="text-align: center; margin-bottom: 15px; padding: 10px; background: #E8F0FE; border-radius: 5px;">' +
    '<div style="font-size: 1.2em; font-weight: bold; color: #4285F4;">Общее покрытие: ' + totalCoveragePercent + '%</div>' +
    '<div style="font-size: 0.9em;">Активное: ' + activeCoveragePercent + '% | Хорошо покрыто: ' + wellCoveredPercent + '%</div>' +
    '</div>' +
    
    '<div style="display: grid; grid-template-columns: 1fr 1fr 1fr; gap: 10px; margin-bottom: 15px;">' +
    '<div style="padding: 10px; background: #E6F4EA; border-radius: 5px; text-align: center;">' +
    '<div style="font-size: 1.1em; font-weight: bold; color: #137333;">' + totalCovered + '</div>' +
    '<div style="font-size: 0.8em;">Всего покрыто</div>' +
    '</div>' +
    '<div style="padding: 10px; background: #4CAF50; border-radius: 5px; text-align: center;">' +
    '<div style="font-size: 1.1em; font-weight: bold; color: white;">' + activeCovered + '</div>' +
    '<div style="font-size: 0.8em; color: white;">Активно</div>' +
    '</div>' +
    '<div style="padding: 10px; background: #FFF9C4; border-radius: 5px; text-align: center;">' +
    '<div style="font-size: 1.1em; font-weight: bold; color: #7B6D00;">' + blockedCovered + '</div>' +
    '<div style="font-size: 0.8em;">Заблокировано</div>' +
    '</div>' +
    '</div>' +
    
    '<div style="display: grid; grid-template-columns: 1fr 1fr; gap: 10px;">' +
    '<div style="padding: 10px; background: #FFF3CD; border-radius: 5px; text-align: center;">' +
    '<div style="font-size: 1.1em; font-weight: bold; color: #856404;">' + coverageStats.total.wellCovered + '</div>' +
    '<div style="font-size: 0.8em;">Хорошо покрыто</div>' +
    '</div>' +
    '<div style="padding: 10px; background: #FCE8E6; border-radius: 5px; text-align: center;">' +
    '<div style="font-size: 1.1em; font-weight: bold; color: #C5221F;">' + notCovered + '</div>' +
    '<div style="font-size: 0.8em;">Не покрыто</div>' +
    '</div>' +
    '</div>' +
    '</div>' +
    
    // ЗАБЛОКИРОВАННЫЕ ТРЕБОВАНИЯ
    (blockedDetails.length > 0 ? 
    '<div style="background: #FFF9C4; padding: 15px; border-radius: 8px; margin-bottom: 15px; border: 1px solid #FFD54F;">' +
    '<h3 style="margin: 0 0 10px 0; color: #7B6D00;">🟡 Заблокированные требования:</h3>' +
    '<div style="font-size: 0.9em; max-height: 80px; overflow-y: auto; background: white; padding: 10px; border-radius: 5px;">' + blockedList + '</div>' +
    '</div>' : '') +
    
    // СТАТИСТИКА ПО ТИПАМ
    '<div style="background: white; padding: 15px; border-radius: 8px; margin-bottom: 15px; box-shadow: 0 2px 4px rgba(0,0,0,0.1);">' +
    '<h3 style="margin: 0 0 15px 0; color: #4285F4;">📈 Детали по типам требований:</h3>' +
    '<div style="max-height: 200px; overflow-y: auto;">' + typeDetails + '</div>' +
    '</div>' +
    
    // КНОПКИ
    '<div style="text-align: center;">' +
    '<button onclick="google.script.run.withSuccessHandler(google.script.host.close).suggestTestCases()" style="background: #4CAF50; color: white; border: none; padding: 10px 20px; border-radius: 5px; cursor: pointer; margin-right: 10px; font-size: 14px;">💡 Предложить тесты</button>' +
    '<button onclick="google.script.run.withSuccessHandler(google.script.host.close).showUncoveredDetails()" style="background: #F44336; color: white; border: none; padding: 10px 20px; border-radius: 5px; cursor: pointer; font-size: 14px;">🔍 Детали</button>' +
    '</div>' +
    '</div>'
  )
  .setWidth(680)
  .setHeight(650);
  
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'Умная статистика покрытия');
}

// УМНАЯ ФУНКЦИЯ ПЕРЕНУМЕРАЦИИ ТРЕБОВАНИЙ
function autoRenumberRequirements() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var matrixSheet = spreadsheet.getSheetByName("Матрица трассировки");
  
  var startRow = 8;
  var lastRow = matrixSheet.getLastRow();
  var requirementRange = matrixSheet.getRange(startRow, 1, lastRow - startRow + 1, 1);
  var requirementData = requirementRange.getValues();
  
  var headerColor = matrixSheet.getRange("B6").getBackground();
  
  var newNumbers = [];
  var hasChanges = false;
  var currentNumber = 1;
  
  for (var i = 0; i < requirementData.length; i++) {
    var currentValue = requirementData[i][0];
    var cell = matrixSheet.getRange(startRow + i, 1);
    var cellColor = cell.getBackground();
    var newValue = currentValue;
    
    if (cellColor === headerColor) {
      newNumbers.push([newValue]);
      continue;
    }
    
    var isEmpty = currentValue === "" || currentValue === null || currentValue === undefined;
    var numericValue = parseFloat(currentValue);
    var isNumber = !isNaN(numericValue) && currentValue !== "";
    
    if (isNumber || isEmpty) {
      if (currentValue !== currentNumber) {
        hasChanges = true;
      }
      newValue = currentNumber;
      currentNumber++;
    }
    
    newNumbers.push([newValue]);
  }
  
  if (hasChanges) {
    requirementRange.setValues(newNumbers);
    updateRegressionRequirements(matrixSheet, startRow, requirementData, newNumbers);
    
    // АВТОМАТИЧЕСКОЕ ОКРАШИВАНИЕ после перенумерации
    autoColorColumnA();
    
    SpreadsheetApp.getUi().alert('✅ Умная перенумерация завершена!');
  } else {
    SpreadsheetApp.getUi().alert('ℹ️ Перенумерация не требуется.');
  }
}

// ФУНКЦИЯ ПЕРЕНУМЕРАЦИИ ПРЕФИКСОВ В СТОЛБЦЕ B
function autoRenumberPrefixesB() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var matrixSheet = spreadsheet.getSheetByName("Матрица трассировки");
  
  var startRow = 8;
  var lastRow = matrixSheet.getLastRow();
  var requirementRange = matrixSheet.getRange(startRow, 2, lastRow - startRow + 1, 1);
  var requirementData = requirementRange.getValues();
  
  var headerColor = matrixSheet.getRange("B6").getBackground();
  
  var newValues = [];
  var hasChanges = false;
  
  var prefixCounters = {
    'UI': 1, 'FUNC': 1, 'NAV': 1, 'DAT': 1, 'ST': 1, 
    'NOT': 1, 'A11Y': 1, 'REQ': 1, 'APP': 1, 'FIL': 1
  };
  
  for (var i = 0; i < requirementData.length; i++) {
    var currentValue = requirementData[i][0];
    var cellA = matrixSheet.getRange(startRow + i, 1);
    var cellAColor = cellA.getBackground();
    var newValue = currentValue;
    
    if (cellAColor === headerColor) {
      newValues.push([newValue]);
      continue;
    }
    
    if (typeof currentValue === 'string' && currentValue.includes('_')) {
      var match = currentValue.match(/^([A-Z]+)_(\d+):\s*(.*)$/) || 
                  currentValue.match(/^([A-Z]+)_(\d+)\s*:\s*(.*)$/) ||
                  currentValue.match(/^([A-Z]+)_(\d+)\s*(.*)$/);
      
      if (match) {
        var prefix = match[1];
        var oldNumber = match[2];
        var description = match[3] ? match[3].trim() : "";
        
        prefix = normalizePrefix(prefix);
        
        if (prefixCounters.hasOwnProperty(prefix)) {
          var newNumber = prefixCounters[prefix];
          var newPrefixedValue = prefix + '_' + newNumber + (description ? ': ' + description : '');
          
          if (currentValue !== newPrefixedValue) {
            hasChanges = true;
          }
          
          newValue = newPrefixedValue;
          prefixCounters[prefix]++;
        } else {
          prefixCounters[prefix] = 1;
          newValue = prefix + '_1' + (description ? ': ' + description : '');
          prefixCounters[prefix]++;
          hasChanges = true;
        }
      }
    }
    
    newValues.push([newValue]);
  }
  
  if (hasChanges) {
    requirementRange.setValues(newValues);
    SpreadsheetApp.getUi().alert('✅ Перенумерация префиксов завершена!');
  } else {
    SpreadsheetApp.getUi().alert('ℹ️ Перенумерация не требуется.');
  }
}

// ФУНКЦИЯ АВТОЗАПОЛНЕНИЯ МЕТАДАННЫХ
function autoFillTestMetadata(overwriteExisting) {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var matrixSheet = spreadsheet.getSheetByName("Матрица трассировки");
  
  var lastRow = matrixSheet.getLastRow();
  var testNames = matrixSheet.getRange(8, 2, lastRow - 7, 1).getValues();
  var existingTypes = matrixSheet.getRange(8, 3, lastRow - 7, 1).getValues();
  var existingPriorities = matrixSheet.getRange(8, 4, lastRow - 7, 1).getValues();
  
  var newTypes = [];
  var newPriorities = [];
  
  var typeMapping = {
    'UI': 'Визуальная проверка', 'FUNC': 'Функциональная проверка', 
    'A11Y': 'Проверка доступности', 'NAV': 'Навигационная проверка',
    'ST': 'Проверка состояния', 'DAT': 'Проверка данных',
    'NOT': 'Проверка уведомлений', 'APP': 'Проверка автозаполнения',
    'FIL': 'Проверка фильтрации'
  };
  
  var priorityMapping = {
    'Визуальная проверка': 'Низкий', 'Функциональная проверка': 'Высокий',
    'Проверка доступности': 'Низкий', 'Навигационная проверка': 'Средний',
    'Проверка состояния': 'Средний', 'Проверка данных': 'Высокий',
    'Проверка уведомлений': 'Высокий', 'Проверка автозаполнения': 'Низкий',
    'Проверка фильтрации': 'Средний'
  };
  
  var headerColor = matrixSheet.getRange("B6").getBackground();
  
  for (var i = 0; i < testNames.length; i++) {
    var testName = testNames[i][0];
    var currentType = existingTypes[i][0];
    var currentPriority = existingPriorities[i][0];
    var cellA = matrixSheet.getRange(i + 8, 1);
    var cellAColor = cellA.getBackground();
    
    var detectedType = '';
    var detectedPriority = '';
    
    if (cellAColor === headerColor) {
      newTypes.push([currentType]);
      newPriorities.push([currentPriority]);
      continue;
    }
    
    if (testName) {
      var upperTestName = testName.toString().toUpperCase();
      
      if (upperTestName.includes('FUNC')) detectedType = typeMapping['FUNC'];
      else if (upperTestName.includes('DAT')) detectedType = typeMapping['DAT'];
      else if (upperTestName.includes('NOT')) detectedType = typeMapping['NOT'];
      else if (upperTestName.includes('NAV')) detectedType = typeMapping['NAV'];
      else if (upperTestName.includes('ST')) detectedType = typeMapping['ST'];
      else if (upperTestName.includes('A11Y')) detectedType = typeMapping['A11Y'];
      else if (upperTestName.includes('APP')) detectedType = typeMapping['APP'];
      else if (upperTestName.includes('FIL')) detectedType = typeMapping['FIL'];
      else if (upperTestName.includes('UI')) detectedType = typeMapping['UI'];
    }
    
    if (detectedType) {
      detectedPriority = priorityMapping[detectedType];
    }
    
    if (overwriteExisting) {
      newTypes.push([detectedType]);
      newPriorities.push([detectedPriority]);
    } else {
      newTypes.push([currentType || detectedType]);
      newPriorities.push([currentPriority || detectedPriority]);
    }
  }
  
  matrixSheet.getRange(8, 3, newTypes.length, 1).setValues(newTypes);
  matrixSheet.getRange(8, 4, newPriorities.length, 1).setValues(newPriorities);
  
  return {
    updatedTypes: newTypes.filter(row => row[0]).length,
    updatedPriorities: newPriorities.filter(row => row[0]).length
  };
}

// ВСПОМОГАТЕЛЬНЫЕ ФУНКЦИИ
function parseRequirementIds(reqString) {
  if (!reqString) return [];
  return reqString.toString()
    .split(',')
    .map(req => parseInt(req.trim()))
    .filter(req => !isNaN(req));
}

function detectRequirementType(existingType, requirementName) {
  if (existingType) {
    var typeMap = {
      'Визуальная проверка': 'UI', 'Проверка доступности': 'A11Y',
      'Навигационная проверка': 'NAV', 'Проверка состояния': 'ST',
      'Функциональная проверка': 'FUNC', 'Проверка данных': 'DAT',
      'Проверка уведомлений': 'NOT', 'Проверка автозаполнения': 'APP',
      'Проверка фильтрации': 'FIL'
    };
    return typeMap[existingType] || null;
  }
  
  var upperName = requirementName.toString().toUpperCase();
  if (upperName.includes('UI') || upperName.includes('ВИЗУАЛ') || upperName.includes('ЦВЕТ') || upperName.includes('ШРИФТ')) return 'UI';
  if (upperName.includes('A11Y') || upperName.includes('ДОСТУПН')) return 'A11Y';
  if (upperName.includes('NAV') || upperName.includes('НАВИГАЦ') || upperName.includes('ПЕРЕХОД')) return 'NAV';
  if (upperName.includes('ST') || upperName.includes('СОСТОЯН')) return 'ST';
  if (upperName.includes('DAT') || upperName.includes('ДАНН')) return 'DAT';
  if (upperName.includes('NOT') || upperName.includes('УВЕДОМЛ')) return 'NOT';
  if (upperName.includes('APP') || upperName.includes('АВТОЗАПОЛН')) return 'APP';
  if (upperName.includes('FIL') || upperName.includes('ФИЛЬТР')) return 'FIL';
  if (upperName.includes('FUNC') || upperName.includes('ФУНКЦИОНАЛ')) return 'FUNC';
  
  return null;
}

function normalizePrefix(prefix) {
  var corrections = { 'FUNС': 'FUNC' };
  return corrections[prefix] || prefix;
}

function applySimpleVisuals(matrixSheet) {
  var headerRange = matrixSheet.getRange(4, 5, 1, matrixSheet.getLastColumn() - 4);
  headerRange.setBackground("#4285F4");
  headerRange.setFontColor("#FFFFFF");
  headerRange.setFontWeight("bold");
  headerRange.setHorizontalAlignment("center");
  
  var fullDataRange = matrixSheet.getDataRange();
  fullDataRange.setBorder(false, false, false, false, false, false);
  
  matrixSheet.setFrozenRows(4);
  matrixSheet.setFrozenColumns(4);
}

// ОБНОВЛЕННОЕ МЕНЮ С ПОДМЕНЮ
function createMenu() {
  var menu = SpreadsheetApp.getUi().createMenu('🚀 Автоматизация тестирования')
    .addItem('🔄 Обновить матрицу', 'updateTraceabilityMatrixWithBlocked')
    .addItem('📊 Умная статистика', 'showSmartCoverageStats')
    .addItem('🔢 Перенумеровать всё', 'autoRenumberRequirements')
    .addItem('🔤 Перенумеровать префиксы', 'autoRenumberPrefixesB')
    .addItem('🎨 Авто-цвет столбца A', 'autoColorColumnA') // НОВАЯ КНОПКА
    .addItem('🏷️ Автозаполнить метаданные', 'autoFillMetadataManual')
    .addItem('💡 Предложить тесты', 'suggestTestCases')
    .addSeparator();
  
  // Подменю отладки
  menu.addSubMenu(
    SpreadsheetApp.getUi().createMenu('🐛 Отладка')
      .addItem('Проверить требование 54', 'checkRequirement54')
      .addItem('Диагностика цветов', 'debugTestColors')
      .addItem('Диагностика требований', 'debugRequirements')
      .addItem('Очистить столбец покрытия', 'fixColorsNow')
  );
  
  menu.addSeparator()
    .addItem('🧹 Очистить матрицу', 'clearMatrix')
    .addItem('🔄 Сбросить настройки', 'resetRememberChoice')
    .addToUi();
}

// ФУНКЦИИ ДЛЯ МЕНЮ
function autoFillMetadataManual() {
  var result = autoFillTestMetadata(false);
  SpreadsheetApp.getUi().alert('✅ Автозаполнение завершено!\nОбновлено типов: ' + result.updatedTypes + '\nОбновлено приоритетов: ' + result.updatedPriorities);
}

function suggestTestCases() {
  var matrixSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Матрица трассировки");
  var dataRange = matrixSheet.getRange(8, 5, matrixSheet.getLastRow() - 7, matrixSheet.getLastColumn() - 4);
  var data = dataRange.getValues();
  var requirementData = matrixSheet.getRange(8, 1, matrixSheet.getLastRow() - 7, 1).getValues();
  
  var uncoveredReqs = [];
  for (var i = 0; i < data.length; i++) {
    var currentReqId = requirementData[i][0];
    if (typeof currentReqId !== 'number' || isNaN(currentReqId)) continue;
    
    var covered = false;
    for (var j = 0; j < data[i].length; j++) {
      if (data[i][j] === "✓" || data[i][j] === "⏸️") {
        covered = true;
        break;
      }
    }
    if (!covered) {
      uncoveredReqs.push(currentReqId);
    }
  }
  
  if (uncoveredReqs.length > 0) {
    var suggestions = "💡 Предлагаемые тест-кейсы для непокрытых требований:\n\n";
    uncoveredReqs.forEach(reqId => {
      suggestions += `📝 Для требования ${reqId}:\n`;
      suggestions += `• Создайте позитивный тест-кейс\n`;
      suggestions += `• Создайте 2-3 негативных тест-кейса\n`;
      suggestions += `• Проверьте граничные значения\n\n`;
    });
    SpreadsheetApp.getUi().alert(suggestions);
  } else {
    SpreadsheetApp.getUi().alert('🎉 Все требования покрыты!');
  }
}

function clearMatrix() {
  var matrixSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Матрица трассировки");
  var dataRange = matrixSheet.getRange(8, 5, matrixSheet.getLastRow() - 7, matrixSheet.getLastColumn() - 4);
  dataRange.clearContent();
  
  var requirementRange = matrixSheet.getRange(8, 1, matrixSheet.getLastRow() - 7, 1);
  requirementRange.setBackground("#FFFFFF");
  requirementRange.setFontColor("#000000");
  requirementRange.setFontWeight("normal");
  
  var fullDataRange = matrixSheet.getDataRange();
  fullDataRange.setBorder(false, false, false, false, false, false);
  
  SpreadsheetApp.getUi().alert('✅ Матрица трассировки очищена!');
}

// ОСТАЛЬНЫЕ ВСПОМОГАТЕЛЬНЫЕ ФУНКЦИИ
function onEdit(e) {
  handleAutoRenumber(e);
  
  if (rememberChoice.enabled && rememberChoice.choice === 'NO') {
    return;
  }
  
  var sheet = e.source.getActiveSheet();
  var range = e.range;
  
  if (sheet.getName() === "Регрессионное тестирование" && 
      (range.getColumn() === 1 || range.getColumn() === 12 || 
       range.getColumn() === 13 || range.getColumn() === 24)) {
    
    var ui = SpreadsheetApp.getUi();
    var htmlOutput = HtmlService.createHtmlOutput(
      '<div style="font-family: Arial; width: 300px; padding: 20px;">' +
      '<h3>🔄 Обновить матрицу трассировки?</h3>' +
      '<p>Вы изменили данные регрессионного тестирования.</p>' +
      '<label>' +
      '<input type="checkbox" id="remember" style="margin-right: 10px;">' +
      'Запомнить выбор и больше не спрашивать' +
      '</label>' +
      '<div style="margin-top: 20px; text-align: right;">' +
      '<button onclick="google.script.run.withSuccessHandler(google.script.host.close).handleDialogResponse(true, document.getElementById(\'remember\').checked)" style="margin-right: 10px; padding: 8px 16px;">Да</button>' +
      '<button onclick="google.script.run.withSuccessHandler(google.script.host.close).handleDialogResponse(false, document.getElementById(\'remember\').checked)" style="padding: 8px 16px;">Нет</button>' +
      '</div>' +
      '</div>'
    ).setWidth(350).setHeight(200);
    
    ui.showModalDialog(htmlOutput, 'Авто-обновление матрицы');
  }
}

function handleDialogResponse(update, remember) {
  if (remember) {
    rememberChoice.enabled = true;
    rememberChoice.choice = update ? 'YES' : 'NO';
  }
  
  if (update) {
    updateTraceabilityMatrixWithBlocked();
  }
}

function resetRememberChoice() {
  rememberChoice.enabled = false;
  rememberChoice.choice = null;
  SpreadsheetApp.getUi().alert('✅ Настройки запоминания сброшены!');
}

function onOpen() {
  createMenu();
}

// ДИАГНОСТИЧЕСКИЕ ФУНКЦИИ (для подменю отладки)
function checkRequirement54() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var regressionSheet = spreadsheet.getSheetByName("Регрессионное тестирование");
  var matrixSheet = spreadsheet.getSheetByName("Матрица трассировки");
  
  var requirementId = 54;
  var regressionData = regressionSheet.getDataRange().getValues();
  var coveringTests = [];
  
  for (var i = 1; i < regressionData.length; i++) {
    var row = regressionData[i];
    var rowNum = i + 1;
    
    var leftTestId = parseInt(row[0]);
    var leftReqIds = parseRequirementIds(row[11]);
    var rightTestId = parseInt(row[12]);
    var rightReqIds = parseRequirementIds(row[23]);
    
    if ((leftReqIds.includes(requirementId)) || (rightReqIds.includes(requirementId))) {
      var testId = leftReqIds.includes(requirementId) ? leftTestId : rightTestId;
      var colToCheck = leftReqIds.includes(requirementId) ? 6 : 23;
      var isBlocked = isTestBlocked(regressionSheet, rowNum, colToCheck);
      var cellColor = regressionSheet.getRange(rowNum, colToCheck).getBackground();
      
      coveringTests.push({
        testId: testId,
        row: rowNum,
        blocked: isBlocked,
        location: leftReqIds.includes(requirementId) ? "Левый блок (F)" : "Правый блок (W)",
        color: cellColor
      });
    }
  }
  
  if (coveringTests.length > 0) {
    var message = "✅ Требование 54 покрыто тестами:\n\n";
    coveringTests.forEach(test => {
      message += `• Тест ${test.testId} (строка ${test.row}, ${test.location})\n`;
      message += `  Цвет: ${test.color}\n`;
      message += `  Статус: ${test.blocked ? "ЗАБЛОКИРОВАН ⏸️" : "АКТИВЕН ✓"}\n\n`;
    });
    SpreadsheetApp.getUi().alert(message);
  } else {
    SpreadsheetApp.getUi().alert("❌ Требование 54 не покрыто ни одним тестом!");
  }
}

function debugTestColors() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var regressionSheet = spreadsheet.getSheetByName("Регрессионное тестирование");
  var testsToCheck = [35, 39];
  var debugInfo = "=== ДИАГНОСТИКА ЦВЕТОВ ТЕСТОВ ===\n\n";
  var regressionData = regressionSheet.getDataRange().getValues();
  
  for (var i = 1; i < regressionData.length; i++) {
    var row = regressionData[i];
    var leftTestId = parseInt(row[0]);
    var rightTestId = parseInt(row[12]);
    
    if (testsToCheck.includes(leftTestId) || testsToCheck.includes(rightTestId)) {
      var testId = testsToCheck.includes(leftTestId) ? leftTestId : rightTestId;
      var colToCheck = testsToCheck.includes(leftTestId) ? 6 : 23;
      var rowNum = i + 1;
      var cell = regressionSheet.getRange(rowNum, colToCheck);
      var backgroundColor = cell.getBackground();
      var isBlocked = isTestBlocked(regressionSheet, rowNum, colToCheck);
      
      debugInfo += "Тест " + testId + " (строка " + rowNum + ", столбец " + (colToCheck === 6 ? "F" : "W") + "):\n";
      debugInfo += "  Цвет: " + backgroundColor + "\n";
      debugInfo += "  Заблокирован: " + (isBlocked ? "ДА" : "НЕТ") + "\n\n";
    }
  }
  SpreadsheetApp.getUi().alert(debugInfo);
}

function debugRequirements() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var matrixSheet = spreadsheet.getSheetByName("Матрица трассировки");
  var startRow = 8;
  var lastRow = matrixSheet.getLastRow();
  var requirementRange = matrixSheet.getRange(startRow, 1, lastRow - startRow + 1, 1);
  var requirementData = requirementRange.getValues();
  var debugInfo = "=== ДИАГНОСТИКА ТРЕБОВАНИЙ ===\n\n";
  
  for (var i = 0; i < requirementData.length; i++) {
    var value = requirementData[i][0];
    var rowNum = startRow + i;
    debugInfo += "Строка " + rowNum + ": ";
    debugInfo += "Значение: '" + value + "', ";
    debugInfo += "Тип: " + typeof value + ", ";
    debugInfo += "isNaN: " + isNaN(value) + "\n";
  }
  debugInfo += "\nВсего строк: " + requirementData.length;
  SpreadsheetApp.getUi().alert(debugInfo);
}

// ФУНКЦИИ ДЛЯ УМНОЙ СТАТИСТИКИ
function showSmartStatsDialog(coverageStats) {
  var totalPercent = (coverageStats.total.covered / coverageStats.total.requirements * 100).toFixed(2);
  var wellCoveredPercent = (coverageStats.total.wellCovered / coverageStats.total.requirements * 100).toFixed(2);
  
  var typeDetails = '';
  for (var type in coverageStats.byType) {
    var typeData = coverageStats.byType[type];
    if (typeData.requirements > 0) {
      var coveredPercent = (typeData.covered / typeData.requirements * 100).toFixed(1);
      var wellCoveredPercentType = (typeData.wellCovered / typeData.requirements * 100).toFixed(1);
      typeDetails += `
        <div style="margin: 8px 0; padding: 8px; background: #f8f9fa; border-radius: 5px;">
          <strong>${getTypeDisplayName(type)}</strong><br>
          <span style="font-size: 0.9em;">
            Требований: ${typeData.requirements} | 
            Покрыто: ${typeData.covered} (${coveredPercent}%) |
            Хорошо: ${typeData.wellCovered} (${wellCoveredPercentType}%)
          </span>
        </div>`;
    }
  }
  
  var htmlOutput = HtmlService.createHtmlOutput(
    '<div style="font-family: Arial; width: 600px; padding: 20px; background: #f8f9fa; color: #333; border-radius: 10px; border: 1px solid #ddd;">' +
    '<h2 style="margin: 0 0 20px 0; text-align: center; color: #4285F4;">🧠 Умная статистика покрытия</h2>' +
    '<div style="background: white; padding: 15px; border-radius: 8px; margin-bottom: 15px; box-shadow: 0 2px 4px rgba(0,0,0,0.1);">' +
    '<div style="text-align: center; margin-bottom: 15px; padding: 10px; background: #E8F0FE; border-radius: 5px;">' +
    '<div style="font-size: 1.2em; font-weight: bold; color: #4285F4;">Общее покрытие: ' + totalPercent + '%</div>' +
    '<div style="font-size: 0.9em;">Хорошо покрыто: ' + wellCoveredPercent + '% требований</div>' +
    '</div>' +
    '<div style="display: grid; grid-template-columns: 1fr 1fr; gap: 10px; margin-bottom: 15px;">' +
    '<div style="padding: 10px; background: #E6F4EA; border-radius: 5px; text-align: center;">' +
    '<div style="font-size: 1.1em; font-weight: bold; color: #137333;">' + coverageStats.total.covered + '</div>' +
    '<div style="font-size: 0.8em;">Покрыто требований</div>' +
    '</div>' +
    '<div style="padding: 10px; background: #FFF3CD; border-radius: 5px; text-align: center;">' +
    '<div style="font-size: 1.1em; font-weight: bold; color: #856404;">' + coverageStats.total.wellCovered + '</div>' +
    '<div style="font-size: 0.8em;">Хорошо покрыто</div>' +
    '</div>' +
    '</div>' +
    '<h3 style="margin: 15px 0 10px 0; font-size: 1.1em;">Детали по типам требований:</h3>' +
    typeDetails +
    '</div>' +
    '<div style="text-align: center; margin-top: 15px;">' +
    '<button onclick="google.script.run.withSuccessHandler(google.script.host.close).showUncoveredDetails()" style="background: #F44336; color: white; border: none; padding: 10px 20px; border-radius: 5px; cursor: pointer; margin-right: 10px; font-size: 14px;">🔍 Детали непокрытых</button>' +
    '<button onclick="google.script.run.withSuccessHandler(google.script.host.close).suggestTestCases()" style="background: #4CAF50; color: white; border: none; padding: 10px 20px; border-radius: 5px; cursor: pointer; font-size: 14px;">💡 Предложить тесты</button>' +
    '</div>' +
    '</div>'
  ).setWidth(640).setHeight(500);
  
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'Умная статистика покрытия');
}

function getTypeDisplayName(typeCode) {
  var names = {
    'UI': '🎨 Визуальные', 'A11Y': '♿ Доступность', 'NAV': '🧭 Навигационные',
    'ST': '🔄 Состояния', 'FUNC': '⚙️ Функциональные', 'DAT': '💾 Данные',
    'NOT': '🔔 Уведомления', 'APP': '🤖 Автозаполнение', 'FIL': '🔍 Фильтрация'
  };
  return names[typeCode] || typeCode;
}

function showUncoveredDetails() {
  var matrixSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Матрица трассировки");
  var dataRange = matrixSheet.getRange(8, 5, matrixSheet.getLastRow() - 7, matrixSheet.getLastColumn() - 4);
  var data = dataRange.getValues();
  var requirementData = matrixSheet.getRange(8, 1, matrixSheet.getLastRow() - 7, 1).getValues();
  var uncoveredReqs = [];
  
  for (var i = 0; i < data.length; i++) {
    var currentReqId = requirementData[i][0];
    if (typeof currentReqId !== 'number' || isNaN(currentReqId)) continue;
    
    var covered = false;
    for (var j = 0; j < data[i].length; j++) {
      if (data[i][j] === "✓" || data[i][j] === "⏸️") {
        covered = true;
        break;
      }
    }
    if (!covered) {
      uncoveredReqs.push(currentReqId);
    }
  }
  
  if (uncoveredReqs.length > 0) {
    var message = "🔍 НЕПОКРЫТЫЕ ТРЕБОВАНИЯ:\n\n";
    uncoveredReqs.forEach(reqId => { message += `• Требование ${reqId}\n`; });
    message += `\n💡 Всего непокрытых: ${uncoveredReqs.length}`;
    SpreadsheetApp.getUi().alert(message);
  } else {
    SpreadsheetApp.getUi().alert('🎉 Все требования покрыты!');
  }
}

//Автоматическое окрашивание столбца A
function autoColorColumnA() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var matrixSheet = spreadsheet.getSheetByName("Матрица трассировки");
  
  var headerColor = matrixSheet.getRange("B6").getBackground();
  var lastRow = matrixSheet.getLastRow();
  var startRow = 8;
  
  // Получаем цвета столбца B
  var columnBRange = matrixSheet.getRange(startRow, 2, lastRow - startRow + 1, 1);
  var columnBColors = columnBRange.getBackgrounds();
  
  // Подготавливаем цвета для столбца A
  var columnAColors = [];
  
  for (var i = 0; i < columnBColors.length; i++) {
    var currentBColor = columnBColors[i][0];
    
    // КРАСИМ В СИНИЙ ТОЛЬКО если ячейка B имеет цвет заголовка
    if (currentBColor === headerColor) {
      columnAColors.push([headerColor]);
    } else {
      // Иначе - белый цвет
      columnAColors.push(["#ffffff"]);
    }
  }
  
  // Применяем цвета к столбцу A
  var columnARange = matrixSheet.getRange(startRow, 1, lastRow - startRow + 1, 1);
  columnARange.setBackgrounds(columnAColors);
}

// ФУНКЦИИ ДЛЯ ПЕРЕНУМЕРАЦИИ ССЫЛОК
function updateRegressionRequirements(matrixSheet, startRow, oldData, newData) {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var regressionSheet = spreadsheet.getSheetByName("Регрессионное тестирование");
  var regressionData = regressionSheet.getDataRange().getValues();
  var changesMade = false;
  var oldToNewMap = {};
  
  for (var i = 0; i < oldData.length; i++) {
    var oldValue = oldData[i][0];
    var newValue = newData[i][0];
    if (oldValue !== newValue) {
      oldToNewMap[oldValue] = newValue;
      if (typeof oldValue === 'number' && !isNaN(oldValue)) {
        oldToNewMap[oldValue.toString()] = newValue;
      }
    }
  }
  
  for (var i = 1; i < regressionData.length; i++) {
    var row = regressionData[i];
    var updatedLeft = updateRequirementIdsUniversal(row[11], oldToNewMap);
    var updatedRight = updateRequirementIdsUniversal(row[23], oldToNewMap);
    
    if (updatedLeft !== row[11]) {
      regressionSheet.getRange(i + 1, 12).setValue(updatedLeft);
      changesMade = true;
    }
    if (updatedRight !== row[23]) {
      regressionSheet.getRange(i + 1, 24).setValue(updatedRight);
      changesMade = true;
    }
  }
  return changesMade;
}

function updateRequirementIdsUniversal(requirementString, mapping) {
  if (!requirementString) return requirementString;
  return requirementString.toString().split(',').map(req => {
    var trimmedReq = req.trim();
    if (mapping[trimmedReq] !== undefined) return mapping[trimmedReq];
    var numValue = parseInt(trimmedReq);
    if (!isNaN(numValue) && mapping[numValue] !== undefined) return mapping[numValue];
    return trimmedReq;
  }).join(', ');
}

function handleAutoRenumber(e) {
  var sheet = e.source.getActiveSheet();
  var range = e.range;
  if (sheet.getName() === "Матрица трассировки" && range.getNumRows() > 1) {
    var ui = SpreadsheetApp.getUi();
    var response = ui.alert('🆕 Обнаружены новые строки', 'Вы добавили новые строки в матрицу. Хотите автоматически перенумеровать требования?', ui.ButtonSet.YES_NO);
    if (response == ui.Button.YES) autoRenumberRequirements();
  }
}

function freezeCoverageColumn() {
  var matrixSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Матрица трассировки");
  var lastColumn = matrixSheet.getLastColumn();
  
  // Закрепляем первые 4 столбца + столбец покрытия
  matrixSheet.setFrozenColumns(4);
  
  // Если тестов много - показываем подсказку как быстро найти столбец покрытия
  if (lastColumn > 20) {
    // Можно добавить визуальный маркер
    var coverageColumn = lastColumn;
    var headerCell = matrixSheet.getRange(4, coverageColumn);
    headerCell.setNote("🎯 Столбец с количеством покрывающих тестов");
    headerCell.setBackground("#FFEB3B"); // Яркий цвет для заметности
  }
}


function addNavigationHelp() {
  var matrixSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Матрица трассировки");
  var lastColumn = matrixSheet.getLastColumn();
  
  // Добавляем кнопку-подсказку в меню
  if (lastColumn > 15) {
    var ui = SpreadsheetApp.getUi();
    var response = ui.alert('💡 Навигация', 
      'Столбец с покрытием находится в колонке ' + lastColumn + 
      '. Хочешь быстро перейти к нему?', 
      ui.ButtonSet.YES_NO);
    
    if (response == ui.Button.YES) {
      matrixSheet.getRange(1, lastColumn).activate();
    }
  }
}
