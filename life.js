// VARS

var tasksSheetName = "tasks";
var dashboardSheetName = "dashboard";
var financeSheetName = "finance";
var dataSheetName = "data";
var auxiliarDataSheetName = "auxiliar_data";

var alertColor = "#FF5950";

var professionalType = "professional";
var personalType = "personal";
var logosophyType = "logosophy";

var activitiesKey = "activities";
var howKey = "hows";
var whyKey = "whys";
var scoresKey = "scores";
var leftKey = "left";
var doneKey = "done";
var percentageKey = "percentage";
var notesKey = "notes";
var tasksSumKey = "tasksSum";

var missing = "MISSING";
var checkEmoji = "✅";

// Columns and Rows
// Dashboard
var personalTaskStatusIndex = "1";
var personalTitleColumn = "B";
var personalWhyColumn = "C";
var personalHowColumn = "D";
var personalScoreColumn = "E";
var professionalTaskStatusIndex = "6";
var professionalTitleColumn = "G";
var professionalWhyColumn = "H";
var professionalHowColumn = "I";
var professionalScoreColumn = "J";
var logosophyTaskStatusIndex = "11";
var logosophyTitleColumn = "L";
var logosophyWhyColumn = "M";
var logosophyHowColumn = "N";
var logosophyScoreColumn = "O";
var basicsColumnA = "B";
var basicsColumnB = "C";
var basicsColumnC = "D";
var routineColumnA = "G";
var routineColumnB = "H";
var routineColumnC = "I";
var challengesColumnA = "L";
var challengesColumnB = "M";
var challengesColumnC = "N";
var auxiliarDataColumns = [
  basicsColumnA,
  basicsColumnB,
  basicsColumnC,
  routineColumnA,
  routineColumnB,
  routineColumnC,
  challengesColumnA,
  challengesColumnB,
  challengesColumnC
];
var tasksLeftColumn = "B";
var doneColumn = "G";
var percentageColumn = "L";
var notesColumn = "L";
var notesRow = "2";
var tasksDataRow = "8";
var newExpenseTypeCell = "H5";
var newExpenseValueCell = "I5";
var spentValueCell = "H2";
var goalValueCell = "H4";
var notesCell = "L2";
var pointsDoneCell = "G8";
var refreshCheckboxCell = "Q2";
var auxiliarDataRow = 13;
var tasksFirstRow = 16;

//Data
var dateColumn = "A";
var leftColumn = "B";
var doneColumn = "C";
var percentageColumn = "D";
var notesColumn = "E";
var tasksSumColumn = "F";

//Tasks
var typeColumn = "A";
var activityColumn = "B";
var howColumn = "G";
var whyColumn = "F";
var scoreColumn = "H";

// TRIGGERS

function onEdit(e) {
  if (shouldUpdateFinance()) {
    addNewExpense();
    updateDailyExpense();
  }

  if (shouldResetDashboard()) {
    var data = getTodaysData();
    populateDataSheet(data);
    var auxiliarData = getAuxiliarData();
    populateAuxiliarDataSheet(auxiliarData);
    resetSheet();
  }

  if (shouldUpdateScoreSum()) {
    var scoreSum = getScoreSum();
    updateScoreSum(scoreSum);
  }

  if (getCurrentSheet().getSheetName() == tasksSheetName) {
    var types = [personalType, professionalType, logosophyType];

    for (var i = 0; i < types.length; i++) {
      var type = types[i];
      var info = getMappedInfoForType(type);
      populateType(type, info);
    }
  }
}

// MODULES

// Finance

function shouldUpdateFinance() {
  var sheet = getCurrentSheet().getSheetByName(dashboardSheetName);
  return (
    getCurrentSheet().getSheetName() == dashboardSheetName &&
    sheet.getRange(newExpenseTypeCell).getValue() !== "" &&
    checkForValidNumber(sheet.getRange(newExpenseValueCell).getValue())
  );
}

function addNewExpense() {
  var kind = getStringForPos("H", 5);
  var value = getStringForPos("I", 5);
  var financeSheet = getCurrentSheet().getSheetByName(financeSheetName);
  var row = getLastPopulatedRow(financeSheet) + 1;
  var dateColumn = "A";
  var kindColumn = "B";
  var valueColumn = "C";

  financeSheet.getRange(dateColumn + row).setValue(getTodaysDate());
  financeSheet.getRange(kindColumn + row).setValue(kind);
  financeSheet.getRange(valueColumn + row).setValue(value);
}

function updateDailyExpense() {
  var sheet = getCurrentSheet().getSheetByName(dashboardSheetName);
  resetValueForRange(
    sheet.getRange(newExpenseTypeCell + ":" + newExpenseValueCell)
  );
  sheet.getRange(spentValueCell).setValue(getDailyExpensesSum());
}

function getDailyExpensesSum() {
  var financeSheet = getCurrentSheet().getSheetByName(financeSheetName);
  var row = getLastPopulatedRow(financeSheet);
  var sum = 0;

  for (var i = 0; i < row - 1; i++) {
    var currentRow = i + 3;
    var date = getStringOnSheetForPos(financeSheet, "A", currentRow);
    if (date == String(getTodaysDate())) {
      sum += financeSheet.getRange("C" + currentRow).getValue();
    }
  }

  return sum;
}

function calculateDailyGoal() {
  var financeSheet = getCurrentSheet().getSheetByName(financeSheetName);
  var monthlyGoal = financeSheet.getRange("B1").getValue();
  var amountSpent = 0;
  var lastRow = getLastPopulatedRow(financeSheet);

  for (var i = 0; i < lastRow - 1; i++) {
    var currentRow = i + 3;
    var rowsDate = getStringOnSheetForPos(financeSheet, "A", currentRow);
    if (checkForCurrentMonth(rowsDate)) {
      amountSpent += financeSheet.getRange("C" + currentRow).getValue();
    }
  }

  var remainingValue = monthlyGoal - amountSpent;
  var today = parseInt(String(getTodaysDate()).substring(0, 2));
  var remainingDays = 31 - parseInt(String(getTodaysDate()).substring(0, 2));

  return remainingValue / remainingDays;
}

// Score

function shouldUpdateScoreSum() {
  if (getCurrentSheet().getSheetName() == dashboardSheetName) {
    var currentCell = getCurrentSheet().getActiveCell();
    var value = String(currentCell.getValue());
    var column = currentCell.getColumn();

    if (
      column == personalTaskStatusIndex ||
      column == professionalTaskStatusIndex ||
      column == logosophyTaskStatusIndex
    ) {
      return true;
    }
  }

  return false;
}

function getScoreSumForColumns(checkColumn, scoreColumn) {
  var sheet = getCurrentSheet().getSheetByName(dashboardSheetName);
  var avals = sheet
    .getRange(scoreColumn + tasksFirstRow + ":" + scoreColumn)
    .getValues();
  var alast = avals.filter(String).length;
  var scoreSum = 0;

  if (alast > 0) {
    for (var i = 0; i <= alast; i++) {
      if (getStringForPos(checkColumn, i + tasksFirstRow) == checkEmoji) {
        var scoreToSum = parseInt(
          sheet.getRange(scoreColumn + String(i + tasksFirstRow)).getValue()
        );

        if (scoreToSum === scoreToSum) {
          scoreSum += scoreToSum;
        }
      }
    }
  }

  return scoreSum;
}

function getScoreSum() {
  var scoreSum = 0;

  scoreSum += getScoreSumForColumns("A", "E");
  scoreSum += getScoreSumForColumns("F", "J");
  scoreSum += getScoreSumForColumns("K", "O");

  return scoreSum;
}

function updateScoreSum(scoreSum) {
  var sheet = getCurrentSheet().getSheetByName(dashboardSheetName);
  sheet.getRange(pointsDoneCell).setValue(scoreSum);
}

// Reset

function shouldResetDashboard() {
  if (getCurrentSheet().getSheetName() == dashboardSheetName) {
    if (
      getCurrentSheet()
        .getRange(refreshCheckboxCell)
        .getValue() == checkEmoji
    ) {
      return true;
    }
  }

  return false;
}

function populateDataSheet(data) {
  dataSheet = getCurrentSheet().getSheetByName(dataSheetName);

  var row = getLastPopulatedRow(dataSheet) + 1;

  var tasksLeft = data[leftKey];
  var done = data[doneKey];
  var percentage = data[percentageKey];
  var notes = data[notesKey];
  var tasksSum = data[tasksSumKey];

  dataSheet.getRange(dateColumn + row).setValue(getTodaysDate());
  dataSheet.getRange(leftColumn + row).setValue(tasksLeft);
  dataSheet.getRange(doneColumn + row).setValue(done);
  dataSheet.getRange(percentageColumn + row).setValue(percentage);
  dataSheet.getRange(notesColumn + row).setValue(notes);
  dataSheet.getRange(tasksSumColumn + row).setValue(tasksSum);
}

function getTodaysData() {
  dashboardSheet = getCurrentSheet().getSheetByName(dashboardSheetName);

  var tasksLeft = getStringOnSheetForPos(
    dashboardSheet,
    tasksLeftColumn,
    tasksDataRow
  );
  var doneTasks = getStringOnSheetForPos(
    dashboardSheet,
    doneColumn,
    tasksDataRow
  );
  var percentage = getStringOnSheetForPos(
    dashboardSheet,
    percentageColumn,
    tasksDataRow
  );

  var notes = getStringOnSheetForPos(dashboardSheet, notesColumn, notesRow);

  var tasksSheet = getCurrentSheet().getSheetByName(tasksSheetName);
  var tasksSum = String(tasksSheet.getRange("A2:G").getValues());

  return getMappedData(tasksLeft, doneTasks, percentage, notes, tasksSum);
}

function populateAuxiliarDataSheet(data) {
  var sheet = getCurrentSheet().getSheetByName(auxiliarDataSheetName);
  var row = getLastPopulatedRow(sheet) + 1;

  for (var i = 0; i < auxiliarDataColumns.length; i++) {
    var key = auxiliarDataColumns[i];
    var title = data[0][key];
    var result = data[1][key];

    var titleColumn = (i + 1) * 2;
    var resultColumn = titleColumn + 1;

    sheet.getRange(row, titleColumn).setValue(title);
    sheet.getRange(row, resultColumn).setValue(result);
  }

  sheet.getRange("A" + row).setValue(getTodaysDate());
}

function getAuxiliarData() {
  var sheet = getCurrentSheet().getSheetByName(dashboardSheetName);
  var titlesMap = {};
  var resultsMap = {};

  for (var i = 0; i < auxiliarDataColumns.length; i++) {
    var column = auxiliarDataColumns[i];
    var title = getStringOnSheetForPos(sheet, column, auxiliarDataRow - 1);
    var result = getStringOnSheetForPos(sheet, column, auxiliarDataRow);

    titlesMap[column] = title;
    resultsMap[column] = result;
  }

  return [titlesMap, resultsMap];
}

function getMappedData(left, done, percentage, notes, tasksSum) {
  var data = {};

  data[leftKey] = left;
  data[doneKey] = done;
  data[percentageKey] = percentage;
  data[notesKey] = notes;
  data[tasksSumKey] = tasksSum;

  return data;
}

function resetSheet() {
  dashboard = getCurrentSheet().getSheetByName(dashboardSheetName);
  resetValueForRange(dashboard.getRange("A" + tasksFirstRow + ":N"));
  resetValueForRange(
    dashboard.getRange(spentValueCell + ":" + newExpenseValueCell)
  );
  resetValueForRange(
    dashboard.getRange(
      basicsColumnA +
        auxiliarDataRow +
        ":" +
        challengesColumnC +
        auxiliarDataRow
    )
  );
  resetValueForRange(dashboard.getRange(notesCell));
  resetValueForRange(dashboard.getRange(refreshCheckboxCell));

  var goal = calculateDailyGoal();
  dashboard.getRange(goalValueCell).setValue(goal);
  dashboard.getRange(spentValueCell).setValue(0);

  var scoreSum = getScoreSum();
  updateScoreSum(scoreSum);

  tasks = getCurrentSheet().getSheetByName(tasksSheetName);
  resetValueForRange(tasks.getRange("A2:G"));
}

// Populate

function getMappedInfoForType(type) {
  var lastIndex = getCurrentSheet().getLastRow() - 2;
  var arraySize = 0;
  var scores = new Array();

  for (var i = 0; i <= lastIndex; i++) {
    var row = i + 2;
    var activityType = getStringForPos(typeColumn, row);

    if (activityType == type) {
      scores.push(getStringForPos(scoreColumn, row));

      arraySize++;
    }
  }

  var sortedActivities = new Array(arraySize);
  var sortedHows = new Array(arraySize);
  var sortedWhys = new Array(arraySize);

  var usedIndexes = new Array();

  scores.sort(function(a, b) {
    return b - a;
  });

  for (var i = 0; i <= lastIndex; i++) {
    var row = i + 2;

    if (getStringForPos(typeColumn, row) != type) continue;

    var activity = getStringForPos(activityColumn, row);
    var how = getStringForPos(howColumn, row);
    var why = getStringForPos(whyColumn, row);
    var score = getStringForPos(scoreColumn, row);

    var index = getIndexForScore(scores, score, usedIndexes, arraySize - 1);
    usedIndexes.push(index);

    if (activity != "" && (how == "" || why == "")) {
      var missing = missing;
      sortedActivities[index] = missing;
      sortedHows[index] = missing;
      sortedWhys[index] = missing;
    } else {
      sortedActivities[index] = activity;
      sortedHows[index] = how;
      sortedWhys[index] = why;
    }
  }

  return getMappedInfo(sortedActivities, sortedHows, sortedWhys, scores);
}

function getMappedInfo(activities, hows, whys, scores) {
  var info = {};

  info[activitiesKey] = activities;
  info[howKey] = hows;
  info[whyKey] = whys;
  info[scoresKey] = scores;

  return info;
}

function getIndexForScore(scores, score, usedIndexes, maxIndex) {
  var index = scores.indexOf(score);

  for (var j = 0; j <= maxIndex; j++) {
    if (usedIndexes.indexOf(index) == -1) {
      break;
    }
    index++;
  }

  return index;
}

function populateType(type, mappedInfo) {
  var rankSheet = getCurrentSheet().getSheetByName(dashboardSheetName);
  var sortedActivities = mappedInfo[activitiesKey];
  var sortedHows = mappedInfo[howKey];
  var sortedWhys = mappedInfo[whyKey];
  var sortedScores = mappedInfo[scoresKey];
  var lastIndex = sortedActivities.length - 1;

  var first = personalTitleColumn;
  var second = personalWhyColumn;
  var third = personalHowColumn;
  var last = personalScoreColumn;

  if (type == professionalType) {
    first = professionalTitleColumn;
    second = professionalWhyColumn;
    third = professionalHowColumn;
    last = professionalScoreColumn;
  } else if (type == logosophyType) {
    first = logosophyTitleColumn;
    second = logosophyWhyColumn;
    third = logosophyHowColumn;
    last = logosophyScoreColumn;
  }

  for (var i = 0; i <= lastIndex; i++) {
    var row = i + tasksFirstRow;
    var activity = sortedActivities[i];
    var how = sortedHows[i];
    var why = sortedWhys[i];
    var score = sortedScores;

    var bgRange = first + row + ":" + last + row;
    if (activity == missing) {
      rankSheet.getRange(bgRange).setBackgroundColor(alertColor);
    } else {
      rankSheet.getRange(bgRange).setBackgroundColor(null);
    }

    rankSheet.getRange(first + row).setValue(activity);
    rankSheet.getRange(second + row).setValue(why);
    rankSheet.getRange(third + row).setValue(how);
    rankSheet.getRange(last + row).setValue(score);
  }
}

// UTILS

function getStringOnSheetForPos(sheet, column, row) {
  return String(sheet.getRange(column + row).getValue());
}

function getStringForPos(column, row) {
  return String(
    getCurrentSheet()
      .getRange(column + row)
      .getValue()
  );
}

function getCurrentSheet() {
  return SpreadsheetApp.getActive();
}

function getTodaysDate() {
  return Utilities.formatDate(new Date(), "GMT-3", "dd/MM/yyyy");
}

function getLastPopulatedRow(sheet) {
  var avals = sheet.getRange("A1:A").getValues();
  var alast = avals.filter(String).length;
  return alast;
}

function checkForValidNumber(value) {
  return !isNaN(value) && value > 0;
}

function resetValueForRange(range) {
  range.setValue(null);
}

function checkForCurrentMonth(date) {
  var month = parseInt(date.substring(3, 5));
  var currentMonth = parseInt(String(getTodaysDate()).substring(3, 5));

  return month == currentMonth;
}
