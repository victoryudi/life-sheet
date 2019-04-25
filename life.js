function onEdit(e) {
  if (shouldResetDashboard()) {
    var data = getTodaysData();
    populateDataSheet(data);
    resetSheet();
  }

  if (shouldUpdateScoreSum()) {
    var scoreSum = getScoreSum();
    updateScoreSum(scoreSum);
  }

  if (getCurrentSheet().getSheetName() == "tasks") {
    var types = ["personal", "professional", "logosophy"];

    for (var i = 0; i < types.length; i++) {
      var type = types[i];
      var info = getMappedInfoForType(type);
      populateType(type, info);
    }
  }
}

function updateScoreSum(scoreSum) {
  var sheet = getCurrentSheet().getSheetByName("dashboard");
  sheet.getRange("G8").setValue(scoreSum);
}

function getScoreSum() {
  var scoreSum = 0;

  scoreSum += getScoreSumForColumns("A", "E");
  scoreSum += getScoreSumForColumns("F", "J");
  scoreSum += getScoreSumForColumns("K", "O");

  return scoreSum;
}

function getScoreSumForColumns(checkColumn, scoreColumn) {
  var sheet = getCurrentSheet().getSheetByName("dashboard");
  var avals = sheet.getRange(scoreColumn + "13:" + scoreColumn).getValues();
  var alast = avals.filter(String).length;
  var scoreSum = 0;

  if (alast > 0) {
    for (var i = 0; i <= alast; i++) {
      if (getStringForPos(checkColumn, i + 13) == "✅") {
        var scoreToSum = parseInt(
          sheet.getRange(scoreColumn + String(i + 13)).getValue()
        );

        if (scoreToSum === scoreToSum) {
          scoreSum += scoreToSum;
        }
      }
    }
  }

  return scoreSum;
}

function shouldUpdateScoreSum() {
  if (getCurrentSheet().getSheetName() == "dashboard") {
    var currentCell = getCurrentSheet().getActiveCell();
    var value = String(currentCell.getValue());
    var column = currentCell.getColumn();

    if (column == "1" || column == "6" || column == "11") {
      return true;
    }
  }

  return false;
}

function resetSheet() {
  dashboard = getCurrentSheet().getSheetByName("dashboard");
  resetValueForRange(dashboard.getRange("A13:N"));
  resetValueForRange(dashboard.getRange("L2"));
  resetValueForRange(dashboard.getRange("Q2"));

  tasks = getCurrentSheet().getSheetByName("tasks");
  resetValueForRange(tasks.getRange("A2:G"));
}

function resetValueForRange(range) {
  range.setValue(null);
}

function shouldResetDashboard() {
  if (getCurrentSheet().getSheetName() == "dashboard") {
    if (
      getCurrentSheet()
        .getRange("Q2")
        .getValue() == "✅"
    ) {
      return true;
    }
  }

  return false;
}

function populateDataSheet(data) {
  dataSheet = getCurrentSheet().getSheetByName("data");

  var row = getLastPopulatedRow(dataSheet) + 1;
  var leftColumn = "B";
  var doneColumn = "C";
  var percentageColumn = "D";
  var notesColumn = "E";
  var tasksSumColumn = "F";

  var tasksLeft = data["left"];
  var done = data["done"];
  var percentage = data["percentage"];
  var notes = data["notes"];
  var tasksSum = data["tasksSum"];

  dataSheet.getRange(leftColumn + row).setValue(tasksLeft);
  dataSheet.getRange(doneColumn + row).setValue(done);
  dataSheet.getRange(percentageColumn + row).setValue(percentage);
  dataSheet.getRange(notesColumn + row).setValue(notes);
  dataSheet.getRange(tasksSumColumn + row).setValue(tasksSum);
}

function getLastPopulatedRow(sheet) {
  var avals = sheet.getRange("A1:A").getValues();
  var alast = avals.filter(String).length;
  return alast;
}

function getTodaysData() {
  dashboardSheet = getCurrentSheet().getSheetByName("dashboard");

  var tasksDataRow = "8";
  var tasksLeftColumn = "B";
  var doneColumn = "G";
  var percentageColumn = "L";
  var notesColumn = "L";
  var notesRow = "2";

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

  var tasksSheet = getCurrentSheet().getSheetByName("tasks");
  var tasksSum = String(tasksSheet.getRange("A2:G").getValues());

  return getMappedData(tasksLeft, doneTasks, percentage, notes, tasksSum);
}

function getMappedData(left, done, percentage, notes, tasksSum) {
  var data = {};

  data["left"] = left;
  data["done"] = done;
  data["percentage"] = percentage;
  data["notes"] = notes;
  data["tasksSum"] = tasksSum;

  return data;
}

function getMappedInfoForType(type) {
  var activityColumn = "B";
  var howColumn = "G";
  var whyColumn = "F";
  var scoreColumn = "H";
  var typeColumn = "A";

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
      var missing = "MISSING";
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

  info["activities"] = activities;
  info["hows"] = hows;
  info["whys"] = whys;
  info["scores"] = scores;

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
  var missingBgColor = "#FF5950";
  var rankSheet = getCurrentSheet().getSheetByName("dashboard");
  var sortedActivities = mappedInfo["activities"];
  var sortedHows = mappedInfo["hows"];
  var sortedWhys = mappedInfo["whys"];
  var sortedScores = mappedInfo["scores"];
  var lastIndex = sortedActivities.length - 1;

  var first = "B";
  var second = "C";
  var third = "D";
  var last = "E";

  if (type == "professional") {
    first = "G";
    second = "H";
    third = "I";
    last = "J";
  } else if (type == "logosophy") {
    first = "L";
    second = "M";
    third = "N";
    last = "O";
  }

  for (var i = 0; i <= lastIndex; i++) {
    var row = i + 13;
    var activity = sortedActivities[i];
    var how = sortedHows[i];
    var why = sortedWhys[i];
    var score = sortedScores;

    var bgRange = first + row + ":" + last + row;
    if (activity == "MISSING") {
      rankSheet.getRange(bgRange).setBackgroundColor(missingBgColor);
    } else {
      rankSheet.getRange(bgRange).setBackgroundColor(null);
    }

    rankSheet.getRange(first + row).setValue(activity);
    rankSheet.getRange(second + row).setValue(why);
    rankSheet.getRange(third + row).setValue(how);
    rankSheet.getRange(last + row).setValue(score);
  }
}

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
