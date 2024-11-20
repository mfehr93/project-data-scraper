function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Generate Project Data')
      .addItem('Run Script', 'runScript')
      .addToUi();
}

function runScript() {
  var ui = SpreadsheetApp.getUi();
  var response = ui.prompt('Enter start row and end row in the format "startRow, endRow"', 'Format: startRow,endRow', ui.ButtonSet.OK_CANCEL);

  if (response.getSelectedButton() == ui.Button.OK) {
    var input = response.getResponseText().split(',');
    var startRow = parseInt(input[0]);
    var endRow = parseInt(input[1]);

    if (isNaN(startRow) || isNaN(endRow) || startRow > endRow) {
      ui.alert('Invalid input. Please enter valid row numbers.');
      return;
    }

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getActiveSheet();
    var data = sheet.getRange(startRow, 1, endRow - startRow + 1, sheet.getLastColumn()).getValues();

    // Create new sheet with today's date
    var today = new Date();
    var formattedDate = Utilities.formatDate(today, Session.getScriptTimeZone(), 'yyyy-MM-dd');
    var newSheetName = 'Project Data ' + formattedDate;
    var newSheet = ss.getSheetByName(newSheetName) || ss.insertSheet(newSheetName);

    // Clear any existing content in the cells of the sheet
    newSheet.clear();
    // Clear any existing charts or graphs from previous run
    var charts = newSheet.getCharts();
    for (var i = 0; i < charts.length; i++) {
      newSheet.removeChart(charts[i]);
    }

    var avgEstTime = getAverageEstimatedProjectTime(data);
    var avgTime = getAverageActualProjectTime(data);
    var medianTime = getMedianProjectTime(data);
    var estLaborHours = getEstimatedLaborHours(data);
    var actLaborHours = getActualLaborHours(data);
    getPriorityGraph(data, newSheet);
    graphAvgEmployeeProjectTime(data, newSheet);
    getDepartmentOriginGraph(data, newSheet);
    getProjectTypeGraph(data, newSheet);
    newSheet.getRange('B4').setValue(String(estLaborHours.toFixed(2)));
    newSheet.getRange('B5').setValue(String(actLaborHours.toFixed(2)));
    getEstimatedVsActualGraph(newSheet);

    newSheet.getRange('A1').setValue('Average Estimated Project Completion Time: ' + avgEstTime.toFixed(2));
    newSheet.getRange('A2').setValue('Average Project Completion Time: ' + avgTime.toFixed(2));
    newSheet.getRange('A3').setValue('Median Project Time: ' + medianTime.toFixed(2));
    newSheet.getRange('A4').setValue('Estimated LH:');
    newSheet.getRange('A5').setValue('Actual LH:');

    Logger.log('Average Estimated Project Completion Time: ' + avgEstTime);
    Logger.log('Average Project Completion Time: ' + avgTime);
    Logger.log('Median Project Time: ' + medianTime);
    Logger.log('Estimated Labor Hours: ' + estLaborHours);
    Logger.log('Actual Labor Hours: ' + actLaborHours);
    
    

    ui.alert('Script executed successfully. Check the Logger for details.');
  }
}

function getAverageEstimatedProjectTime(data) {
  var sum = 0;
  for (var i = 0; i < data.length; i++) {
    if canConvertToFloat(data[i][5]){
      sum += data[i][5]; // Column F
    }
    else{
      throw new Error("Estimated Labor Hours is not a number for project: \"".concat(data[i][3]).concat("\""));
    }
  }
  return sum / data.length;
}

function getAverageActualProjectTime(data) {
  var sum = 0;
  for (var i = 0; i < data.length; i++) {
    if canConvertToFloat(data[i][6]){
      sum += data[i][6]; // Column G
    }
    else{
      throw new Error("Actual Labor Hours is not a number for project: \"".concat(data[i][3]).concat("\""));
    }
  }
  return sum / data.length;
}

function getMedianProjectTime(data) {
  var times = [];
  for (var i = 0; i < data.length; i++) {
    times.push(data[i][6]); // Column G
  }
  times.sort(function(a, b) { return a - b; });

  var half = Math.floor(times.length / 2);
  if (times.length % 2) {
    return times[half];
  } else {
    return (times[half - 1] + times[half]) / 2.0;
  }
}

function canConvertToFloat(value) {
  const parsedValue = parseFloat(value);
  return !isNaN(parsedValue);
}

function getEstimatedLaborHours(data) {
  var sum = 0;
  for (var i = 0; i < data.length; i++) {
    if (canConvertToFloat(data[i][5])){ // Column F
      sum += parseFloat(data[i][5])
    } 
    else {
      throw new Error("Estimated Labor Hours is not a number for project ".concat(data[i][3]));
    }
  }
  return sum;
}

function getActualLaborHours(data) {
  var sum = 0;
  for (var i = 0; i < data.length; i++) {
    if (canConvertToFloat(data[i][6])){ // Column G
      sum += parseFloat(data[i][6])
    } 
    else {
      throw new Error("Actual Labor Hours is not a number for project ".concat(data[i][3]));
    }
  }
  return sum;
}

function getPriorityGraph(data, newSheet) {
  var priorityCounts = [0, 0, 0, 0, 0];
  for (var i = 0; i < data.length; i++) {
    var priority = data[i][8]; // Column I
    if (priority >= 1 && priority <= 5) {
      priorityCounts[priority - 1]++;
    }
  }

  newSheet.getRange('A7').setValue('Priority');
  newSheet.getRange('B7').setValue('Count');

  for (var j = 0; j < 5; j++) {
    newSheet.getRange('A' + (j + 8)).setValue(j + 1);
    newSheet.getRange('B' + (j + 8)).setValue(priorityCounts[j]);
  }

  var chartBuilder = newSheet.newChart()
      .setChartType(Charts.ChartType.COLUMN)
      .addRange(newSheet.getRange('A7:B12'))
      .setPosition(55, 8, 0, 0)
      .setOption('title', 'Priority Levels of Projects')
      .build();

  newSheet.insertChart(chartBuilder);
}

function graphAvgEmployeeProjectTime(data, newSheet) {
  var employeeTimes = {};
  var employeeProjects = {};

  for (var i = 0; i < data.length; i++) {
    if (data[i][7] == ""){
      throw new Error("Something is funny with the completed by field for the project titled: ".concat(data[i][3]));
    }
    var employees = getEmployeesFromColumnH(data[i][7]); // Column H
    for (var j = 0; j < employees.length; j++) {
      var employee = employees[j];
      if (!employeeTimes[employee]) {
        employeeTimes[employee] = 0;
        employeeProjects[employee] = 0;
      }
      employeeTimes[employee] += data[i][6]; // Column G
      employeeProjects[employee]++;
    }
  }

  var avgTimes = [];
  for (var employee in employeeTimes) {
    avgTimes.push([employee, employeeProjects[employee],employeeTimes[employee] / employeeProjects[employee]]);
  }

  newSheet.getRange('D7').setValue('Employee');
  newSheet.getRange('E7').setValue('# of Projects');
  newSheet.getRange('F7').setValue('Avg Time');

  for (var k = 0; k < avgTimes.length; k++) {
    newSheet.getRange('D' + (k + 8)).setValue(avgTimes[k][0]);
    newSheet.getRange('E' + (k + 8)).setValue(avgTimes[k][1]);
    newSheet.getRange('F' + (k + 8)).setValue(avgTimes[k][2].toFixed(2));
  }

  var chartBuilderAverage = newSheet.newChart()
      .setChartType(Charts.ChartType.BAR)
      .addRange(newSheet.getRange('D8:D' + (avgTimes.length + 7)))
      .addRange(newSheet.getRange('F8:F' + (avgTimes.length + 7)))
      .setPosition(17, 8, 0, 0)
      .setOption('title', 'Average Hours Spent Per Project')
      .build();

  var chartBuilderTotal = newSheet.newChart()
      .setChartType(Charts.ChartType.BAR)
      .addRange(newSheet.getRange('D8:D' + (avgTimes.length + 7)))
      .addRange(newSheet.getRange('E8:E' + (avgTimes.length + 7)))
      .setPosition(17, 1, 0, 0)
      .setOption('title', 'Total Number of Projects Completed')
      .build();

  var chartBuilderPie = newSheet.newChart()
      .setChartType(Charts.ChartType.PIE)
      .addRange(newSheet.getRange('D8:E' + (avgTimes.length + 7)))
      .setPosition(74, 1, 0, 0)
      .setOption('title', 'Share of the Workload (Number of Projects)')
      .build();

  newSheet.insertChart(chartBuilderPie);
  newSheet.insertChart(chartBuilderAverage);
  newSheet.insertChart(chartBuilderTotal);
}

function getEmployeesFromColumnH(cellValue) {
  var matches = cellValue.toLowerCase().match(/[a-zA-Z\s]*by (.+)/);
  const splitRegex = /,\s+and\s|\s+and\s+|,\s+|\s+/
  if (matches) {
    return matches[1].split(splitRegex).map(function(name) { return cleanText(name.trim()); });
  }
  return [];
}

function cleanText(inputString) {
  return inputString.charAt(0).toUpperCase() + inputString.slice(1).toLowerCase();
}

function getDepartmentOriginGraph(data, newSheet) {
  var departmentCounts = {};
  for (var i = 0; i < data.length; i++) {
    var department = cleanText(data[i][1]); // Column B
    if (department == "") {
      throw new Error("The department field is empty for project ".concat(data[i][3]));
    }
    if (!departmentCounts[department]) { // Create a new field for this department if we haven't encountered it yet
      departmentCounts[department] = 0;
    }
    departmentCounts[department]++;
  }

  var departmentData = [];
  for (var department in departmentCounts) {
    departmentData.push([department, departmentCounts[department]]);
  }

  newSheet.getRange('H7').setValue('Department');
  newSheet.getRange('I7').setValue('Count');

  for (var j = 0; j < departmentData.length; j++) {
    newSheet.getRange('H' + (j + 8)).setValue(departmentData[j][0]);
    newSheet.getRange('I' + (j + 8)).setValue(departmentData[j][1]);
  }

  var chartBuilder = newSheet.newChart()
      .setChartType(Charts.ChartType.PIE)
      .addRange(newSheet.getRange('H7:I' + (departmentData.length + 7)))
      .setOption('title', 'Department Origin of Projects')
      .setPosition(36, 1, 0, 0)
      .build();

  newSheet.insertChart(chartBuilder);
}

function getProjectTypeGraph(data, newSheet) {
  var typeCounts = {
    'Repair': 0,
    'Improvement': 0,
    'PM': 0,
  };
  for (var i = 0; i < data.length; i++) {
    var type = data[i][2].trim(); // Column C
    if (typeCounts[type] !== undefined) {
      typeCounts[type]++;
    }
    else {
      throw new Error("The following type of project is not accounted for: ".concat(String(type)));
    }
  }

  var typeData = [];
  for (var type in typeCounts) {
    typeData.push([type, typeCounts[type]]);
  }

  newSheet.getRange('K7').setValue('Type');
  newSheet.getRange('L7').setValue('Count');

  for (var j = 0; j < typeData.length; j++) {
    newSheet.getRange('K' + (j + 8)).setValue(typeData[j][0]);
    newSheet.getRange('L' + (j + 8)).setValue(typeData[j][1]);
  }

  var chartBuilder = newSheet.newChart()
      .setChartType(Charts.ChartType.PIE)
      .addRange(newSheet.getRange('K7:L' + (typeData.length + 7)))
      .setOption('title', 'Type of Project')
      .setPosition(36, 8, 0, 0)
      .build();

  newSheet.insertChart(chartBuilder);
}

function getEstimatedVsActualGraph(newSheet) {
  var chartBuilder = newSheet.newChart()
      .setChartType(Charts.ChartType.BAR)
      .setOption('title', 'Estimated vs Actual Project Labor Hours')
      .addRange(newSheet.getRange('A4:B5'))
      .setPosition(55, 1, 0, 0)
      .build();
  
  newSheet.insertChart(chartBuilder);
}