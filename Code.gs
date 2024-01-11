function formatdate(date) {
  if (date) {
    var originalDate = new Date(date);
    if (!isNaN(originalDate.getTime())) {
      return Utilities.formatDate(originalDate, 'GMT', 'MM/dd/yyyy');
    }
  }
  return "-";
}

function createConditionalFormatRule(sheet) {
  var range = sheet.getRange("B2:G" + (sheet.getMaxRows() - 1));
  var rule = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied("=AND($C2<TODAY(),$G2=FALSE)")
    .setBackground("#EA9999")
    .setRanges([range])
    .build();
  var rules = sheet.getConditionalFormatRules();
  rules.push(rule);
  sheet.setConditionalFormatRules(rules);
}

function applyAlternatingColors(sheet) {
  var range = sheet.getRange("B2:G" + (sheet.getMaxRows() - 1));
  var banding = range.applyRowBanding(SpreadsheetApp.BandingTheme.LIGHT_GREY, false, false); // apply a default banding theme with no header and no footer
  banding.setFirstRowColor("#FFFFFF"); // set the first row color to white
  banding.setSecondRowColor("#F3F3F3"); // set the second row color to light grey
}

function run(token, spreadsheet, namespace) {
  var headers = {
  "Authorization": "Bearer " + token
};

// Set the options for the request
var options = {
  "method": "get",
  "headers": headers
};
__main(options, spreadsheet, namespace);
}

function __main(options, spreadsheet, namespace) {
  spreadsheet.toast("", "Refreshing Data...")
  var overviewSheet = spreadsheet.getSheetByName("Overview");

  overviewSheet.getRange("A:F").clear()

  // Apply background colors
  var headerRange = overviewSheet.getRange("1:1");
  var completeRange = overviewSheet.getRange("B:B");
  var incompleteRange = overviewSheet.getRange("C:C");
  var duesoonRange = overviewSheet.getRange("D:D");
  var dueoneweekRange = overviewSheet.getRange("E:E");
  var overdueRange = overviewSheet.getRange("F:F");
  headerRange.setFontWeight("bold");
  completeRange.setBackground("#93C47D")
  incompleteRange.setBackground("#FFD966")
  duesoonRange.setBackground("#F6B26B")
  dueoneweekRange.setBackground("#76A5AF")
  overdueRange.setBackground("#E06666")

  // Set header row
  overviewSheet.getRange("A1").setValue("Class");
  overviewSheet.getRange("B1").setValue("Complete");
  overviewSheet.getRange("C1").setValue("Incomplete");
  overviewSheet.getRange("D1").setValue("Due Soon");
  overviewSheet.getRange("E1").setValue("Due in 1 week");
  overviewSheet.getRange("F1").setValue("Overdue");

  // Initialize variables for calculating total rows and row index
  var totalRows = 2;
  var rowIndex = 2;

  var url = "https://" + namespace + "/api/v1/courses"

  // Make the request and log the response
  var response = UrlFetchApp.fetch(url, options);
  var data = JSON.parse(response.getContentText());
  const date = new Date();

  const utcTime = date.getTime();
  const mstTime = new Date(utcTime);

  const mstMonth = mstTime.getMonth() + 1;
  const mstYear = mstTime.getFullYear();
  const month = mstMonth.toString();
  const year = mstYear.toString().slice(-2);
  if (month < 6) {
    var termCode = "(" + year + "W)";
  } else if (month > 6 && month < 8) {
    var termCode = "(" + year + "S)";
  } else {
    var termCode = "(" + year + "F)";
  }

  for (var i = 1; i < Object.keys(data).length; i++) {
    var course_id = data[i]["id"];
    var course_name = data[i]["name"]
    if (course_name.startsWith(termCode)) {
      Logger.log(course_id);
      Logger.log(course_name);
      Logger.log("------------------------------")
      var existingSheet = spreadsheet.getSheetByName(course_name);

      if (existingSheet) {
        // Delete the existing sheet
        spreadsheet.deleteSheet(existingSheet);
      }

      var s = spreadsheet.insertSheet(course_name);

      s.setColumnWidth(1, 21);
      s.setColumnWidth(8, 21);

      // Bold and center cells B2:G2
      var headerRange = s.getRange('B2:G2');
      headerRange.setFontWeight('bold');
      headerRange.setHorizontalAlignment('center');
      headerRange.setVerticalAlignment('middle');

      var rangeToFormat = s.getRange('A:A');
      rangeToFormat.setBackground('black');
      rangeToFormat.setFontColor('black');
      var rangeToFormat = s.getRange('H:H');
      rangeToFormat.setBackground('black');
      rangeToFormat.setFontColor('black');
      var rangeToFormat = s.getRange('1:1');
      rangeToFormat.setBackground('black');
      rangeToFormat.setFontColor('black');

      s.getRange('A1').setValue(course_id)

      s.getRange('B2').setValue('Assignment');
      s.getRange('C2').setValue('Due Date');
      s.getRange('D2').setValue('Submission Date');
      s.getRange('E2').setValue('Pts Rcvd');
      s.getRange('F2').setValue('Worth');
      s.getRange('G2').setValue('Submitted?');

      s.deleteColumns(9, 18);
      var row = 3;
      var url = "https://" + namespace + "/api/v1/courses/" + course_id + "/assignments?per_page=100"
      var response = UrlFetchApp.fetch(url, options);
      var adata = JSON.parse(response.getContentText());

      for (var i1 = 1; i1 < adata.length; i1++) {
        var assignment_id = adata[i1]["id"];
        var submission_url = "https://" + namespace + "/api/v1/courses/" + course_id + "/assignments/" + assignment_id + "/submissions/self"
        var response = UrlFetchApp.fetch(submission_url, options);
        var sdata = JSON.parse(response.getContentText());

        var assignment_name = adata[i1]["name"];
        var assignment_due_at = formatdate(adata[i1]["due_at"]);
        var assignment_worth = adata[i1]["points_possible"];
        var assignment_submitted = Boolean(adata[i1]["has_submitted_submissions"]);
        if (assignment_submitted) {
          var assignment_submitted_at = formatdate(sdata["submitted_at"]);
          var assignment_score = sdata["score"];
        } else {
          var assignment_submitted_at = "-";
          var assignment_score = "-";
        }
        var assignment_type = adata[i1]["is_quiz_assignment"];
        s.getRange('A' + row).setValue(assignment_id);
        s.getRange('B' + row).setValue(assignment_name);
        s.getRange('C' + row).setValue(assignment_due_at);
        s.getRange('D' + row).setValue(assignment_submitted_at);
        s.getRange('E' + row).setValue(assignment_score);
        s.getRange('F' + row).setValue(assignment_worth);
        if (assignment_submitted_at == "-") {
          assignment_submitted = false;
        }
        s.getRange('G' + row).setValue(assignment_submitted);
        s.getRange('G' + row).insertCheckboxes();
        row++
      }

      var pagesurl = "https://" + namespace + "/api/v1/courses/" + course_id + "/pages?per_page=100"
      var response = UrlFetchApp.fetch(pagesurl, options);
      var pages = JSON.parse(response.getContentText());
      for (var i2 = 1; i2 < pages.length; i2++) {
        var page_id = pages[i2]["url"];
        var page_name = pages[i2]["title"]
        var page_due_at = "-"
        if (pages[i2]["todo_date"]) {
          var page_due_at = formatdate(pages[i2]["todo_date"])
        }
        s.getRange('A' + row).setValue(page_id);
        s.getRange('B' + row).setValue(page_name);
        s.getRange('C' + row).setValue(page_due_at);
        s.getRange('D' + row).setValue("-");
        s.getRange('E' + row).setValue("-");
        s.getRange('F' + row).setValue("-");
        s.getRange('G' + row).setValue("-");
        row++
      }
      s.autoResizeColumns(2, s.getLastColumn());
      var columnsToAdjust = [4, 5, 6, 7];
      for (var i3 = 0; i3 < columnsToAdjust.length; i3++) {
        var columnWidth = s.getColumnWidth(columnsToAdjust[i3]);
        s.setColumnWidth(columnsToAdjust[i3], columnWidth + 30);
        s.setColumnWidth(2, s.getColumnWidth(2) + 10)
        s.setColumnWidth(3, s.getColumnWidth(3) + 10)
      }


      // Calculate the first empty row based on assignments, pages, and 3
      var totalRowsNeeded = adata.length + pages.length + 3;

      // Get the first empty row
      var firstEmptyRow = s.getLastRow() + 1;

      // Set the background of the row below the first empty row to black
      s.getRange('A' + firstEmptyRow + ':G' + firstEmptyRow).setBackground('black');

      // Delete all rows from the one below the final black one
      if (firstEmptyRow > 0) {
        s.deleteRows(firstEmptyRow + 1, s.getMaxRows() - firstEmptyRow);
      }
      createConditionalFormatRule(s);
      applyAlternatingColors(s);

      overviewSheet.getRange("A" + rowIndex).setValue(course_name);
      overviewSheet.getRange("B" + rowIndex).setFormula('=COUNTIF(\'' + course_name + '\'!$G$3:$G$' + (rowIndex + 2) + ', "TRUE")');
      overviewSheet.getRange("C" + rowIndex).setFormula('=COUNTIF(\'' + course_name + '\'!$G$3:$G$' + (rowIndex + 2) + ', "FALSE")');
      overviewSheet.getRange("D" + rowIndex).setFormula('=COUNTIFS(\'' + course_name + '\'!$G$3:$G$' + (rowIndex + 2) + ', FALSE, \'' + course_name + '\'!$C$3:$C$' + (rowIndex + 2) + ', ">"&(TODAY()+3))');
      overviewSheet.getRange("E" + rowIndex).setFormula('=COUNTIFS(\'' + course_name + '\'!$G$3:$G$' + (rowIndex + 2) + ', FALSE, \'' + course_name + '\'!$C$3:$C$' + (rowIndex + 2) + ', ">="&(TODAY()+3)), \'' + course_name + '\'!$C$3:$C$' + (rowIndex + 2) + ', "<="&(TODAY()+7))');
      overviewSheet.getRange("F" + rowIndex).setFormula('=COUNTIFS(\'' + course_name + '\'!$G$3:$G$' + (rowIndex + 2) + ', FALSE, \'' + course_name + '\'!$C$3:$C$' + (rowIndex + 2) + ', "<"&(TODAY()))');
      totalRows++;
      rowIndex++;

      var dataRange = s.getRange(2, 2, s.getLastRow() - 1, s.getLastColumn() - 1);
      dataRange.createFilter();

      Logger.log('Refreshed at ' + spreadsheet.getUrl() + '#gid=' + s.getSheetId());
    }
    // Auto-resize columns
    overviewSheet.autoResizeColumns(1, 6);
    var columnsToAdjust = [1, 2, 3, 4, 5, 6];
    for (var i5 = 0; i5 < columnsToAdjust.length; i5++) {
      var columnWidth = overviewSheet.getColumnWidth(columnsToAdjust[i5]);
      overviewSheet.setColumnWidth(columnsToAdjust[i5], columnWidth + 10);
    }
  }
  spreadsheet.toast("", "Refreshed", 2)
}
