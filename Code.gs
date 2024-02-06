var Utils = {

  formatdate: function (date) {
    if (date) {
      var originalDate = new Date(date);
      if (!isNaN(originalDate.getTime())) {
        return Utilities.formatDate(originalDate, 'GMT', 'MM/dd/yyyy');
      }
    }
    return "-";
  },

  createConditionalFormatRule: function (sheet) {
    var range = sheet.getRange("B2:G" + (sheet.getMaxRows() - 1));
    var rule = SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied("=AND($C2<TODAY(),$G2=FALSE)")
      .setBackground("#EA9999")
      .setRanges([range])
      .build();
    var rules = sheet.getConditionalFormatRules();
    rules.push(rule);
    sheet.setConditionalFormatRules(rules);
  },

  applyAlternatingColors: function (sheet) {
    var range = sheet.getRange("B2:G" + (sheet.getMaxRows() - 1));
    var banding = range.applyRowBanding(SpreadsheetApp.BandingTheme.LIGHT_GREY, false, false);
    banding.setFirstRowColor("#FFFFFF");
    banding.setSecondRowColor("#F3F3F3");
  },

  checkUserStatus: function (user) {
    var url = "https://docs.veillax.com/docs/canvas-spreadsheet-integration/banned-users.json"
    var response = UrlFetchApp.fetch(url, { "method": "get" });
    var data = JSON.parse(response.getContentText());

    const isBlocked = Object.keys(data).includes(user);

    if (isBlocked) {
      const blockReason = data[user]; // Replace with the actual property representing the block reason
      throw new MissingPermissions(user, blockReason);
    }
  },
  betaCheck: function (user) {
    var url = "https://docs.veillax.com/docs/canvas-spreadsheet-integration/beta-testers.json"
    var response = UrlFetchApp.fetch(url, { "method": "get" });
    var data = JSON.parse(response.getContentText());

    const isBeta = data["isBeta"]
    const email_from_betalist = Object.keys(data).toString().toLowerCase()

    const beta = email_from_betalist.includes(user);

    if (!beta && isBeta) {
      throw new MissingPermissions(user, "User is not enrolled in the Closed Beta Program");
    }
  },
  main: function (options, spreadsheet, namespace, url_override) {
    spreadsheet.toast("", "Refreshing Data...")
    var overviewSheet = spreadsheet.getSheetByName("Overview");
    Logger.log(overviewSheet)
    var sheet1 = spreadsheet.getSheetByName("Sheet1");

    if (!overviewSheet) {
      spreadsheet.insertSheet("Overview")
      overviewSheet = spreadsheet.getSheetByName("Overview");
    }
    try {
      overviewSheet.getRange("A:F").clear()
    } catch { }
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

    var url = "https://" + namespace + "." + url_override + "/api/v1/courses"

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

    for (var i = 0; i < Object.keys(data).length; i++) {
      var course_id = data[i]["id"];
      var course_name = data[i]["name"]
      if (course_name && !course_name.includes("Homeroom")) {
        Logger.log(course_id);
        Logger.log(course_name);
        Logger.log("------------------------------")
        var existingSheet = spreadsheet.getSheetByName(course_name);

        if (existingSheet) {
          // Delete the existing sheet
          spreadsheet.deleteSheet(existingSheet);
        }

        var s = spreadsheet.insertSheet(course_name);
        spreadsheet.toast("Refreshing " + course_name, "Refreshing Data")

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
        var url = "https://" + namespace + "." + url_override + "/api/v1/courses/" + course_id + "/assignments?per_page=100"
        var response = UrlFetchApp.fetch(url, options);
        var adata = JSON.parse(response.getContentText());

        for (var i1 = 0; i1 < adata.length; i1++) {
          try {
            var assignment_id = adata[i1]["id"];
            var submission_url = "https://" + namespace + "." + url_override + "/api/v1/courses/" + course_id + "/assignments/" + assignment_id + "/submissions/self"
            var response = UrlFetchApp.fetch(submission_url, options);
            var sdata = JSON.parse(response.getContentText());

            var assignment_name = adata[i1]["name"];
            var assignment_due_at = Utils.formatdate(adata[i1]["due_at"]);
            var assignment_worth = adata[i1]["points_possible"];
            var assignment_submitted = Boolean(adata[i1]["has_submitted_submissions"]);
            if (assignment_submitted) {
              var assignment_submitted_at = Utils.formatdate(sdata["submitted_at"]);
              var assignment_score = sdata["score"];
            } else {
              var assignment_submitted_at = "-";
              var assignment_score = "-";
            }
            var assignment_type = adata[i1]["is_quiz_assignment"];
            s.getRange('A' + row).setValue(assignment_id);
            s.getRange('B' + row).setFormula('=HYPERLINK("https://' + namespace + '.' + url_override + '/courses/' + course_id + '/assignments/' + assignment_id + '", "' + assignment_name + '")');
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
          } catch { }
        }


        var pagesurl = "https://" + namespace + "." + url_override + "/api/v1/courses/" + course_id + "/pages?per_page=100"
        try {
          var response = UrlFetchApp.fetch(pagesurl, options);
        } catch { }
        var pages = JSON.parse(response.getContentText());
        for (var i2 = 0; i2 < pages.length; i2++) {
          try {
            var page_id = pages[i2]["url"];
            var page_name = pages[i2]["title"]
            var page_due_at = "-"
            if (pages[i2]["todo_date"]) {
              var page_due_at = Utils.formatdate(pages[i2]["todo_date"])
            }
            s.getRange('A' + row).setValue(page_id);
            s.getRange('B' + row).setFormula('=HYPERLINK("https://' + namespace + '.' + url_override + '/courses/' + course_id + '/pages/' + page_id + '", "' + page_name + '")');
            s.getRange('C' + row).setValue(page_due_at);
            s.getRange('D' + row).setValue("-");
            s.getRange('E' + row).setValue("-");
            s.getRange('F' + row).setValue("-");
            s.getRange('G' + row).setValue("-");
            row++
          } catch { }
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
        var totalRowsNeeded = adata.length + pages.length + 2;
        s.getRange('A' + (totalRowsNeeded) + ':G' + (totalRowsNeeded + 1)).setBackground('black');
        // Delete all rows from the one below the final black one
        if (totalRowsNeeded > 0) {
          s.deleteRows(totalRowsNeeded + 2, s.getMaxRows() - totalRowsNeeded - 1);
        }
        Utils.createConditionalFormatRule(s);
        Utils.applyAlternatingColors(s);
        s.getRange("A3:H" + totalRowsNeeded).sort({column: 3, ascending: true})

        overviewSheet.getRange("A" + rowIndex).setFormula('=HYPERLINK("https://' + namespace + '.' + url_override + '/courses/' + course_id + '", "' + course_name + '")')
        overviewSheet.getRange("B" + rowIndex).setFormula('=COUNTIF(\'' + course_name + '\'!$G:$G, "TRUE")');
        overviewSheet.getRange("C" + rowIndex).setFormula('=COUNTIF(\'' + course_name + '\'!$G:$G, "FALSE")');
        overviewSheet.getRange("D" + rowIndex).setFormula('=COUNTIFS(\'' + course_name + '\'!$G:$G, "FALSE", \'' + course_name + '\'!$C:$C, "<"&(TODAY()+3))');
        overviewSheet.getRange("E" + rowIndex).setFormula('=COUNTIFS(\'' + course_name + '\'!$G:$G, "FALSE", \'' + course_name + '\'!$C:$C, "<"&(TODAY()+7))');
        overviewSheet.getRange("F" + rowIndex).setFormula('=COUNTIFS(\'' + course_name + '\'!$G:$G, "FALSE", \'' + course_name + '\'!$C:$C, "<"&(TODAY()))');
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
};

class MissingPermissions extends Error {
  constructor(userName, blockReason) {
    const message = `Access to this service for '${userName}' has been denied. Reason: ${blockReason}`;
    super(message);
    this.name = 'MissingPermissionsError';
    this.userName = userName;
    this.blockReason = blockReason;
  }
}

function run(token, spreadsheet, namespace, override = false, url_override = undefined) {
  var email = Session.getEffectiveUser().getEmail();

  Utils.betaCheck(email);
  Utils.checkUserStatus(email);
  email = ""
  var headers = {
    "Authorization": "Bearer " + token
  };

  // Set the options for the request
  var options = {
    "method": "get",
    "headers": headers
  };
  if (url_override == undefined && override) {
    throw ("Please provide a url_override")
  }
  if (override) {
    Utils.main(options, spreadsheet, namespace, url_override);
  } else {
    Utils.main(options, spreadsheet, namespace, "instructure.com")
  }
}
