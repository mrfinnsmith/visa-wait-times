/*
This works.

What it does: Checks the State department site daily (should be a trigger) to get visa appointment wait times from consulates around the world.

To do:
* Make it only update data if the data for that consulate is different from the preceding data (not sure if this is a good idea)
* Put important variables like sheet ID as script properties
* Make it more OOP
* More logs and email alerts if something fails.
* More statuses — if they're making appointments, make a "open" status or something

*/

function updateVisaWaitTimeHTML() {
  const dayOfWeek = new Date().getDay();
  if (dayOfWeek == 0 || dayOfWeek == 6) {
    console.log('Today is a weekend day. Exiting updateVisaWaitTimeHTML.');
    return;
  }
  importDataFromWeb();
}

function importDataFromWeb() {
  var propertiesService = PropertiesService.getScriptProperties();
  var url = propertiesService.getProperty('VISA_WAIT_TIME_PAGE');
  var response = UrlFetchApp.fetch(url);
  var html = response.getContentText();

  var tableData = parseHtmlTable(html);

  var htmlSheetId = propertiesService.getProperty('HTML_SHEET_ID');
  if (!htmlSheetId) {
    throw new Error('Sheet ID not found in script properties');
  }
  htmlSheetId = parseInt(htmlSheetId, 10);

  var sheets = SpreadsheetApp.getActiveSpreadsheet().getSheets();
  var htmlSheet;

  for (var i = 0; i < sheets.length; i++) {
    if (sheets[i].getSheetId() === htmlSheetId) {
      htmlSheet = sheets[i];
      break;
    }
  }

  if (!htmlSheet) {
    throw new Error('Did not find a sheet with an ID that matches the script property HTML_SHEET_ID');
  }

  var headerRow = findHeaderRow(htmlSheet);
  htmlSheet.insertRowsAfter(headerRow, tableData.length);
  htmlSheet.getRange(headerRow + 1, 1, tableData.length, tableData[0].length).setValues(tableData);

  updateConditionalFormatting(htmlSheet);
}

function findHeaderRow(sheet) {
  var column = sheet.getRange("A:A");
  var values = column.getValues();
  for (var i = 0; i < values.length; i++) {
    if (values[i][0].toString().trim().toLowerCase() === "date") {
      return i + 1;
    }
  }
  throw new Error("Header row not found");
}

function parseHtmlTable(html) {
  var tableRegex = /<table[\s\S]*?<\/table>/;
  var tableMatch = html.match(tableRegex);
  if (!tableMatch) {
    throw new Error("No table found in the HTML content");
  }

  var tableHtml = tableMatch[0];
  var rowRegex = /<tr[\s\S]*?<\/tr>/g;
  var rows = tableHtml.match(rowRegex);

  var headerRow = rows[0];
  var cellRegex = /<t[dh][^>]*>([\s\S]*?)<\/t[dh]>/g;
  var headers = headerRow.match(cellRegex).map(function(cell) {
    return cell.replace(/<[^>]+>/g, '').trim();
  });

  var visaTypes = headers.slice(1);

  var tableData = [];
  var currentDate = new Date();
  var formattedDate = Utilities.formatDate(currentDate, Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm:ss");

  for (var i = 1; i < rows.length; i++) {
    var rowHtml = rows[i];
    var cells = rowHtml.match(cellRegex);
    if (cells) {
      var cityPost = cells[0].replace(/<[^>]+>/g, '').trim();

      for (var j = 1; j < cells.length; j++) {
        if (cells[j]) {
          var appointmentWaitTimeResponse = cells[j].replace(/<[^>]+>/g, '').replace(/&nbsp;/g, ' ').trim();
          if (appointmentWaitTimeResponse == '') {
            continue;
          }
          var appointmentWaitTimeCalendarDays = parseAppointmentWaitTime(appointmentWaitTimeResponse);
          if (appointmentWaitTimeResponse.toLowerCase() === 'same day') {
            appointmentWaitTimeCalendarDays = 0;
          }

          var visaTypeResponse = visaTypes[j - 1].replace(/&nbsp;/g, ' ');
          var visaType = visaTypeResponse;
          if (visaType.toLowerCase().startsWith('interview required')) {
            visaType = visaType.replace(/interview required\s*/i, '').trim();
          }

          var status = '';
          if (appointmentWaitTimeResponse.length > 0 && (appointmentWaitTimeCalendarDays === ''|| appointmentWaitTimeCalendarDays == 0)) {
            status = appointmentWaitTimeResponse;
          }

          var rowData = [
            formattedDate,
            cityPost,
            visaTypeResponse,
            visaType,
            appointmentWaitTimeResponse,
            appointmentWaitTimeCalendarDays,
            status
          ];
          tableData.push(rowData);
        }
      }
    }
  }

  return tableData;
}

function parseAppointmentWaitTime(text) {
  var daysMatch = text.match(/(\d+)\s*(day|days)/i);
  if (daysMatch) {
    return parseInt(daysMatch[1], 10);
  } else {
    return '';
  }
}

function updateConditionalFormatting(sheet) {
  var lastRow = sheet.getLastRow();
  var range = sheet.getRange("A2:B" + lastRow);

  var rule = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied("=AND($A2=$A1,$B2=$B1)")
    .setBackground("#FFFFFF")
    .setFontColor("#FFFFFF")
    .setRanges([range])
    .build();

  sheet.clearConditionalFormatRules();
  sheet.setConditionalFormatRules([rule]);
}