function myFunction() {

  // Main function to update the calendar sheet
  function updateCalendarSheet() {
    var calendarId = '1312'; // Replace with your calendar ID
    var sheetId = '1312'; // Replace with your sheet ID
    var sheet = SpreadsheetApp.openById(sheetId).getSheetByName('Sheet1');
    
    // Fetch events from Google Calendar
    var now = new Date(2024, 8, 1);  // Start date
    var future = new Date(2025, 11, 31);  // End date
    var events = CalendarApp.getCalendarById(calendarId).getEvents(now, future);
    
    // Process events to gather data
    var eventData = [];
    var totalHours = 0; // Total hours for the year
    var monthlyHours = {}; // to store hours per month
    for (var i = 0; i < events.length; i++) {
      var event = events[i];
      var start = event.getStartTime();
      var end = event.getEndTime();
      var duration = (end - start) / (1000 * 60 * 60);  // Duration in hours

      eventData.push([ 
        event.getTitle(),
        Utilities.formatDate(start, Session.getScriptTimeZone(), 'dd/MM/yyyy'),
        Utilities.formatDate(start, Session.getScriptTimeZone(), 'HH.mm'),
        Utilities.formatDate(end, Session.getScriptTimeZone(), 'HH.mm'),
        duration
      ]);

      // Accumulate total and monthly hours
      totalHours += duration;
      var month = Utilities.formatDate(start, Session.getScriptTimeZone(), 'MMMM');
      monthlyHours[month] = (monthlyHours[month] || 0) + duration;
    }

    // Clear existing data and formatting in the sheet
    sheet.clear();
    sheet.clearFormats();

    // Insert Events Table (B2:F2)
    sheet.getRange('B2:F2').setValues([['Event Name', 'Date', 'Start Time', 'End Time', 'Hours']]);
    sheet.getRange(3, 2, eventData.length, eventData[0].length).setValues(eventData);
    sheet.getRange(2, 2, eventData.length + 1, eventData[0].length)
      .setBorder(true, true, true, true, true, true)
      .setBackground('#f4cccc');
    sheet.getRange('B2:F2')
      .setFontWeight('bold')
      .setBorder(true, true, true, true, true, true)
      .setBackground('#ea9999');

    // Insert Yearly Summary Table (H3:I3)
    sheet.getRange('H3:I3').setValues([['Total Hours', 'Total Euro']])
      .setFontWeight('bold')
      .setBackground('#93c47d')
      .setBorder(true, true, true, true, true, true);
    sheet.getRange('H4').setValue(totalHours)
      .setBackground('#b6d7a8')
      .setBorder(true, true, true, true, true, true);
    sheet.getRange('I4').setValue(totalHours * 17)
      .setBackground('#b6d7a8')
      .setBorder(true, true, true, true, true, true);

    // Insert Monthly Summary Table (K3:M3)
    sheet.getRange('K3:M3').setValues([['Month', 'Hours', 'Euro']])
      .setFontWeight('bold')
      .setBackground('#b4a7d6')
      .setBorder(true, true, true, true, true, true);

    // Prepare monthly data for insertion
    var monthData = [];
    for (var month in monthlyHours) {
      var monthTotal = monthlyHours[month];  // Total hours for the month
      var monthEuro = monthTotal * 17;  // Calculate Euro (hours * 17)
      monthData.push([month, monthTotal, monthEuro]);  // Add both total hours and Euro to the array
    }

    // Insert Monthly Data (K4:M...)
    if (monthData.length > 0) {
      sheet.getRange(4, 11, monthData.length, 3).setValues(monthData);
      sheet.getRange(4, 11, monthData.length, 3)
        .setBorder(true, true, true, true, true, true)
        .setBackground('#d9d2e9');
    }

    // Format text alignment
    sheet.getRange(2, 2, sheet.getMaxRows() - 1, sheet.getMaxColumns() - 1)
      .setHorizontalAlignment('center')
      .setVerticalAlignment('middle');

    // Set font family and size for the entire sheet
    sheet.getRange(1, 1, sheet.getMaxRows(), sheet.getMaxColumns())
      .setFontFamily('Arial')
      .setFontSize(10);

    // Call the function to adjust specific column widths
    adjustSpecificColumnWidths(sheet);
  }

  // Function to adjust specific column widths
  function adjustSpecificColumnWidths(sheet) {
    // Array of columns to adjust: B=2, C=3, D=4, E=5, F=6, H=8, I=9, K=11, L=12, M=13
    var columnsToAdjust = [2, 3, 4, 5, 6, 8, 9, 11, 12, 13];

    // Loop through the specified columns
    columnsToAdjust.forEach(function(col) {
      // Auto resize the column
      sheet.autoResizeColumn(col);

      // Get current column width
      var currentWidth = sheet.getColumnWidth(col);

      // Increase the width by 50%
      var newWidth = currentWidth * 1.5;

      // Set the new width for the column
      sheet.setColumnWidth(col, newWidth);
    });
  }

  // Execute the main function
  updateCalendarSheet();
}
