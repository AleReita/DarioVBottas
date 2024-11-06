function myFunction() {

  // Main function to update the calendar sheet
  function updateCalendarSheet() {
    var calendarId = '1312'; // Calendar ID
    var sheetId = '1312'; // Sheet ID
    var sheet = SpreadsheetApp.openById(sheetId).getSheetByName('S1312');
    
    // Extract events from Google Calendar
    var now = new Date(2024, 8, 1);  // Start date
    var future = new Date(2025, 11, 31);  // End date
    var events = CalendarApp.getCalendarById(calendarId).getEvents(now, future);
    
    // Process events to extract data
    var eventData = [];
    var totalHours = 0; // Total hours for the year
    var monthlyHours = {}; // Hours per month
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

      // Add total and monthly hours
      totalHours += duration;
      var month = Utilities.formatDate(start, Session.getScriptTimeZone(), 'MMMM');
      monthlyHours[month] = (monthlyHours[month] || 0) + duration;
    }

    // Clear existing data and formatting from the sheet
    sheet.clear();
    sheet.clearFormats();

    // Insert event table
    sheet.getRange('B2:F2').setValues([['Event Name', 'Date', 'Start Time', 'End Time', 'Hours']]);
    sheet.getRange(3, 2, eventData.length, eventData[0].length).setValues(eventData);
    sheet.getRange(2, 2, eventData.length + 1, eventData[0].length)
      .setBorder(true, true, true, true, true, true)
      .setBackground('#f4cccc');
    sheet.getRange('B2:F2')
      .setFontWeight('bold')
      .setBorder(true, true, true, true, true, true)
      .setBackground('#ea9999');

    // Insert annual summary table
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

    // Insert monthly summary table
    sheet.getRange('K3:M3').setValues([['Month', 'Hours', 'Euro']])
      .setFontWeight('bold')
      .setBackground('#b4a7d6')
      .setBorder(true, true, true, true, true, true);

    // Prepare monthly data for insertion
    var monthData = [];
    for (var month in monthlyHours) {
      var monthTotal = monthlyHours[month];  // Total hours for the month
      var monthEuro = monthTotal * 17;  // Calculate Euro
      monthData.push([month, monthTotal, monthEuro]);  // Add hours and Euro to the array
    }

    // Insert monthly data
    if (monthData.length > 0) {
      sheet.getRange(4, 11, monthData.length, 3).setValues(monthData);
      sheet.getRange(4, 11, monthData.length, 3)
        .setBorder(true, true, true, true, true, true)
        .setBackground('#d9d2e9');
    }

    // Format text to be centered
    sheet.getRange(2, 2, sheet.getMaxRows() - 1, sheet.getMaxColumns() - 1)
      .setHorizontalAlignment('center')
      .setVerticalAlignment('middle');

    // Set the font for the entire sheet
    sheet.getRange(1, 1, sheet.getMaxRows(), sheet.getMaxColumns())
      .setFontFamily('Arial')
      .setFontSize(10);

    // Call the function to adjust the column widths
    adjustSpecificColumnWidths(sheet);

    // Create the monthly hours chart
    createMonthlyHoursChart(sheet); // Pass the sheet to the chart function
  }

  // Function to create a bar chart for monthly hours
  function createMonthlyHoursChart(sheet) {
    // Remove any existing charts in the chosen position
    var existingCharts = sheet.getCharts();
    existingCharts.forEach(function(chart) {
      sheet.removeChart(chart);
    });

    // Create a new bar chart for monthly hours
    var chartBuilder = sheet.newChart()
      .setChartType(Charts.ChartType.COLUMN)  // Use a column chart
      .addRange(sheet.getRange("K4:K"))  // Range for months (X-axis)
      .addRange(sheet.getRange("L4:L"))  // Range for hours (Y-axis)
      .setPosition(10, 8, -75, 0)  // Position of the chart
      .setOption('title', 'Monthly Hours')  // Chart title
      .setOption('hAxis', { title: 'Month', textStyle: { color: '#1c4587', fontSize: 12 } })  // Horizontal axis (months)
      .setOption('vAxis', { title: 'Hours', textStyle: { color: '#1c4587', fontSize: 12 } })  // Vertical axis (hours)
      .setOption('width', 350)  // Chart width
      .setOption('height', 250)  // Chart height
      .setOption('colors', ['#6aa84f'])  // Column color
      .setOption('legend', { position: 'none' })  // Hide legend
      .setOption('titleTextStyle', { color: '#0b5394', fontSize: 14, bold: true });  // Title style

    sheet.insertChart(chartBuilder.build());

    // Insert the chart into the sheet
    sheet.insertChart(chartBuilder.build());
  }

  // Function to adjust specific column widths
  function adjustSpecificColumnWidths(sheet) {
    // Array of columns to adjust: B=2, C=3, D=4, E=5, F=6, H=8, I=9, K=11, L=12, M=13
    var columnsToAdjust = [2, 3, 4, 5, 6, 8, 9, 11, 12, 13];

    // Loop through the specified columns
    columnsToAdjust.forEach(function(col) {
      // Auto resize the column
      sheet.autoResizeColumn(col);

      // Get the current width of the column
      var currentWidth = sheet.getColumnWidth(col);

      // Increase the width by 50%
      var newWidth = currentWidth * 1.5;

      // Set the new width for the column
      sheet.setColumnWidth(col, newWidth);
    });
  }

  // Run the main function to update the sheet
  updateCalendarSheet();
}
