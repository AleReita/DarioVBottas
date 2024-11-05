function myFunction() {

    function updateCalendarSheet() {
      var calendarId = '1312'; // Replace with your calendar ID
      var sheetId = '1312'; // Replace with your sheet ID
      var sheet = SpreadsheetApp.openById(sheetId).getSheetByName('Calendar');
      
      // Clear existing data and formatting in the sheet
      sheet.clear();
      sheet.clearFormats();
      
      // Fetch events from Google Calendar
      var now = new Date(2024, 9, 1);  // Start date
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
      
      // Insert Events Table
      sheet.getRange('B2:F2').setValues([['Nome Evento', 'Data', 'Ora Inizio', 'Ora Fine', 'Ore']]);
      sheet.getRange(3, 2, eventData.length, eventData[0].length).setValues(eventData);
      sheet.getRange(2, 2, eventData.length + 1, eventData[0].length).setBorder(true, true, true, true, true, true).setBackground('#f4cccc');
      sheet.getRange('B2:F2').setFontWeight('bold').setBorder(true, true, true, true, true, true).setBackground('#ea9999');
  
      // Insert Yearly Summary Table in columns H and I
      sheet.getRange('H3:I3').setValues([['Ore', 'Euro']]).setFontWeight('bold').setBackground('#93c47d').setBorder(true, true, true, true, true, true);
      sheet.getRange('H4').setValue(totalHours).setBackground('#b6d7a8').setBorder(true, true, true, true, true, true);
      sheet.getRange('I4').setValue(totalHours * 17).setBackground('#b6d7a8').setBorder(true, true, true, true, true, true);
  
      // Insert Monthly Summary Table in columns K and L
      sheet.getRange('K3:M3').setValues([['Mese', 'Ore', 'Euro']])
        .setFontWeight('bold')
        .setBorder(true, true, true, true, true, true)
        .setBackground('#b4a7d6');
  
      // Prepare monthly data for insertion
      var monthData = [];
      for (var month in monthlyHours) {
        var monthTotal = monthlyHours[month];  // Total hours for the month
        var monthEuro = monthTotal * 17;  // Calculate Euro (hours * 17)
        monthData.push([month, monthTotal, monthEuro]);  // Add both total hours and Euro to the array
      }
  
      // Insert Monthly Data
      if (monthData.length > 0) {
        sheet.getRange(4, 11, monthData.length, 3).setValues(monthData); // Insert Month, Hours, and Euro
        sheet.getRange(4, 11, monthData.length, 3).setBorder(true, true, true, true, true, true)
          .setBackground('#d9d2e9'); // Apply border and background color to the data cells
      }
  
         // Formatting all text to be centered
      sheet.getRange(2, 2, sheet.getMaxRows() - 1, sheet.getMaxColumns() - 1)
        .setHorizontalAlignment('center')
        .setVerticalAlignment('middle');
     
      // Set font to Arial and size to 11 for all characters in the sheet
      sheet.getRange(1, 1, sheet.getMaxRows(), sheet.getMaxColumns())
        .setFontFamily('Arial')
        .setFontSize(11);
  
      // Regola la larghezza delle colonne alla fine
      adjustColumnWidth(sheet);
      
  
    }
  
    function adjustColumnWidth(sheet) {
      // Auto resize columns based on content
      sheet.autoResizeColumns(1, sheet.getMaxColumns());
    
      // Get the current column widths and increase each by 50%
      var numColumns = sheet.getMaxColumns();
      for (var col = 1; col <= numColumns; col++) {
        var currentWidth = sheet.getColumnWidth(col);
        var newWidth = currentWidth * 1.5;  // Increase by 50%
        sheet.setColumnWidth(col, newWidth);
      }
    }
  
    function resetEmptyColumnWidths(sheet) {
      var numColumns = sheet.getMaxColumns();
      var defaultWidth = 100;  // Default column width in pixels
  
      // Loop through all columns and reset the width for empty ones
      for (var col = 1; col <= numColumns; col++) {
        var range = sheet.getRange(1, col, sheet.getMaxRows());
        if (range.isBlank()) {  // Check if column is empty
          sheet.setColumnWidth(col, defaultWidth);  // Reset to default width
        }
      }
    }
  
    // Execute the main function
    updateCalendarSheet();
  }
  