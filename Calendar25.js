function myFunction() {

  // Funzione principale per aggiornare il foglio del calendario
  function updateCalendarSheet() {
    var calendarId = '1312'; // ID del calendario
    var sheetId = '1312'; // ID del foglio
    var sheet = SpreadsheetApp.openById(sheetId).getSheetByName('1312');
    
    // Estrai gli eventi da Google Calendar
    var now = new Date(2024, 8, 1);  // Data di inizio
    var future = new Date(2025, 11, 31);  // Data di fine
    var events = CalendarApp.getCalendarById(calendarId).getEvents(now, future);
    
    // Processa gli eventi per ottenere i dati
    var eventData = [];
    var totalHours = 0; // Ore totali per l'anno
    var monthlyHours = {}; // Ore per mese
    for (var i = 0; i < events.length; i++) {
      var event = events[i];
      var start = event.getStartTime();
      var end = event.getEndTime();
      var duration = (end - start) / (1000 * 60 * 60);  // Durata in ore

      eventData.push([ 
        event.getTitle(),
        Utilities.formatDate(start, Session.getScriptTimeZone(), 'dd/MM/yyyy'),
        Utilities.formatDate(start, Session.getScriptTimeZone(), 'HH.mm'),
        Utilities.formatDate(end, Session.getScriptTimeZone(), 'HH.mm'),
        duration
      ]);

      // Aggiungi ore totali e mensili
      totalHours += duration;
      var month = Utilities.formatDate(start, Session.getScriptTimeZone(), 'MMMM');
      monthlyHours[month] = (monthlyHours[month] || 0) + duration;
    }

    // Pulisci dati e formattazioni esistenti nel foglio
    sheet.clear();
    sheet.clearFormats();

    // Inserisci tabella degli eventi
    sheet.getRange('B2:F2').setValues([['Event Name', 'Date', 'Start Time', 'End Time', 'Hours']]);
    sheet.getRange(3, 2, eventData.length, eventData[0].length).setValues(eventData);
    sheet.getRange(2, 2, eventData.length + 1, eventData[0].length)
      .setBorder(true, true, true, true, true, true)
      .setBackground('#f4cccc');
    sheet.getRange('B2:F2')
      .setFontWeight('bold')
      .setBorder(true, true, true, true, true, true)
      .setBackground('#ea9999');

    // Inserisci tabella di riepilogo annuale
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

    // Inserisci tabella di riepilogo mensile
    sheet.getRange('H10:J10').setValues([['Month', 'Hours', 'Euro']])
      .setFontWeight('bold')
      .setBackground('#b4a7d6')
      .setBorder(true, true, true, true, true, true);

    // Prepara i dati mensili per l'inserimento
    var monthData = [];
    for (var month in monthlyHours) {
      var monthTotal = monthlyHours[month];  // Ore totali per il mese
      var monthEuro = monthTotal * 17;  // Calcolo Euro
      monthData.push([month, monthTotal, monthEuro]);  // Aggiunge ore e Euro all'array
    }

    // Inserisci dati mensili
    if (monthData.length > 0) {
      sheet.getRange(11, 8, monthData.length, 3).setValues(monthData);
      sheet.getRange(11, 8, monthData.length, 3)
        .setBorder(true, true, true, true, true, true)
        .setBackground('#d9d2e9');
    }

    // Formattazione testo allineato al centro
    sheet.getRange(2, 2, sheet.getMaxRows() - 1, sheet.getMaxColumns() - 1)
      .setHorizontalAlignment('center')
      .setVerticalAlignment('middle');

    // Imposta il font per l'intero foglio
    sheet.getRange(1, 1, sheet.getMaxRows(), sheet.getMaxColumns())
      .setFontFamily('Arial')
      .setFontSize(10);

    // Richiama la funzione per regolare la larghezza delle colonne
    adjustSpecificColumnWidths(sheet);

    // Crea il grafico mensile
    createMonthlyHoursChart(sheet); // Passa il foglio alla funzione del grafico
  }

  // Funzione per creare un grafico a barre delle ore mensili
  function createMonthlyHoursChart(sheet) {
    // Rimuovi eventuali grafici esistenti nella posizione scelta
    var existingCharts = sheet.getCharts();
    existingCharts.forEach(function(chart) {
      sheet.removeChart(chart);
    });

    // Crea un nuovo grafico a barre per le ore mensili
var chartBuilder = sheet.newChart()
  .setChartType(Charts.ChartType.COLUMN)  // Usa un grafico a colonne
  .addRange(sheet.getRange("H11:H"))  // Intervallo per i mesi (asse X)
  .addRange(sheet.getRange("I11:I"))  // Intervallo per le ore (asse Y)
  .setPosition(3, 12, -50, 0)  // Posizione del grafico
  .setOption('title', 'Monthly Hours')  // Titolo del grafico
  .setOption('hAxis', { title: '', textStyle: { color: '#1c4587', fontSize: 12 } })  // Asse orizzontale (mesi)
  .setOption('vAxis', { title: '', textStyle: { color: '#1c4587', fontSize: 12 } })  // Asse verticale (ore)
  .setOption('width', 400)  // Larghezza del grafico
  .setOption('height', 300)  // Altezza del grafico
  .setOption('colors', ['#6aa84f'])  // Colore delle colonne
  .setOption('legend', { position: 'none' })  // Nasconde la legenda
  .setOption('titleTextStyle', { color: '#0b5394', fontSize: 14, bold: true });  // Stile del titolo

sheet.insertChart(chartBuilder.build());

    // Inserisce il grafico nel foglio
    sheet.insertChart(chartBuilder.build());
  }

// Funzione per regolare le larghezze delle colonne
function adjustSpecificColumnWidths(sheet) {
  // Array di colonne da regolare: B=2, C=3, D=4, E=5, F=6, H=8, I=9
  var columnsToAdjust = [2, 3, 4, 5, 6, 8, 9];

  // Ciclo attraverso le colonne specificate (escludendo K, L, M)
  columnsToAdjust.forEach(function(col) {
    // Auto resize della colonna
    sheet.autoResizeColumn(col);

    // Ottieni la larghezza corrente della colonna
    var currentWidth = sheet.getColumnWidth(col);

    // Aumenta la larghezza del 50%
    var newWidth = currentWidth * 1.5;

    // Imposta la nuova larghezza per la colonna
    sheet.setColumnWidth(col, newWidth);
  });

  // Imposta larghezza fissa per le colonne K, L, M
  sheet.setColumnWidth(1, 50);  // Colonna M (imposta larghezza fissa)
  sheet.setColumnWidth(7, 50);  // Colonna M (imposta larghezza fissa)
  sheet.setColumnWidth(11, 100);  // Colonna K (imposta larghezza fissa)
  sheet.setColumnWidth(12, 100);  // Colonna L (imposta larghezza fissa)
  sheet.setColumnWidth(13, 100);  // Colonna M (imposta larghezza fissa)
}


  // Esegui la funzione principale per aggiornare il foglio
  updateCalendarSheet();
}
