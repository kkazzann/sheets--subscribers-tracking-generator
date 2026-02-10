var Service = {
  
  getAvailableMonths: function(showArchived) {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(Config.SHEETS.SOURCE);
    if (!sheet) return [];

    const lastRow = sheet.getLastRow();
    const data = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
    const uniqueMonths = new Set();
    const currentMonthStart = new Date(new Date().getFullYear(), new Date().getMonth(), 1);

    data.forEach(row => {
      const dateCell = row[0];
      if (dateCell instanceof Date) {
        if (!showArchived && dateCell < currentMonthStart) return;
        const formatted = Utilities.formatDate(dateCell, ss.getSpreadsheetTimeZone(), "MM/yyyy");
        uniqueMonths.add(formatted);
      }
    });

    return Array.from(uniqueMonths).sort((a, b) => {
      const [m1, y1] = a.split('/').map(Number);
      const [m2, y2] = b.split('/').map(Number);
      return y1 !== y2 ? y1 - y2 : m1 - m2;
    });
  },

  generateTable: function(monthYear) {
    if (!monthYear) throw new Error("No month selected.");

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sourceSheet = ss.getSheetByName(Config.SHEETS.SOURCE);
    const [month, year] = monthYear.split('/');
    const targetSheetName = Config.SHEETS.TARGET_PREFIX + year;

    let targetSheet = ss.getSheetByName(targetSheetName);
    if (!targetSheet) {
      targetSheet = ss.insertSheet(targetSheetName);
    }

    // Fetch range covering all necessary columns (A to H)
    const lastRowSrc = sourceSheet.getLastRow();
    // Getting 8 columns to reach index 7 (Column H)
    const sourceData = sourceSheet.getRange(2, 1, lastRowSrc - 1, 8).getValues();
    const validDates = [];

    // Filter Logic
    sourceData.forEach(row => {
      const dateCell = row[Config.COLUMNS.DATE];
      if (!(dateCell instanceof Date)) return;

      const rowMonth = Utilities.formatDate(dateCell, ss.getSpreadsheetTimeZone(), "MM/yyyy");
      if (rowMonth !== monthYear) return;

      // Extract and sanitize data
      const dayName = (row[Config.COLUMNS.DAY] || "").toString().trim();
      const nameVal = (row[Config.COLUMNS.NAME] || "").toString().trim();
      const personVal = (row[Config.COLUMNS.PERSON] || "").toString().trim();

      // Rule 1: Skip if Name is CGB
      if (nameVal.toUpperCase() === "CGB") return;

      // Rule 2: Skip empty Tuesdays (Tuesday + Empty Name + Empty Person)
      if (dayName === "Tuesday" && nameVal === "" && personVal === "") return;

      validDates.push(dateCell);
    });

    if (validDates.length === 0) return "No valid dates found for " + monthYear;

    this.writeTableToSheet(targetSheet, validDates);
    return "Table created successfully.";
  },

  writeTableToSheet: function(sheet, dates) {
    const countryNames = Object.keys(Config.COUNTRY_MAP);
    
    // Build Matrix
    const headerRow = ["Country", ...dates];
    const matrix = [headerRow];
    const emptyCols = new Array(dates.length).fill("");

    countryNames.forEach(country => {
      const url = Config.URLS.BASE + (Config.COUNTRY_MAP[country] || "");
      const formula = `=HYPERLINK("${url}"; "${country}")`;
      matrix.push([formula, ...emptyCols]);
    });

    matrix.push(["Package count", ...emptyCols]);

    // Dimensions
    const lastRow = sheet.getLastRow();
    const startRow = lastRow === 0 ? 1 : lastRow + 3;
    const numRows = matrix.length;
    const numCols = headerRow.length;

    // Write Data
    const range = sheet.getRange(startRow, 1, numRows, numCols);
    range.setValues(matrix);

    // Styling
    this.applyStyling(sheet, range, startRow, numRows, numCols, countryNames.length);
  },

  applyStyling: function(sheet, range, startRow, numRows, numCols, numCountries) {
    // 1. General Styles
    range.setFontFamily(Config.STYLE.FONT_NAME)
         .setFontSize(Config.STYLE.FONT_SIZE)
         .setFontWeight("normal")
         .setVerticalAlignment("middle")
         .setHorizontalAlignment("center")
         .setBorder(false, false, false, false, false, false);

    // 2. Country Column (Align Left, Link Color)
    sheet.getRange(startRow, 1, numRows, 1)
         .setHorizontalAlignment("left")
         .setFontColor(Config.STYLE.LINK_COLOR);

    // 3. Header (Bold, Bg Color, Date Format, Black Text)
    sheet.getRange(startRow, 1, 1, numCols)
         .setFontWeight("bold")
         .setBackground(Config.STYLE.HEADER_BG)
         .setNumberFormat("dd/MM/yyyy")
         .setFontColor("black");

    // 4. Footer (Bold, Black Text)
    sheet.getRange(startRow + numRows - 1, 1, 1, numCols)
         .setFontWeight("bold")
         .setFontColor("black");

    // 5. Conditional Formatting
    const rules = sheet.getConditionalFormatRules();
    
    // Sunday Rule (Gray Background for country rows only)
    const bodyRange = sheet.getRange(startRow + 1, 2, numCountries, numCols - 1);
    const sundayFormula = `=WEEKDAY(B$${startRow}; 2)=7`;
    rules.push(SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied(sundayFormula)
      .setBackground(Config.STYLE.SUNDAY_BG)
      .setRanges([bodyRange])
      .build());

    // Empty Package Rule (Red Background for footer)
    const footerRange = sheet.getRange(startRow + numRows - 1, 2, 1, numCols - 1);
    rules.push(SpreadsheetApp.newConditionalFormatRule()
      .whenCellEmpty()
      .setBackground(Config.STYLE.ERROR_BG)
      .setRanges([footerRange])
      .build());

    sheet.setConditionalFormatRules(rules);
  }
};