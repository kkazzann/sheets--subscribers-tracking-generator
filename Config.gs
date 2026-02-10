var Config = {
  // Spreadsheet Settings
  SHEETS: {
    SOURCE: "HTML Schedule",
    TARGET_PREFIX: "QA - Subscribers "
  },

  // Data Source Column Mapping (0-based index)
  // A=0, B=1, C=2, D=3, E=4, F=5, G=6, H=7
  COLUMNS: {
    DATE: 0,   // Column A
    DAY: 5,    // Column F (Actual day name, skipping hidden cols)
    NAME: 6,   // Column G
    PERSON: 7  // Column H
  },

  // Visual Styling
  STYLE: {
    FONT_NAME: "Poppins",
    FONT_SIZE: 10,
    HEADER_BG: "#d9ead3",      // Light Green
    SUNDAY_BG: "#eeeeee",      // Light Gray
    ERROR_BG: "#ff0000",       // Red
    LINK_COLOR: "#434343",     // Dark Gray 4
    BORDER: false
  },

  // URLs
  URLS: {
    BASE: "https://www.prologistics.info/react/reports_page/customers_newsletter/?filter_id="
  },

  // Country to Filter ID Mapping
  COUNTRY_MAP: {
    "CH-DE": "11607",
    "CH-DE RICARDO": "11606",
    "CH-FR": "11604",
    "CH-FR RICARDO": "11605",
    "AT": "46175",
    "BENL": "1309711",
    "BEFR": "1309715",
    "CZ": "11619",
    "DE": "11608",
    "DE AVANDEO": "11609",
    "DK": "11618",
    "FI": "11616",
    "FR": "11612",
    "HU": "11615",
    "IT": "11610",
    "NL": "11614",
    "NO": "79358",
    "PL": "11620",
    "PT": "11617",
    "RO": "1323241",
    "SE": "11603",
    "SK": "165840",
    "ES": "11613",
    "UK": "11621"
  }
};