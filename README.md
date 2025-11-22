# Sheets Wizard

A production-ready Model Context Protocol (MCP) server for Google Sheets automation.

## Features

- **create_sheet** - Create new Google Spreadsheets
- **get_data** - Retrieve data from spreadsheet ranges
- **update_cells** - Update cell values in spreadsheets
- **add_formula** - Add formulas to cells
- **create_chart** - Create charts and visualizations
- **export_csv** - Export sheets as CSV format

## Installation

```bash
npm install
npm run build
```

## Configuration

### Google Cloud Setup

1. Create a project in [Google Cloud Console](https://console.cloud.google.com/)
2. Enable the Google Sheets API
3. Create OAuth 2.0 credentials
4. Download the credentials JSON

### Credentials Format

Tools require credentials in the following JSON format:

```json
{
  "client_id": "your-client-id",
  "client_secret": "your-client-secret",
  "redirect_uri": "your-redirect-uri",
  "access_token": "your-access-token",
  "refresh_token": "your-refresh-token"
}
```

## Usage

### Claude Desktop Configuration

Add to your `claude_desktop_config.json`:

```json
{
  "mcpServers": {
    "sheets-wizard": {
      "command": "node",
      "args": ["/path/to/sheets-wizard/dist/index.js"]
    }
  }
}
```

### Tool Examples

#### Create a New Sheet

```javascript
{
  "title": "My New Spreadsheet",
  "credentials": "{...}"
}
```

#### Get Data from Range

```javascript
{
  "spreadsheetId": "1abc...",
  "range": "Sheet1!A1:D10",
  "credentials": "{...}"
}
```

#### Update Cells

```javascript
{
  "spreadsheetId": "1abc...",
  "range": "Sheet1!A1:B2",
  "values": [["Name", "Age"], ["John", "30"]],
  "credentials": "{...}"
}
```

#### Add Formula

```javascript
{
  "spreadsheetId": "1abc...",
  "range": "Sheet1!C1",
  "formula": "=SUM(A1:A10)",
  "credentials": "{...}"
}
```

#### Create Chart

```javascript
{
  "spreadsheetId": "1abc...",
  "sheetId": 0,
  "chartType": "COLUMN",
  "sourceRange": "Sheet1!A1:B10",
  "credentials": "{...}"
}
```

#### Export as CSV

```javascript
{
  "spreadsheetId": "1abc...",
  "sheetId": 0,
  "credentials": "{...}"
}
```

## Development

```bash
# Install dependencies
npm install

# Build
npm run build

# Run
npm start
```

## Requirements

- Node.js 18+
- Google Cloud account with Sheets API enabled
- OAuth 2.0 credentials

## License

MIT

## Author

consigcody94
