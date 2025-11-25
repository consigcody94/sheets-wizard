# Sheets Wizard

A Model Context Protocol (MCP) server for Google Sheets automation. Create spreadsheets, read and write data, add formulas, create charts, and export to CSV.

## Overview

Sheets Wizard brings Google Sheets into your AI workflow. Create spreadsheets, manipulate data, add formulas and charts, all through natural language.

### Why Use Sheets Wizard?

**Traditional workflow:**
- Open Google Sheets in browser
- Manually navigate and click through UI
- Write formulas one cell at a time
- Export manually for data analysis

**With Sheets Wizard:**
```
"Create a new spreadsheet for Q4 expenses"
"Get all data from the Sales sheet"
"Update cells A1:B10 with the monthly totals"
"Add a SUM formula for the revenue column"
"Create a bar chart from the sales data"
```

## Features

- **Spreadsheet Creation** - Create new Google Spreadsheets
- **Data Reading** - Retrieve data from any range
- **Data Writing** - Update cells with values
- **Formula Support** - Add formulas to cells
- **Chart Creation** - Create various chart types
- **CSV Export** - Export sheets as CSV format

## Installation

```bash
# Clone the repository
git clone https://github.com/consigcody94/sheets-wizard.git
cd sheets-wizard

# Install dependencies
npm install

# Build the project
npm run build
```

## Configuration

### Google Cloud Setup

1. Go to [Google Cloud Console](https://console.cloud.google.com/)
2. Create a new project or select existing
3. Enable the **Google Sheets API**:
   - Go to "APIs & Services" → "Library"
   - Search "Google Sheets API"
   - Click "Enable"
4. Create OAuth 2.0 credentials:
   - Go to "APIs & Services" → "Credentials"
   - Click "Create Credentials" → "OAuth client ID"
   - Select "Web application"
   - Add authorized redirect URI
   - Download the credentials JSON

### Getting OAuth Tokens

After creating OAuth credentials, you need to get access tokens:

1. Use the OAuth 2.0 Playground or your own OAuth flow
2. Authorize with scopes: `https://www.googleapis.com/auth/spreadsheets`
3. Exchange authorization code for tokens
4. Save access_token and refresh_token

### Credentials Format

```json
{
  "client_id": "your-client-id.apps.googleusercontent.com",
  "client_secret": "your-client-secret",
  "redirect_uri": "your-redirect-uri",
  "access_token": "ya29.your-access-token",
  "refresh_token": "1//your-refresh-token"
}
```

### Claude Desktop Integration

Add to your `claude_desktop_config.json`:

```json
{
  "mcpServers": {
    "sheets-wizard": {
      "command": "node",
      "args": ["/absolute/path/to/sheets-wizard/dist/index.js"]
    }
  }
}
```

**Config file locations:**
- macOS: `~/Library/Application Support/Claude/claude_desktop_config.json`
- Windows: `%APPDATA%\Claude\claude_desktop_config.json`
- Linux: `~/.config/Claude/claude_desktop_config.json`

## Tools Reference

### create_sheet

Create a new Google Spreadsheet.

**Parameters:**

| Parameter | Type | Required | Description |
|-----------|------|----------|-------------|
| `title` | string | Yes | Spreadsheet title |
| `credentials` | string | Yes | JSON string of OAuth credentials |

**Example:**

```json
{
  "title": "Q4 2024 Budget",
  "credentials": "{\"client_id\":\"...\",\"client_secret\":\"...\",\"access_token\":\"...\",\"refresh_token\":\"...\"}"
}
```

**Response includes:**
- Spreadsheet ID (for future operations)
- Spreadsheet URL (direct link)
- Title confirmation

### get_data

Retrieve data from a spreadsheet range.

**Parameters:**

| Parameter | Type | Required | Description |
|-----------|------|----------|-------------|
| `spreadsheetId` | string | Yes | Spreadsheet ID |
| `range` | string | Yes | A1 notation range |
| `credentials` | string | Yes | OAuth credentials |

**Example:**

```json
{
  "spreadsheetId": "1abc123def456",
  "range": "Sheet1!A1:D10",
  "credentials": "{...}"
}
```

**A1 Notation examples:**
- `Sheet1!A1:D10` - Specific range on Sheet1
- `Sheet1!A:D` - Columns A through D
- `Sheet1!1:10` - Rows 1 through 10
- `A1:D10` - Range on first sheet
- `Sheet1` - Entire sheet

**Response includes:**
- Range retrieved
- 2D array of values
- Row count

### update_cells

Update cells in a spreadsheet range.

**Parameters:**

| Parameter | Type | Required | Description |
|-----------|------|----------|-------------|
| `spreadsheetId` | string | Yes | Spreadsheet ID |
| `range` | string | Yes | A1 notation range |
| `values` | string[][] | Yes | 2D array of values |
| `credentials` | string | Yes | OAuth credentials |

**Example:**

```json
{
  "spreadsheetId": "1abc123def456",
  "range": "Sheet1!A1:C3",
  "values": [
    ["Name", "Revenue", "Growth"],
    ["Product A", "50000", "15%"],
    ["Product B", "75000", "22%"]
  ],
  "credentials": "{...}"
}
```

**Response includes:**
- Updated range
- Rows updated
- Columns updated
- Total cells updated

### add_formula

Add a formula to a cell or range.

**Parameters:**

| Parameter | Type | Required | Description |
|-----------|------|----------|-------------|
| `spreadsheetId` | string | Yes | Spreadsheet ID |
| `range` | string | Yes | Target cell (A1 notation) |
| `formula` | string | Yes | Formula (include `=`) |
| `credentials` | string | Yes | OAuth credentials |

**Example:**

```json
{
  "spreadsheetId": "1abc123def456",
  "range": "Sheet1!D1",
  "formula": "=SUM(B2:B100)",
  "credentials": "{...}"
}
```

**Common formulas:**
- `=SUM(A1:A10)` - Sum of range
- `=AVERAGE(B1:B10)` - Average
- `=COUNT(C1:C10)` - Count numbers
- `=IF(A1>100,"High","Low")` - Conditional
- `=VLOOKUP(A1,B:C,2,FALSE)` - Lookup
- `=CONCATENATE(A1," ",B1)` - Join text

### create_chart

Create a chart in the spreadsheet.

**Parameters:**

| Parameter | Type | Required | Description |
|-----------|------|----------|-------------|
| `spreadsheetId` | string | Yes | Spreadsheet ID |
| `sheetId` | number | Yes | Sheet ID (usually 0 for first) |
| `chartType` | string | Yes | Chart type |
| `sourceRange` | string | Yes | Data range for chart |
| `credentials` | string | Yes | OAuth credentials |

**Chart types:**
- `COLUMN` - Vertical bar chart
- `BAR` - Horizontal bar chart
- `LINE` - Line chart
- `PIE` - Pie chart
- `AREA` - Area chart
- `SCATTER` - Scatter plot

**Example:**

```json
{
  "spreadsheetId": "1abc123def456",
  "sheetId": 0,
  "chartType": "COLUMN",
  "sourceRange": "Sheet1!A1:B10",
  "credentials": "{...}"
}
```

**Response includes:**
- Chart type created
- Success confirmation
- Spreadsheet ID

### export_csv

Export a sheet as CSV format.

**Parameters:**

| Parameter | Type | Required | Description |
|-----------|------|----------|-------------|
| `spreadsheetId` | string | Yes | Spreadsheet ID |
| `sheetId` | number | Yes | Sheet ID to export |
| `credentials` | string | Yes | OAuth credentials |

**Example:**

```json
{
  "spreadsheetId": "1abc123def456",
  "sheetId": 0,
  "credentials": "{...}"
}
```

**Response includes:**
- CSV content as string
- Row count
- Sheet title

## Finding IDs

### Spreadsheet ID

From URL: `https://docs.google.com/spreadsheets/d/1abc123def456/edit`

The ID is: `1abc123def456`

### Sheet ID

- First sheet is usually `0`
- Use `get_data` on the sheet to verify
- Or check sheet properties in Sheets UI

## Workflow Examples

### Creating a Report

1. **Create spreadsheet:**
   ```
   create_sheet with title: "Monthly Report - January 2024", credentials: "..."
   ```

2. **Add headers:**
   ```
   update_cells with range: "A1:D1", values: [["Date", "Revenue", "Expenses", "Profit"]], ...
   ```

3. **Add data:**
   ```
   update_cells with range: "A2:C10", values: [...data...], ...
   ```

4. **Add formulas:**
   ```
   add_formula with range: "D2", formula: "=B2-C2", ...
   ```

5. **Create chart:**
   ```
   create_chart with chartType: "LINE", sourceRange: "A1:B10", ...
   ```

### Data Analysis

1. **Read existing data:**
   ```
   get_data with range: "Sheet1!A:D", ...
   ```

2. **Add summary formulas:**
   ```
   add_formula with range: "E1", formula: "=SUM(B:B)", ...
   add_formula with range: "E2", formula: "=AVERAGE(B:B)", ...
   ```

3. **Export for analysis:**
   ```
   export_csv with sheetId: 0, ...
   ```

### Budget Tracking

1. **Create budget sheet:**
   ```
   create_sheet with title: "2024 Budget", ...
   ```

2. **Set up categories:**
   ```
   update_cells with range: "A1:B5", values: [["Category", "Amount"], ["Rent", "2000"], ...], ...
   ```

3. **Add total formula:**
   ```
   add_formula with range: "B10", formula: "=SUM(B2:B9)", ...
   ```

4. **Create pie chart:**
   ```
   create_chart with chartType: "PIE", sourceRange: "A2:B9", ...
   ```

## API Quotas

Google Sheets API has usage limits:

| Quota | Limit |
|-------|-------|
| Read requests | 300/minute per project |
| Write requests | 300/minute per project |
| Per-user limit | 60/minute |

If you hit limits:
- Wait 60 seconds
- Batch multiple updates into single requests
- Use larger ranges instead of cell-by-cell

## Requirements

- Node.js 18 or higher
- Google Cloud account
- Sheets API enabled
- OAuth 2.0 credentials with valid tokens

## Troubleshooting

### "Invalid credentials"

1. Verify OAuth credentials are correctly formatted
2. Check access_token hasn't expired
3. Ensure refresh_token is included for auto-refresh

### "Spreadsheet not found"

1. Verify spreadsheet ID is correct
2. Ensure credentials have access to the spreadsheet
3. Check if spreadsheet is shared with the OAuth client

### "Range not found"

1. Verify sheet name is correct (case-sensitive)
2. Check range format uses A1 notation
3. Ensure range exists in the spreadsheet

### Token refresh issues

1. Include refresh_token in credentials
2. Verify client_id and client_secret are correct
3. Check OAuth consent screen is properly configured

## Security Notes

- Never commit credentials to version control
- Access tokens expire (usually 1 hour)
- Refresh tokens should be stored securely
- Use minimum required OAuth scopes
- Rotate credentials periodically

## License

MIT License - see [LICENSE](LICENSE) file for details.

## Contributing

Contributions are welcome! Please feel free to submit a Pull Request.

## Author

consigcody94
