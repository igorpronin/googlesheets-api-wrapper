# Google Sheets API Wrapper

Google Sheets API wrapper for internal purposes.

## Install

```bash
npm i @igorpronin/googlesheets-api-wrapper
```

## Usage

```typescript
import { GoogleSheetsClient } from '@igorpronin/googlesheets-api-wrapper';

const client = GoogleSheetsClient.get_instance({ keyFilePath: './path/to/keyfile.json', isSilent: false });

client.append_row_v2({spreadsheetId, sheetName, data});
```

## API

- `get_instance(keyFilePath: string)`: Returns a singleton instance of the GoogleSheetsClient. Initializes with the path to the Google service account key file.

- `read_tab(spreadsheetId: string, range: string)`: Reads data from a specified range in a spreadsheet. Returns a 2D array of values.

- `read_tab_by_name(spreadsheetId: string, tabName: string, range: string)`: Reads data from a specified range in a spreadsheet. Returns a 2D array of values.

- `write_to_tab(spreadsheetId: string, range: string, values: any[][])`: Writes data to a specified range in a spreadsheet.

- `read_entire_tab(spreadsheetId: string, sheetName: string)`: Reads all data from a specified sheet in a spreadsheet. Returns a 2D array of values.

- [deprecated] `append_row(spreadsheetId: string, sheetName: string, values: any[])`: Appends a row of data to the first empty row after the last non-empty row in a specified sheet. Method is deprecated, use `append_row_v2` instead.

- `append_row_v2({spreadsheetId: string, sheetName: string, values: any[]})`: Appends a row of data to the first empty row after the last non-empty row in a specified sheet.

- `get_cell(spreadsheetId: string, sheetName: string, column: string, row: number, force?: boolean)`: Reads the value of a specific cell in a sheet. Returns the value of the cell.

- [deprecated] `fill_cell(spreadsheetId: string, sheetName: string, column: string, row: number, value: any)`: Fills a specific cell in a sheet with the provided value. Method is deprecated, use `fill_cell_v2` instead.

- `fill_cell_v2({spreadsheetId: string, sheetName: string, column: string, row: number, value: any})`: Fills a specific cell in a sheet with the provided value.

- `get_row_by_column_value(spreadsheetId: string, sheetName: string, column: string, value: any)`: Returns the row that contains the specified value in the specified column.

- `get_rows_map_of_values(spreadsheetId: string, sheetName: string, column: string)`: Returns a map of values in a specified column to their corresponding rows.

- `get_column_letter_map_of_row(spreadsheetId: string, sheetName: string, row: number)`: Returns a map of column letters to values in a specified row.

- `get_filename(spreadsheetId: string)`: Returns the filename of a specified spreadsheet.

- `change_filename(spreadsheetId: string, new_name: string)`: Changes the name of a specified sheet in a spreadsheet.

- `check_read_permissions(spreadsheetId: string)`: Checks if the client has read permissions for a specified spreadsheet. Returns a boolean.

- `check_read_write_permissions(spreadsheetId: string)`: Checks if the client has read and write permissions for a specified spreadsheet. Returns a boolean.

- `clear_tab(spreadsheetId: string, tabName: string, from_row: number = 1, force?: boolean)`: Clears a specified tab in a spreadsheet.

All methods return Promises and are queued to avoid rate limiting issues.
