import { google, sheets_v4 } from 'googleapis';
import { has_duplicates, to_console } from './helpers';
import { AxiosError } from 'axios';
type QueueItem = {
  operation: () => Promise<any>;
  resolve: (value: any) => void;
  reject: (reason: any) => void;
};

export class GoogleSheetsClient {
  private static instance: GoogleSheetsClient;
  private sheets: sheets_v4.Sheets | null = null;
  private last_init_time = 0;
  private readonly REFRESH_INTERVAL = 50 * 60 * 1000; // 50 minutes
  private queue: QueueItem[] = [];
  private is_processing = false;

  private constructor(private keyFilePath: string) {}

  public static get_instance(keyFilePath: string): GoogleSheetsClient {
    if (!GoogleSheetsClient.instance) {
      GoogleSheetsClient.instance = new GoogleSheetsClient(keyFilePath);
    }
    return GoogleSheetsClient.instance;
  }

  private async init_client(): Promise<sheets_v4.Sheets> {
    try {
      const auth = new google.auth.GoogleAuth({
        keyFilename: this.keyFilePath,
        scopes: ['https://www.googleapis.com/auth/spreadsheets'],
      });

      const authClient = (await auth.getClient()) as any;
      return google.sheets({ version: 'v4', auth: authClient });
    } catch (error) {
      console.error('Error initializing Google Sheets client:', error);
      throw error;
    }
  }

  private async get_client(): Promise<sheets_v4.Sheets> {
    const now = Date.now();
    if (!this.sheets || now - this.last_init_time > this.REFRESH_INTERVAL) {
      this.sheets = await this.init_client();
      this.last_init_time = now;
    }
    return this.sheets;
  }

  private async process_queue() {
    if (this.is_processing) return;
    this.is_processing = true;

    while (this.queue.length > 0) {
      const item = this.queue.shift()!;
      try {
        const result = await item.operation();
        item.resolve(result);
      } catch (error) {
        item.reject(error);
      }
      // Add a small delay between requests to avoid hitting rate limits
      await new Promise((resolve) => setTimeout(resolve, 100));
    }

    this.is_processing = false;
  }

  private enqueue<T>(operation: () => Promise<T>): Promise<T> {
    return new Promise((resolve, reject) => {
      this.queue.push({ operation, resolve, reject });
      this.process_queue();
    });
  }

  public read_sheet(spreadsheetId: string, range: string): Promise<any[][]> {
    return this.enqueue(async () => {
      try {
        const sheets = await this.get_client();
        const response = await sheets.spreadsheets.values.get({
          spreadsheetId,
          range,
        });
        return response.data.values || [];
      } catch (error) {
        console.error('Error reading sheet:', error);
        throw error;
      }
    });
  }

  public write_to_sheet(spreadsheetId: string, range: string, values: any[][]): Promise<void> {
    return this.enqueue(async () => {
      try {
        const sheets = await this.get_client();
        await sheets.spreadsheets.values.update({
          spreadsheetId,
          range,
          valueInputOption: 'USER_ENTERED',
          requestBody: { values },
        });
      } catch (error) {
        console.error('Error writing to sheet:', error);
        throw error;
      }
    });
  }

  public read_entire_sheet(spreadsheetId: string, sheetName: string): Promise<any[][]> {
    return this.enqueue(async () => {
      try {
        const sheets = await this.get_client();

        const sheetProperties = await sheets.spreadsheets.get({
          spreadsheetId,
          ranges: [sheetName],
          fields: 'sheets.properties',
        });

        const sheetProps = sheetProperties.data.sheets?.[0].properties;
        if (!sheetProps) {
          throw new Error(`Sheet "${sheetName}" not found`);
        }

        const rowCount = sheetProps.gridProperties?.rowCount || 0;
        const columnCount = sheetProps.gridProperties?.columnCount || 0;

        const range = `${sheetName}!A1:${this.column_to_letter(columnCount)}${rowCount}`;
        const response = await sheets.spreadsheets.values.get({
          spreadsheetId,
          range,
        });

        return response.data.values || [];
      } catch (error) {
        console.error('Error reading entire sheet:', error);
        throw error;
      }
    });
  }

  public append_row(spreadsheetId: string, sheetName: string, values: any[]): Promise<void> {
    return this.enqueue(async () => {
      try {
        const sheets = await this.get_client();

        // Get the sheet's properties to determine its dimensions
        const sheetProperties = await sheets.spreadsheets.get({
          spreadsheetId,
          ranges: [sheetName],
          fields: 'sheets.properties',
        });

        const sheetProps = sheetProperties.data.sheets?.[0].properties;
        if (!sheetProps) {
          throw new Error(`Sheet "${sheetName}" not found`);
        }

        const rowCount = sheetProps.gridProperties?.rowCount || 0;
        const columnCount = sheetProps.gridProperties?.columnCount || 0;

        // Read all the data
        const response = await sheets.spreadsheets.values.get({
          spreadsheetId,
          range: `${sheetName}!A1:${this.column_to_letter(columnCount)}${rowCount}`,
        });

        const sheetData = response.data.values || [];

        // Find the last row with any data
        let lastRowIndex = 0;
        for (let i = sheetData.length - 1; i >= 0; i--) {
          if (sheetData[i] && sheetData[i].some((cell) => cell !== null && cell !== '')) {
            lastRowIndex = i + 1;
            break;
          }
        }

        // Append the new row
        await sheets.spreadsheets.values.append({
          spreadsheetId,
          range: `${sheetName}!A${lastRowIndex + 1}`,
          valueInputOption: 'USER_ENTERED',
          insertDataOption: 'INSERT_ROWS',
          requestBody: { values: [values] },
        });

        console.log(`Appended row at position ${lastRowIndex + 1}`);
      } catch (error) {
        console.error('Error appending row:', error);
        throw error;
      }
    });
  }

  public fill_cell(
    spreadsheetId: string,
    sheetName: string,
    column: string,
    row: number,
    value: any,
  ): Promise<void> {
    return this.enqueue(async () => {
      try {
        const cellAddress = `${column}${row}`;
        const sheets = await this.get_client();
        await sheets.spreadsheets.values.update({
          spreadsheetId,
          range: `${sheetName}!${cellAddress}`,
          valueInputOption: 'USER_ENTERED',
          requestBody: { values: [[value]] },
        });
        to_console(`Cell ${cellAddress} in sheet ${sheetName} updated successfully`);
      } catch (error) {
        console.error('Error updating cell:', error);
        throw error;
      }
    });
  }

  public get_row_by_column_value(
    spreadsheetId: string,
    sheetName: string,
    column: string,
    value: any,
  ): Promise<number | null> {
    return this.enqueue(async () => {
      try {
        const sheets = await this.get_client();
        const col_val = `${column.toUpperCase()}:${column.toUpperCase()}`;
        const response = await sheets.spreadsheets.values.get({
          spreadsheetId,
          range: `${sheetName}!${col_val}`,
        });
        const rows = response.data.values || [];

        const values = rows.map((row) => row[0]);

        if (has_duplicates(values)) {
          to_console('❗️ Column has duplicates');
        }

        const column_index = values.indexOf(value);

        if (column_index === -1) {
          return null;
        }

        return column_index + 1; 
      } catch (error) {
        console.error('Error getting row by column value:', error);
        throw error;
      }
    });
  }

  private column_to_letter(column: number): string {
    let temp,
      letter = '';
    while (column > 0) {
      temp = (column - 1) % 26;
      letter = String.fromCharCode(temp + 65) + letter;
      column = (column - temp - 1) / 26;
    }
    return letter;
  }

  public async check_read_write_permissions(spreadsheetId: string): Promise<boolean> {
    try {
      const sheets = await this.get_client();
      
      // Attempt to get minimal spreadsheet information
      await sheets.spreadsheets.get({
        spreadsheetId: spreadsheetId,
        fields: 'spreadsheetId'
      });

      // If the above doesn't throw an error, we have at least read access
      // Now check for write access by attempting to add a developer metadata
      await sheets.spreadsheets.batchUpdate({
        spreadsheetId: spreadsheetId,
        requestBody: {
          requests: [{
            createDeveloperMetadata: {
              developerMetadata: {
                metadataId: 1,
                metadataKey: 'test_write_access',
                metadataValue: 'test',
                location: { spreadsheet: true },
                visibility: 'DOCUMENT'
              }
            }
          }]
        }
      });

      // If both operations succeed, we have read-write access
      return true;
    } catch (error) {
      console.error('Error checking read-write access');
      return false;
    } finally {
      // Clean up: remove the test metadata
      try {
        const sheets = await this.get_client();
        await sheets.spreadsheets.batchUpdate({
          spreadsheetId: spreadsheetId,
          requestBody: {
            requests: [{
              deleteDeveloperMetadata: {
                dataFilter: {
                  developerMetadataLookup: {
                    metadataKey: 'test_write_access'
                  }
                }
              }
            }]
          }
        });
      } catch (cleanupError) {
        // Ignore cleanup errors
      }
    }
  }
}
