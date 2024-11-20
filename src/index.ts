import { google, sheets_v4 } from 'googleapis';
import { has_duplicates, to_console, wait } from './helpers';

type QueueItem = {
  operation: () => Promise<any>;
  resolve: (value: any) => void;
  reject: (reason: any) => void;
};

type Params = {
  keyFilePath: string;
  isSilent?: boolean;
  isDebug?: boolean;
};

export class GoogleSheetsClient {
  private static instance: GoogleSheetsClient;
  private sheets: sheets_v4.Sheets | null = null;
  private last_init_time = 0;
  private readonly REFRESH_INTERVAL = 50 * 60 * 1000; // 50 minutes
  private queue: QueueItem[] = [];
  private is_processing = false;

  private keyFilePath: string;
  private is_silent: boolean = false;
  private is_debug: boolean = false;

  private constructor({ keyFilePath, isSilent, isDebug }: Params) {
    this.keyFilePath = keyFilePath;
    this.is_silent = isSilent || false;
    this.is_debug = isDebug || false;
  }

  private readonly MAX_RETRIES = 8;
  private readonly RETRY_DELAY = 60000;

  public static get_instance({ keyFilePath, isSilent }: Params): GoogleSheetsClient {
    if (!GoogleSheetsClient.instance) {
      GoogleSheetsClient.instance = new GoogleSheetsClient({ keyFilePath, isSilent });
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

  public read_tab(
    spreadsheetId: string,
    rangeWithTabName: string,
    force?: boolean,
  ): Promise<any[][]> {
    const operation = async () => {
      try {
        const sheets = await this.get_client();
        const response = await sheets.spreadsheets.values.get({
          spreadsheetId,
          range: rangeWithTabName,
        });
        return response.data.values || [];
      } catch (error) {
        to_console('Error reading tab', this.is_silent, true);
        console.error(error);
        throw error;
      }
    };

    if (force) {
      return operation();
    } else {
      return this.enqueue(operation);
    }
  }

  public read_tab_by_name(
    spreadsheetId: string,
    tabName: string,
    range: string,
    force?: boolean,
  ): Promise<any[][]> {
    const fullRange = `${tabName}!${range}`;
    return this.read_tab(spreadsheetId, fullRange, force);
  }

  public read_sheet(
    spreadsheetId: string,
    rangeWithTabName: string,
    force?: boolean,
  ): Promise<any[][]> {
    to_console(
      '❗️ Method "read_sheet" is deprecated due to incorrect naming and will be removed soon, use "read_tab" instead',
      false,
      true,
    );
    return this.read_tab(spreadsheetId, rangeWithTabName, force);
  }

  public async write_to_tab(
    spreadsheetId: string,
    tabName: string,
    range: string,
    values: any[][],
    force?: boolean,
  ): Promise<boolean> {
    if (force) {
      to_console(
        '❗️ Method "write_to_tab" does not support "force" parameter anymore, remove it for code purity',
        false,
        true,
      );
    }
    const destination_range = `${tabName}!${range}`;
    const operation = async () => {
      const sheets = await this.get_client();
      await sheets.spreadsheets.values.update({
        spreadsheetId,
        range: destination_range,
        valueInputOption: 'USER_ENTERED',
        requestBody: { values },
      });
      to_console(`Wrote to ${destination_range}`, this.is_silent);
      return true;
    };

    let attempt = 0;
    for (attempt = 0; attempt < this.MAX_RETRIES; attempt++) {
      if (await this.process_cycle(operation, attempt, 'Error writing to tab')) {
        break;
      }
    }
    return true;
  }

  public write_to_sheet(
    spreadsheetId: string,
    tabName: string,
    range: string,
    values: any[][],
    force?: boolean,
  ): Promise<boolean> {
    to_console(
      '❗️ Method "write_to_sheet" is deprecated due to incorrect naming and will be removed soon, use "write_to_tab" instead',
      false,
      true,
    );
    return this.write_to_tab(spreadsheetId, tabName, range, values, force);
  }

  public async read_entire_tab(
    spreadsheetId: string,
    tabName: string,
    force?: boolean, // Deprecated
  ): Promise<any[][]> {
    if (force) {
      to_console(
        '❗️ Method "read_entire_tab" does not support "force" parameter anymore, remove it for code purity',
        false,
        true,
      );
    }

    const operation = async () => {
      const sheets = await this.get_client();

      const sheetProperties = await sheets.spreadsheets.get({
        spreadsheetId,
        ranges: [tabName],
        fields: 'sheets.properties',
      });

      const sheetProps = sheetProperties.data.sheets?.[0].properties;
      if (!sheetProps) {
        throw new Error(`Tab "${tabName}" not found`);
      }

      const rowCount = sheetProps.gridProperties?.rowCount || 0;
      const columnCount = sheetProps.gridProperties?.columnCount || 0;

      const range = `${tabName}!A1:${this.column_to_letter(columnCount)}${rowCount}`;
      const response = await sheets.spreadsheets.values.get({
        spreadsheetId,
        range,
      });

      return response.data.values || [];
    };

    let attempt = 0;
    let result: any;
    for (attempt = 0; attempt < this.MAX_RETRIES; attempt++) {
      result = await this.process_cycle(operation, attempt, 'Error reading entire tab');
      if (result) {
        break;
      }
    }
    return result;
  }

  public read_entire_sheet(
    spreadsheetId: string,
    tabName: string,
    force?: boolean,
  ): Promise<any[][]> {
    to_console(
      '❗️ Method "read_entire_sheet" is deprecated due to incorrect naming and will be removed soon, use "read_entire_tab" instead',
      false,
      true,
    );
    return this.read_entire_tab(spreadsheetId, tabName, force);
  }

  public append_row(
    spreadsheetId: string,
    tabName: string,
    values: any[],
    force?: boolean,
  ): Promise<void> {
    to_console(
      '❗️ Method "append_row" is deprecated, use "append_row_v2" instead (it has a retry mechanism and does not uses "force" parameter)',
      false,
      true,
    );

    const operation = async () => {
      try {
        const sheets = await this.get_client();

        // Get the sheet's properties to determine its dimensions
        const sheetProperties = await sheets.spreadsheets.get({
          spreadsheetId,
          ranges: [tabName],
          fields: 'sheets.properties',
        });

        const sheetProps = sheetProperties.data.sheets?.[0].properties;
        if (!sheetProps) {
          throw new Error(`Tab "${tabName}" not found`);
        }

        const rowCount = sheetProps.gridProperties?.rowCount || 0;
        const columnCount = sheetProps.gridProperties?.columnCount || 0;

        // Read all the data
        const response = await sheets.spreadsheets.values.get({
          spreadsheetId,
          range: `${tabName}!A1:${this.column_to_letter(columnCount)}${rowCount}`,
        });

        const tabData = response.data.values || [];

        // Find the last row with any data
        let lastRowIndex = 0;
        for (let i = tabData.length - 1; i >= 0; i--) {
          if (tabData[i] && tabData[i].some((cell) => cell !== null && cell !== '')) {
            lastRowIndex = i + 1;
            break;
          }
        }

        // Append the new row
        await sheets.spreadsheets.values.append({
          spreadsheetId,
          range: `${tabName}!A${lastRowIndex + 1}`,
          valueInputOption: 'USER_ENTERED',
          insertDataOption: 'INSERT_ROWS',
          requestBody: { values: [values] },
        });

        to_console(`Appended row at position ${lastRowIndex + 1}`, this.is_silent);
      } catch (error) {
        to_console('Error appending row', this.is_silent, true);
        console.error(error);
        throw error;
      }
    };

    if (force) {
      return operation();
    } else {
      return this.enqueue(operation);
    }
  }

  public async append_row_v2({
    spreadsheetId,
    tabName,
    values,
  }: {
    spreadsheetId: string;
    tabName: string;
    values: any[];
  }): Promise<boolean> {
    const operation = async () => {
      const sheets = await this.get_client();

      // Get the sheet's properties to determine its dimensions
      const sheetProperties = await sheets.spreadsheets.get({
        spreadsheetId,
        ranges: [tabName],
        fields: 'sheets.properties',
      });

      const sheetProps = sheetProperties.data.sheets?.[0].properties;
      if (!sheetProps) {
        throw new Error(`Tab "${tabName}" not found`);
      }

      const rowCount = sheetProps.gridProperties?.rowCount || 0;
      const columnCount = sheetProps.gridProperties?.columnCount || 0;

      // Read all the data
      const response = await sheets.spreadsheets.values.get({
        spreadsheetId,
        range: `${tabName}!A1:${this.column_to_letter(columnCount)}${rowCount}`,
      });

      const tabData = response.data.values || [];

      // Find the last row with any data
      let lastRowIndex = 0;
      for (let i = tabData.length - 1; i >= 0; i--) {
        if (tabData[i] && tabData[i].some((cell) => cell !== null && cell !== '')) {
          lastRowIndex = i + 1;
          break;
        }
      }

      // Append the new row
      await sheets.spreadsheets.values.append({
        spreadsheetId,
        range: `${tabName}!A${lastRowIndex + 1}`,
        valueInputOption: 'USER_ENTERED',
        insertDataOption: 'INSERT_ROWS',
        requestBody: { values: [values] },
      });

      to_console(`Appended row at position ${lastRowIndex + 1}`, this.is_silent);
      return true;
    };

    let attempt = 0;
    for (attempt = 0; attempt < this.MAX_RETRIES; attempt++) {
      if (await this.process_cycle(operation, attempt, 'Error appending row')) {
        break;
      }
    }
    return true;
  }

  public get_cell(
    spreadsheetId: string,
    tabName: string,
    column: string,
    row: number,
    force?: boolean,
  ): Promise<any> {
    const operation = async (): Promise<any> => {
      try {
        const sheets = await this.get_client();
        const cellAddress = `${column}${row}`;
        const response = await sheets.spreadsheets.values.get({
          spreadsheetId,
          range: `${tabName}!${cellAddress}`,
        });
        return response.data.values?.[0]?.[0];
      } catch (error) {
        to_console('Error getting cell', this.is_silent, true);
        console.error(error);
        throw error;
      }
    };

    if (force) {
      return operation();
    } else {
      return this.enqueue(operation);
    }
  }

  public fill_cell(
    spreadsheetId: string,
    tabName: string,
    column: string,
    row: number,
    value: any,
    force?: boolean,
  ): Promise<void> {
    to_console(
      '❗️ Method "fill_cell" is deprecated, use "fill_cell_v2" instead (it has a retry mechanism and does not uses "force" parameter)',
      false,
      true,
    );

    const operation = async () => {
      try {
        const cellAddress = `${column}${row}`;
        const sheets = await this.get_client();
        await sheets.spreadsheets.values.update({
          spreadsheetId,
          range: `${tabName}!${cellAddress}`,
          valueInputOption: 'USER_ENTERED',
          requestBody: { values: [[value]] },
        });
        to_console(`Cell ${cellAddress} in tab ${tabName} updated successfully`, this.is_silent);
      } catch (error) {
        to_console('Error updating cell', this.is_silent, true);
        console.error(error);
        throw error;
      }
    };

    if (force) {
      return operation();
    } else {
      return this.enqueue(operation);
    }
  }

  public async fill_cell_v2({
    spreadsheetId,
    tabName,
    column,
    row,
    value,
  }: {
    spreadsheetId: string;
    tabName: string;
    column: string;
    row: number;
    value: any;
  }): Promise<boolean> {
    const operation = async () => {
      const cellAddress = `${column}${row}`;
      const sheets = await this.get_client();
      await sheets.spreadsheets.values.update({
        spreadsheetId,
        range: `${tabName}!${cellAddress}`,
        valueInputOption: 'USER_ENTERED',
        requestBody: { values: [[value]] },
      });
      to_console(`Cell ${cellAddress} in tab ${tabName} updated successfully`, this.is_silent);
      return true;
    };

    let attempt = 0;
    for (attempt = 0; attempt < this.MAX_RETRIES; attempt++) {
      if (await this.process_cycle(operation, attempt, 'Error updating cell')) {
        break;
      }
    }
    return true;
  }

  public async clear_tab(
    spreadsheetId: string,
    tabName: string,
    from_row: number = 1,
    force?: boolean, // Deprecated
  ): Promise<boolean> {
    if (force) {
      to_console(
        '❗️ Method "clear_tab" does not support "force" parameter anymore, remove it for code purity',
        false,
        true,
      );
    }
    const operation = async () => {
      const sheets = await this.get_client();
      await sheets.spreadsheets.values.clear({
        spreadsheetId,
        range: `${tabName}!A${from_row}:Z`,
      });
      to_console(`Cleared tab ${tabName} from row ${from_row}`, this.is_silent);
      return true;
    };

    let attempt = 0;
    for (attempt = 0; attempt < this.MAX_RETRIES; attempt++) {
      if (await this.process_cycle(operation, attempt, 'Error clearing tab')) {
        break;
      }
    }
    return true;
  }

  public get_row_by_column_value(
    spreadsheetId: string,
    tabName: string,
    column: string,
    value: any,
    force?: boolean,
  ): Promise<number | null> {
    const operation = async () => {
      try {
        const sheets = await this.get_client();
        const col_val = `${column.toUpperCase()}:${column.toUpperCase()}`;
        const response = await sheets.spreadsheets.values.get({
          spreadsheetId,
          range: `${tabName}!${col_val}`,
        });
        const rows = response.data.values || [];

        const values = rows.map((row) => row[0]);

        if (has_duplicates(values, this.is_silent)) {
          to_console('❗️ Column has duplicates', this.is_silent);
        }

        const column_index = values.indexOf(value);

        if (column_index === -1) {
          return null;
        }

        return column_index + 1;
      } catch (error) {
        to_console('Error getting row by column value', this.is_silent, true);
        console.error(error);
        throw error;
      }
    };

    if (force) {
      return operation();
    } else {
      return this.enqueue(operation);
    }
  }

  public get_rows_map_of_values(
    spreadsheetId: string,
    tabName: string,
    column: string,
    values: string[],
    force?: boolean,
  ): Promise<{ [key: string]: number | null }> {
    const operation = async () => {
      try {
        const sheets = await this.get_client();
        const col_val = `${column.toUpperCase()}:${column.toUpperCase()}`;
        const response = await sheets.spreadsheets.values.get({
          spreadsheetId,
          range: `${tabName}!${col_val}`,
        });

        const rows = response.data.values || [];
        const raw_values = rows.map((row) => row[0]);

        if (has_duplicates(values, this.is_silent)) {
          to_console('❗️ Column has duplicates', this.is_silent);
        }

        const rows_map: { [key: string]: number } = {};
        values.forEach((value, index) => {
          rows_map[value] = raw_values.indexOf(value) + 1;
        });

        return rows_map;
      } catch (error) {
        to_console('Error getting rows map of values', this.is_silent, true);
        console.error(error);
        throw error;
      }
    };

    if (force) {
      return operation();
    } else {
      return this.enqueue(operation);
    }
  }

  public get_column_letter_map_of_row(
    spreadsheetId: string,
    tabName: string,
    row: number,
    force?: boolean,
  ): Promise<{ [key: string]: string | null }> {
    const operation = async () => {
      try {
        const sheets = await this.get_client();
        const response = await sheets.spreadsheets.values.get({
          spreadsheetId,
          range: `${tabName}!A${row}:Z${row}`,
        });

        const columns = response.data.values?.[0] || [];

        const column_letter_map: { [key: string]: string | null } = {};
        columns.forEach((column, index) => {
          column_letter_map[column] = this.column_to_letter(index + 1);
        });

        return column_letter_map;
      } catch (error) {
        to_console('Error getting column letter map of row', this.is_silent, true);
        console.error(error);
        throw error;
      }
    };

    if (force) {
      return operation();
    } else {
      return this.enqueue(operation);
    }
  }

  public column_to_letter(column: number): string {
    let temp,
      letter = '';
    while (column > 0) {
      temp = (column - 1) % 26;
      letter = String.fromCharCode(temp + 65) + letter;
      column = (column - temp - 1) / 26;
    }
    return letter;
  }

  private async process_cycle(
    operation: any,
    attempt: number,
    error_message: string,
  ): Promise<boolean> {
    try {
      return await operation(); // True (for write methods) or result data (for read methods) is expected
    } catch (error) {
      let message = `${error_message}, attempt ${attempt + 1} of ${this.MAX_RETRIES}`;
      if (attempt < this.MAX_RETRIES - 1) {
        message += `, retrying in ${this.RETRY_DELAY / 1000} seconds...`;
      }
      to_console(message, this.is_silent, true);
      if (this.is_debug) {
        console.error(error);
      }
      if (attempt === this.MAX_RETRIES - 1) {
        to_console(error_message, this.is_silent, true);
        console.error(error);
        throw error;
      }
      await wait(this.RETRY_DELAY);
      return false;
    }
  }

  public async get_filename(spreadsheetId: string): Promise<string> {
    try {
      const sheets = await this.get_client();
      const response = await sheets.spreadsheets.get({
        spreadsheetId,
        fields: 'properties.title',
      });

      return response.data.properties?.title || '';
    } catch (error) {
      to_console('Error getting filename', this.is_silent, true);
      console.error(error);
      throw error;
    }
  }

  public async change_filename({
    spreadsheetId,
    new_name,
  }: {
    spreadsheetId: string;
    new_name: string;
  }): Promise<boolean> {
    const operation = async () => {
      const sheets = await this.get_client();
      await sheets.spreadsheets.batchUpdate({
        spreadsheetId,
        requestBody: {
          requests: [
            {
              updateSpreadsheetProperties: {
                properties: {
                  title: new_name,
                },
                fields: 'title',
              },
            },
          ],
        },
      });
      to_console(`Sheet name changed to "${new_name}"`, this.is_silent);
      return true;
    };

    let attempt = 0;
    for (attempt = 0; attempt < this.MAX_RETRIES; attempt++) {
      if (await this.process_cycle(operation, attempt, 'Error changing sheet name')) {
        break;
      }
    }
    return true;
  }

  public async check_read_permissions(spreadsheetId: string): Promise<boolean> {
    try {
      const sheets = await this.get_client();

      // Attempt to get minimal spreadsheet information
      await sheets.spreadsheets.get({
        spreadsheetId: spreadsheetId,
        fields: 'spreadsheetId',
      });

      // If the above doesn't throw an error, we have read access
      return true;
    } catch (error) {
      to_console('Error checking read access', this.is_silent, true);
      return false;
    }
  }

  public async check_read_write_permissions(spreadsheetId: string): Promise<boolean> {
    try {
      const sheets = await this.get_client();

      // Attempt to get minimal spreadsheet information
      await sheets.spreadsheets.get({
        spreadsheetId: spreadsheetId,
        fields: 'spreadsheetId',
      });

      // If the above doesn't throw an error, we have at least read access
      // Now check for write access by attempting to add a developer metadata
      await sheets.spreadsheets.batchUpdate({
        spreadsheetId: spreadsheetId,
        requestBody: {
          requests: [
            {
              createDeveloperMetadata: {
                developerMetadata: {
                  metadataId: 1,
                  metadataKey: 'test_write_access',
                  metadataValue: 'test',
                  location: { spreadsheet: true },
                  visibility: 'DOCUMENT',
                },
              },
            },
          ],
        },
      });

      // If both operations succeed, we have read-write access
      return true;
    } catch (error) {
      to_console('Error checking read-write access', this.is_silent, true);
      return false;
    } finally {
      // Clean up: remove the test metadata
      try {
        const sheets = await this.get_client();
        await sheets.spreadsheets.batchUpdate({
          spreadsheetId: spreadsheetId,
          requestBody: {
            requests: [
              {
                deleteDeveloperMetadata: {
                  dataFilter: {
                    developerMetadataLookup: {
                      metadataKey: 'test_write_access',
                    },
                  },
                },
              },
            ],
          },
        });
      } catch (cleanupError) {
        // Ignore cleanup errors
        to_console('Error cleaning up test metadata', this.is_silent, true);
      }
    }
  }
}
