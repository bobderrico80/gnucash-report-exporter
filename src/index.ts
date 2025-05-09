import dotenv from 'dotenv';
import * as cheerio from 'cheerio';
import fs from 'fs';
import path from 'path';
import { google } from 'googleapis';

interface BudgetEntry {
  category: string;
  spent: number;
  code: string;
}

const logger = console;

dotenv.config();

const { GOOGLE_SPREADSHEET_ID } = process.env;
const SHEET_RANGE = 'A:J';

const MONTH_MAP = new Map([
  [1, 'January'],
  [2, 'February'],
  [3, 'March'],
  [4, 'April'],
  [5, 'May'],
  [6, 'June'],
  [7, 'July'],
  [8, 'August'],
  [9, 'September'],
  [10, 'October'],
  [11, 'November'],
  [12, 'December'],
]);

const codeRowMap = new Map<string, number>();

const CODE_COLUMN_INDEX = 0;
const ACTUAL_COLUMN = 'G';
const LAST_UPDATED_RANGE = 'K1';

const updateSpreadsheet = async (entries: BudgetEntry[], month: number) => {
  logger.info(`Updating spreadsheet with data for month #${month}`);
  try {
    const auth = new google.auth.GoogleAuth({
      keyFilename: './keyFile.json',
      scopes: ['https://www.googleapis.com/auth/spreadsheets'],
    });

    const monthString = MONTH_MAP.get(month);

    if (!monthString) {
      throw new Error(`No month string mapped for value: ${month}`);
    }

    const range = `${monthString}!${SHEET_RANGE}`;

    logger.info(`Reading data from range: ${range}`)

    const sheets = google.sheets({ version: 'v4', auth });
    const { data } = await sheets.spreadsheets.values.get({
      spreadsheetId: GOOGLE_SPREADSHEET_ID,
      range,
    });

    const { values } = data;

    if (!values) {
      throw new Error(`No values found for range: ${SHEET_RANGE}`);
    }

    // Create a mapping of line item codes to the rows they appear on in the sheet
    values.forEach((columns, rowIndex) => {
      const code = columns[CODE_COLUMN_INDEX];

      if (code) {
        codeRowMap.set(code, rowIndex + 1);
      }
    });

    const dataToInsert: { range: string; values: any[][] }[] = [];

    entries.forEach((entry) => {
      const row = codeRowMap.get(entry.code);

      if (!row) {
        return;
      }

      const range = `${monthString}!${ACTUAL_COLUMN}${row}`;

      logger.info(
        `Will insert ${entry.spent} into ${range} for ${entry.category} (${entry.code})`
      );

      dataToInsert.push({
        range,
        values: [[entry.spent]],
      });
    });

    // Add last updated timestamp
    const timestamp = new Date().toLocaleString();
    const timeStampRange = `${monthString}!${LAST_UPDATED_RANGE}`;
    logger.info(`Will insert timestamp ${timestamp} into ${timeStampRange}`);
    dataToInsert.push({
      range: timeStampRange,
      values: [[timestamp]]
    })

    if (!process.env.DRY_RUN) {
      await sheets.spreadsheets.values.batchUpdate({
        spreadsheetId: GOOGLE_SPREADSHEET_ID,
        requestBody: {
          valueInputOption: 'USER_ENTERED',
          data: dataToInsert,
        },
      });
    } else {
      logger.warn('Performing dry run!');
    }

    logger.info('Done!');
  } catch (error) {
    logger.error(error);
  }
};

(async () => {
  const html = fs.readFileSync(
      path.resolve(__dirname, '../resources/export.html')
  );
  const $ = cheerio.load(html);

  const totalRows = $('.total-label-cell').parent();

  const entries: BudgetEntry[] = [];
  let month = parseInt(
      $('h3')
          .text()
          .replace(/^.*From (\d\d).*$/, '$1'),
      10
  );

  // Compensate for the fact that April reporting will include the tail end of March
  if (month === 3) {
    month = 4
  }

  totalRows.each((_, row) => {
    const $row = $(row);
    const category = $row
        .find('.total-label-cell')
        .text()
        .replace('Total For ', '');

    const code = $row
        .prev()
        .find('td')
        .eq(5)
        .text()
        .trim()
        .replace(/^(\d*).*$/, '$1');
    const spent = Math.abs(parseFloat(
        $row.find('.total-number-cell').text().replace(/[$,]/g, '')
    ));

    if (category !== 'Grand Total') {
      entries.push({ category, spent, code });
    }
  });

  await updateSpreadsheet(entries, month);
  logger.info('Update complete!')
})()
