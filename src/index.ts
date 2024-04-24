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
const RANGE = 'A:AQ';
const CODE_COLUMN = 'AQ';

const MONTH_MAP = new Map([
  [1, 'January 2025'],
  [2, 'February 2025'],
  [3, 'March 2025'],
  [4, 'April 2024'],
  [5, 'May 2024'],
  [6, 'June 2024'],
  [7, 'July 2024'],
  [8, 'August 2024'],
  [9, 'September 2024'],
  [10, 'October 2024'],
  [11, 'November 2024'],
  [12, 'December 2024'],
]);

const LETTERS = 'ABCDEFGHIJKLMNOPQRSTUVWXYZ';
const DAY_OF_YEAR_LABEL = 'Day of Year:';

const codeRowMap = new Map<string, number>();

const getDayOfYear = () => {
  const now = new Date();
  const nowInMs = Date.UTC(now.getFullYear(), now.getMonth(), now.getDate());
  const startOfYear = Date.UTC(2024, 3, 0); // April 1st, 2024, start of budget year
  const msSinceStartOfYear = nowInMs - startOfYear;
  const daysSinceStartOfYear = msSinceStartOfYear / 24 / 60 / 60 / 1000;
  return daysSinceStartOfYear;
}

const getColumnFromIndex = (index: number) => {
  if (index >= LETTERS.length * LETTERS.length + LETTERS.length) {
    throw new Error(`Column index conversion exceeds maximum length: ${index}`);
  }

  let column = '';

  const magnitude = Math.floor(index / LETTERS.length);

  if (magnitude) {
    column += LETTERS[magnitude - 1];
  }

  column += LETTERS[index % LETTERS.length];

  return column;
};

const getRangeFromIndices = ([rowIndex, columnIndex]: number[]) => {
  const column = getColumnFromIndex(columnIndex);
  const row = (rowIndex + 1).toString();

  return column + row;
};

const findIndices = (searchString: string, values: any[][]) => {
  for (let rowIndex = 0; rowIndex < values.length; rowIndex++) {
    const columns = values[rowIndex];

    for (let columnIndex = 0; columnIndex < columns.length; columnIndex++) {
      const value = columns[columnIndex];

      if (value === searchString) {
        return [rowIndex, columnIndex];
      }
    }
  }

  return null;
};

const findRange = (searchString: string, values: any[][]) => {
  if (!searchString) {
    return null;
  }

  const indices = findIndices(searchString, values);

  if (!indices) {
    return null;
  }

  return getRangeFromIndices(indices);
};

const getIndexFromColumn = (column: string) => {
  const placeValues = column.split('').reverse();

  if (placeValues.length > 2) {
    throw new Error(
      `Index to column conversion exceeds maximum input: ${column}`
    );
  }

  let index = 0;

  index += LETTERS.indexOf(placeValues[0]);

  if (placeValues[1]) {
    index += (LETTERS.indexOf(placeValues[1]) + 1) * LETTERS.length;
  }

  return index;
};

const getNextColumn = (startingRange: string | null, offset: number) => {
  if (!offset || !startingRange) {
    return startingRange;
  }

  const column = startingRange.replace(/\d+/, '');
  const row = startingRange.replace(/\D+/, '');

  const columnIndex = getIndexFromColumn(column);
  const offsetIndex = columnIndex + offset;

  return getColumnFromIndex(offsetIndex) + row.toString();
};

const CODE_COLUMN_INDEX = getIndexFromColumn(CODE_COLUMN);

const updateSpreadsheet = async (entries: BudgetEntry[], month: number) => {

  logger.info(`Updating spreadsheet with data for month #${month}`);
  try {
    const auth = new google.auth.GoogleAuth({
      keyFilename: './keyFile.json',
      scopes: ['https://www.googleapis.com/auth/spreadsheets'],
    });

    const sheets = google.sheets({ version: 'v4', auth });
    const { data } = await sheets.spreadsheets.values.get({
      spreadsheetId: GOOGLE_SPREADSHEET_ID,
      range: RANGE,
    });

    const { values } = data;

    if (!values) {
      throw new Error(`No values found for range: ${RANGE}`);
    }

    const monthString = MONTH_MAP.get(month);

    if (!monthString) {
      throw new Error(`No month string mapped for value: ${month}`);
    }

    const monthRange = getNextColumn(findRange(monthString, values), 1);

    if (!monthRange) {
      throw new Error(`Could not locate column for month: ${month}`);
    }

    const monthColumn = monthRange.replace(/\d/g, '');

    values.forEach((columns, rowIndex) => {
      if (
        columns.length === CODE_COLUMN_INDEX + 1 &&
        columns[columns.length - 1]
      ) {
        codeRowMap.set(columns[columns.length - 1], rowIndex + 1);
      }
    });

    const dataToInsert: { range: string; values: any[][] }[] = [];

    entries.forEach((entry) => {
      const row = codeRowMap.get(entry.code);

      if (!row) {
        return;
      }

      const range = monthColumn + row;

      logger.info(
        `Will insert ${entry.spent} into ${range} for ${entry.category} (${entry.code})`
      );

      dataToInsert.push({
        range,
        values: [[entry.spent]],
      });
    });

    const dayOfYear = getDayOfYear();
    const dayOfYearRange = getNextColumn(findRange(DAY_OF_YEAR_LABEL, values), 1);

    if (!dayOfYearRange) {
      throw new Error(`Day of year could not be found for range: ${dayOfYearRange}`);
    }

    logger.info(
      `Will insert day of year ${dayOfYear} into ${dayOfYearRange}`
    );

    dataToInsert.push({
      range: dayOfYearRange,
      values: [[dayOfYear]]
    })

    await sheets.spreadsheets.values.batchUpdate({
      spreadsheetId: GOOGLE_SPREADSHEET_ID,
      requestBody: {
        valueInputOption: 'USER_ENTERED',
        data: dataToInsert,
      },
    });

    logger.info('Done!');
  } catch (error) {
    logger.error(error);
  }
};

const html = fs.readFileSync(
  path.resolve(__dirname, '../resources/export.html')
);
const $ = cheerio.load(html);

const totalRows = $('.total-label-cell').parent();

const entries: BudgetEntry[] = [];
const month = parseInt(
  $('h3')
    .text()
    .replace(/^.*From (\d\d).*$/, '$1'),
  10
);

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

updateSpreadsheet(entries, month);
