'use strict';

const ExcelJS = require('exceljs');

/**
 * Convert an XLSX file to a JSON object.
 *
 * Each worksheet in the workbook becomes a key in the returned object whose
 * value is an array of row objects.  The first row of every sheet is treated
 * as the header row – its cell values become the property names for all
 * subsequent rows.
 *
 * @param {string} filePath  Absolute or relative path to the .xlsx / .xls file.
 * @returns {Promise<Object>} A plain object keyed by sheet name.
 *
 * @example
 * const result = await xlsxToJson('./data.xlsx');
 * // {
 * //   Sheet1: [ { Name: 'Alice', Age: 30 }, { Name: 'Bob', Age: 25 } ],
 * //   Sheet2: [ { City: 'Paris', Country: 'France' } ]
 * // }
 */
async function xlsxToJson(filePath) {
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.readFile(filePath);

  const result = {};

  workbook.eachSheet((worksheet) => {
    const rows = [];
    let headers = null;

    worksheet.eachRow((row, rowNumber) => {
      const values = row.values.slice(1); // ExcelJS uses 1-based index; index 0 is always null

      if (rowNumber === 1) {
        headers = values.map((v) => (v == null ? '' : String(v)));
      } else {
        if (headers === null) return;
        const obj = {};
        headers.forEach((header, index) => {
          const cellValue = values[index];
          obj[header] = cellValue == null ? null : cellValue;
        });
        rows.push(obj);
      }
    });

    result[worksheet.name] = rows;
  });

  return result;
}

module.exports = xlsxToJson;
