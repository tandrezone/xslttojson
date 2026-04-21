'use strict';

const assert = require('assert');
const path = require('path');
const xlsxToJson = require('../index');

const FIXTURES = path.join(__dirname, 'fixtures');

let passed = 0;
let failed = 0;

async function test(name, fn) {
  try {
    await fn();
    console.log(`  ✓ ${name}`);
    passed++;
  } catch (err) {
    console.error(`  ✗ ${name}`);
    console.error(`    ${err.message}`);
    failed++;
  }
}

(async () => {
  console.log('xlsxToJson');

  await test('returns an object with a key per sheet', async () => {
    const result = await xlsxToJson(path.join(FIXTURES, 'multi-sheet.xlsx'));
    assert.ok(typeof result === 'object' && result !== null);
    assert.deepStrictEqual(Object.keys(result), ['Employees', 'Offices']);
  });

  await test('each sheet value is an array', async () => {
    const result = await xlsxToJson(path.join(FIXTURES, 'multi-sheet.xlsx'));
    assert.ok(Array.isArray(result.Employees));
    assert.ok(Array.isArray(result.Offices));
  });

  await test('rows use the first row as headers', async () => {
    const result = await xlsxToJson(path.join(FIXTURES, 'multi-sheet.xlsx'));
    const [first] = result.Employees;
    assert.ok(Object.prototype.hasOwnProperty.call(first, 'Name'));
    assert.ok(Object.prototype.hasOwnProperty.call(first, 'Department'));
    assert.ok(Object.prototype.hasOwnProperty.call(first, 'Salary'));
  });

  await test('row data is correct for first sheet', async () => {
    const result = await xlsxToJson(path.join(FIXTURES, 'multi-sheet.xlsx'));
    assert.strictEqual(result.Employees.length, 3);
    assert.strictEqual(result.Employees[0].Name, 'Alice');
    assert.strictEqual(result.Employees[0].Department, 'Engineering');
    assert.strictEqual(result.Employees[0].Salary, 90000);
  });

  await test('row data is correct for second sheet', async () => {
    const result = await xlsxToJson(path.join(FIXTURES, 'multi-sheet.xlsx'));
    assert.strictEqual(result.Offices.length, 2);
    assert.strictEqual(result.Offices[0].City, 'Paris');
    assert.strictEqual(result.Offices[1].City, 'London');
  });

  await test('works with a single-sheet workbook', async () => {
    const result = await xlsxToJson(path.join(FIXTURES, 'single-sheet.xlsx'));
    assert.deepStrictEqual(Object.keys(result), ['Products']);
    assert.strictEqual(result.Products.length, 2);
    assert.strictEqual(result.Products[0].Name, 'Widget');
    assert.strictEqual(result.Products[1].Price, 19.99);
  });

  await test('rejects with an error for a missing file', async () => {
    await assert.rejects(
      () => xlsxToJson(path.join(FIXTURES, 'does-not-exist.xlsx')),
      /file not found|ENOENT/i
    );
  });

  console.log(`\n${passed} passing, ${failed} failing`);
  process.exit(failed > 0 ? 1 : 0);
})();
