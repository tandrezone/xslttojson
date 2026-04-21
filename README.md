# xslttojson

Convert an XLSX file (with one or more sheets) to JSON.  Every worksheet
becomes a key in the output object whose value is an array of row objects.
The first row of each sheet is used as the header – its cell values become
the property names for all subsequent rows.

## Installation

```bash
npm install xslttojson
```

Or install globally to use the CLI:

```bash
npm install -g xslttojson
```

## Library usage

```js
const xlsxToJson = require('xslttojson');

const result = await xlsxToJson('./data.xlsx');
console.log(JSON.stringify(result, null, 2));
```

### Output format

Given a workbook with two sheets:

| **Employees** | Name    | Department  | Salary |
|---------------|---------|-------------|--------|
|               | Alice   | Engineering | 90000  |
|               | Bob     | Marketing   | 75000  |

| **Offices** | City   | Country | Headcount |
|-------------|--------|---------|-----------|
|             | Paris  | France  | 42        |
|             | London | UK      | 65        |

`xlsxToJson` returns:

```json
{
  "Employees": [
    { "Name": "Alice", "Department": "Engineering", "Salary": 90000 },
    { "Name": "Bob",   "Department": "Marketing",   "Salary": 75000 }
  ],
  "Offices": [
    { "City": "Paris",  "Country": "France", "Headcount": 42 },
    { "City": "London", "Country": "UK",     "Headcount": 65 }
  ]
}
```

## CLI usage

```
Usage: xslttojson <file> [options]

Arguments:
  file        Path to the .xlsx file to convert

Options:
  -o, --output  Path to write the JSON output (defaults to stdout)
  -h, --help    Show this help message
```

### Examples

Print JSON to stdout:

```bash
xslttojson data.xlsx
```

Write JSON to a file:

```bash
xslttojson data.xlsx -o output.json
```

## Running tests

```bash
npm test
```
