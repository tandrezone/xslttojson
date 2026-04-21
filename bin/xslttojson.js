#!/usr/bin/env node
'use strict';

const path = require('path');
const xlsxToJson = require('../index');

const args = process.argv.slice(2);

if (args.length === 0 || args.includes('--help') || args.includes('-h')) {
  console.log(`
Usage: xslttojson <file> [options]

Convert an XLSX file (with one or more sheets) to JSON.

Arguments:
  file        Path to the .xlsx file to convert

Options:
  -o, --output  Path to write the JSON output (defaults to stdout)
  -h, --help    Show this help message

Examples:
  xslttojson data.xlsx
  xslttojson data.xlsx -o output.json
`);
  process.exit(args.length === 0 ? 1 : 0);
}

const filePath = path.resolve(args[0]);

let outputPath = null;
const outputIndex = args.findIndex((a) => a === '-o' || a === '--output');
if (outputIndex !== -1 && args[outputIndex + 1]) {
  outputPath = path.resolve(args[outputIndex + 1]);
}

xlsxToJson(filePath)
  .then((json) => {
    const jsonString = JSON.stringify(json, null, 2);
    if (outputPath) {
      const fs = require('fs');
      fs.writeFileSync(outputPath, jsonString, 'utf8');
      console.log(`JSON written to ${outputPath}`);
    } else {
      process.stdout.write(jsonString + '\n');
    }
  })
  .catch((err) => {
    console.error(`Error: ${err.message}`);
    process.exit(1);
  });
