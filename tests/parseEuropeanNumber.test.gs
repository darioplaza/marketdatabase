const assert = require('assert');
const fs = require('fs');
const vm = require('vm');

// Load the Apps Script code
const code = fs.readFileSync(require('path').join(__dirname, '..', 'codigo.gs'), 'utf8');
const sandbox = {};
vm.runInNewContext(code, sandbox);
const parseEuropeanNumber = sandbox._parseEuropeanNumber;

// Valores con punto de miles y coma decimal
assert.strictEqual(parseEuropeanNumber('1.234,56'), 1234.56);

// Números negativos
assert.strictEqual(parseEuropeanNumber('-123,45'), -123.45);

// Entradas inválidas que deben devolver cadena vacía
assert.strictEqual(parseEuropeanNumber('abc'), '');
assert.strictEqual(parseEuropeanNumber(null), '');

console.log('All tests passed.');
