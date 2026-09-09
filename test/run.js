// Small test runner — no framework, so `npm test` works from a fresh clone
// with only jsdom installed.
const path = require('path');

const tests = [];
let currentFile = '';

function test(name, fn) {
  tests.push({ name, fn, file: currentFile });
}

function assert(condition, message) {
  if (!condition) throw new Error(message || 'assertion failed');
}

function equal(actual, expected, message) {
  const a = JSON.stringify(actual);
  const e = JSON.stringify(expected);
  if (a !== e) {
    throw new Error(`${message ? message + ': ' : ''}expected ${e}, got ${a}`);
  }
}

global.test = test;
global.assert = assert;
global.equal = equal;

const FILES = ['matcher.test.js', 'xlsx-sheets.test.js', 'pipeline.test.js'];

FILES.forEach((file) => {
  currentFile = file;
  require(path.join(__dirname, file));
});

(async () => {
  let passed = 0;
  const failures = [];

  for (const t of tests) {
    try {
      await t.fn();
      passed++;
      process.stdout.write('.');
    } catch (err) {
      failures.push({ ...t, err });
      process.stdout.write('F');
    }
  }

  process.stdout.write('\n\n');
  failures.forEach((f) => {
    console.error(`FAIL ${f.file} — ${f.name}`);
    console.error(`     ${f.err.message}`);
    if (process.env.VERBOSE) console.error(f.err.stack);
    console.error('');
  });

  console.log(`${passed} passed, ${failures.length} failed`);
  process.exit(failures.length ? 1 : 0);
})();
