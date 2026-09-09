const { createPage } = require('./harness');

const { D365IA } = createPage();
const { matcher } = D365IA;
const { sourceFormatFor } = D365IA.queue;

// Real file names from the batch this was built for.
test('strips a classification code prefix', () => {
  equal(
    matcher.cleanFileName('02B.03SYS-Inventory adjustment journal names.xlsx'),
    'Inventory Adjustment Journal Names'
  );
});

test('strips a plain sequence prefix', () => {
  equal(matcher.cleanFileName('01_Vendors_V2.xlsx'), 'Vendors V2');
});

test('strips a trailing date', () => {
  equal(matcher.cleanFileName('Customers_20240115.xlsx'), 'Customers');
});

test('leaves an embedded version alone', () => {
  equal(matcher.cleanFileName('CustomersV3.xlsx'), 'CustomersV3');
});

test('never strips the name down to nothing', () => {
  equal(matcher.cleanFileName('01.xlsx'), '01');
});

// Files depend on each other, so the queue must run in numbered order —
// which is not the order the OS hands a multi-file drop over in.
test('sorts numerically, not lexically', () => {
  const names = ['10_Customers.xlsx', '2_Vendors.xlsx', '1_Sites.xlsx'];
  const sorted = names.slice().sort(matcher.compareNaturally);
  equal(sorted, ['1_Sites.xlsx', '2_Vendors.xlsx', '10_Customers.xlsx']);
});

test('sorts dotted classification codes in order', () => {
  const names = ['02B.03SYS-B.xlsx', '01.02-A.xlsx', '02B.01SYS-C.xlsx'];
  const sorted = names.slice().sort(matcher.compareNaturally);
  equal(sorted, ['01.02-A.xlsx', '02B.01SYS-C.xlsx', '02B.03SYS-B.xlsx']);
});

// The bug the user hit: a data package went in as Excel.
test('detects the source data format from the extension', () => {
  equal(sourceFormatFor('Vendors.xlsx'), 'Excel');
  equal(sourceFormatFor('Vendors.xlsm'), 'Excel');
  equal(sourceFormatFor('Vendors.csv'), 'CSV');
  equal(sourceFormatFor('DataPackage.zip'), 'Package');
});

test('detects format case-insensitively', () => {
  equal(sourceFormatFor('VENDORS.XLSX'), 'Excel');
  equal(sourceFormatFor('PACKAGE.ZIP'), 'Package');
});

test('matches an entity name exactly when it exists', () => {
  const { candidate, score } = matcher.bestMatch('Vendors V2', ['Vendors V2', 'Customers V3']);
  equal(candidate, 'Vendors V2');
  assert(score >= 1, `expected an exact match, scored ${score}`);
});

test('ignores case and punctuation when matching', () => {
  const { score } = matcher.bestMatch('vendors-v2', ['VendorsV2']);
  assert(score >= 1, `expected punctuation to be ignored, scored ${score}`);
});
