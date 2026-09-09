const zlib = require('zlib');
const { createPage } = require('./harness');

const { window, D365IA } = createPage();

// Builds a real .xlsx-shaped zip (local header + central directory + EOCD)
// containing one entry, so the parser is exercised against actual bytes
// rather than a mock.
function makeZip(entries, { streaming = false } = {}) {
  const parts = [];
  const central = [];
  let offset = 0;

  entries.forEach(({ name, content }) => {
    const nameBytes = Buffer.from(name, 'utf8');
    const raw = Buffer.from(content, 'utf8');
    const deflated = zlib.deflateRawSync(raw);
    const crc = crc32(raw);

    const local = Buffer.alloc(30);
    local.writeUInt32LE(0x04034b50, 0);
    local.writeUInt16LE(20, 4);
    // Streaming zips set bit 3 and zero the sizes in the local header,
    // putting the real ones in a trailing data descriptor — the case that
    // made some workbooks read back truncated.
    local.writeUInt16LE(streaming ? 0x08 : 0, 6);
    local.writeUInt16LE(8, 8);
    local.writeUInt32LE(streaming ? 0 : crc, 14);
    local.writeUInt32LE(streaming ? 0 : deflated.length, 18);
    local.writeUInt32LE(streaming ? 0 : raw.length, 22);
    local.writeUInt16LE(nameBytes.length, 26);

    parts.push(local, nameBytes, deflated);

    const cd = Buffer.alloc(46);
    cd.writeUInt32LE(0x02014b50, 0);
    cd.writeUInt16LE(20, 4);
    cd.writeUInt16LE(20, 6);
    cd.writeUInt16LE(streaming ? 0x08 : 0, 8);
    cd.writeUInt16LE(8, 10);
    cd.writeUInt32LE(crc, 16);
    cd.writeUInt32LE(deflated.length, 20);
    cd.writeUInt32LE(raw.length, 24);
    cd.writeUInt16LE(nameBytes.length, 28);
    cd.writeUInt32LE(offset, 42);
    central.push(cd, nameBytes);

    offset += local.length + nameBytes.length + deflated.length;
  });

  const centralBuf = Buffer.concat(central);
  const eocd = Buffer.alloc(22);
  eocd.writeUInt32LE(0x06054b50, 0);
  eocd.writeUInt16LE(entries.length, 8);
  eocd.writeUInt16LE(entries.length, 10);
  eocd.writeUInt32LE(centralBuf.length, 12);
  eocd.writeUInt32LE(offset, 16);

  return Buffer.concat([Buffer.concat(parts), centralBuf, eocd]);
}

function crc32(buf) {
  let c;
  const table = [];
  for (let n = 0; n < 256; n++) {
    c = n;
    for (let k = 0; k < 8; k++) c = c & 1 ? 0xedb88320 ^ (c >>> 1) : c >>> 1;
    table[n] = c >>> 0;
  }
  let crc = 0 ^ -1;
  for (let i = 0; i < buf.length; i++) crc = (crc >>> 8) ^ table[(crc ^ buf[i]) & 0xff];
  return (crc ^ -1) >>> 0;
}

// A workbook.xml as Excel really writes it, including the definedName noise
// that a naive name="..." scan would pick up as extra sheets.
const WORKBOOK_XML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <fileVersion appName="xl" lastEdited="7" lowestEdited="7"/>
  <workbookPr defaultThemeVersion="166925"/>
  <sheets>
    <sheet name="en_us" sheetId="1" r:id="rId1"/>
    <sheet name="DataSheet1" sheetId="2" r:id="rId2"/>
  </sheets>
  <definedNames>
    <definedName name="_xlnm._FilterDatabase" localSheetId="0" hidden="1">en_us!$A$1:$D$99</definedName>
    <definedName name="NotASheet">en_us!$A$1</definedName>
  </definedNames>
</workbook>`;

function asFile(buffer, name) {
  return new window.File([new Uint8Array(buffer)], name);
}

test('reads sheet names from a workbook', async () => {
  const zip = makeZip([{ name: 'xl/workbook.xml', content: WORKBOOK_XML }]);
  const names = await D365IA.xlsxSheets.readSheetNames(asFile(zip, 'book.xlsx'));
  equal(names, ['en_us', 'DataSheet1']);
});

test('ignores definedName entries', async () => {
  const zip = makeZip([{ name: 'xl/workbook.xml', content: WORKBOOK_XML }]);
  const names = await D365IA.xlsxSheets.readSheetNames(asFile(zip, 'book.xlsx'));
  assert(!names.includes('NotASheet'), `definedName leaked into ${JSON.stringify(names)}`);
  assert(!names.includes('_xlnm._FilterDatabase'), 'filter range leaked in');
});

// The local header carries zeroed sizes here, so the reader has to fall back
// to the central directory's — otherwise the entry reads back truncated and
// inflating it fails.
test('reads a streaming-mode zip whose local header has no sizes', async () => {
  const zip = makeZip([{ name: 'xl/workbook.xml', content: WORKBOOK_XML }], { streaming: true });
  const names = await D365IA.xlsxSheets.readSheetNames(asFile(zip, 'book.xlsx'));
  equal(names, ['en_us', 'DataSheet1']);
});

test('decodes XML entities in sheet names', async () => {
  const xml =
    '<workbook><sheets><sheet name="Q&amp;A &lt;draft&gt;" sheetId="1"/></sheets></workbook>';
  const zip = makeZip([{ name: 'xl/workbook.xml', content: xml }]);
  const names = await D365IA.xlsxSheets.readSheetNames(asFile(zip, 'book.xlsx'));
  equal(names, ['Q&A <draft>']);
});

// A CSV or a corrupt file must not throw — the queue treats "no sheets" as
// "never ask about sheets" and carries on.
test('returns nothing for a file that is not a workbook', async () => {
  const names = await D365IA.xlsxSheets.readSheetNames(
    new window.File(['id,name\n1,a\n'], 'data.csv')
  );
  equal(names, []);
});

test('finds the workbook among several zip entries', async () => {
  const zip = makeZip([
    { name: '[Content_Types].xml', content: '<Types/>' },
    { name: 'xl/workbook.xml', content: WORKBOOK_XML },
    { name: 'xl/styles.xml', content: '<styleSheet/>' }
  ]);
  const names = await D365IA.xlsxSheets.readSheetNames(asFile(zip, 'book.xlsx'));
  equal(names, ['en_us', 'DataSheet1']);
});
