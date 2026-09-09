(function () {
  const D365IA = (window.D365IA = window.D365IA || {});

  const EOCD_SIGNATURE = 0x06054b50;
  const CENTRAL_HEADER_SIGNATURE = 0x02014b50;
  const LOCAL_HEADER_SIGNATURE = 0x04034b50;
  const DEFLATE = 8;
  const STORED = 0;

  // Only the zip-based formats can be read this way; .xls is a binary blob.
  const ZIP_WORKBOOK_RE = /\.(xlsx|xlsm)$/i;

  // The end-of-central-directory record sits at the very end, behind an
  // optional comment of up to 64KB.
  function findEndOfCentralDirectory(view) {
    const earliest = Math.max(0, view.byteLength - 65557);
    for (let i = view.byteLength - 22; i >= earliest; i--) {
      if (view.getUint32(i, true) === EOCD_SIGNATURE) return i;
    }
    return -1;
  }

  async function inflateRaw(bytes) {
    const stream = new Blob([bytes]).stream().pipeThrough(new DecompressionStream('deflate-raw'));
    return new Uint8Array(await new Response(stream).arrayBuffer());
  }

  async function readZipEntry(buffer, wantedName) {
    const view = new DataView(buffer);
    const eocd = findEndOfCentralDirectory(view);
    if (eocd < 0) return null;

    const entryCount = view.getUint16(eocd + 10, true);
    let offset = view.getUint32(eocd + 16, true);
    const decoder = new TextDecoder();

    for (let i = 0; i < entryCount; i++) {
      if (offset + 46 > buffer.byteLength) return null;
      if (view.getUint32(offset, true) !== CENTRAL_HEADER_SIGNATURE) return null;

      const method = view.getUint16(offset + 10, true);
      const compressedSize = view.getUint32(offset + 20, true);
      const nameLength = view.getUint16(offset + 28, true);
      const extraLength = view.getUint16(offset + 30, true);
      const commentLength = view.getUint16(offset + 32, true);
      const localOffset = view.getUint32(offset + 42, true);
      const name = decoder.decode(new Uint8Array(buffer, offset + 46, nameLength));

      if (name === wantedName) {
        if (view.getUint32(localOffset, true) !== LOCAL_HEADER_SIGNATURE) return null;
        // The local header repeats these lengths and they can differ from
        // the central directory's, so read them again here. Also read the
        // local compressed size — some generators put the size in the local
        // header and it may differ from the central directory entry.
        const localNameLength = view.getUint16(localOffset + 26, true);
        const localExtraLength = view.getUint16(localOffset + 28, true);
        const localCompressedSize = view.getUint32(localOffset + 18, true);
        const dataStart = localOffset + 30 + localNameLength + localExtraLength;
        const data = new Uint8Array(buffer, dataStart, localCompressedSize);

        if (method === STORED) return data;
        if (method === DEFLATE) return inflateRaw(data);
        return null;
      }

      offset += 46 + nameLength + extraLength + commentLength;
    }
    return null;
  }

  function decodeXmlEntities(text) {
    return text
      .replace(/&lt;/g, '<')
      .replace(/&gt;/g, '>')
      .replace(/&quot;/g, '"')
      .replace(/&apos;/g, "'")
      .replace(/&amp;/g, '&');
  }

  // Reads worksheet names straight out of the workbook so a multi-sheet file
  // can be spotted before D365 asks which sheet to import. Returns [] for
  // .xls and for anything that can't be parsed — callers treat that as
  // "don't know", not as "one sheet".
  async function readSheetNames(file) {
    if (!ZIP_WORKBOOK_RE.test(file.name)) return [];
    try {
      const workbook = await readZipEntry(await file.arrayBuffer(), 'xl/workbook.xml');
      if (!workbook) return [];

      const xml = new TextDecoder().decode(workbook);
      const names = [];
      const sheetRe = /<sheet\b[^>]*\bname="([^"]*)"/g;
      let match;
      while ((match = sheetRe.exec(xml))) names.push(decodeXmlEntities(match[1]));
      return names;
    } catch (e) {
      return [];
    }
  }

  D365IA.xlsxSheets = { readSheetNames };
})();
