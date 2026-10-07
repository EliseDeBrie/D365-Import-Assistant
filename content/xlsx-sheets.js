(function () {
  const D365IA = (window.D365IA = window.D365IA || {});

  const EOCD_SIGNATURE = 0x06054b50;
  const CENTRAL_HEADER_SIGNATURE = 0x02014b50;
  const LOCAL_HEADER_SIGNATURE = 0x04034b50;
  const DEFLATE = 8;
  const STORED = 0;

  // Only the zip-based formats can be read this way; .xls is a binary blob.
  const ZIP_WORKBOOK_RE = /\.(xlsx|xlsm)$/i;
  const MAX_ARCHIVE_BYTES = 256 * 1024 * 1024;
  const MAX_ENTRY_COUNT = 10000;
  const MAX_CENTRAL_DIRECTORY_BYTES = 16 * 1024 * 1024;
  const MAX_COMPRESSED_WORKBOOK_BYTES = 4 * 1024 * 1024;
  const MAX_PARSE_TIME_MS = 8000;
  const MAX_SHEETS = 1000;

  // The end-of-central-directory record sits at the very end, behind an
  // optional comment of up to 64KB.
  function findEndOfCentralDirectory(view) {
    const earliest = Math.max(0, view.byteLength - 65557);
    for (let i = view.byteLength - 22; i >= earliest; i--) {
      if (view.getUint32(i, true) === EOCD_SIGNATURE) return i;
    }
    return -1;
  }

  // workbook.xml is a few KB even for huge workbooks. Files arrive from
  // third parties, and a crafted entry that inflates to gigabytes would take
  // the whole D365 tab down with it, so inflation stops at this ceiling.
  const MAX_INFLATED_BYTES = 4 * 1024 * 1024;

  async function inflateRaw(bytes) {
    const reader = new Blob([bytes])
      .stream()
      .pipeThrough(new DecompressionStream('deflate-raw'))
      .getReader();
    const chunks = [];
    let total = 0;
    let timeoutId;
    const timedOut = new Promise((_, reject) => {
      timeoutId = setTimeout(() => reject(new Error('workbook index parsing timed out')), MAX_PARSE_TIME_MS);
    });

    try {
      for (;;) {
        const { done, value } = await Promise.race([reader.read(), timedOut]);
        if (done) break;
        total += value.byteLength;
        if (total > MAX_INFLATED_BYTES) {
          throw new Error('workbook index exceeds the size limit');
        }
        chunks.push(value);
      }
    } catch (e) {
      await reader.cancel().catch(() => {});
      throw e;
    } finally {
      clearTimeout(timeoutId);
      reader.releaseLock();
    }

    const out = new Uint8Array(total);
    let offset = 0;
    chunks.forEach((chunk) => {
      out.set(chunk, offset);
      offset += chunk.byteLength;
    });
    return out;
  }

  async function readZipEntry(buffer, wantedName) {
    const view = new DataView(buffer);
    const eocd = findEndOfCentralDirectory(view);
    if (eocd < 0 || eocd + 22 > buffer.byteLength) return null;

    const entryCount = view.getUint16(eocd + 10, true);
    const centralDirectoryBytes = view.getUint32(eocd + 12, true);
    let offset = view.getUint32(eocd + 16, true);
    const centralDirectoryEnd = offset + centralDirectoryBytes;
    if (
      entryCount === 0xffff ||
      entryCount > MAX_ENTRY_COUNT ||
      centralDirectoryBytes > MAX_CENTRAL_DIRECTORY_BYTES ||
      centralDirectoryEnd > eocd
    ) {
      return null;
    }
    const decoder = new TextDecoder();

    for (let i = 0; i < entryCount; i++) {
      if (offset + 46 > centralDirectoryEnd) return null;
      if (view.getUint32(offset, true) !== CENTRAL_HEADER_SIGNATURE) return null;

      const method = view.getUint16(offset + 10, true);
      const compressedSize = view.getUint32(offset + 20, true);
      const uncompressedSize = view.getUint32(offset + 24, true);
      const nameLength = view.getUint16(offset + 28, true);
      const extraLength = view.getUint16(offset + 30, true);
      const commentLength = view.getUint16(offset + 32, true);
      const localOffset = view.getUint32(offset + 42, true);
      const nextOffset = offset + 46 + nameLength + extraLength + commentLength;
      if (nextOffset > centralDirectoryEnd) return null;
      const name = decoder.decode(new Uint8Array(buffer, offset + 46, nameLength));

      if (name === wantedName) {
        if (
          compressedSize === 0xffffffff ||
          uncompressedSize === 0xffffffff ||
          localOffset === 0xffffffff ||
          compressedSize > MAX_COMPRESSED_WORKBOOK_BYTES ||
          uncompressedSize > MAX_INFLATED_BYTES ||
          localOffset + 30 > buffer.byteLength
        ) {
          return null;
        }
        if (view.getUint32(localOffset, true) !== LOCAL_HEADER_SIGNATURE) return null;
        // A streaming zip may leave local sizes empty, so the central
        // directory is authoritative for the compressed size.
        const localNameLength = view.getUint16(localOffset + 26, true);
        const localExtraLength = view.getUint16(localOffset + 28, true);
        const dataStart = localOffset + 30 + localNameLength + localExtraLength;
        if (dataStart + compressedSize > buffer.byteLength) return null;
        const data = new Uint8Array(buffer, dataStart, compressedSize);

        if (method === STORED) {
          return data.byteLength <= MAX_INFLATED_BYTES ? data : null;
        }
        if (method === DEFLATE) return inflateRaw(data);
        return null;
      }

      offset = nextOffset;
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
    if (
      !file ||
      !ZIP_WORKBOOK_RE.test(file.name) ||
      (Number.isFinite(file.size) && file.size > MAX_ARCHIVE_BYTES)
    ) {
      return [];
    }
    try {
      const workbook = await readZipEntry(await file.arrayBuffer(), 'xl/workbook.xml');
      if (!workbook) return [];

      const xml = new TextDecoder().decode(workbook);
      const names = [];
      const sheetRe = /<sheet\b[^>]*\bname="([^"]*)"/g;
      let match;
      while ((match = sheetRe.exec(xml)) && names.length < MAX_SHEETS) {
        names.push(decodeXmlEntities(match[1]));
      }
      return names;
    } catch (e) {
      return [];
    }
  }

  D365IA.xlsxSheets = { readSheetNames };
})();
