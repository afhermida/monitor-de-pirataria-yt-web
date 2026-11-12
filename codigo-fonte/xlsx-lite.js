(function (global) {
  'use strict';

  const REL_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
  const MAIN_NS = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';

  function requireZip() {
    if (!global.JSZip) throw new Error('JSZip não foi carregado.');
    return global.JSZip;
  }

  function parseXml(text, fileName) {
    const doc = new DOMParser().parseFromString(text, 'application/xml');
    const errors = doc.getElementsByTagName('parsererror');
    if (errors.length) throw new Error(`XML inválido em ${fileName || 'arquivo XLSX'}.`);
    return doc;
  }

  function elements(node, localName) {
    return Array.from(node.getElementsByTagNameNS('*', localName));
  }

  function textFromTs(node) {
    return elements(node, 't').map((item) => item.textContent || '').join('');
  }

  function columnIndexFromReference(reference) {
    const match = /^([A-Z]+)\d+$/i.exec(String(reference || ''));
    if (!match) return -1;
    let result = 0;
    for (const char of match[1].toUpperCase()) result = (result * 26) + (char.charCodeAt(0) - 64);
    return result - 1;
  }

  function columnName(index) {
    let number = Number(index) + 1;
    let label = '';
    while (number > 0) {
      const remainder = (number - 1) % 26;
      label = String.fromCharCode(65 + remainder) + label;
      number = Math.floor((number - 1) / 26);
    }
    return label;
  }

  function normalizeZipPath(base, target) {
    let value = String(target || '').replace(/\\/g, '/');
    if (value.startsWith('/')) return value.slice(1);
    const parts = `${base}/${value}`.split('/');
    const normalized = [];
    for (const part of parts) {
      if (!part || part === '.') continue;
      if (part === '..') normalized.pop();
      else normalized.push(part);
    }
    return normalized.join('/');
  }

  async function loadSharedStrings(zip) {
    const file = zip.file('xl/sharedStrings.xml');
    if (!file) return [];
    const doc = parseXml(await file.async('string'), 'xl/sharedStrings.xml');
    return elements(doc, 'si').map(textFromTs);
  }

  function readCell(cell, sharedStrings) {
    const type = cell.getAttribute('t') || '';
    if (type === 'inlineStr') {
      const inline = elements(cell, 'is')[0];
      return inline ? textFromTs(inline) : '';
    }
    const valueNode = elements(cell, 'v')[0];
    const raw = valueNode ? valueNode.textContent || '' : '';
    if (type === 's') return sharedStrings[Number(raw)] == null ? '' : sharedStrings[Number(raw)];
    if (type === 'b') return raw === '1';
    if (type === 'str' || type === 'e') return raw;
    if (raw === '') return '';
    const numeric = Number(raw);
    return Number.isFinite(numeric) ? numeric : raw;
  }

  async function readWorksheet(zip, path, sharedStrings) {
    const file = zip.file(path);
    if (!file) throw new Error(`A planilha referencia um arquivo ausente: ${path}`);
    const doc = parseXml(await file.async('string'), path);
    const rows = [];
    for (const rowElement of elements(doc, 'row')) {
      const explicitRow = Number(rowElement.getAttribute('r') || 0);
      const rowIndex = explicitRow > 0 ? explicitRow - 1 : rows.length;
      const row = rows[rowIndex] || [];
      for (const cell of elements(rowElement, 'c')) {
        const ref = cell.getAttribute('r') || '';
        const columnIndex = columnIndexFromReference(ref);
        if (columnIndex >= 0) row[columnIndex] = readCell(cell, sharedStrings);
      }
      rows[rowIndex] = row;
    }
    for (let i = 0; i < rows.length; i += 1) if (!rows[i]) rows[i] = [];
    return rows;
  }

  async function readWorkbook(arrayBuffer) {
    const JSZip = requireZip();
    const zip = await JSZip.loadAsync(arrayBuffer);
    const workbookFile = zip.file('xl/workbook.xml');
    const relsFile = zip.file('xl/_rels/workbook.xml.rels');
    if (!workbookFile || !relsFile) throw new Error('O arquivo não parece ser uma planilha XLSX válida.');

    const workbookDoc = parseXml(await workbookFile.async('string'), 'xl/workbook.xml');
    const relsDoc = parseXml(await relsFile.async('string'), 'xl/_rels/workbook.xml.rels');
    const relationships = new Map();
    for (const relationship of elements(relsDoc, 'Relationship')) {
      relationships.set(relationship.getAttribute('Id'), relationship.getAttribute('Target'));
    }
    const sharedStrings = await loadSharedStrings(zip);
    const sheets = [];
    for (const sheetElement of elements(workbookDoc, 'sheet')) {
      const name = sheetElement.getAttribute('name') || `Planilha ${sheets.length + 1}`;
      const relId = sheetElement.getAttributeNS(REL_NS, 'id') || sheetElement.getAttribute('r:id');
      const target = relationships.get(relId);
      if (!target) continue;
      const path = normalizeZipPath('xl', target);
      const rows = await readWorksheet(zip, path, sharedStrings);
      sheets.push({ name, rows });
    }
    if (!sheets.length) throw new Error('Nenhuma aba legível foi encontrada na planilha.');
    return { sheets };
  }

  function xmlEscape(value) {
    return String(value == null ? '' : value)
      .replace(/[\u0000-\u0008\u000B\u000C\u000E-\u001F]/g, '')
      .replace(/&/g, '&amp;')
      .replace(/</g, '&lt;')
      .replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;')
      .replace(/'/g, '&apos;');
  }

  function sanitizeSheetName(name, usedNames) {
    let cleaned = String(name || 'Planilha').replace(/[\\/?*\[\]:]/g, ' ').replace(/\s+/g, ' ').trim();
    if (!cleaned) cleaned = 'Planilha';
    cleaned = cleaned.slice(0, 31);
    let candidate = cleaned;
    let counter = 2;
    while (usedNames.has(candidate.toLowerCase())) {
      const suffix = ` (${counter})`;
      candidate = `${cleaned.slice(0, 31 - suffix.length)}${suffix}`;
      counter += 1;
    }
    usedNames.add(candidate.toLowerCase());
    return candidate;
  }

  const STYLE_IDS = Object.freeze({
    default: 0,
    header: 1,
    title: 2,
    subtitle: 3,
    summaryLabel: 4,
    summaryValue: 5,
    summaryPercent: 6,
    wrap: 7,
    integer: 8,
    scoreRed: 9,
    scoreOrange: 10,
    scoreGreen: 11,
    hyperlink: 12,
    statusRed: 13,
    statusOrange: 14,
    statusGreen: 15,
    textAlt: 16,
    wrapAlt: 17,
    integerAlt: 18,
    scoreRedAlt: 19,
    scoreOrangeAlt: 20,
    scoreGreenAlt: 21,
    hyperlinkAlt: 22,
    center: 23,
    centerAlt: 24,
    decimal: 25,
    decimalAlt: 26,
    dateText: 27,
    dateTextAlt: 28,
  });

  function styleId(style) {
    if (typeof style === 'number') return style;
    return Object.prototype.hasOwnProperty.call(STYLE_IDS, style) ? STYLE_IDS[style] : 0;
  }

  function stylesXml() {
    const fonts = [
      '<font><sz val="11"/><color theme="1"/><name val="Calibri"/><family val="2"/><scheme val="minor"/></font>',
      '<font><b/><sz val="11"/><color rgb="FFFFFFFF"/><name val="Calibri"/><family val="2"/></font>',
      '<font><b/><sz val="16"/><color rgb="FF24528A"/><name val="Calibri"/><family val="2"/></font>',
      '<font><sz val="10"/><color rgb="FF808080"/><name val="Calibri"/><family val="2"/></font>',
      '<font><b/><sz val="11"/><color theme="1"/><name val="Calibri"/><family val="2"/></font>',
      '<font><b/><sz val="11"/><color rgb="FFD93025"/><name val="Calibri"/><family val="2"/></font>',
      '<font><b/><sz val="11"/><color rgb="FFEF6C00"/><name val="Calibri"/><family val="2"/></font>',
      '<font><sz val="11"/><color rgb="FF188038"/><name val="Calibri"/><family val="2"/></font>',
      '<font><u/><sz val="11"/><color rgb="FF0563C1"/><name val="Calibri"/><family val="2"/></font>',
    ];
    const fills = [
      '<fill><patternFill patternType="none"/></fill>',
      '<fill><patternFill patternType="gray125"/></fill>',
      '<fill><patternFill patternType="solid"><fgColor rgb="FF28588F"/><bgColor indexed="64"/></patternFill></fill>',
      '<fill><patternFill patternType="solid"><fgColor rgb="FFF2F2F2"/><bgColor indexed="64"/></patternFill></fill>',
      '<fill><patternFill patternType="solid"><fgColor rgb="FFEAF3F8"/><bgColor indexed="64"/></patternFill></fill>',
    ];
    const alignWrap = '<alignment vertical="top" wrapText="1"/>';
    const alignCenterWrap = '<alignment horizontal="center" vertical="center" wrapText="1"/>';
    const alignCenter = '<alignment horizontal="center" vertical="top"/>';
    const xf = (fontId = 0, fillId = 0, numFmtId = 0, alignment = '', options = '') => (
      `<xf numFmtId="${numFmtId}" fontId="${fontId}" fillId="${fillId}" borderId="0" xfId="0"${fontId ? ' applyFont="1"' : ''}${fillId ? ' applyFill="1"' : ''}${numFmtId ? ' applyNumberFormat="1"' : ''}${alignment ? ' applyAlignment="1"' : ''}${options}>${alignment}</xf>`
    );
    const cellXfs = [
      xf(),
      xf(1, 2, 0, alignCenterWrap),
      xf(2),
      xf(3),
      xf(0, 3),
      xf(4),
      xf(4, 0, 165),
      xf(0, 0, 0, alignWrap),
      xf(0, 0, 166),
      xf(5, 0, 164),
      xf(6, 0, 164),
      xf(7, 0, 164),
      xf(8),
      xf(5, 0, 0, alignCenter),
      xf(6, 0, 0, alignCenter),
      xf(7, 0, 0, alignCenter),
      xf(0, 4),
      xf(0, 4, 0, alignWrap),
      xf(0, 4, 166),
      xf(5, 4, 164),
      xf(6, 4, 164),
      xf(7, 4, 164),
      xf(8, 4),
      xf(0, 0, 0, alignCenter),
      xf(0, 4, 0, alignCenter),
      xf(0, 0, 164),
      xf(0, 4, 164),
      xf(0, 0, 167),
      xf(0, 4, 167),
    ];
    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="${MAIN_NS}">
  <numFmts count="4"><numFmt numFmtId="164" formatCode="0.00"/><numFmt numFmtId="165" formatCode="0.0%"/><numFmt numFmtId="166" formatCode="#,##0"/><numFmt numFmtId="167" formatCode="dd/mm/yyyy"/></numFmts>
  <fonts count="${fonts.length}">${fonts.join('')}</fonts>
  <fills count="${fills.length}">${fills.join('')}</fills>
  <borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>
  <cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>
  <cellXfs count="${cellXfs.length}">${cellXfs.join('')}</cellXfs>
  <cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>
  <dxfs count="0"/>
  <tableStyles count="0" defaultTableStyle="TableStyleMedium2" defaultPivotStyle="PivotStyleLight16"/>
</styleSheet>`;
  }

  function cellXml(reference, cell) {
    const spec = cell && typeof cell === 'object' && !Array.isArray(cell) && Object.prototype.hasOwnProperty.call(cell, 'value')
      ? cell : { value: cell };
    const value = spec.value;
    const style = styleId(spec.style);
    const styleAttr = style ? ` s="${style}"` : '';
    if (value == null || value === '') return `<c r="${reference}"${styleAttr}/>`;
    if (typeof value === 'number' && Number.isFinite(value)) return `<c r="${reference}"${styleAttr}><v>${value}</v></c>`;
    if (typeof value === 'boolean') return `<c r="${reference}" t="b"${styleAttr}><v>${value ? 1 : 0}</v></c>`;
    const text = xmlEscape(value);
    return `<c r="${reference}" t="inlineStr"${styleAttr}><is><t xml:space="preserve">${text}</t></is></c>`;
  }

  function worksheetXml(sheet) {
    const rows = Array.isArray(sheet.rows) ? sheet.rows : [];
    const columnCount = Math.max(1, ...rows.map((row) => Array.isArray(row) ? row.length : 0));
    const rowCount = Math.max(1, rows.length);
    const lastRef = `${columnName(columnCount - 1)}${rowCount}`;
    const columns = (sheet.columns || []).map((width, index) => (
      `<col min="${index + 1}" max="${index + 1}" width="${Number(width) || 10}" customWidth="1"/>`
    )).join('');
    const hyperlinks = [];
    const relationshipItems = [];
    const rowXml = rows.map((row, rowIndex) => {
      const cells = (row || []).map((cell, columnIndex) => {
        const reference = `${columnName(columnIndex)}${rowIndex + 1}`;
        const spec = cell && typeof cell === 'object' && !Array.isArray(cell) && Object.prototype.hasOwnProperty.call(cell, 'value')
          ? cell : { value: cell };
        if (spec.link) {
          const relId = `rId${relationshipItems.length + 1}`;
          relationshipItems.push({ id: relId, target: String(spec.link) });
          hyperlinks.push(`<hyperlink ref="${reference}" r:id="${relId}"/>`);
        }
        return cellXml(reference, cell);
      }).join('');
      const height = sheet.rowHeights && sheet.rowHeights[rowIndex + 1];
      const heightAttr = height ? ` ht="${Number(height)}" customHeight="1"` : '';
      return `<row r="${rowIndex + 1}"${heightAttr}>${cells}</row>`;
    }).join('');

    const freezeRows = Math.max(0, Number(sheet.freezeRows || 0));
    const sheetViews = freezeRows
      ? `<sheetViews><sheetView workbookViewId="0"><pane ySplit="${freezeRows}" topLeftCell="A${freezeRows + 1}" activePane="bottomLeft" state="frozen"/><selection pane="bottomLeft" activeCell="A${freezeRows + 1}" sqref="A${freezeRows + 1}"/></sheetView></sheetViews>`
      : '<sheetViews><sheetView workbookViewId="0"/></sheetViews>';
    const merges = (sheet.merges || []).length
      ? `<mergeCells count="${sheet.merges.length}">${sheet.merges.map((ref) => `<mergeCell ref="${xmlEscape(ref)}"/>`).join('')}</mergeCells>` : '';
    const autoFilter = sheet.autoFilter
      ? `<autoFilter ref="${typeof sheet.autoFilter === 'string' ? xmlEscape(sheet.autoFilter) : `A1:${lastRef}` }"/>` : '';
    const hyperlinkXml = hyperlinks.length ? `<hyperlinks>${hyperlinks.join('')}</hyperlinks>` : '';

    const xml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="${MAIN_NS}" xmlns:r="${REL_NS}">
  <dimension ref="A1:${lastRef}"/>
  ${sheetViews}
  <sheetFormatPr defaultRowHeight="15"/>
  ${columns ? `<cols>${columns}</cols>` : ''}
  <sheetData>${rowXml}</sheetData>
  ${merges}
  ${autoFilter}
  ${hyperlinkXml}
  <pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>
</worksheet>`;
    const rels = relationshipItems.length
      ? `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${relationshipItems.map((item) => `<Relationship Id="${item.id}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="${xmlEscape(item.target)}" TargetMode="External"/>`).join('')}</Relationships>`
      : null;
    return { xml, rels };
  }

  function workbookXml(sheets) {
    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="${MAIN_NS}" xmlns:r="${REL_NS}"><bookViews><workbookView xWindow="0" yWindow="0" windowWidth="24000" windowHeight="12000"/></bookViews><sheets>${sheets.map((sheet, index) => `<sheet name="${xmlEscape(sheet.name)}" sheetId="${index + 1}" r:id="rId${index + 1}"/>`).join('')}</sheets><calcPr calcId="191029" fullCalcOnLoad="1"/></workbook>`;
  }

  function workbookRelsXml(sheetCount) {
    const rels = [];
    for (let index = 0; index < sheetCount; index += 1) {
      rels.push(`<Relationship Id="rId${index + 1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet${index + 1}.xml"/>`);
    }
    rels.push(`<Relationship Id="rId${sheetCount + 1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>`);
    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${rels.join('')}</Relationships>`;
  }

  function contentTypesXml(sheetCount) {
    const overrides = [];
    for (let index = 0; index < sheetCount; index += 1) {
      overrides.push(`<Override PartName="/xl/worksheets/sheet${index + 1}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>`);
    }
    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/><Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/><Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>${overrides.join('')}</Types>`;
  }

  function rootRelsXml() {
    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/><Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/></Relationships>`;
  }

  function corePropertiesXml(metadata) {
    const now = new Date().toISOString();
    const creator = xmlEscape(metadata.creator || 'Monitor de Audiolivros no YouTube');
    const title = xmlEscape(metadata.title || 'Relatório de monitoramento');
    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:dcmitype="http://purl.org/dc/dcmitype/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"><dc:title>${title}</dc:title><dc:creator>${creator}</dc:creator><cp:lastModifiedBy>${creator}</cp:lastModifiedBy><dcterms:created xsi:type="dcterms:W3CDTF">${now}</dcterms:created><dcterms:modified xsi:type="dcterms:W3CDTF">${now}</dcterms:modified></cp:coreProperties>`;
  }

  function appPropertiesXml(sheets) {
    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"><Application>Monitor de Audiolivros no YouTube</Application><DocSecurity>0</DocSecurity><ScaleCrop>false</ScaleCrop><HeadingPairs><vt:vector size="2" baseType="variant"><vt:variant><vt:lpstr>Worksheets</vt:lpstr></vt:variant><vt:variant><vt:i4>${sheets.length}</vt:i4></vt:variant></vt:vector></HeadingPairs><TitlesOfParts><vt:vector size="${sheets.length}" baseType="lpstr">${sheets.map((sheet) => `<vt:lpstr>${xmlEscape(sheet.name)}</vt:lpstr>`).join('')}</vt:vector></TitlesOfParts><Company></Company><LinksUpToDate>false</LinksUpToDate><SharedDoc>false</SharedDoc><HyperlinksChanged>false</HyperlinksChanged><AppVersion>16.0300</AppVersion></Properties>`;
  }

  async function createWorkbookData(sheetsInput, metadata = {}) {
    const JSZip = requireZip();
    const usedNames = new Set();
    const sheets = (sheetsInput || []).map((sheet, index) => ({
      ...sheet,
      name: sanitizeSheetName(sheet.name || `Planilha ${index + 1}`, usedNames),
    }));
    if (!sheets.length) throw new Error('É necessário informar pelo menos uma aba.');

    const zip = new JSZip();
    zip.file('[Content_Types].xml', contentTypesXml(sheets.length));
    zip.folder('_rels').file('.rels', rootRelsXml());
    zip.folder('docProps').file('core.xml', corePropertiesXml(metadata));
    zip.folder('docProps').file('app.xml', appPropertiesXml(sheets));
    zip.folder('xl').file('workbook.xml', workbookXml(sheets));
    zip.folder('xl').file('styles.xml', stylesXml());
    zip.folder('xl').folder('_rels').file('workbook.xml.rels', workbookRelsXml(sheets.length));

    const worksheetsFolder = zip.folder('xl').folder('worksheets');
    const worksheetRelsFolder = worksheetsFolder.folder('_rels');
    sheets.forEach((sheet, index) => {
      const generated = worksheetXml(sheet);
      worksheetsFolder.file(`sheet${index + 1}.xml`, generated.xml);
      if (generated.rels) worksheetRelsFolder.file(`sheet${index + 1}.xml.rels`, generated.rels);
    });

    return zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE', compressionOptions: { level: 6 } });
  }

  async function createWorkbookBlob(sheets, metadata = {}) {
    const data = await createWorkbookData(sheets, metadata);
    return new Blob([data], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
  }

  const api = {
    readWorkbook, createWorkbookData, createWorkbookBlob, columnName, columnIndexFromReference,
    STYLE_IDS,
  };
  global.XlsxLite = api;
  if (typeof module !== 'undefined' && module.exports) module.exports = api;
}(typeof window !== 'undefined' ? window : globalThis));
