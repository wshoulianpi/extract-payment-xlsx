#!/usr/bin/env node
/*
 * 从 JSON 数据生成收款记录 xlsx（零依赖，Node >= 14）。
 * 用法: node build_payment_xlsx.js <data.json> <output.xlsx>
 *
 * data.json 格式:
 * {
 *   "title": "收款记录",            // 可选，表名
 *   "records": [
 *     { "payer": "某某公司", "amount": 1600.00, "type": "银行电子回执", "date": "2026-08-17", "source": "xxx.png" }
 *   ]
 * }
 * type/date/source 均可省略。
 */
const fs = require('fs');
const path = require('path');

// ---------- CRC32 ----------
const crcTable = (() => {
  const t = new Uint32Array(256);
  for (let n = 0; n < 256; n++) {
    let c = n;
    for (let k = 0; k < 8; k++) c = c & 1 ? 0xedb88320 ^ (c >>> 1) : c >>> 1;
    t[n] = c >>> 0;
  }
  return t;
})();
function crc32(buf) {
  let c = 0xffffffff;
  for (let i = 0; i < buf.length; i++) c = crcTable[(c ^ buf[i]) & 0xff] ^ (c >>> 8);
  return (c ^ 0xffffffff) >>> 0;
}

// ---------- ZIP (stored, 不压缩) ----------
function buildZip(entries) {
  const locals = [];
  const centrals = [];
  let offset = 0;
  for (const { name, data } of entries) {
    const nameBuf = Buffer.from(name, 'utf8');
    const crc = crc32(data);
    const local = Buffer.alloc(30);
    local.writeUInt32LE(0x04034b50, 0);
    local.writeUInt16LE(20, 4);          // version needed
    local.writeUInt16LE(0x0800, 6);      // UTF-8 flag
    local.writeUInt16LE(0, 8);           // method: stored
    local.writeUInt16LE(0, 10);          // time
    local.writeUInt16LE(0x21, 12);       // date 1980-01-01
    local.writeUInt32LE(crc, 14);
    local.writeUInt32LE(data.length, 18);
    local.writeUInt32LE(data.length, 22);
    local.writeUInt16LE(nameBuf.length, 26);
    local.writeUInt16LE(0, 28);
    locals.push(local, nameBuf, data);

    const central = Buffer.alloc(46);
    central.writeUInt32LE(0x02014b50, 0);
    central.writeUInt16LE(20, 4);
    central.writeUInt16LE(20, 6);
    central.writeUInt16LE(0x0800, 8);
    central.writeUInt16LE(0, 10);
    central.writeUInt16LE(0, 12);
    central.writeUInt16LE(0x21, 14);
    central.writeUInt32LE(crc, 16);
    central.writeUInt32LE(data.length, 20);
    central.writeUInt32LE(data.length, 24);
    central.writeUInt16LE(nameBuf.length, 28);
    central.writeUInt32LE(offset, 42);
    centrals.push(central, nameBuf);
    offset += 30 + nameBuf.length + data.length;
  }
  const centralSize = centrals.reduce((s, b) => s + b.length, 0);
  const end = Buffer.alloc(22);
  end.writeUInt32LE(0x06054b50, 0);
  end.writeUInt16LE(entries.length, 8);
  end.writeUInt16LE(entries.length, 10);
  end.writeUInt32LE(centralSize, 12);
  end.writeUInt32LE(offset, 16);
  return Buffer.concat([...locals, ...centrals, end]);
}

// ---------- XLSX ----------
function colName(i) {
  let s = '';
  i += 1;
  while (i > 0) { s = String.fromCharCode(65 + ((i - 1) % 26)) + s; i = Math.floor((i - 1) / 26); }
  return s;
}
function xmlEscape(s) {
  return String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');
}
function cell(ref, value) {
  if (typeof value === 'number') {
    return `<c r="${ref}"><v>${value}</v></c>`;
  }
  return `<c r="${ref}" t="inlineStr"><is><t xml:space="preserve">${xmlEscape(value)}</t></is></c>`;
}

function buildSheet(headers, rows) {
  let xml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
    '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData>';
  xml += '<row r="1">' + headers.map((h, i) => cell(colName(i) + '1', h)).join('') + '</row>';
  rows.forEach((row, r) => {
    const rn = r + 2;
    xml += `<row r="${rn}">` + row.map((v, i) => cell(colName(i) + rn, v)).join('') + '</row>';
  });
  xml += '</sheetData></worksheet>';
  return Buffer.from(xml, 'utf8');
}

function main() {
  const [jsonPath, outPath] = process.argv.slice(2);
  if (!jsonPath || !outPath) {
    console.error('用法: node build_payment_xlsx.js <data.json> <output.xlsx>');
    process.exit(1);
  }
  const data = JSON.parse(fs.readFileSync(jsonPath, 'utf8'));
  const records = data.records || [];

  const headers = ['序号', '付款方', '金额(元)', '类型', '日期', '来源文件'];
  const rows = records.map((r, i) => [
    i + 1, r.payer, Number(r.amount), r.type || '', r.date || '', r.source || '',
  ]);
  const total = records.reduce((s, r) => s + Number(r.amount), 0);
  rows.push(['', '合计', Number(total.toFixed(2)), '', '', '']);

  const sheet = buildSheet(headers, rows);
  const str = s => Buffer.from(s, 'utf8');
  const entries = [
    { name: '[Content_Types].xml', data: str('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
      '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
      '<Default Extension="xml" ContentType="application/xml"/>' +
      '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>' +
      '<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>' +
      '</Types>') },
    { name: '_rels/.rels', data: str('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
      '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>' +
      '</Relationships>') },
    { name: 'xl/workbook.xml', data: str('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" ' +
      'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">' +
      `<sheets><sheet name="${xmlEscape(data.title || '收款记录')}" sheetId="1" r:id="rId1"/></sheets></workbook>`) },
    { name: 'xl/_rels/workbook.xml.rels', data: str('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
      '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>' +
      '</Relationships>') },
    { name: 'xl/worksheets/sheet1.xml', data: sheet },
  ];
  fs.writeFileSync(outPath, buildZip(entries));
  console.log(`已生成 ${path.resolve(outPath)}（${records.length} 条记录，合计 ${total.toFixed(2)} 元）`);
}

main();
