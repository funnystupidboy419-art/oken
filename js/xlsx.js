/* =========================================================================
 * xlsx.js  依存ライブラリなしの最小 XLSX(Excel) 書き出し
 *   - ZIP(無圧縮 store) + SpreadsheetML を自前で組み立てます。
 *   - 文字列は inlineStr で書き出すため sharedStrings は使いません。
 *   使い方:
 *     Xlsx.build([{ name:'シート名', cols:[8,40], rows:[['A','B'],[1,'text']] }])
 *       -> Blob
 * ========================================================================= */
(function (global) {
  'use strict';

  /* ---------------- CRC32 ---------------- */
  var CRC_TABLE = (function () {
    var table = new Uint32Array(256);
    for (var n = 0; n < 256; n++) {
      var c = n;
      for (var k = 0; k < 8; k++) {
        c = (c & 1) ? (0xEDB88320 ^ (c >>> 1)) : (c >>> 1);
      }
      table[n] = c >>> 0;
    }
    return table;
  })();

  function crc32(bytes) {
    var c = 0xFFFFFFFF;
    for (var i = 0; i < bytes.length; i++) {
      c = CRC_TABLE[(c ^ bytes[i]) & 0xFF] ^ (c >>> 8);
    }
    return (c ^ 0xFFFFFFFF) >>> 0;
  }

  var encoder = new TextEncoder();

  function utf8(str) { return encoder.encode(str); }

  /* ---------------- ZIP(store) ---------------- */
  function dosDateTime(d) {
    var time = (d.getHours() << 11) | (d.getMinutes() << 5) | (Math.floor(d.getSeconds() / 2));
    var date = ((d.getFullYear() - 1980) << 9) | ((d.getMonth() + 1) << 5) | d.getDate();
    return { time: time & 0xFFFF, date: date & 0xFFFF };
  }

  function zip(files) {
    var now = dosDateTime(new Date());
    var chunks = [];
    var central = [];
    var offset = 0;

    files.forEach(function (f) {
      var nameBytes = utf8(f.name);
      var dataBytes = f.data instanceof Uint8Array ? f.data : utf8(f.data);
      var crc = crc32(dataBytes);

      var local = new Uint8Array(30 + nameBytes.length);
      var lv = new DataView(local.buffer);
      lv.setUint32(0, 0x04034B50, true);
      lv.setUint16(4, 20, true);          /* version needed */
      lv.setUint16(6, 0x0800, true);      /* UTF-8 flag */
      lv.setUint16(8, 0, true);           /* method: store */
      lv.setUint16(10, now.time, true);
      lv.setUint16(12, now.date, true);
      lv.setUint32(14, crc, true);
      lv.setUint32(18, dataBytes.length, true);
      lv.setUint32(22, dataBytes.length, true);
      lv.setUint16(26, nameBytes.length, true);
      lv.setUint16(28, 0, true);
      local.set(nameBytes, 30);

      chunks.push(local, dataBytes);

      var cd = new Uint8Array(46 + nameBytes.length);
      var cv = new DataView(cd.buffer);
      cv.setUint32(0, 0x02014B50, true);
      cv.setUint16(4, 20, true);
      cv.setUint16(6, 20, true);
      cv.setUint16(8, 0x0800, true);
      cv.setUint16(10, 0, true);
      cv.setUint16(12, now.time, true);
      cv.setUint16(14, now.date, true);
      cv.setUint32(16, crc, true);
      cv.setUint32(20, dataBytes.length, true);
      cv.setUint32(24, dataBytes.length, true);
      cv.setUint16(28, nameBytes.length, true);
      cv.setUint16(30, 0, true);
      cv.setUint16(32, 0, true);
      cv.setUint16(34, 0, true);
      cv.setUint16(36, 0, true);
      cv.setUint32(38, 0, true);
      cv.setUint32(42, offset, true);
      cd.set(nameBytes, 46);
      central.push(cd);

      offset += local.length + dataBytes.length;
    });

    var centralSize = central.reduce(function (a, b) { return a + b.length; }, 0);
    var end = new Uint8Array(22);
    var ev = new DataView(end.buffer);
    ev.setUint32(0, 0x06054B50, true);
    ev.setUint16(4, 0, true);
    ev.setUint16(6, 0, true);
    ev.setUint16(8, files.length, true);
    ev.setUint16(10, files.length, true);
    ev.setUint32(12, centralSize, true);
    ev.setUint32(16, offset, true);
    ev.setUint16(20, 0, true);

    return new Blob(chunks.concat(central, [end]), {
      type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    });
  }

  /* ---------------- XML ---------------- */
  function esc(s) {
    return String(s == null ? '' : s)
      .replace(/[\x00-\x08\x0B\x0C\x0E-\x1F]/g, '')
      .replace(/&/g, '&amp;')
      .replace(/</g, '&lt;')
      .replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;');
  }

  function colName(index) {
    var s = '';
    var n = index + 1;
    while (n > 0) {
      var r = (n - 1) % 26;
      s = String.fromCharCode(65 + r) + s;
      n = Math.floor((n - 1) / 26);
    }
    return s;
  }

  function safeSheetName(name, fallback) {
    var s = String(name || fallback).replace(/[\[\]:*?\/\\]/g, ' ').slice(0, 31);
    return s === '' ? fallback : s;
  }

  function sheetXml(sheet) {
    var xml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">';
    xml += '<sheetViews><sheetView workbookViewId="0">' +
      '<pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/>' +
      '</sheetView></sheetViews>';
    if (sheet.cols && sheet.cols.length) {
      xml += '<cols>';
      sheet.cols.forEach(function (w, i) {
        xml += '<col min="' + (i + 1) + '" max="' + (i + 1) + '" width="' + w + '" customWidth="1"/>';
      });
      xml += '</cols>';
    }
    xml += '<sheetData>';
    (sheet.rows || []).forEach(function (row, r) {
      xml += '<row r="' + (r + 1) + '">';
      row.forEach(function (value, c) {
        if (value === null || value === undefined || value === '') return;
        var ref = colName(c) + (r + 1);
        var style = r === 0 ? 1 : (sheet.wrapFrom !== undefined && c >= sheet.wrapFrom ? 2 : 0);
        if (typeof value === 'number' && isFinite(value)) {
          xml += '<c r="' + ref + '" s="' + style + '"><v>' + value + '</v></c>';
        } else {
          xml += '<c r="' + ref + '" s="' + style + '" t="inlineStr"><is><t xml:space="preserve">' +
            esc(value) + '</t></is></c>';
        }
      });
      xml += '</row>';
    });
    xml += '</sheetData></worksheet>';
    return xml;
  }

  var STYLES_XML = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
    '<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">' +
    '<fonts count="2">' +
    '<font><sz val="11"/><color theme="1"/><name val="Yu Gothic"/></font>' +
    '<font><b/><sz val="11"/><color theme="1"/><name val="Yu Gothic"/></font>' +
    '</fonts>' +
    '<fills count="3">' +
    '<fill><patternFill patternType="none"/></fill>' +
    '<fill><patternFill patternType="gray125"/></fill>' +
    '<fill><patternFill patternType="solid"><fgColor rgb="FFEFF3F8"/><bgColor indexed="64"/></patternFill></fill>' +
    '</fills>' +
    '<borders count="2">' +
    '<border><left/><right/><top/><bottom/><diagonal/></border>' +
    '<border><left style="thin"><color rgb="FFBBBBBB"/></left><right style="thin"><color rgb="FFBBBBBB"/></right>' +
    '<top style="thin"><color rgb="FFBBBBBB"/></top><bottom style="thin"><color rgb="FFBBBBBB"/></bottom><diagonal/></border>' +
    '</borders>' +
    '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>' +
    '<cellXfs count="3">' +
    '<xf numFmtId="0" fontId="0" fillId="0" borderId="1" xfId="0" applyBorder="1">' +
    '<alignment vertical="top"/></xf>' +
    '<xf numFmtId="0" fontId="1" fillId="2" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1">' +
    '<alignment horizontal="center" vertical="center"/></xf>' +
    '<xf numFmtId="0" fontId="0" fillId="0" borderId="1" xfId="0" applyBorder="1" applyAlignment="1">' +
    '<alignment vertical="top" wrapText="1"/></xf>' +
    '</cellXfs>' +
    '<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>' +
    '</styleSheet>';

  var Xlsx = {};

  /**
   * sheets: [{ name, rows:[[cell,...],...], cols:[width,...], wrapFrom:Number }]
   * 1行目はヘッダー行として書式が付きます。
   */
  Xlsx.build = function (sheets) {
    var list = (sheets || []).map(function (s, i) {
      return {
        name: safeSheetName(s.name, 'Sheet' + (i + 1)),
        rows: s.rows || [],
        cols: s.cols || [],
        wrapFrom: s.wrapFrom
      };
    });
    if (!list.length) list = [{ name: 'Sheet1', rows: [], cols: [] }];

    var contentTypes = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
      '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
      '<Default Extension="xml" ContentType="application/xml"/>' +
      '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>' +
      '<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>' +
      list.map(function (s, i) {
        return '<Override PartName="/xl/worksheets/sheet' + (i + 1) + '.xml" ' +
          'ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>';
      }).join('') +
      '</Types>';

    var rootRels = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
      '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>' +
      '</Relationships>';

    var workbook = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" ' +
      'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets>' +
      list.map(function (s, i) {
        return '<sheet name="' + esc(s.name) + '" sheetId="' + (i + 1) + '" r:id="rId' + (i + 1) + '"/>';
      }).join('') +
      '</sheets></workbook>';

    var workbookRels = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
      list.map(function (s, i) {
        return '<Relationship Id="rId' + (i + 1) + '" ' +
          'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" ' +
          'Target="worksheets/sheet' + (i + 1) + '.xml"/>';
      }).join('') +
      '<Relationship Id="rId' + (list.length + 1) + '" ' +
      'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>' +
      '</Relationships>';

    var files = [
      { name: '[Content_Types].xml', data: contentTypes },
      { name: '_rels/.rels', data: rootRels },
      { name: 'xl/workbook.xml', data: workbook },
      { name: 'xl/_rels/workbook.xml.rels', data: workbookRels },
      { name: 'xl/styles.xml', data: STYLES_XML }
    ];
    list.forEach(function (s, i) {
      files.push({ name: 'xl/worksheets/sheet' + (i + 1) + '.xml', data: sheetXml(s) });
    });

    return zip(files);
  };

  /* CSV(UTF-8 BOM付き。Excelでそのまま開けます) */
  Xlsx.toCsv = function (rows) {
    var body = (rows || []).map(function (row) {
      return row.map(function (v) {
        var s = String(v == null ? '' : v);
        return /[",\n]/.test(s) ? '"' + s.replace(/"/g, '""') + '"' : s;
      }).join(',');
    }).join('\r\n');
    return new Blob([String.fromCharCode(0xFEFF) + body], { type: 'text/csv;charset=utf-8;' });
  };

  global.Xlsx = Xlsx;
})(window);
