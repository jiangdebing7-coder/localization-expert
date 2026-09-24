"use strict";

function countLiteral(text, token) {
  let count = 0;
  let offset = 0;
  while ((offset = text.indexOf(token, offset)) !== -1) {
    count += 1;
    offset += token.length;
  }
  return count;
}

cindy.onHostMessage(async function (msg) {
  if (msg.type !== "tool-call") return;
  if (msg.tool === "deliver_translation_excel" || msg.tool === "deliver_proofread_excel") {
    await deliverExcel(msg);
    return;
  }
  if (msg.tool !== "check_protected_tokens") {
    await cindy.send({type:"tool-result",callId:msg.callId,ok:false,errorCode:"UNSUPPORTED_TOOL",message:"请使用 ghost_info 返回的翻译交付、校对交付或固定片段检查工具。"});
    return;
  }
  const args = msg.args || {};
  if (typeof args.source !== "string" || typeof args.translation !== "string" ||
      args.source.length > 100000 || args.translation.length > 100000 ||
      !Array.isArray(args.tokens) || args.tokens.length < 1 || args.tokens.length > 64 ||
      args.tokens.some(token => typeof token !== "string" || token.length < 1 || token.length > 512) ||
      new Set(args.tokens).size !== args.tokens.length) {
    await cindy.send({type:"tool-result",callId:msg.callId,ok:false,errorCode:"INVALID_INPUT",message:"请提供不超过 100000 字符的原文和译文，以及 1 至 64 个不重复的非空保护片段，每个最多 512 字符。"});
    return;
  }
  const checks = args.tokens.map(token => {
    const sourceCount = countLiteral(args.source, token);
    const translationCount = countLiteral(args.translation, token);
    const status = sourceCount === 0 ? "not_in_source" :
      translationCount < sourceCount ? "missing" :
      translationCount > sourceCount ? "extra" : "same_count";
    return {token, sourceCount, translationCount, status};
  });
  await cindy.send({type:"tool-result",callId:msg.callId,ok:true,result:{
    checks,
    issues:checks.filter(item => item.status !== "same_count"),
    scope:"仅比较指定片段的逐字出现次数；未检查语义、格式语法、嵌套、位置或未指定片段。"
  }});
});

// A deliberately small OOXML writer for text-only localization tables.
// Runs inside Cindy's browser sandbox; no Node, Python, network or dependencies.
function exportShape(mode) {
  return mode === "translation"
    ? {headers:["原文", "译文"], keys:["source", "translation"], sheet:"翻译结果"}
    : {headers:["原文", "修改前译文", "修改后译文", "修改原因"], keys:["source", "before", "after", "reason"], sheet:"校对结果"};
}

function validateRows(mode, args) {
  if (!args || !Array.isArray(args.rows) || args.rows.length < 1 || args.rows.length > 1000) {
    throw new Error("请提供 1 至 1000 条已完成翻译或校对的 rows；更多条目请分批交付。");
  }
  const shape = exportShape(mode);
  let characters = 0;
  args.rows.forEach(function (row, index) {
    if (!row || typeof row !== "object" || Array.isArray(row) ||
        Object.keys(row).length !== shape.keys.length ||
        !shape.keys.every(key => Object.prototype.hasOwnProperty.call(row, key))) {
      throw new Error("第 " + (index + 1) + " 行字段必须是 " + shape.keys.join(", ") + "。");
    }
    shape.keys.forEach(function (key) {
      const value = row[key];
      if (typeof value !== "string" || value.length > 32767) {
        throw new Error("第 " + (index + 1) + " 行 " + key + " 必须是最多 32767 个 UTF-16 单位的原样字符串；禁止自动截断。");
      }
      for (const ch of value) {
        const n = ch.codePointAt(0);
        if ((n < 32 && n !== 9 && n !== 10 && n !== 13) ||
            (n >= 0xD800 && n <= 0xDFFF) || n === 0xFFFE || n === 0xFFFF) {
          throw new Error("第 " + (index + 1) + " 行 " + key + " 含 Excel XML 不能原样保存的字符；请先明确处理方式。");
        }
      }
      characters += value.length;
    });
    if (mode === "proofread" && !row.reason.trim()) {
      throw new Error("第 " + (index + 1) + " 行缺少修改原因；未修改时填写“无需修改”。");
    }
  });
  if (characters > 1000000) throw new Error("本批文本超过 100 万字符，请分批交付，不能省略条目。");
  return shape;
}

function xmlText(value) {
  return value.replace(/&/g, "&amp;").replace(/</g, "&lt;")
    .replace(/>/g, "&gt;").replace(/"/g, "&quot;")
    .replace(/'/g, "&apos;").replace(/\r/g, "&#13;");
}

function spreadsheetText(value) {
  // Escape literal OOXML escape sequences so Excel does not reinterpret text.
  return xmlText(value.replace(/_x[0-9A-Fa-f]{4}_/g, token => "_x005F_" + token.slice(1)));
}

function zipStore(entries) {
  const encoder = new TextEncoder();
  const parts = [];
  const directory = [];
  let offset = 0;
  const table = new Uint32Array(256);
  for (let n = 0; n < 256; n += 1) {
    let c = n;
    for (let bit = 0; bit < 8; bit += 1) c = (c & 1) ? 0xEDB88320 ^ (c >>> 1) : c >>> 1;
    table[n] = c >>> 0;
  }
  entries.forEach(function (entry) {
    const name = encoder.encode(entry[0]);
    const data = encoder.encode(entry[1]);
    let crc = 0xFFFFFFFF;
    for (let i = 0; i < data.length; i += 1) crc = table[(crc ^ data[i]) & 255] ^ (crc >>> 8);
    crc = (crc ^ 0xFFFFFFFF) >>> 0;
    const local = new Uint8Array(30 + name.length);
    const lv = new DataView(local.buffer);
    lv.setUint32(0, 0x04034B50, true);
    lv.setUint16(4, 20, true);
    lv.setUint16(6, 0x0800, true);
    lv.setUint16(12, 33, true); // 1980-01-01, a valid DOS ZIP date.
    lv.setUint32(14, crc, true);
    lv.setUint32(18, data.length, true);
    lv.setUint32(22, data.length, true);
    lv.setUint16(26, name.length, true);
    local.set(name, 30);
    parts.push(local, data);
    const central = new Uint8Array(46 + name.length);
    const cv = new DataView(central.buffer);
    cv.setUint32(0, 0x02014B50, true);
    cv.setUint16(4, 20, true);
    cv.setUint16(6, 20, true);
    cv.setUint16(8, 0x0800, true);
    cv.setUint16(14, 33, true);
    cv.setUint32(16, crc, true);
    cv.setUint32(20, data.length, true);
    cv.setUint32(24, data.length, true);
    cv.setUint16(28, name.length, true);
    cv.setUint32(42, offset, true);
    central.set(name, 46);
    directory.push(central);
    offset += local.length + data.length;
  });
  const directorySize = directory.reduce((total, part) => total + part.length, 0);
  const end = new Uint8Array(22);
  const ev = new DataView(end.buffer);
  ev.setUint32(0, 0x06054B50, true);
  ev.setUint16(8, entries.length, true);
  ev.setUint16(10, entries.length, true);
  ev.setUint32(12, directorySize, true);
  ev.setUint32(16, offset, true);
  const output = new Uint8Array(offset + directorySize + end.length);
  let cursor = 0;
  parts.concat(directory, [end]).forEach(part => {output.set(part, cursor); cursor += part.length;});
  if (output.length > 12 * 1024 * 1024) throw new Error("Excel 超过本次文件大小限制，请将条目分批交付。");
  return output;
}

function makeWorkbook(mode, args) {
  const shape = validateRows(mode, args);
  const ns = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
  const rel = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
  const pre = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';
  const strings = [];
  const indexes = new Map();
  let stringCount = 0;
  const range = "A1:" + String.fromCharCode(64 + shape.keys.length) + (args.rows.length + 1);
  function rowXml(values, index) {
    const cells = values.map(function (value, column) {
      if (!indexes.has(value)) {indexes.set(value, strings.length); strings.push(value);}
      stringCount += 1;
      const position = String.fromCharCode(65 + column) + index;
      return '<c r="' + position + '" s="' + (index === 1 ? 1 : 2) + '" t="s"><v>' + indexes.get(value) + '</v></c>';
    });
    return '<row r="' + index + '">' + cells.join("") + '</row>';
  }
  const sheetRows = [rowXml(shape.headers, 1)].concat(args.rows.map((row, index) => rowXml(shape.keys.map(key => row[key]), index + 2)));
  const cols = shape.keys.map((key, index) => '<col min="' + (index + 1) + '" max="' + (index + 1) + '" width="' + (key === "reason" ? 48 : 60) + '" customWidth="1"/>').join("");
  const sheet = pre + '<worksheet xmlns="' + ns + '"><dimension ref="' + range + '"/>' +
    '<sheetViews><sheetView workbookViewId="0"><pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/><selection pane="bottomLeft" activeCell="A2" sqref="A2"/></sheetView></sheetViews>' +
    '<sheetFormatPr defaultRowHeight="18"/><cols>' + cols + '</cols><sheetData>' + sheetRows.join("") + '</sheetData><autoFilter ref="' + range + '"/></worksheet>';
  const shared = pre + '<sst xmlns="' + ns + '" count="' + stringCount + '" uniqueCount="' + strings.length + '">' +
    strings.map(value => '<si><t xml:space="preserve">' + spreadsheetText(value) + '</t></si>').join("") + '</sst>';
  const styles = pre + '<styleSheet xmlns="' + ns + '">' +
    '<fonts count="2"><font><sz val="11"/><name val="Calibri"/></font><font><b/><sz val="11"/><color rgb="FFFFFFFF"/><name val="Calibri"/></font></fonts>' +
    '<fills count="3"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill><fill><patternFill patternType="solid"><fgColor rgb="FF28566B"/><bgColor indexed="64"/></patternFill></fill></fills>' +
    '<borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>' +
    '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>' +
    '<cellXfs count="3"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>' +
    '<xf numFmtId="49" fontId="1" fillId="2" borderId="0" xfId="0" applyAlignment="1" applyNumberFormat="1"><alignment vertical="top" wrapText="1"/></xf>' +
    '<xf numFmtId="49" fontId="0" fillId="0" borderId="0" xfId="0" applyAlignment="1" applyNumberFormat="1"><alignment vertical="top" wrapText="1"/></xf></cellXfs>' +
    '<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles></styleSheet>';
  const entries = [
    ['[Content_Types].xml', pre + '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/><Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/></Types>'],
    ['_rels/.rels', pre + '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="' + rel + '/officeDocument" Target="xl/workbook.xml"/></Relationships>'],
    ['xl/workbook.xml', pre + '<workbook xmlns="' + ns + '" xmlns:r="' + rel + '"><sheets><sheet name="' + shape.sheet + '" sheetId="1" r:id="rId1"/></sheets></workbook>'],
    ['xl/_rels/workbook.xml.rels', pre + '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="' + rel + '/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="rId2" Type="' + rel + '/styles" Target="styles.xml"/><Relationship Id="rId3" Type="' + rel + '/sharedStrings" Target="sharedStrings.xml"/></Relationships>'],
    ['xl/styles.xml', styles], ['xl/sharedStrings.xml', shared], ['xl/worksheets/sheet1.xml', sheet]
  ];
  return {bytes:zipStore(entries), shape, rowCount:args.rows.length};
}

function bytesBase64(bytes) {
  const chunks = [];
  for (let start = 0; start < bytes.length; start += 32768) {
    chunks.push(String.fromCharCode.apply(null, bytes.subarray(start, start + 32768)));
  }
  return btoa(chunks.join(""));
}

async function deliverExcel(msg) {
  const mode = msg.tool === "deliver_translation_excel" ? "translation" : "proofread";
  let workbook;
  try {
    workbook = makeWorkbook(mode, msg.args);
  } catch (error) {
    await cindy.send({type:"tool-result",callId:msg.callId,ok:false,errorCode:"EXCEL_INPUT_INVALID",message:error.message});
    return;
  }
  const filename = mode + "-" + Date.now().toString(36) + "-" + Math.random().toString(36).slice(2, 10) + ".xlsx";
  const relativePath = "out/" + filename;
  let saved;
  try {
    saved = await cindy.fs({op:"write",root:"workdir",path:relativePath,encoding:"base64",content:bytesBase64(workbook.bytes),callId:msg.callId});
  } catch (error) {
    await cindy.send({type:"tool-result",callId:msg.callId,ok:false,errorCode:"EXCEL_WRITE_FAILED",message:"写入 Excel 未成功：" + error.message + "。请检查工作目录权限；不要把文本答复当作文件交付。"});
    return;
  }
  if (!saved || !saved.ok) {
    await cindy.send({type:"tool-result",callId:msg.callId,ok:false,errorCode:"EXCEL_WRITE_FAILED",message:"写入 Excel 未成功：" + (saved && saved.message || "未获得保存成功回执") + "。请解决写入问题后再交付文件。"});
    return;
  }
  await cindy.send({type:"tool-result",callId:msg.callId,ok:true,result:{
    format:"xlsx",filename,relative_path:relativePath,host_path:saved.path || relativePath,
    row_count:workbook.rowCount,columns:workbook.shape.headers,bytes:workbook.bytes.length,
    file_link:"[下载" + workbook.shape.sheet + " Excel](" + relativePath + ")",
    delivery:"Excel 已实际写入当前会话工作目录。最终回复请提供此文件的可点击链接；可结合当前工作目录转成绝对路径，不要仅复述译文。"
  }});
}
