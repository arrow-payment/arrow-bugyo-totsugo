'use strict';

// 共通処理(ほっか時給更新ツールの common.js をもとにしたもの)

const OUTPUT_SHEET = 'インポート用';

class ConvertError extends Error {
  constructor(message, details = []) {
    super(message);
    this.details = details;
  }
}

const str = (v) => (v === null || v === undefined ? '' : String(v).trim());
// 見出しは空白(全角含む)を除き、全角かっこを半角に揃えて完全一致で探す
const normalizeHeader = (v) => str(v).replace(/\s/g, '').replace(/（/g, '(').replace(/）/g, ')');
// 店舗名の「ＪＲ」「JR」のような全角・半角の違いを吸収する
const normalizeName = (v) => str(v).normalize('NFKC').replace(/\s/g, '');

// columns: { key: 見出し名 } → { key: 列index }
function findColumns(headerRow, columns, fileLabel) {
  const names = (headerRow || []).map(normalizeHeader);
  const col = {};
  const missing = [];
  for (const [key, name] of Object.entries(columns)) {
    const i = names.indexOf(name);
    if (i < 0) missing.push(name);
    col[key] = i;
  }
  if (missing.length) throw new ConvertError(`${fileLabel}に必要な列が見つかりません`, missing);
  return col;
}

// 「04:30:00」「4:30」→ 分。Excelの時刻セル(1日=1の小数)にも対応
function parseMinutes(v) {
  if (typeof v === 'number') return Math.round(v < 1 ? v * 24 * 60 : v * 60);
  const m = str(v).match(/^(\d{1,2}):(\d{2})(?::(\d{2}))?$/);
  return m ? Number(m[1]) * 60 + Number(m[2]) : null;
}

// 分 → 「05:00」
function formatHHMM(minutes) {
  const p = (n) => String(n).padStart(2, '0');
  return `${p(Math.floor(minutes / 60))}:${p(minutes % 60)}`;
}

// XLSX: SheetJS。先頭のシートを、見出しが1行目・A列始まりになるよう読む
function readSheetRows(XLSX, data) {
  const wb = XLSX.read(data, { type: 'array' });
  const ws = wb.Sheets[wb.SheetNames[0]];
  if (!ws || !ws['!ref']) return [];
  const range = XLSX.utils.decode_range(ws['!ref']);
  range.s = { r: 0, c: 0 };
  return XLSX.utils.sheet_to_json(ws, { header: 1, raw: true, defval: null, range });
}

// CSVは文字コードがUTF-8かShift_JISのどちらか。値は文字列のまま読む
function readCsvRows(XLSX, data) {
  let text;
  try {
    text = new TextDecoder('utf-8', { fatal: true }).decode(data);
  } catch {
    text = new TextDecoder('shift_jis').decode(data);
  }
  const wb = XLSX.read(text.replace(/^﻿/, ''), { type: 'string', raw: true });
  return XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]], { header: 1, raw: true, defval: null });
}

function buildWorkbook(XLSX, { formatVersion, header, colWidths, outRows }) {
  const ws = XLSX.utils.aoa_to_sheet([['__format_version', formatVersion], header, ...outRows]);
  ws['!cols'] = colWidths.map((wch) => ({ wch }));
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, OUTPUT_SHEET);
  return wb;
}

function outputFileName(prefix, d = new Date()) {
  const p = (n) => String(n).padStart(2, '0');
  return `${prefix}_${d.getFullYear()}${p(d.getMonth() + 1)}${p(d.getDate())}.xlsx`;
}
