'use strict';

// 給与データはブラウザのメモリ上でだけ処理し、外部送信も保存(localStorage等)もしない

const $ = (id) => document.getElementById(id);
const yen = (n) => n.toLocaleString('ja-JP');
const signedYen = (n) => (n > 0 ? '+' : '') + yen(n);

// sort: 並び替え中の列 { id, dir }(dir は 1=昇順 / -1=降順)。null なら社員コード順
// errors: ファイルごとの読み込みエラー。もう一方のファイルを読み込んでも消えないよう別々に持つ
const state = { bugyo: null, arrow: null, outcome: null, sort: null, errors: { bugyo: null, arrow: null } };
const FILE_LABELS = { bugyo: '奉行の給与データ', arrow: 'ARROWの給与データ' };

function fillMessage(box, { message, details }) {
  box.querySelector('.message').textContent = message;
  box.querySelector('ul').replaceChildren(...(details || []).map((d) => {
    const li = document.createElement('li');
    li.textContent = d;
    return li;
  }));
  box.hidden = false;
}

function renderErrors() {
  const failed = Object.entries(state.errors).filter(([, err]) => err);
  for (const key of Object.keys(state.errors)) $(`drop-${key}`).classList.toggle('failed', !!state.errors[key]);
  $('error').hidden = true;
  if (!failed.length) return;
  $('result').hidden = true;
  fillMessage($('error'), {
    message: failed.map(([key, err]) => `${FILE_LABELS[key]}: ${err.message}`).join(' / '),
    details: failed.flatMap(([, err]) => err.details || []),
  });
}

function setupDropZone(zone, onFile) {
  const input = zone.querySelector('input');
  input.addEventListener('change', () => {
    onFile(input.files[0]);
    input.value = '';
  });
  zone.addEventListener('dragover', (e) => {
    e.preventDefault();
    zone.classList.add('over');
  });
  zone.addEventListener('dragleave', () => zone.classList.remove('over'));
  zone.addEventListener('drop', (e) => {
    e.preventDefault();
    zone.classList.remove('over');
    onFile(e.dataTransfer.files[0]);
  });
}

const sheetRows = (ws) => XLSX.utils.sheet_to_json(ws, { header: 1, raw: true, defval: null });

function readBugyoFile(data) {
  const wb = XLSX.read(data, { type: 'array' });
  return readBugyo(wb.SheetNames.map((name) => ({ name, rows: sheetRows(wb.Sheets[name]) })));
}

// ARROWのCSVはShift_JIS。UTF-8で読めるファイルならそのまま読む
function readArrowFile(data) {
  let text;
  try {
    text = new TextDecoder('utf-8', { fatal: true }).decode(data);
  } catch {
    text = new TextDecoder('shift_jis').decode(data);
  }
  const wb = XLSX.read(text.replace(/^﻿/, ''), { type: 'string', raw: true });
  return readArrow(sheetRows(wb.Sheets[wb.SheetNames[0]]));
}

function loader(key, read) {
  return async (file) => {
    if (!file) return;
    const zone = $(`drop-${key}`);
    zone.querySelector('.drop-sub').textContent = file.name;
    zone.classList.add('loaded');
    try {
      state[key] = read(await file.arrayBuffer());
      state.errors[key] = null;
    } catch (err) {
      console.error(err);
      state[key] = null;
      state.errors[key] = err instanceof ReconcileError ? err : { message: '読み込めませんでした' };
    }
    renderErrors();
    run();
  };
}

function tolerance() {
  const v = Number($('tolerance').value);
  return Number.isFinite(v) && v > 0 ? v : 10;
}

function run() {
  if (!state.bugyo || !state.arrow) return;
  const outcome = reconcile(state.bugyo, state.arrow, tolerance());
  state.outcome = outcome;
  const { results, mismatchCount, bugyoOnly, arrowOnly, mergedCount } = outcome;
  const allOk = results.filter((r) => r.items.every((it) => it.ok)).length;
  $('matched').textContent = `${results.length}名`;
  $('all-ok').textContent = `${allOk}名`;
  $('ng').textContent = `${results.length - allOk}名`;

  const notes = [];
  if (arrowOnly.length) notes.push(`ARROWのみ ${arrowOnly.length}名は対象外`);
  if (mergedCount) notes.push(`奉行で複数行の ${mergedCount}名は合算`);
  $('breakdown').textContent = notes.join(' / ');

  $('notice').hidden = true;
  if (bugyoOnly.length) {
    fillMessage($('notice'), {
      message: `ARROWにいない社員がいます(${bugyoOnly.length}名)`,
      details: bugyoOnly.map(({ emp, name }) => (name ? `${emp} ${name}` : emp)),
    });
  }

  const select = $('item-filter');
  const current = select.value;
  select.replaceChildren(new Option(`すべての項目`, ''), ...ITEMS.map((it) => (
    new Option(`${it.label} (${mismatchCount[it.key]})`, it.key)
  )));
  select.value = current;

  // 固定列の幅を測るので、表示してから一覧を作る
  $('result').hidden = false;
  renderRows();
}

// 一覧・Excelの列。グループ(見出し1段目)ごとに列(見出し2段目)を持つ。
// value は Excel に書く値・並び替えのキー、text は画面に出す文字列。item があるグループは不一致なら色を付ける
// 社員名はどちらのデータにも入っていなければ列ごと出さない
function buildColumns(itemKey) {
  const hasName = state.outcome.results.some((r) => r.name);
  const groups = [{
    label: '',
    cols: [
      { label: '社員コード', pin: true, value: (r) => r.emp },
      ...(hasName ? [{ label: '社員名', pin: true, ellipsis: true, value: (r) => r.name }] : []),
      { label: '店舗', value: (r) => r.store },
      { label: '奉行シート', value: (r) => r.area },
    ],
  }];
  for (const c of groups[0].cols) c.id = c.label;

  ITEMS.forEach((it, i) => {
    if (itemKey && it.key !== itemKey) return;
    const get = (r) => r.items[i];
    const bugyoNames = [...it.bugyo, ...(it.bugyoLess || []).map((n) => `−${n}`)];
    const cols = [];
    // 元の列が1つだけなら合計列は出さない
    if (it.arrow.length > 1) {
      it.arrow.forEach((n, j) => cols.push({ label: n, side: 'arrow', value: (r) => get(r).arrowParts[j] }));
      cols.push({ label: 'ARROW計', side: 'arrow', total: true, value: (r) => get(r).arrow });
    } else {
      cols.push({ label: 'ARROW', side: 'arrow', total: true, value: (r) => get(r).arrow });
    }
    if (bugyoNames.length > 1) {
      bugyoNames.forEach((n, j) => cols.push({ label: n, side: 'bugyo', value: (r) => get(r).bugyoParts[j] }));
      cols.push({ label: '奉行計', side: 'bugyo', total: true, value: (r) => get(r).bugyo });
    } else {
      cols.push({ label: '奉行', side: 'bugyo', total: true, value: (r) => get(r).bugyo });
    }
    cols.push({ label: '差', side: 'diff', value: (r) => get(r).diff, text: (r) => signedYen(get(r).diff) });
    for (const c of cols) {
      c.id = `${it.key}:${c.label}`;
      if (!c.text) c.text = (r) => yen(c.value(r));
    }
    groups.push({ label: it.label, item: get, cols });
  });
  return groups;
}

const cellText = (c, r) => (c.text ? c.text(r) : String(c.value(r) ?? ''));

function th(text, attrs = {}) {
  const el = document.createElement('th');
  el.textContent = text;
  Object.assign(el, attrs);
  return el;
}

// 見出しをクリックすると並び替える。同じ列をもう一度押すと逆順。金額の列は降順から始める
function sortableTh(c, attrs = {}) {
  const cell = th(c.label, attrs);
  cell.classList.add('sortable');
  if (state.sort?.id === c.id) cell.ariaSort = state.sort.dir > 0 ? 'ascending' : 'descending';
  cell.addEventListener('click', () => {
    state.sort = state.sort?.id === c.id
      ? { id: c.id, dir: -state.sort.dir }
      : { id: c.id, dir: c.side ? -1 : 1 };
    renderRows();
  });
  return cell;
}

function sortRows(rows, groups) {
  const col = state.sort && groups.flatMap((g) => g.cols).find((c) => c.id === state.sort.id);
  if (!col) return rows;
  const { dir } = state.sort;
  const compare = (a, b) => (typeof a === 'number' && typeof b === 'number'
    ? a - b
    : String(a ?? '').localeCompare(String(b ?? ''), 'ja'));
  // 同じ値の中では社員コード順のまま(安定ソート)
  return [...rows].sort((x, y) => dir * compare(col.value(x), col.value(y)));
}

function renderRows() {
  const itemKey = $('item-filter').value;
  const q = $('emp-filter').value.trim().toUpperCase();
  const onlyNg = $('only-ng').checked;
  const groups = buildColumns(itemKey);

  // 1段目はグループ名。先頭グループ(社員コード等)は2段分の高さにする
  const top = document.createElement('tr');
  const sub = document.createElement('tr');
  for (const [gi, g] of groups.entries()) {
    if (gi === 0) {
      g.cols.forEach((c, ci) => {
        const cell = sortableTh(c, { rowSpan: 2 });
        if (c.pin) cell.classList.add('pin', `pin-${ci}`);
        top.append(cell);
      });
      continue;
    }
    top.append(th(g.label, { colSpan: g.cols.length, className: 'group' }));
    g.cols.forEach((c, ci) => {
      const cell = sortableTh(c);
      if (ci === 0) cell.classList.add('group-start');
      if (c.side) cell.classList.add(c.side);
      if (c.total) cell.classList.add('total');
      sub.append(cell);
    });
  }
  $('head').replaceChildren(top, sub);

  const body = document.createDocumentFragment();
  for (const r of sortRows(state.outcome.results, groups)) {
    if (q && !r.emp.toUpperCase().includes(q)) continue;
    const ngItems = r.items.filter((it) => !it.ok && (!itemKey || it.key === itemKey));
    if (onlyNg && !ngItems.length) continue;
    const tr = document.createElement('tr');
    if (ngItems.length) tr.classList.add('has-ng');
    for (const [gi, g] of groups.entries()) {
      const ng = g.item && !g.item(r).ok;
      g.cols.forEach((c, ci) => {
        const td = document.createElement('td');
        if (c.ellipsis) {
          // 長い社員名は省略し、全体はマウスを乗せると見られるようにする
          const span = document.createElement('span');
          span.className = 'ellipsis';
          span.textContent = cellText(c, r);
          span.title = span.textContent;
          td.append(span);
        } else {
          td.textContent = cellText(c, r);
        }
        if (c.pin) td.classList.add('pin', `pin-${ci}`);
        if (gi > 0 && ci === 0) td.classList.add('group-start');
        if (c.side) td.classList.add('num');
        if (c.total) td.classList.add('total');
        if (ng) td.classList.add(c.side === 'diff' ? 'ng-diff' : 'ng');
        tr.append(td);
      });
    }
    body.append(tr);
  }
  $('rows').replaceChildren(body);
  $('empty').hidden = $('rows').rows.length > 0;

  // 固定する2列目(社員名)は、1列目(社員コード)の幅のぶん右にずらして固定する
  const firstPin = $('head').querySelector('.pin-0');
  $('head').closest('table').style.setProperty('--pin-1-left', `${firstPin.offsetWidth}px`);
}

function buildWorkbook({ results, bugyoOnly, arrowOnly }) {
  // 突合結果シートは画面の一覧と同じ列。見出しは「グループ｜列名」の1段にする
  const groups = buildColumns('');
  const cols = groups.flatMap((g) => g.cols.map((c) => ({ ...c, header: g.label ? `${g.label}｜${c.label}` : c.label })));
  const fixed = groups[0].cols.length;
  const hasName = cols[1].label === '社員名';
  const person = (r) => (hasName ? [r.emp, r.name] : [r.emp]);
  const personHeader = hasName ? ['社員コード', '社員名'] : ['社員コード'];
  const wide = [
    [...cols.slice(0, fixed).map((c) => c.header), '不一致項目', ...cols.slice(fixed).map((c) => c.header)],
    ...results.map((r) => {
      const values = cols.map((c) => c.value(r));
      const ngLabels = r.items.filter((it) => !it.ok).map((it) => it.label).join('、');
      return [...values.slice(0, fixed), ngLabels, ...values.slice(fixed)];
    }),
  ];
  const long = [
    [...personHeader, '店舗', '項目', 'ARROW', '奉行', '差'],
    ...results.flatMap((r) => r.items.filter((it) => !it.ok).map((it) => [...person(r), r.store, it.label, it.arrow, it.bugyo, it.diff])),
  ];
  const excluded = [
    [...personHeader, '理由'],
    ...bugyoOnly.map((e) => [...person(e), 'ARROWにいない']),
    ...arrowOnly.map((e) => [...person(e), '奉行にいない']),
  ];
  const personWidths = hasName ? [12, 14] : [12];

  const wb = XLSX.utils.book_new();
  for (const [name, aoa, widths] of [
    ['突合結果', wide, [...personWidths, 16, 10, 30, ...cols.slice(fixed).map(() => 12)]],
    ['不一致一覧', long, [...personWidths, 16, 14, 10, 10, 10]],
    ['対象外', excluded, [...personWidths, 14]],
  ]) {
    const ws = XLSX.utils.aoa_to_sheet(aoa);
    ws['!cols'] = widths.map((wch) => ({ wch }));
    ws['!autofilter'] = { ref: XLSX.utils.encode_range({ s: { r: 0, c: 0 }, e: { r: aoa.length - 1, c: aoa[0].length - 1 } }) };
    XLSX.utils.book_append_sheet(wb, ws, name);
  }
  return wb;
}

function outputFileName(d = new Date()) {
  const p = (n) => String(n).padStart(2, '0');
  return `給与突合結果_${d.getFullYear()}${p(d.getMonth() + 1)}${p(d.getDate())}.xlsx`;
}

setupDropZone($('drop-bugyo'), loader('bugyo', readBugyoFile));
setupDropZone($('drop-arrow'), loader('arrow', readArrowFile));
$('tolerance').addEventListener('input', run);
$('item-filter').addEventListener('change', renderRows);
$('emp-filter').addEventListener('input', renderRows);
$('only-ng').addEventListener('change', renderRows);
$('download').addEventListener('click', () => XLSX.writeFile(buildWorkbook(state.outcome), outputFileName()));
