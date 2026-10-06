'use strict';

// 入力ファイルはブラウザのメモリ上でだけ処理し、外部送信も保存(localStorage等)もしない
// 4種類の入力がそろったら、タブごとに各インポート用データを作る

const CONVERTERS = [HOURLY, PAYROLL];
const $ = (id) => document.getElementById(id);
// kind → { fileName, rows }
const sources = {};
// メイン時給にする条件を画面で選んだ結果: 社員コード → 店舗ID
const choices = new Map();

function fillMessage(box, { message, details }) {
  box.querySelector('.message').textContent = message;
  box.querySelector('ul').replaceChildren(...(details || []).map((d) => {
    const li = document.createElement('li');
    li.textContent = d;
    return li;
  }));
  box.hidden = false;
}

const toConvertError = (err) => (err instanceof ConvertError ? err : new Error('変換できませんでした'));

function renderSources() {
  $('sources').replaceChildren(...SOURCE_KINDS.map(({ key, label }) => {
    const li = document.createElement('li');
    const name = document.createElement('span');
    const file = document.createElement('span');
    name.textContent = label;
    file.className = 'source-file';
    file.textContent = sources[key]?.fileName || '未選択';
    li.classList.toggle('ready', Boolean(sources[key]));
    li.append(name, file);
    return li;
  }));
}

function renderPreview(view, conv, result) {
  const { outRows } = result;
  const header = result.header || conv.header;
  // 社員コードの列(従業員マスタは先頭列が姓のため見出しで探す)
  const empCol = Math.max(0, header.map(normalizeHeader).indexOf('社員コード'));
  const head = document.createElement('tr');
  for (const [i, name] of header.entries()) {
    const th = document.createElement('th');
    th.textContent = conv.previewLabels[i] || name;
    head.append(th);
  }
  view.querySelector('thead').replaceChildren(head);

  // 従業員ごとに背景を交互にし、highlight 対象の行は色を変える。
  // チェックボックスの絞り込みは従業員単位(対象行が1つでもある人の全行を残す)
  const highlight = conv.highlighter(outRows, result);
  const flaggedEmps = new Set(outRows.filter(highlight).map((r) => r[empCol]));
  let band = false;
  const body = document.createDocumentFragment();
  outRows.forEach((row, n) => {
    const tr = document.createElement('tr');
    if (n > 0 && row[empCol] !== outRows[n - 1][empCol]) {
      band = !band;
      tr.classList.add('emp-start');
    }
    tr.dataset.emp = row[empCol];
    if (flaggedEmps.has(row[empCol])) tr.dataset.flagged = '';
    if (band) tr.classList.add('band');
    if (highlight(row)) tr.classList.add('exception');
    const changedCols = result.changed?.get(row);
    row.forEach((v, i) => {
      const td = document.createElement('td');
      td.textContent = v ?? '';
      if (changedCols?.has(i)) td.classList.add('changed');
      tr.append(td);
    });
    body.append(tr);
  });
  view.querySelector('tbody').replaceChildren(body);
}

// メイン時給を決められない人ごとに、どの店舗の条件にするかのラジオボタンを出す
function renderChoices(choosable) {
  $('choices').hidden = !choosable.length;
  $('choice-list').replaceChildren(...choosable.map(({ emp, stores, chosen, main }) => {
    const fieldset = document.createElement('fieldset');
    fieldset.classList.toggle('pending', !chosen);
    const legend = document.createElement('legend');
    legend.textContent = emp;
    fieldset.append(legend);
    for (const s of stores) {
      const label = document.createElement('label');
      const radio = document.createElement('input');
      radio.type = 'radio';
      radio.name = `main-${emp}`;
      radio.checked = chosen && s === main;
      radio.addEventListener('change', () => {
        choices.set(emp, s.storeId);
        renderOutputs();
      });
      const text = document.createElement('span');
      text.textContent = `${s.store}  ${describeCond(s)}`;
      label.append(radio, text);
      fieldset.append(label);
    }
    return fieldset;
  }));
}

function buildPanel(conv, rowsByKind, planned) {
  const view = $('result-template').content.firstElementChild.cloneNode(true);
  view.id = `panel-${conv.key}`;
  try {
    const result = conv.convert(rowsByKind, planned);
    const { outRows, empCount, stats, notice } = result;
    view.querySelector('.emp-count').textContent = `${empCount}名`;
    view.querySelector('.out-count').textContent = `${outRows.length}行`;
    view.querySelector('.breakdown').textContent = Object.entries(stats)
      .filter(([, v]) => v)
      .map(([k, v]) => `${k} ${v}名`)
      .join(' / ');
    if (notice) fillMessage(view.querySelector('.notice'), notice);
    renderPreview(view, conv, result);

    const filter = view.querySelector('.filter');
    const onlyFlagged = view.querySelector('.only-flagged input');
    view.querySelector('.only-flagged span').textContent = conv.highlightLabel;
    const applyFilter = () => {
      const q = filter.value.trim().toUpperCase();
      for (const tr of view.querySelector('tbody').rows) {
        tr.hidden = (q !== '' && !tr.dataset.emp.toUpperCase().includes(q))
          || (onlyFlagged.checked && !('flagged' in tr.dataset));
      }
    };
    filter.addEventListener('input', applyFilter);
    onlyFlagged.addEventListener('change', applyFilter);
    // メイン時給の選択が残っている間はダウンロードできない
    view.querySelector('.download').disabled = planned.pending > 0;
    view.querySelector('.download-hint').hidden = planned.pending === 0;
    view.querySelector('.download').addEventListener('click', () => {
      XLSX.writeFile(buildWorkbook(XLSX, { ...conv, ...result }), outputFileName(conv.filePrefix));
    });
  } catch (err) {
    console.error(err);
    view.querySelector('.result').hidden = true;
    fillMessage(view.querySelector('.error'), toConvertError(err));
  }
  return view;
}

function selectTab(key) {
  for (const tab of $('tabs').children) tab.setAttribute('aria-selected', tab.dataset.key === key);
  for (const panel of $('panels').children) panel.hidden = panel.id !== `panel-${key}`;
}

function renderOutputs() {
  const selected = document.querySelector('.tab[aria-selected="true"]')?.dataset.key;
  $('tabs').replaceChildren();
  $('panels').replaceChildren();
  $('outputs').hidden = true;
  $('choices').hidden = true;
  $('plan-error').hidden = true;
  if (!SOURCE_KINDS.every(({ key }) => sources[key])) return;

  const rowsByKind = Object.fromEntries(SOURCE_KINDS.map(({ key }) => [key, sources[key].rows]));
  let planned;
  try {
    planned = buildPlans(rowsByKind, choices);
  } catch (err) {
    console.error(err);
    fillMessage($('plan-error'), toConvertError(err));
    return;
  }
  renderChoices(planned.choosable);

  for (const conv of CONVERTERS) {
    const tab = document.createElement('button');
    tab.type = 'button';
    tab.role = 'tab';
    tab.className = 'tab';
    tab.dataset.key = conv.key;
    tab.textContent = conv.label;
    tab.addEventListener('click', () => selectTab(conv.key));
    $('tabs').append(tab);
    $('panels').append(buildPanel(conv, rowsByKind, planned));
  }
  selectTab(selected || CONVERTERS[0].key);
  $('outputs').hidden = false;
}

// 同じ種類のファイルを選び直したら差し替える
async function handleFiles(files) {
  $('error').hidden = true;
  const errors = [];
  for (const file of files) {
    try {
      const { kind, rows } = identifySource(XLSX, file.name, await file.arrayBuffer());
      sources[kind] = { fileName: file.name, rows };
    } catch (err) {
      console.error(err);
      errors.push(toConvertError(err).message);
    }
  }
  if (errors.length) fillMessage($('error'), { message: '読み込めないファイルがあります', details: errors });
  renderSources();
  renderOutputs();
}

const drop = $('drop');
const input = drop.querySelector('input');
input.addEventListener('change', () => {
  handleFiles([...input.files]);
  input.value = '';
});
drop.addEventListener('dragover', (e) => {
  e.preventDefault();
  drop.classList.add('over');
});
drop.addEventListener('dragleave', () => drop.classList.remove('over'));
drop.addEventListener('drop', (e) => {
  e.preventDefault();
  drop.classList.remove('over');
  handleFiles([...e.dataTransfer.files]);
});

renderSources();
