'use strict';

// 給与突合ロジック
//
// 奉行の支給控除項目一覧表(正解)と、ARROWの給与データCSVを社員コードで突き合わせ、
// 項目ごとに「差が許容誤差未満なら一致」と判定する。
//
// - 奉行: エリアごとのシート。見出し行の位置・列の並びはシートごとに違うので列名で探す。
//   店舗見出し行(【 …)・合計行は社員番号が空なので読み飛ばす。
//   同じ社員番号が複数行ある人(複数店舗勤務)は金額を合算する
//   列がない項目(BUGYO_OPTIONAL)は0とみなす
// - ARROW: 「合計」が3回出てくるので、最初の「合計」を支給合計として使う。
//   「残業手当(資格手当分)」は月によって列がないので、ないときは0とみなす
// - 片方にしかいない社員は突合対象外として件数・社員コードだけ出す
// - 社員名は ARROW の「名前」、空なら奉行の「氏名」を使う(どちらも空・列なしなら空欄)

class ReconcileError extends Error {
  constructor(message, details = []) {
    super(message);
    this.details = details;
  }
}

// 比較項目。arrow / bugyo に複数列あるときは合計して比べる。bugyoLess の列は奉行側から差し引く
const ITEMS = [
  // ARROWの基本給には奉行の日別手当・早朝手当にあたる分も含まれる(2026年8月支給分で確認)。
  // 違っていたら bugyo から '日別手当', '早朝手当' を外せば元の比べ方に戻る
  { key: 'base', label: '基本給+残業', arrow: ['基本給/役員報酬', '残業手当', '残業手当(資格手当分)'], bugyo: ['基本給', '残業手当', '日別手当', '早朝手当'] },
  { key: 'night', label: '深夜手当', arrow: ['深夜手当'], bugyo: ['深夜手当'] },
  { key: 'paid', label: '有給手当', arrow: ['有給手当'], bugyo: ['その他２'] },
  { key: 'commute', label: '通勤手当', arrow: ['通勤手当'], bugyo: ['通勤手当'] },
  { key: 'qual', label: '資格手当', arrow: ['資格手当'], bugyo: ['資格手当'] },
  { key: 'delivery', label: '宅配手当', arrow: ['宅配手当'], bugyo: ['宅配手当'] },
  // 奉行の「その他手当」は仮払いで、同額を「仮払い戻入」で控除している(差引支給額には影響しない)。
  // ARROWにはない項目なので総支給額からは除いて比べる
  { key: 'gross', label: '総支給額', arrow: ['合計'], bugyo: ['総支給金額'], bugyoLess: ['その他手当'] },
  { key: 'health', label: '健康保険', arrow: ['健康保険', '子ども子育て支援金'], bugyo: ['健康保険料'] },
  { key: 'care', label: '介護保険', arrow: ['介護保険'], bugyo: ['介護保険料'] },
  { key: 'pension', label: '厚生年金', arrow: ['厚生年金保険'], bugyo: ['厚生年金'] },
  { key: 'employment', label: '雇用保険', arrow: ['雇用保険'], bugyo: ['雇用保険'] },
  { key: 'incomeTax', label: '所得税', arrow: ['所得税'], bugyo: ['所得税'] },
  { key: 'residentTax', label: '住民税', arrow: ['住民税'], bugyo: ['住民税'] },
  { key: 'net', label: '差引支給額', arrow: ['差引支給額'], bugyo: ['差引支給額'] },
];
const ARROW_OPTIONAL = new Set(['残業手当(資格手当分)']);
// シートによって列がない項目(例: 山陰シートには早朝手当がない)
const BUGYO_OPTIONAL = new Set(['早朝手当', '日別手当']);
const BUGYO_EMP = '社員番号';
const BUGYO_NAME = '氏名';
const ARROW_NAME = '名前';
const ARROW_EMP = '社員コード';
const ARROW_STORE = '店舗';

const str = (v) => (v === null || v === undefined ? '' : String(v).trim());
// 見出しは空白(全角含む)を除き、全角かっこを半角に揃えて完全一致で探す
const normalizeHeader = (v) => str(v).replace(/\s/g, '').replace(/（/g, '(').replace(/）/g, ')');

// 見出し名 → 列index(同名の列は最初のもの)。optional にない列が欠けていたら missing に入れる
function findColumns(headerRow, names, optional = new Set()) {
  const header = (headerRow || []).map(normalizeHeader);
  const col = {};
  const missing = [];
  for (const name of names) {
    const i = header.indexOf(name);
    if (i < 0 && !optional.has(name)) missing.push(name);
    col[name] = i;
  }
  return { col, missing };
}

// 空欄は0。数値にできない値は null
function toAmount(v) {
  if (v === null || v === undefined || str(v) === '') return 0;
  if (typeof v === 'number') return v;
  const n = Number(str(v).replace(/,/g, ''));
  return Number.isFinite(n) ? n : null;
}

const uniq = (arrays) => [...new Set(arrays.flat())];

// sheets: [{ name, rows }] → Map(社員コード → { name, sheets, values: { 列名: 金額 } })
function readBugyo(sheets) {
  const names = uniq(ITEMS.map((it) => [...it.bugyo, ...(it.bugyoLess || [])]));
  const byEmp = new Map();
  const problems = [];
  let found = false;
  for (const { name, rows } of sheets) {
    const headerIdx = rows.slice(0, 30).findIndex((r) => (r || []).some((v) => normalizeHeader(v) === BUGYO_EMP));
    if (headerIdx < 0) continue;
    found = true;
    const { col, missing } = findColumns(rows[headerIdx], [BUGYO_EMP, BUGYO_NAME, ...names], new Set([...BUGYO_OPTIONAL, BUGYO_NAME]));
    if (missing.length) throw new ReconcileError(`奉行の「${name}」シートに必要な列が見つかりません`, missing);

    for (const r of rows.slice(headerIdx + 1)) {
      const emp = str(r[col[BUGYO_EMP]]);
      if (!emp || emp.startsWith('【')) continue;
      if (!byEmp.has(emp)) byEmp.set(emp, { name: '', sheets: [], values: Object.fromEntries(names.map((n) => [n, 0])) });
      const entry = byEmp.get(emp);
      entry.sheets.push(name);
      if (!entry.name && col[BUGYO_NAME] >= 0) entry.name = str(r[col[BUGYO_NAME]]);
      for (const n of names) {
        const v = toAmount(r[col[n]]);
        if (v === null) problems.push(`${name} ${emp}: ${n}が数値ではありません`);
        else entry.values[n] += v;
      }
    }
  }
  if (!found) throw new ReconcileError('奉行の支給控除項目一覧表ではないようです(「社員番号」の見出しがありません)');
  if (problems.length) throw new ReconcileError('奉行のデータに読めない金額があります', problems);
  if (!byEmp.size) throw new ReconcileError('奉行のデータに社員がいません');
  return byEmp;
}

// rows: CSVの全行(1行目が見出し) → Map(社員コード → { name, store, values })
function readArrow(rows) {
  const names = uniq(ITEMS.map((it) => it.arrow));
  const { col, missing } = findColumns(rows[0], [ARROW_EMP, ARROW_NAME, ARROW_STORE, ...names], new Set([...ARROW_OPTIONAL, ARROW_NAME]));
  if (missing.length) throw new ReconcileError('ARROWの給与データに必要な列が見つかりません', missing);

  const byEmp = new Map();
  const problems = [];
  const duplicates = [];
  for (const r of rows.slice(1)) {
    const emp = str(r[col[ARROW_EMP]]);
    if (!emp) continue;
    if (byEmp.has(emp)) {
      duplicates.push(emp);
      continue;
    }
    const values = {};
    for (const n of names) {
      const v = col[n] < 0 ? 0 : toAmount(r[col[n]]);
      if (v === null) problems.push(`${emp}: ${n}が数値ではありません`);
      values[n] = v ?? 0;
    }
    byEmp.set(emp, { name: col[ARROW_NAME] >= 0 ? str(r[col[ARROW_NAME]]) : '', store: str(r[col[ARROW_STORE]]), values });
  }
  if (duplicates.length) throw new ReconcileError('ARROWの給与データに同じ社員コードが複数あります', duplicates);
  if (problems.length) throw new ReconcileError('ARROWの給与データに読めない金額があります', problems);
  if (!byEmp.size) throw new ReconcileError('ARROWの給与データに社員がいません');
  return byEmp;
}

const sumOf = (values, names) => names.reduce((s, n) => s + values[n], 0);

// tolerance: 差の絶対値がこれ未満なら一致
function reconcile(bugyo, arrow, tolerance) {
  const results = [];
  for (const [emp, b] of bugyo) {
    const a = arrow.get(emp);
    if (!a) continue;
    const items = ITEMS.map((it) => {
      const arrowValue = sumOf(a.values, it.arrow);
      const bugyoValue = sumOf(b.values, it.bugyo) - sumOf(b.values, it.bugyoLess || []);
      const diff = arrowValue - bugyoValue;
      return {
        key: it.key,
        label: it.label,
        arrow: arrowValue,
        bugyo: bugyoValue,
        // 合計する前の元の列の値。bugyoLess の列は値を負にして持つ
        arrowParts: it.arrow.map((n) => a.values[n]),
        bugyoParts: [...it.bugyo.map((n) => b.values[n]), ...(it.bugyoLess || []).map((n) => -b.values[n])],
        diff,
        ok: Math.abs(diff) < tolerance,
      };
    });
    results.push({
      emp,
      name: a.name || b.name,
      store: a.store,
      area: uniq([b.sheets]).join('・'),
      merged: b.sheets.length > 1,
      items,
    });
  }
  results.sort((x, y) => x.emp.localeCompare(y.emp));

  const mismatchCount = Object.fromEntries(ITEMS.map((it) => [it.key, 0]));
  for (const r of results) for (const it of r.items) if (!it.ok) mismatchCount[it.key]++;
  return {
    results,
    mismatchCount,
    // 片方にしかいない社員: [{ emp, name }]
    bugyoOnly: [...bugyo].filter(([e]) => !arrow.has(e)).map(([emp, b]) => ({ emp, name: b.name })).sort((x, y) => x.emp.localeCompare(y.emp)),
    arrowOnly: [...arrow].filter(([e]) => !bugyo.has(e)).map(([emp, a]) => ({ emp, name: a.name })).sort((x, y) => x.emp.localeCompare(y.emp)),
    mergedCount: results.filter((r) => r.merged).length,
  };
}
