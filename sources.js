'use strict';

// 入力ファイルの判別・読み込みと、更新する人ごとの条件(メイン時給・例外時給)の決定
//
// 入力は4種類で、どのファイルかは中身(見出し)で判別する:
// - form    : 労働契約更新元データ(Googleフォームの回答。1行=1人×1店舗)
// - stores  : 店舗・部署一覧.csv(店舗名 → 店舗ID)
// - master  : 従業員マスタ(ARROWの従業員エクスポート。現在の登録)
// - hourly  : ARROWの時給エクスポート(現在の登録)

const SOURCE_KINDS = [
  { key: 'form', label: '労働契約更新元データ' },
  { key: 'stores', label: '店舗・部署一覧' },
  { key: 'master', label: '従業員マスタ' },
  { key: 'hourly', label: 'ARROWの時給' },
];

const FORM_COLUMNS = {
  timestamp: 'タイムスタンプ',
  store: '店舗名',
  emp: 'FPコード',
  renew: '契約更新の有無',
  wage: '時給(数字のみ)',
  sunHoliday: '諸手当[日祝手当]',
  weekendHoliday: '諸手当[土日祝手当]',
  early: '諸手当[早朝手当]',
};
const STORE_COLUMNS = { id: '店舗ID', name: '店舗名' };
const MASTER_COLUMNS = {
  emp: '社員コード',
  lastName: '姓',
  firstName: '名',
  lastKana: '姓(カナ)',
  firstKana: '名(カナ)',
  mainStore: '主所属部署/店舗ID',
  weekDays: '週所定労働日数',
  weekHours: '週所定労働時間',
};
const CURRENT_HOURLY_COLUMNS = {
  emp: '社員コード',
  wage: '時給',
  store: '対象部署/店舗ID',
  start: '開始時間',
  end: '終了時間',
};
const RENEW = '契約更新する';
const RETIRE = '契約更新しない(退職)';
// フォームの早朝手当の時間帯は旧元データと同じ 5:00〜8:00 とする
const EARLY_BAND = '5:00~8:00';

// ファイルの中身から種類を判別して { kind, rows } を返す
function identifySource(XLSX, fileName, data) {
  if (/\.csv$/i.test(fileName)) {
    const rows = readCsvRows(XLSX, data);
    const header = (rows[0] || []).map(normalizeHeader);
    if (header.includes('店舗ID') && header.includes('店舗名')) return { kind: 'stores', rows };
  } else {
    const rows = readSheetRows(XLSX, data);
    const first = (rows[0] || []).map(normalizeHeader);
    if (first.includes('FPコード')) return { kind: 'form', rows };
    if (first[0] === '__format_version') {
      const header = (rows[1] || []).map(normalizeHeader);
      if (header.includes('週所定労働時間') && header.includes('有給休暇全日使用単価')) return { kind: 'master', rows };
      if (header.includes('開始時間') && header.includes('時給')) return { kind: 'hourly', rows };
    }
  }
  throw new ConvertError(`${fileName} がどのファイルか判別できませんでした`);
}

// 「なし」→0、「50円」→50
function parseAllowance(v, emp, label) {
  const s = normalizeName(v);
  if (s === '' || s === 'なし') return 0;
  const m = s.match(/^(\d+)円$/);
  if (!m) throw new ConvertError(`${emp}: 想定外の${label}です(${s})`);
  return Number(m[1]);
}

function parseStores(rows) {
  const col = findColumns(rows[0], STORE_COLUMNS, '店舗・部署一覧');
  const byName = new Map();
  for (const r of rows.slice(1)) {
    if (str(r[col.name])) byName.set(normalizeName(r[col.name]), str(r[col.id]));
  }
  return byName;
}

// フォームの回答 → byEmp: 契約更新する人の { emp: [{ store, storeId, cond }] }
//                  answered: { emp: 回答があった店舗IDの集合(更新・退職とも) }
function parseForm(rows, storeIds) {
  const col = findColumns(rows[0], FORM_COLUMNS, '労働契約更新元データ');

  // 同じ人・同じ店舗で回答が複数あれば、タイムスタンプが一番新しいものを使う
  const latest = new Map();
  for (const r of rows.slice(1)) {
    const emp = str(r[col.emp]);
    if (!emp) continue;
    const key = `${emp}\t${normalizeName(r[col.store])}`;
    const prev = latest.get(key);
    if (!prev || r[col.timestamp] > prev[col.timestamp]) latest.set(key, r);
  }

  const byEmp = new Map();
  const answered = new Map();
  const unknownStores = new Set();
  for (const r of latest.values()) {
    const emp = str(r[col.emp]);
    const store = str(r[col.store]);
    const storeId = storeIds.get(normalizeName(store));
    if (!storeId) {
      unknownStores.add(store);
      continue;
    }
    if (!answered.has(emp)) answered.set(emp, new Set());
    answered.get(emp).add(storeId);

    const renew = normalizeName(r[col.renew]);
    if (renew === normalizeName(RETIRE)) continue;
    if (renew !== RENEW) throw new ConvertError(`${emp}: 想定外の契約更新の有無です(${renew})`);
    const wage = r[col.wage];
    if (typeof wage !== 'number' || wage <= 0) throw new ConvertError(`${emp}: 時給が不正です(${wage})`);

    const sunHoliday = parseAllowance(r[col.sunHoliday], emp, '日祝手当');
    const weekendHoliday = parseAllowance(r[col.weekendHoliday], emp, '土日祝手当');
    const early = parseAllowance(r[col.early], emp, '早朝手当');
    if (sunHoliday && weekendHoliday) {
      throw new ConvertError(`${emp}: 日祝手当と土日祝手当が両方入っています`);
    }
    // 時給ツールの条件と同じ形にそろえる(buildHourlyRows でそのまま行にできる)
    const cond = {
      base: wage,
      early,
      daily: sunHoliday || weekendHoliday,
      holiday: sunHoliday ? '日祝' : weekendHoliday ? '土日祝' : '',
      earlyBand: early ? EARLY_BAND : '',
    };

    if (!byEmp.has(emp)) byEmp.set(emp, []);
    byEmp.get(emp).push({ store, storeId, cond });
  }
  if (unknownStores.size) {
    throw new ConvertError('店舗・部署一覧にない店舗名があります', [...unknownStores]);
  }
  if (!byEmp.size) throw new ConvertError('契約更新する人がいません');
  return { byEmp, answered };
}

// 従業員マスタ → { emp: { name, kana, mainStore, weekDays, weekMinutes, row } }
// row はエクスポートの行そのまま(従業員の出力はこれを上書きして作る)
function parseMaster(rows) {
  const col = findColumns(rows[1], MASTER_COLUMNS, '従業員マスタ');
  const byEmp = new Map();
  const duplicates = [];
  for (const r of rows.slice(2)) {
    const emp = str(r[col.emp]);
    if (!emp) continue;
    if (byEmp.has(emp)) duplicates.push(emp);
    byEmp.set(emp, {
      // ARROWの時給の社員名・カナは姓と名を区切りなしでつなげた形
      name: str(r[col.lastName]) + str(r[col.firstName]) || null,
      kana: str(r[col.lastKana]) + str(r[col.firstKana]) || null,
      mainStore: str(r[col.mainStore]),
      weekDays: Number(str(r[col.weekDays])) || null,
      weekMinutes: parseMinutes(r[col.weekHours]),
      row: r,
    });
  }
  if (duplicates.length) throw new ConvertError('従業員マスタに同じ社員コードが複数あります', duplicates);
  return byEmp;
}

// ARROWの時給エクスポート → { emp: { exceptionStores, rows, storeRows } }
// rows: 全行の比較用キー / storeRows: 店舗ID → その店舗の例外時給の行(社員コード・氏名を除いた出力形式)
function parseCurrentHourly(rows) {
  const header = rows[1] || [];
  const col = findColumns(header, CURRENT_HOURLY_COLUMNS, 'ARROWの時給');
  const dayCols = DAYS.map((d) => header.map(normalizeHeader).indexOf(d === '祝' ? '祝日' : `${d}曜日`));
  const byEmp = new Map();
  for (const r of rows.slice(2)) {
    const emp = str(r[col.emp]);
    if (!emp) continue;
    if (!byEmp.has(emp)) byEmp.set(emp, { exceptionStores: new Set(), rows: [], storeRows: new Map() });
    const cur = byEmp.get(emp);
    const store = str(r[col.store]);
    // エクスポートの数値は「1,200」のようにカンマ区切りで出ることがあるので外してから数値にする
    const wage = str(r[col.wage]).replace(/,/g, '');
    const fragment = [/^\d+(\.\d+)?$/.test(wage) ? Number(wage) : wage, store || null,
      formatTime(r[col.start]), formatTime(r[col.end]),
      ...dayCols.map((i) => (str(r[i]) === '1' ? '1' : null))];
    cur.rows.push(rowKey(fragment));
    if (store) {
      cur.exceptionStores.add(store);
      if (!cur.storeRows.has(store)) cur.storeRows.set(store, []);
      cur.storeRows.get(store).push(fragment);
    }
  }
  return byEmp;
}

// 時刻を「05:00」にそろえる。時刻セル(数値)は1日=1で、24:00 は 1 になる
function formatTime(v) {
  const m = typeof v === 'number' ? Math.round(v * 24 * 60) : parseMinutes(v);
  return m === null ? str(v) : formatHHMM(m);
}

// 時給の行(時給・店舗ID・開始・終了・曜日フラグ)を比較用の文字列にする
const rowKey = (fragment) => fragment.map(str).join('\t');

// 更新する人ごとに、メイン時給(全店舗共通)の条件と例外時給の店舗を決める
// - 更新する店舗が1つ、または全店舗で条件が同じ → その条件だけ(店舗IDなし)
// - 店舗で条件が違う → 従業員マスタの主所属店舗の条件を新しいメイン時給にする。
//   主所属店舗がフォームの店舗にない人は、ARROWで今例外時給が付いていない店舗(=メイン時給で
//   働いている店舗)の条件にする。それも1つに決まらない人は choices(社員コード → 店舗ID)で
//   画面から選んでもらう。選ばれるまでは仮の店舗で計算し、pending として返す
// ARROWの時給インポートはその人の行をすべて置き換えるため、ARROWで例外時給がある店舗のうち
// フォームに回答がない店舗(更新・退職とも)は、今の行をそのまま kept として残す
// 従業員マスタにいない人は出力できないので missing として返す
function planRenewals({ byEmp, answered }, master, currentHourly, choices) {
  const plans = new Map();
  // 画面で選んでもらう人: [{ emp, stores, chosen, main }]
  const choosable = [];
  const missing = [];
  for (const [emp, stores] of byEmp) {
    const person = master.get(emp);
    if (!person) {
      missing.push(emp);
      continue;
    }
    const current = currentHourly.get(emp);
    const same = (a, b) => JSON.stringify(a.cond) === JSON.stringify(b.cond);

    let main = stores[0];
    if (!stores.every((s) => same(s, stores[0]))) {
      const home = stores.find((s) => s.storeId === person.mainStore);
      const candidates = stores.filter((s) => !current?.exceptionStores.has(s.storeId));
      if (home) {
        main = home;
      } else if (candidates.length === 1) {
        main = candidates[0];
      } else {
        const chosen = stores.find((s) => s.storeId === choices.get(emp));
        main = chosen || candidates[0] || stores[0];
        choosable.push({ emp, stores, chosen: Boolean(chosen), main });
      }
    }
    const exceptions = stores.filter((s) => s !== main && !same(s, main));
    const kept = [...(current?.storeRows || [])]
      .filter(([storeId]) => !answered.get(emp).has(storeId))
      .flatMap(([, rows]) => rows);
    plans.set(emp, { main, exceptions, kept, current, person });
  }
  if (!plans.size) throw new ConvertError('元データと従業員マスタで一致する社員コードがありません');
  return { plans, choosable, pending: choosable.filter((c) => !c.chosen).length, missing };
}

function buildPlans(sources, choices) {
  const storeIds = parseStores(sources.stores);
  return planRenewals(parseForm(sources.form, storeIds), parseMaster(sources.master),
    parseCurrentHourly(sources.hourly), choices);
}

// 従業員マスタにいない人の知らせ(各タブで出す)
const missingNotice = (missing) => (missing.length
  ? { message: `従業員マスタに登録がないため出力できない人がいます(${missing.length}名)`, details: missing }
  : null);

// 選択肢の表示用: 「1400円・日祝+50円・早朝+50円」
function describeCond({ cond }) {
  const parts = [`${cond.base}円`];
  if (cond.holiday) parts.push(`${cond.holiday}+${cond.daily}円`);
  if (cond.earlyBand) parts.push(`早朝+${cond.early}円`);
  return parts.join('・');
}
