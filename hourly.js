'use strict';

// 時給インポート用データ
//
// 契約更新する人だけを出力する。行の作り方は「ほっか時給更新ツール」と同じ:
// - メイン時給(E列空欄)は全店舗共通の時給。店舗で条件が違う場合だけ、その店舗IDを入れた例外時給の行を追加
// - 休日手当(日祝/土日祝)あり → 該当曜日以外の行 + 該当曜日は日別手当を加算した行
// - 早朝手当あり → 00:00-05:00 / 05:00-08:00(+早朝手当) / 08:00-24:00 の3分割
// - 社員名・カナは従業員マスタの姓・名から作る
// - インポートはその人の行をすべて置き換えるので、フォームに回答がない店舗の例外時給は今の行のまま出す

const DAYS = ['日', '月', '火', '水', '木', '金', '土', '祝'];
const PREM_MAP = { '日祝': new Set(['日', '祝']), '土日祝': new Set(['土', '日', '祝']) };

function buildHourlyRows([emp, name, kana], storeId, { base, early, daily, holiday, earlyBand }) {
  const prem = PREM_MAP[holiday] || new Set();
  const groups = prem.size
    ? [[new Set(DAYS.filter((d) => !prem.has(d))), 0], [prem, daily]]
    : [[new Set(DAYS), 0]];
  const rows = [];
  for (const [days, add] of groups) {
    const flags = DAYS.map((d) => (days.has(d) ? '1' : null));
    if (earlyBand) {
      rows.push([emp, name, kana, base + add, storeId, '00:00', '05:00', ...flags]);
      rows.push([emp, name, kana, base + add + early, storeId, '05:00', '08:00', ...flags]);
      rows.push([emp, name, kana, base + add, storeId, '08:00', '24:00', ...flags]);
    } else {
      rows.push([emp, name, kana, base + add, storeId, '00:00', '24:00', ...flags]);
    }
  }
  return rows;
}

function convertHourly(sources, { plans, missing }) {

  const outRows = [];
  const changedEmps = new Set();
  const stats = { '変更あり': 0, '例外時給あり': 0, '回答のない店舗の時給を維持': 0, 'ARROWに時給の登録なし': 0 };
  for (const [emp, { main, exceptions, kept, current, person }] of plans) {
    const p = [emp, person.name, person.kana];
    const rows = buildHourlyRows(p, null, main.cond);
    for (const s of exceptions) rows.push(...buildHourlyRows(p, s.storeId, s.cond));
    for (const fragment of kept) rows.push([...p, ...fragment]);
    if (kept.length) stats['回答のない店舗の時給を維持']++;

    // ARROWの現在の行と(順不同で)同じかどうか
    const keys = rows.map((r) => rowKey(r.slice(3))).sort().join('\n');
    if (!current) stats['ARROWに時給の登録なし']++;
    if (!current || keys !== [...current.rows].sort().join('\n')) {
      changedEmps.add(emp);
      stats['変更あり']++;
    }
    if (exceptions.length || kept.length) stats['例外時給あり']++;
    outRows.push(...rows);
  }

  return { outRows, empCount: plans.size, stats, changedEmps, notice: missingNotice(missing) };
}

const HOURLY = {
  key: 'hourly',
  label: '時給',
  formatVersion: '1.1.1',
  header: ['社員コード', '社員名', '社員名(カナ)', '時給', '対象部署/店舗ID',
    '開始時間', '終了時間',
    '日曜日', '月曜日', '火曜日', '水曜日', '木曜日', '金曜日', '土曜日', '祝日'],
  previewLabels: { 2: 'カナ', 4: '店舗ID', 5: '開始', 6: '終了',
    7: '日', 8: '月', 9: '火', 10: '水', 11: '木', 12: '金', 13: '土', 14: '祝' },
  colWidths: [14, 10, 12, 8, 16, 10, 10],
  filePrefix: '時給インポートデータ',
  // ARROWの現在の登録から変わる人の行の色を変える
  highlighter: (outRows, result) => (row) => result.changedEmps.has(row[0]),
  highlightLabel: '変更ありのみ',
  convert: convertHourly,
};
