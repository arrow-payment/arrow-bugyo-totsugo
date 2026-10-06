'use strict';

// 給与設定インポート用データ
//
// 形式は「ほっか時給更新ツール」の docs/アウトプットサンプル_給与設定.xlsx と同じ(ARROWの給与設定インポート)。
// 契約更新する人について、従業員マスタの同じ見出しの列の値をそのまま出し、次の列だけを上書きする。
// - 有給休暇全日使用単価 = メイン時給の条件の時給 × 週所定労働時間 ÷ 週所定労働日数(四捨五入)
// - 有給休暇全日使用単価の計算方法 = 手動(単価を入れるため)
// 週所定労働時間・日数の登録がない人は計算できないので、上書きせず従業員マスタの値のまま出す。
// 社員名・カナは従業員マスタの姓・名から作る(時給と同じ)。

const PAYROLL_HEADER = ['社員コード', '社員名', '社員名(カナ)',
  '基本給', '月給', '基礎単価の計算方法', '基礎単価', '研修時給のカウント方法', '研修時給の使用上限',
  '通勤手当', '通勤方法', '通勤手当の金額', '日額通勤手当の月上限', '日額通勤手当の月上限額',
  'マイカーの片道距離(km)', '通勤経路', '残業手当', '残業手当の対象', '固定残業の上限時間', '固定残業の金額',
  '深夜手当', '休日手当', '休日手当の対象', '有給休暇全日使用単価の計算方法', '有給休暇全日使用単価',
  '賞与', '個人別単価1', '個人別単価2', '個人別単価3'];
// 従業員マスタから値を写す列(社員コード・社員名・カナ以外)
const PAYROLL_COPY_COLUMNS = Object.fromEntries(PAYROLL_HEADER.slice(3).map((h) => [h, h]));
const PAYROLL_UPDATE_COLUMNS = {
  paidLeaveMethod: PAYROLL_HEADER.indexOf('有給休暇全日使用単価の計算方法'),
  paidLeave: PAYROLL_HEADER.indexOf('有給休暇全日使用単価'),
};

function convertPayroll(sources, { plans, missing }) {

  const masterCol = findColumns(sources.master[1], PAYROLL_COPY_COLUMNS, '従業員マスタ');

  const outRows = [];
  // 出力行 → 従業員マスタから値が変わる列index の集合
  const changed = new Map();
  const noWeekly = [];
  for (const [emp, { main, person }] of plans) {
    const row = [emp, person.name, person.kana,
      ...PAYROLL_HEADER.slice(3).map((h) => person.row[masterCol[h]] ?? null)];
    outRows.push(row);
    const { weekDays, weekMinutes } = person;
    if (!weekDays || !weekMinutes) {
      noWeekly.push(emp);
      continue;
    }

    const updates = {
      paidLeaveMethod: '手動',
      paidLeave: String(Math.round((main.cond.base * weekMinutes) / weekDays / 60)),
    };
    const changedCols = new Set();
    for (const [key, value] of Object.entries(updates)) {
      const i = PAYROLL_UPDATE_COLUMNS[key];
      if (str(row[i]) !== value) changedCols.add(i);
      row[i] = value;
    }
    if (changedCols.size) changed.set(row, changedCols);
  }

  return {
    outRows,
    empCount: outRows.length,
    stats: { '有給単価の変更あり': changed.size, '週所定の登録なし': noWeekly.length },
    changed,
    notice: missingNotice(missing),
  };
}

const PAYROLL = {
  key: 'payroll',
  label: '給与設定',
  formatVersion: '1.1.1',
  header: PAYROLL_HEADER,
  previewLabels: { 2: 'カナ' },
  colWidths: PAYROLL_HEADER.map((h) => Math.max(8, str(h).length * 2)),
  filePrefix: '給与設定インポートデータ',
  // 従業員マスタから値が変わる行の色を変える
  highlighter: (outRows, result) => (row) => result.changed.has(row),
  highlightLabel: '変更ありのみ',
  convert: convertPayroll,
};
