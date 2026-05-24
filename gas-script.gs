/**
 * 経営会議タイマー - Google Apps Script
 *
 * 【セットアップ手順】
 * 1. Google Sheetsで新規スプレッドシートを作成
 * 2. 以下の3シートを作成:
 *    - 「設定」シート: 参加者一覧・ジャンル一覧
 *    - 「記録」シート: 会議データの蓄積先
 *    - 「集計」シート: 分析用（任意）
 * 3. 拡張機能 > Apps Script でこのスクリプトを貼り付け
 * 4. デプロイ > 新しいデプロイ > ウェブアプリ
 *    - 実行ユーザー: 自分
 *    - アクセス: 全員
 * 5. デプロイURLをタイマーアプリの設定に貼り付け
 *
 * 【「プルダウン用」シートの構成】
 * A列: 参加者名（A1は「参加者」ヘッダー）
 * B列: ジャンル名（B1は「ジャンル」ヘッダー）
 *
 * 【「記録」シートの構成】（自動作成されます）
 * A: 日付, B: 議題No, C: 議題名, D: 提案者, E: ジャンル,
 * F: 予定_提案(秒), G: 予定_議論(秒), H: 予定_まとめ(秒), I: 予定_合計(秒),
 * J: 実績_提案(秒), K: 実績_議論(秒), L: 実績_まとめ(秒), M: 実績_合計(秒),
 * N: 差分(秒)
 */

// GET: マスタデータ取得
function doGet(e) {
  const action = e.parameter.action;

  if (action === 'getMaster') {
    return jsonResponse(getMasterData());
  }

  if (action === 'getRecords') {
    const from = e.parameter.from || '';
    const to = e.parameter.to || '';
    return jsonResponse(getRecords(from, to));
  }

  if (action === 'getSummary') {
    return jsonResponse(getSummary());
  }

  return jsonResponse({ status: 'error', message: 'Unknown action' });
}

// POST: データ書き込み
function doPost(e) {
  try {
    const body = JSON.parse(e.postData.contents);

    if (body.action === 'saveRecords') {
      saveRecords(body.records);
      return jsonResponse({ status: 'ok', count: body.records.length });
    }

    return jsonResponse({ status: 'error', message: 'Unknown action' });
  } catch (err) {
    return jsonResponse({ status: 'error', message: err.message });
  }
}

// ===== マスタデータ取得 =====
function getMasterData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('プルダウン');
  if (!sheet) return { participants: [], categories: [] };

  const data = sheet.getDataRange().getValues();
  const participants = [];
  const categories = [];

  for (let i = 1; i < data.length; i++) {
    if (data[i][0] && String(data[i][0]).trim()) participants.push(String(data[i][0]).trim());
    if (data[i][1] && String(data[i][1]).trim()) categories.push(String(data[i][1]).trim());
  }

  return { participants, categories };
}

// ===== 記録保存 =====
function saveRecords(records) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName('記録');

  // シートがなければ作成
  if (!sheet) {
    sheet = ss.insertSheet('記録');
    sheet.appendRow([
      '日付', '議題No', '議題名', '提案者', 'ジャンル',
      '予定_提案', '予定_議論', '予定_まとめ', '予定_合計',
      '実績_提案', '実績_議論', '実績_まとめ', '実績_合計',
      '差分'
    ]);
    // ヘッダー書式
    sheet.getRange(1, 1, 1, 14).setFontWeight('bold').setBackground('#4285f4').setFontColor('#ffffff');
    sheet.setFrozenRows(1);
  }

  records.forEach(r => {
    const diff = r.actualTotal - r.plannedTotal;
    sheet.appendRow([
      r.date,
      r.agendaNo,
      r.title,
      r.proposer,
      r.category,
      formatSecToMinSec(r.plannedProposal),
      formatSecToMinSec(r.plannedDiscussion),
      formatSecToMinSec(r.plannedSummary),
      formatSecToMinSec(r.plannedTotal),
      formatSecToMinSec(r.actualProposal),
      formatSecToMinSec(r.actualDiscussion),
      formatSecToMinSec(r.actualSummary),
      formatSecToMinSec(r.actualTotal),
      formatSecToMinSec(diff)
    ]);
  });
}

// ===== 記録取得（日付範囲指定可） =====
function getRecords(from, to) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('記録');
  if (!sheet || sheet.getLastRow() <= 1) return { records: [] };

  const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 14).getValues();
  const records = data
    .filter(row => {
      if (!row[0]) return false;
      const dateStr = String(row[0]);
      if (from && dateStr < from) return false;
      if (to && dateStr > to) return false;
      return true;
    })
    .map(row => ({
      date: String(row[0]),
      agendaNo: row[1],
      title: row[2],
      proposer: row[3],
      category: row[4],
      plannedProposal: row[5],
      plannedDiscussion: row[6],
      plannedSummary: row[7],
      plannedTotal: row[8],
      actualProposal: row[9],
      actualDiscussion: row[10],
      actualSummary: row[11],
      actualTotal: row[12],
      diff: row[13],
    }));

  return { records };
}

// ===== 集計サマリー =====
function getSummary() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('記録');
  if (!sheet || sheet.getLastRow() <= 1) return { summary: null };

  const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 14).getValues();

  // 日付ごとにグループ化
  const byDate = {};
  data.forEach(row => {
    const d = String(row[0]);
    if (!d) return;
    if (!byDate[d]) byDate[d] = [];
    byDate[d].push({
      plannedTotal: row[8],
      actualTotal: row[12],
      diff: row[13],
      proposer: row[3],
      category: row[4],
      actualProposal: row[9],
      actualDiscussion: row[10],
      actualSummary: row[11],
    });
  });

  const meetingCount = Object.keys(byDate).length;
  let totalAgendas = 0;
  let totalPlanned = 0;
  let totalActual = 0;
  let totalOverTime = 0;

  // 提案者別の超過回数
  const proposerOver = {};
  // ジャンル別の平均超過
  const categoryStats = {};

  Object.values(byDate).forEach(agendas => {
    agendas.forEach(a => {
      totalAgendas++;
      totalPlanned += a.plannedTotal;
      totalActual += a.actualTotal;
      if (a.diff > 0) totalOverTime++;

      if (a.proposer) {
        if (!proposerOver[a.proposer]) proposerOver[a.proposer] = { total: 0, over: 0, totalDiff: 0 };
        proposerOver[a.proposer].total++;
        proposerOver[a.proposer].totalDiff += a.diff;
        if (a.diff > 0) proposerOver[a.proposer].over++;
      }

      if (a.category) {
        if (!categoryStats[a.category]) categoryStats[a.category] = { total: 0, totalDiff: 0, totalProposal: 0, totalDiscussion: 0, totalSummary: 0 };
        categoryStats[a.category].total++;
        categoryStats[a.category].totalDiff += a.diff;
        categoryStats[a.category].totalProposal += a.actualProposal;
        categoryStats[a.category].totalDiscussion += a.actualDiscussion;
        categoryStats[a.category].totalSummary += a.actualSummary;
      }
    });
  });

  return {
    summary: {
      meetingCount,
      totalAgendas,
      avgAgendasPerMeeting: totalAgendas / meetingCount,
      totalPlannedSec: totalPlanned,
      totalActualSec: totalActual,
      overTimeRate: totalOverTime / totalAgendas,
      avgDiffSec: (totalActual - totalPlanned) / totalAgendas,
      proposerStats: proposerOver,
      categoryStats,
    }
  };
}

// ===== 秒数を H:MM:SS 形式に変換 =====
function formatSecToMinSec(totalSec) {
  if (totalSec == null || totalSec === '') return '';
  const num = Number(totalSec);
  if (isNaN(num)) return totalSec;
  const negative = num < 0;
  const abs = Math.abs(num);
  const h = Math.floor(abs / 3600);
  const m = Math.floor((abs % 3600) / 60);
  const s = abs % 60;
  const formatted = h + ':' + String(m).padStart(2, '0') + ':' + String(s).padStart(2, '0');
  return negative ? '-' + formatted : formatted;
}

// ===== ユーティリティ =====
function jsonResponse(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}
