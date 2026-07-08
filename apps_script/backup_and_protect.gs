/**
 * 原価管理システム：日次バックアップ ＋ 数式シートの保護
 *
 * このスクリプトは対象のスプレッドシート自身に貼り付けて使います。
 * 機能:
 *   1) dailyBackup()          … 台帳を「バックアップ」フォルダへ日付つきコピー（自動）
 *   2) protectFormulaSheets() … 数式シートを保護（黄色の入力セルだけ編集可）
 *   3) setupDailyTrigger()    … dailyBackup を毎日自動実行するトリガーを登録
 *
 * セットアップ手順は docs/運用ガイド.md を参照してください。
 */

// ===== 設定 =====
var BACKUP_FOLDER_NAME = 'バックアップ';   // 同じ場所に作成／利用
var KEEP_BACKUPS = 30;                      // 保持する世代数（古いものは自動削除）
var BACKUP_HOUR = 2;                        // 日次バックアップの実行時刻（0-23, 深夜帯推奨）

// 保護の方式:
//   true  = 警告のみ（編集しようとすると全員に確認ポップアップ。ロックしない・推奨）
//   false = オーナー以外は編集不可（ハード。社員の編集を完全にブロック）
var WARNING_ONLY = true;

// 保護するシートと、編集を許す例外セル（[]＝シート全体を保護）
// ※ マスタ（工事情報・工種マスタ）と取込データ（見積取込・実績取込）は保護しません。
var PROTECT_CONFIG = {
  '総合管理表':      ['H4:I100'],  // 出来高率・現場見込 だけ編集可
  '工種別内訳':      [],
  '月次集計（月締）': [],
  '支払予定':        []
};

// ===== 1) 日次バックアップ =====
function dailyBackup() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var file = DriveApp.getFileById(ss.getId());
  var parent = file.getParents().hasNext() ? file.getParents().next() : DriveApp.getRootFolder();
  var folder = getOrCreateFolder_(parent, BACKUP_FOLDER_NAME);
  var tz = ss.getSpreadsheetTimeZone() || 'Asia/Tokyo';
  var stamp = Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd_HHmm');
  file.makeCopy(ss.getName() + '_' + stamp, folder);
  pruneBackups_(folder, ss.getName(), KEEP_BACKUPS);
}

function getOrCreateFolder_(parent, name) {
  var it = parent.getFoldersByName(name);
  return it.hasNext() ? it.next() : parent.createFolder(name);
}

function pruneBackups_(folder, baseName, keep) {
  var files = [];
  var it = folder.getFiles();
  while (it.hasNext()) {
    var f = it.next();
    if (f.getName().indexOf(baseName + '_') === 0) files.push(f);
  }
  files.sort(function (a, b) { return b.getDateCreated() - a.getDateCreated(); });
  for (var i = keep; i < files.length; i++) files[i].setTrashed(true);
}

// ===== 2) 数式シートの保護 =====
function protectFormulaSheets() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var applied = [];
  Object.keys(PROTECT_CONFIG).forEach(function (name) {
    var sh = ss.getSheetByName(name);
    if (!sh) { Logger.log('※ シートが見つかりません: ' + name); return; }
    // 既存の自動保護を貼り直す（重複防止）
    sh.getProtections(SpreadsheetApp.ProtectionType.SHEET).forEach(function (p) {
      if (p.getDescription() === '数式保護（自動）') p.remove();
    });
    var prot = sh.protect().setDescription('数式保護（自動）');
    var ex = PROTECT_CONFIG[name];
    if (ex && ex.length) {
      prot.setUnprotectedRanges(ex.map(function (a) { return sh.getRange(a); }));
    }
    if (WARNING_ONLY) {
      // 編集時に確認ポップアップ（全員・オーナー含む）。ロックしない。
      prot.setWarningOnly(true);
    } else {
      // オーナー以外は編集不可
      prot.removeEditors(prot.getEditors());
      if (prot.canDomainEdit()) prot.setDomainEdit(false);
    }
    applied.push(name);
  });
  Logger.log('保護を適用: ' + (applied.join(' / ') || '（対象なし）') +
             '　方式=' + (WARNING_ONLY ? '警告のみ' : 'オーナー以外編集不可'));
}

// 保護を一時解除したいとき（工種の追加・様式変更など）
function unprotectFormulaSheets() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  ss.getSheets().forEach(function (sh) {
    sh.getProtections(SpreadsheetApp.ProtectionType.SHEET).forEach(function (p) {
      if (p.getDescription() === '数式保護（自動）') p.remove();
    });
  });
}

// ===== 3) 日次トリガー登録 =====
function setupDailyTrigger() {
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === 'dailyBackup') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('dailyBackup').timeBased().everyDays(1).atHour(BACKUP_HOUR).create();
}

// ===== 初回セットアップを一括実行 =====
function setupAll() {
  protectFormulaSheets();
  setupDailyTrigger();
  dailyBackup(); // 初回バックアップを1本作成
}
