/**
 * ============================================================
 * 図面承認・採番・出図管理システム
 * ファイル: 05_SearchIndex.gs  ―  検索用索引シートの管理
 * ============================================================
 *
 * 【背景・目的】
 *   図面台帳は改訂のたびに行が増えるappend-only設計のため、運用年数と
 *   ともに行数が際限なく増加する。件数が1万件規模になると、検索・
 *   一覧表示のたびに図面台帳を全件スキャンする現行方式では、Sheets API
 *   の読み込み・Apps Scriptの実行時間の両面で体感速度が悪化していく。
 *
 *   これを避けるため、用途別に「小さく保たれる」索引シートを別途持ち、
 *   検索・一覧系の処理は図面台帳ではなくこれらの索引に対して行う。
 *
 *   - 図面索引（SHEET_NAMES.DRAWING_INDEX）
 *       主図番＋子図番ごとに、最新リビジョンの行だけを複製保持する。
 *       改訂を何度繰り返しても「ユニークな部品点数」以上には増えない。
 *       キーワード検索・類似図面検索（30_DrawingSearch.gs）が使用する。
 *
 *   - 承認待ち索引（SHEET_NAMES.PENDING_INDEX）
 *       ステータスが「検図中／課長承認待ち／部長承認待ち」の行だけを
 *       保持する。承認が完了・差戻しされた図面は自動的に索引から
 *       除外されるため、図面台帳の総件数が何万件になっても、常に
 *       「今まさに進行中の承認件数」程度の小ささを保つ。
 *       承認Webアプリの一覧表示（23_ApprovalWebApp.gs の
 *       getMyPendingApprovals）が使用する。
 *
 *   どちらの索引も列構成は図面台帳（COL_DB / DRAWING_DB_COL_COUNT）と
 *   完全に同じにしてあり、「図面台帳の該当行をそのまま複製する」だけの
 *   単純なデータ構造にしている（スキーマを分けると変換ロジックが増え、
 *   同期漏れのバグを生みやすいため）。
 *
 * 【更新方針：インクリメンタル更新＋定期フル再構築】
 *   図面台帳への書き込み（登録・承認・差戻し・過去図面インポート）の
 *   たびに、その場で該当する索引エントリだけを更新する「インクリメンタル
 *   更新」を基本とする（upsertDrawingIndex_ / upsertPendingIndex_ 等）。
 *   ただし書き込み経路が複数にまたがるため、どこかに更新漏れがあると
 *   索引がズレる可能性がある。そのための保険として、図面台帳を全件
 *   スキャンして索引を1から作り直す rebuildSearchIndexes() を用意し、
 *   時間主導トリガーで毎日深夜に自動実行する（初回セットアップは
 *   setupIndexRebuildTrigger() を参照）。
 *
 * 【FILE_URL列の扱いについての注意】
 *   図面台帳のFILE_URL列（U列）は =HYPERLINK(...) 数式で保存されている。
 *   Range.setValues() は "=" で始まる文字列を自動的に数式として解釈する
 *   ため、newDbRows のように最初から数式文字列として組み立てられた配列は
 *   そのまま索引へ書き込んでよい。一方、getValues() で読み込んだ既存行の
 *   データはHYPERLINK数式が「表示テキスト」に変換されてしまっているため、
 *   索引へ複製する前に buildFileLinkFormula_() で数式文字列を作り直す
 *   必要がある（syncIndexesForFullNo_ / rebuildSearchIndexes を参照）。
 */

// 承認が進行中とみなすステータス一覧（STATUS定数は 00_Config.gs で定義済み）
const PENDING_STATUSES_ = [STATUS.REVIEWING, STATUS.WAITING_MANAGER, STATUS.WAITING_DIRECTOR];

// ============================================================
// インクリメンタル更新：書き込み経路から呼ばれる関数群
// ============================================================

/**
 * 図面索引（最新リビジョンのみ）へ1行分をupsertする。
 * 既存エントリより改訂記号が新しい（またはエントリが存在しない）場合のみ
 * 反映する。既存の方が新しい・同じ場合は何もしない。
 *
 * @param {Array} rowValues - 図面台帳の1行分の値配列（COL_DB準拠、
 *   DRAWING_DB_COL_COUNT要素）。FILE_URL列は数式文字列（"=HYPERLINK(...)"）
 *   であること（新規登録時のnewDbRows等、組み立て時点の配列をそのまま渡す想定）。
 */
function upsertDrawingIndex_(rowValues) {
  try {
    const sheet = ensureIndexSheet_(SHEET_NAMES.DRAWING_INDEX);
    const mainNo  = String(rowValues[COL_DB.MAIN_NO  - 1]);
    const subNo   = String(rowValues[COL_DB.SUB_NO   - 1]);
    const revMark = String(rowValues[COL_DB.REV_MARK - 1]);
    if (!mainNo || !subNo) return;

    const existing = findDrawingIndexRow_(sheet, mainNo, subNo);
    if (existing && existing.revMark.localeCompare(revMark) >= 0) {
      return; // 既存の索引の方が新しい、または同じリビジョン＝更新不要
    }

    if (existing) {
      sheet.getRange(existing.rowIndex, 1, 1, DRAWING_DB_COL_COUNT).setValues([rowValues]);
    } else {
      sheet.appendRow(rowValues);
    }
  } catch (e) {
    console.warn(`図面索引の更新に失敗しました: ${e.message}`);
  }
}

/**
 * 承認待ち索引へ1行分をupsertする。ステータスが承認進行中でなくなった
 * 場合は、索引から自動的に取り除く。
 *
 * @param {Array} rowValues - 図面台帳の1行分の値配列（upsertDrawingIndex_ と同様）
 */
function upsertPendingIndex_(rowValues) {
  try {
    const sheet  = ensureIndexSheet_(SHEET_NAMES.PENDING_INDEX);
    const fullNo = String(rowValues[COL_DB.FULL_NO - 1]);
    const status = String(rowValues[COL_DB.STATUS  - 1]);
    if (!fullNo) return;

    const existingRowIndex = findPendingIndexRowIndex_(sheet, fullNo);
    const isPending = PENDING_STATUSES_.indexOf(status) !== -1;

    if (!isPending) {
      if (existingRowIndex) sheet.deleteRow(existingRowIndex);
      return;
    }

    if (existingRowIndex) {
      sheet.getRange(existingRowIndex, 1, 1, DRAWING_DB_COL_COUNT).setValues([rowValues]);
    } else {
      sheet.appendRow(rowValues);
    }
  } catch (e) {
    console.warn(`承認待ち索引の更新に失敗しました: ${e.message}`);
  }
}

/**
 * 図面台帳上の特定のフル図番について、最新の内容を読み直して両方の
 * 索引へ反映する（承認・差戻し等でステータスや承認日時が更新された後に
 * 呼ぶ想定。23_ApprovalWebApp.gs から使用）。
 * findRow_ が返す data は表示テキスト化されたFILE_URLを含むため、
 * 数式として組み立て直してから索引へ渡す。
 *
 * @param {Sheet} dbSheet - 図面台帳シート
 * @param {string} fullNo - フル図番（10桁）
 */
function syncIndexesForFullNo_(dbSheet, fullNo) {
  try {
    const row = findRow_(dbSheet, fullNo); // 23_ApprovalWebApp.gs で定義済み
    if (!row) return; // 差戻し等で既に行が削除されている場合は何もしない

    const rowValues = row.data.slice();
    rowValues[COL_DB.FILE_URL - 1] = buildFileLinkFormula_(row.fileUrl, '📄図面を開く');

    upsertDrawingIndex_(rowValues);
    upsertPendingIndex_(rowValues);
  } catch (e) {
    console.warn(`索引の同期に失敗しました（${fullNo}）: ${e.message}`);
  }
}

/**
 * 差戻し等で図面台帳から行が削除された際に、承認待ち索引から該当行を
 * 取り除く。
 * @param {string} fullNo
 */
function removeFromPendingIndex_(fullNo) {
  try {
    const sheet = ensureIndexSheet_(SHEET_NAMES.PENDING_INDEX);
    const rowIndex = findPendingIndexRowIndex_(sheet, fullNo);
    if (rowIndex) sheet.deleteRow(rowIndex);
  } catch (e) {
    console.warn(`承認待ち索引からの削除に失敗しました（${fullNo}）: ${e.message}`);
  }
}

/**
 * 差戻し等で図面台帳から行が削除された際、その行が図面索引上で
 * 「最新」として登録されていた場合のみ、図面台帳を再スキャンして
 * 同じグループ（主図番＋子図番）に残っている行の中から最新を再計算する。
 * 差戻しは頻度の低い操作のため、この再計算に限っては全件スキャンを許容する
 * （検索・一覧表示という高頻度パスには一切影響しない）。
 *
 * @param {string} mainNo
 * @param {string} subNo
 * @param {string} deletedFullNo - 削除された行のフル図番
 */
function rebuildDrawingIndexEntryIfNeeded_(mainNo, subNo, deletedFullNo) {
  try {
    const sheet = ensureIndexSheet_(SHEET_NAMES.DRAWING_INDEX);
    const existing = findDrawingIndexRow_(sheet, mainNo, subNo);
    if (!existing) return;

    const existingFullNo = String(sheet.getRange(existing.rowIndex, COL_DB.FULL_NO).getValue());
    if (existingFullNo !== deletedFullNo) return; // 削除された行は索引上の最新ではなかった＝影響なし

    // 索引エントリを一旦削除し、図面台帳の残存行から改めて最新を探す
    sheet.deleteRow(existing.rowIndex);

    const ss     = SpreadsheetApp.getActiveSpreadsheet();
    const dbSheet = ss.getSheetByName(SHEET_NAMES.DRAWING_DB);
    if (!dbSheet) return;
    const lastRow = dbSheet.getLastRow();
    if (lastRow < GLOBAL_ROW.DATA_START) return;

    const range    = dbSheet.getRange(GLOBAL_ROW.DATA_START, 1, lastRow - GLOBAL_ROW.DATA_START + 1, DRAWING_DB_COL_COUNT);
    const rows     = range.getValues();
    const formulas = range.getFormulas();

    let latestIndex = -1;
    let latestRevMark = null;
    rows.forEach((row, i) => {
      if (String(row[COL_DB.MAIN_NO - 1]) === mainNo && String(row[COL_DB.SUB_NO - 1]) === subNo) {
        const revMark = String(row[COL_DB.REV_MARK - 1]);
        if (latestRevMark === null || revMark.localeCompare(latestRevMark) > 0) {
          latestRevMark = revMark;
          latestIndex = i;
        }
      }
    });

    if (latestIndex !== -1) {
      const latestValues = rows[latestIndex].slice();
      const fileUrl = extractUrlFromCellFormula_(formulas[latestIndex][COL_DB.FILE_URL - 1]) ||
                      String(latestValues[COL_DB.FILE_URL - 1]);
      latestValues[COL_DB.FILE_URL - 1] = buildFileLinkFormula_(fileUrl, '📄図面を開く');
      sheet.appendRow(latestValues);
    }
    // 残存行がなければ＝そのグループの図面は消滅＝索引からも消えたままでよい
  } catch (e) {
    console.warn(`図面索引の再計算に失敗しました（${mainNo}/${subNo}）: ${e.message}`);
  }
}

// ============================================================
// 定期フル再構築（時間主導トリガー／手動実行の両対応）
// ============================================================

/**
 * 図面台帳を全件スキャンし、図面索引・承認待ち索引の両方を1から作り直す。
 * インクリメンタル更新の更新漏れを補正するための安全網。
 * 時間主導トリガーで毎日深夜に自動実行する想定（setupIndexRebuildTrigger を参照）。
 * 手動で今すぐ再構築したい場合も、Apps Scriptエディタからこの関数を
 * 直接実行すればよい。
 */
function rebuildSearchIndexes() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const dbSheet = ss.getSheetByName(SHEET_NAMES.DRAWING_DB);
  if (!dbSheet) {
    console.error('図面台帳が見つからないため、索引の再構築を中止しました。');
    return;
  }

  const drawingIndexSheet = ensureIndexSheet_(SHEET_NAMES.DRAWING_INDEX);
  const pendingIndexSheet = ensureIndexSheet_(SHEET_NAMES.PENDING_INDEX);
  clearIndexDataRows_(drawingIndexSheet);
  clearIndexDataRows_(pendingIndexSheet);

  const lastRow = dbSheet.getLastRow();
  if (lastRow < GLOBAL_ROW.DATA_START) {
    console.log('図面台帳にデータがないため、索引は空のままです。');
    return;
  }

  const range    = dbSheet.getRange(GLOBAL_ROW.DATA_START, 1, lastRow - GLOBAL_ROW.DATA_START + 1, DRAWING_DB_COL_COUNT);
  const rows     = range.getValues();
  const formulas = range.getFormulas();

  const latestByGroup = {}; // "主図番|子図番" -> { rowValues, revMark }
  const pendingRows   = [];

  rows.forEach((row, i) => {
    // FILE_URL列は表示テキストではなく数式から実URLを復元し、
    // 索引側では改めてHYPERLINK数式として組み立て直す
    const fileUrl = extractUrlFromCellFormula_(formulas[i][COL_DB.FILE_URL - 1]) ||
                    String(row[COL_DB.FILE_URL - 1]);
    const rowValues = row.slice();
    rowValues[COL_DB.FILE_URL - 1] = buildFileLinkFormula_(fileUrl, '📄図面を開く');

    const mainNo   = String(row[COL_DB.MAIN_NO  - 1]);
    const subNo    = String(row[COL_DB.SUB_NO   - 1]);
    const revMark  = String(row[COL_DB.REV_MARK - 1]);
    const groupKey = mainNo + '|' + subNo;
    if (mainNo && subNo &&
        (!latestByGroup[groupKey] || revMark.localeCompare(latestByGroup[groupKey].revMark) > 0)) {
      latestByGroup[groupKey] = { rowValues, revMark };
    }

    const status = String(row[COL_DB.STATUS - 1]);
    if (PENDING_STATUSES_.indexOf(status) !== -1) {
      pendingRows.push(rowValues);
    }
  });

  const drawingIndexRows = Object.values(latestByGroup).map(g => g.rowValues);
  if (drawingIndexRows.length > 0) {
    drawingIndexSheet.getRange(GLOBAL_ROW.DATA_START, 1, drawingIndexRows.length, DRAWING_DB_COL_COUNT)
      .setValues(drawingIndexRows);
  }
  if (pendingRows.length > 0) {
    pendingIndexSheet.getRange(GLOBAL_ROW.DATA_START, 1, pendingRows.length, DRAWING_DB_COL_COUNT)
      .setValues(pendingRows);
  }

  console.log(
    `索引を再構築しました（図面索引: ${drawingIndexRows.length}件／` +
    `承認待ち索引: ${pendingRows.length}件、図面台帳全${rows.length}件中）。`
  );
}

// ============================================================
// 【初回セットアップ用】索引の毎日自動再構築トリガーを登録する
// ============================================================
/**
 * この関数を、Apps Script エディタから直接1回だけ手動実行してください。
 * 実行時に権限承認のダイアログが出るので、許可してください。
 * これにより、以降は毎日深夜（AM3時ごろ）に自動で索引が再構築されます。
 *
 * 実行方法：
 *   1. Apps Script エディタでこのファイルを開く
 *   2. 関数選択プルダウンで setupIndexRebuildTrigger を選択
 *   3. ▶実行 をクリックし、権限承認ダイアログで許可する
 *
 * ※ 一度登録すれば恒久的に有効。再実行しても重複登録されないよう、
 *   実行前に既存の同名トリガーを削除してから登録し直す。
 * ※ 実行時刻はスクリプトのタイムゾーン設定（appsscript.jsonのtimeZone）
 *   に従う。このプロジェクトでは Asia/Tokyo 設定のため、日本時間の
 *   深夜3時台に実行される。
 */
function setupIndexRebuildTrigger() {
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(t => {
    if (t.getHandlerFunction() === 'rebuildSearchIndexes') {
      ScriptApp.deleteTrigger(t);
    }
  });

  ScriptApp.newTrigger('rebuildSearchIndexes')
    .timeBased()
    .everyDays(1)
    .atHour(3)
    .create();

  console.log('✅ 索引の毎日自動再構築トリガーを登録しました（毎日AM3時ごろに実行）。');
}

// ============================================================
// プライベートヘルパー
// ============================================================

// 索引シートのA2セルに表示する説明文（シートを直接開いた人が、何のための
// シートかをすぐ確認できるようにするための注記）
const DRAWING_INDEX_DESCRIPTION_ =
  '検索（キーワード検索・類似図面検索）を高速化するために自動生成される「図面台帳」の複製です。図面の登録・承認・差戻しのたびに自動更新されます。';

const PENDING_INDEX_DESCRIPTION_ =
  '承認Webアプリの一覧表示を高速化するために自動生成される「図面台帳」の複製です。承認完了または差戻しが行われると自動的に除外されます。';

/**
 * 索引シートのA2セルへ、そのシートの説明文を書き込む（新規作成時・
 * 既存シートへの後付け適用の両方から呼ばれる）
 * @param {Sheet} sheet
 * @param {string} sheetName - SHEET_NAMES.DRAWING_INDEX または SHEET_NAMES.PENDING_INDEX
 */
function applyIndexSheetDescription_(sheet, sheetName) {
  const description = sheetName === SHEET_NAMES.DRAWING_INDEX
    ? DRAWING_INDEX_DESCRIPTION_
    : PENDING_INDEX_DESCRIPTION_;

  sheet.getRange(2, 1)
    .setValue(description)
    .setFontStyle('italic')
    .setFontColor('#888888')
    .setWrap(true);
}

/**
 * 【既存シート向け】図面索引・承認待ち索引のA2セルへ説明文を書き込む。
 * 既にシートが存在する場合はensureIndexSheet_内の自動記入処理が
 * スキップされる（新規作成時にしか通らないため）ため、既存シートへ
 * 後から説明文を追加したい場合はこの関数をApps Scriptエディタから
 * 直接実行してください。
 */
function writeIndexSheetDescriptions() {
  const drawingIndexSheet = ensureIndexSheet_(SHEET_NAMES.DRAWING_INDEX);
  const pendingIndexSheet = ensureIndexSheet_(SHEET_NAMES.PENDING_INDEX);
  applyIndexSheetDescription_(drawingIndexSheet, SHEET_NAMES.DRAWING_INDEX);
  applyIndexSheetDescription_(pendingIndexSheet, SHEET_NAMES.PENDING_INDEX);
  console.log('✅ 図面索引・承認待ち索引のA2セルへ説明文を書き込みました。');
}

/**
 * 索引シートを取得する。存在しなければ、図面台帳と同じ列構成で
 * ヘッダー付きの新規シートを作成する。
 * @param {string} sheetName
 * @returns {Sheet}
 */
function ensureIndexSheet_(sheetName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(sheetName);
  if (sheet) return sheet;

  sheet = ss.insertSheet(sheetName);
  const headers = [
    'フル図番', '主図番', '子図番', '改訂記号', '図名(JPN)', '英名(NAME)', 'ユニット名(UNIT)',
    '機械名(MODEL)', '材質(MATL.)', '縮尺(SCALE)', '図面サイズ', 'ステータス', '申請者', '申請日',
    '検図者', '検図者承認日時', '課長', '課長承認日時', '部長', '部長承認日時',
    '図面リンク', '申請バッチID', 'AIチェック', '特徴属性（JSON）', '特徴属性の要約',
  ];
  sheet.getRange(GLOBAL_ROW.HEADER, 1, 1, headers.length)
    .setValues([headers])
    .setFontWeight('bold')
    .setBackground('#f5f7fa');
  sheet.setFrozenRows(GLOBAL_ROW.HEADER);
  applyIndexSheetDescription_(sheet, sheetName);

  return sheet;
}

/**
 * 索引シートのデータ行（ヘッダーより下）を全てクリアする（フル再構築用）
 * @param {Sheet} sheet
 */
function clearIndexDataRows_(sheet) {
  const lastRow = sheet.getLastRow();
  if (lastRow >= GLOBAL_ROW.DATA_START) {
    sheet.getRange(GLOBAL_ROW.DATA_START, 1, lastRow - GLOBAL_ROW.DATA_START + 1, DRAWING_DB_COL_COUNT).clearContent();
  }
}

/**
 * 図面索引から、指定した主図番＋子図番に一致する行を探す
 * @param {Sheet} sheet - 図面索引シート
 * @param {string} mainNo
 * @param {string} subNo
 * @returns {{rowIndex: number, revMark: string}|null}
 */
function findDrawingIndexRow_(sheet, mainNo, subNo) {
  const lastRow = sheet.getLastRow();
  if (lastRow < GLOBAL_ROW.DATA_START) return null;

  const values = sheet
    .getRange(GLOBAL_ROW.DATA_START, 1, lastRow - GLOBAL_ROW.DATA_START + 1, DRAWING_DB_COL_COUNT)
    .getValues();

  for (let i = 0; i < values.length; i++) {
    if (String(values[i][COL_DB.MAIN_NO - 1]) === mainNo && String(values[i][COL_DB.SUB_NO - 1]) === subNo) {
      return { rowIndex: GLOBAL_ROW.DATA_START + i, revMark: String(values[i][COL_DB.REV_MARK - 1]) };
    }
  }
  return null;
}

/**
 * 承認待ち索引から、指定したフル図番に一致する行番号を探す
 * @param {Sheet} sheet - 承認待ち索引シート
 * @param {string} fullNo
 * @returns {number|null}
 */
function findPendingIndexRowIndex_(sheet, fullNo) {
  const lastRow = sheet.getLastRow();
  if (lastRow < GLOBAL_ROW.DATA_START) return null;

  const values = sheet
    .getRange(GLOBAL_ROW.DATA_START, COL_DB.FULL_NO, lastRow - GLOBAL_ROW.DATA_START + 1, 1)
    .getValues()
    .flat();

  const idx = values.findIndex(v => String(v) === fullNo);
  return idx === -1 ? null : GLOBAL_ROW.DATA_START + idx;
}

// ============================================================
// 【デバッグ用】
// ============================================================
function TEST_rebuildSearchIndexes() {
  rebuildSearchIndexes();
}