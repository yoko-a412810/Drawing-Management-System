/**
 * ============================================================
 * 図面承認・採番・出図管理システム
 * ファイル: 37_RefreshDrawingAttributes.gs  ―  既存図面の特徴属性・要約の再抽出（移行用）
 * ============================================================
 *
 * 【背景】
 *   21_OcrService.gs の OCR_PROMPT（attributesのスキーマ）に、幾何公差・
 *   表面処理・熱処理の項目を追加した（35_LegacyDrawingImport.gs の
 *   LEGACY_OCR_PROMPT も同様）。しかし、これは新規にOCR読み取りを行った
 *   図面にのみ適用され、既に図面台帳へ登録済みの図面（COL_DB.ATTRIBUTES /
 *   ATTRIBUTES_SUMMARY）は旧スキーマのまま残っている。
 *   この移行スクリプトは、図面台帳の各行が指すPDF（COL_DB.FILE_URL）に
 *   対して再度Gemini APIでOCRを実行し、features・summaryを新しいスキーマで
 *   上書きする（表題欄の他の項目・承認状況などは一切変更しない）。
 *
 * 【対象】
 *   図面台帳の全行のうち、FILE_URLが設定されている行
 *   （フル図番の重複やステータスは問わない＝過去図面・承認済み・進行中の
 *   すべての行が対象）。
 *
 * 【処理時間について（再開可能な設計）】
 *   Apps Scriptの1回の実行には時間制限があるため、件数が多い場合は
 *   1回の実行では終わらない。この関数はスクリプトプロパティに処理済み
 *   位置（次に処理する行番号）を1行処理するたびに記録しながら進めるため、
 *   時間切れで打ち切られても、同じ関数を再実行すれば続きから処理を
 *   再開できる（重複処理・重複API呼び出しは発生しない）。
 *   実行数ログに「未完了」と出た場合は、そのままもう一度▶実行してください。
 *   「全件完了」と出るまで繰り返す。
 *
 * 【実行方法】
 *   1. Apps Script エディタでこのファイルを開く
 *   2. 関数選択プルダウンで refreshDrawingAttributesForExistingRows を選択
 *   3. ▶実行 をクリック（実行数ログに進捗が出る）
 *   4.「未完了（あと約N件）」と出たら、再度▶実行を繰り返す
 *   5.「全件完了」と出たら終了（検索用索引も自動的に再構築される）
 *
 * 【やり直したい場合】
 *   途中からではなく全件を最初からやり直したい場合は、
 *   resetAttributesRefreshProgress() を1回実行してから
 *   refreshDrawingAttributesForExistingRows() を実行する。
 *
 * 【注意】
 *   図面1件ごとにGemini APIを呼び出すため、対象件数分のAPI利用料・
 *   処理時間がかかる。テスト用のダミー図面が少数であれば数分で完了するが、
 *   実運用データで件数が多い場合はご注意ください。
 */

// 1回の実行で処理してよい最大時間（Apps Scriptの実行時間制限に対する安全マージン）
const ATTR_REFRESH_TIME_BUDGET_MS_ = 4.5 * 60 * 1000; // 4分30秒
const ATTR_REFRESH_PROP_KEY_ = 'ATTR_REFRESH_NEXT_ROW';

function refreshDrawingAttributesForExistingRows() {
  const ss     = SpreadsheetApp.getActiveSpreadsheet();
  const dbSheet = ss.getSheetByName(SHEET_NAMES.DRAWING_DB);
  if (!dbSheet) {
    console.error(`"${SHEET_NAMES.DRAWING_DB}" シートが見つからないため、処理を中止しました。`);
    return;
  }

  const lastRow = dbSheet.getLastRow();
  if (lastRow < GLOBAL_ROW.DATA_START) {
    console.log('図面台帳にデータがないため、対象はありませんでした。');
    return;
  }

  const props = PropertiesService.getScriptProperties();
  const savedNextRow = Number(props.getProperty(ATTR_REFRESH_PROP_KEY_));
  let currentRow = (savedNextRow && savedNextRow >= GLOBAL_ROW.DATA_START) ? savedNextRow : GLOBAL_ROW.DATA_START;

  if (currentRow > lastRow) {
    console.log(
      '前回までの処理で既に全件完了しています。' +
      '最初からやり直す場合は resetAttributesRefreshProgress() を実行してください。'
    );
    return;
  }

  const startTime = Date.now();
  let updatedCount = 0;
  let skippedCount = 0;
  let failedCount  = 0;

  while (currentRow <= lastRow) {
    if (Date.now() - startTime > ATTR_REFRESH_TIME_BUDGET_MS_) {
      break; // 時間切れ。続きは次回の実行で処理する
    }

    const fullNo = String(dbSheet.getRange(currentRow, COL_DB.FULL_NO).getValue()).trim();

    // FILE_URL列はHYPERLINK数式なので、getFormula()から実URLを取り出す
    const fileUrlFormula = dbSheet.getRange(currentRow, COL_DB.FILE_URL).getFormula();
    const fileUrl = extractUrlFromCellFormula_(fileUrlFormula) || // 00_Config.gs
                    String(dbSheet.getRange(currentRow, COL_DB.FILE_URL).getValue()).trim();

    if (!fileUrl) {
      console.warn(`${currentRow}行目（${fullNo || '図番不明'}）：図面リンクが空のためスキップしました。`);
      skippedCount++;
      currentRow++;
      props.setProperty(ATTR_REFRESH_PROP_KEY_, String(currentRow));
      continue;
    }

    const fileId = extractDriveFileId_(fileUrl); // 25_PdfApprovalStamp.gs
    if (!fileId) {
      console.warn(`${currentRow}行目（${fullNo}）：ファイルIDを取得できないためスキップしました。URL: ${fileUrl}`);
      skippedCount++;
      currentRow++;
      props.setProperty(ATTR_REFRESH_PROP_KEY_, String(currentRow));
      continue;
    }

    let quotaExceeded = false;

    try {
      const file      = DriveApp.getFileById(fileId);
      const extracted = extractDrawingInfoByOcr_(file); // 21_OcrService.gs

      if (!extracted || !extracted.attributes) {
        console.warn(`${currentRow}行目（${fullNo}）：属性の抽出に失敗したためスキップしました。`);
        failedCount++;
      } else {
        const attributesJson = JSON.stringify(extracted.attributes);
        const summary        = String(extracted.attributes.summary || '');
        dbSheet.getRange(currentRow, COL_DB.ATTRIBUTES).setValue(attributesJson);
        dbSheet.getRange(currentRow, COL_DB.ATTRIBUTES_SUMMARY).setValue(summary);
        updatedCount++;
      }
    } catch (e) {
      const msg = String(e && e.message ? e.message : e);
      // Gemini APIのクォータ超過（429 / RESOURCE_EXHAUSTED）は、この行固有の
      // 問題ではなくAPIキー自体がその日使えなくなったことを意味する。
      // 通常のエラーと同じように「失敗として次の行へ進める」扱いにしてしまうと、
      // 以降の行も連続して同じエラーになり、本来は再試行すべき行が次々と
      // 「失敗」のまま読み飛ばされてしまう（＝進捗カーソルだけ進み、実際には
      // 再抽出されないまま放置される）。これを避けるため、クォータ超過を
      // 検知した場合はこの行の進捗を進めずに（＝次回はこの行から再開する）
      // ループ全体を中断する。
      if (msg.indexOf('429') !== -1 || msg.indexOf('RESOURCE_EXHAUSTED') !== -1 || msg.indexOf('quota') !== -1) {
        console.error(
          `${currentRow}行目（${fullNo}）でGemini APIのクォータ上限に達したため、処理を中断しました。\n` +
          `クォータがリセットされるまで（無料枠は通常24時間以内）待つか、Google Cloud Console側で` +
          `課金設定を有効化してから、再度 refreshDrawingAttributesForExistingRows を実行してください。\n` +
          `この行はまだ処理していない扱いとして、次回はここから再開します。`
        );
        quotaExceeded = true;
      } else {
        console.warn(`${currentRow}行目（${fullNo}）：処理中にエラーが発生したためスキップしました：${e.message}`);
        failedCount++;
      }
    }

    if (quotaExceeded) break; // 行カーソルを進めずにループを抜ける（この行は次回再試行される）

    currentRow++;
    props.setProperty(ATTR_REFRESH_PROP_KEY_, String(currentRow));

    Utilities.sleep(1000); // Gemini APIのレート制限対策（既存のOCR一括処理と同じ間隔）
  }

  const remaining = Math.max(0, lastRow - currentRow + 1);
  console.log(
    `今回の実行：更新 ${updatedCount} 件／スキップ ${skippedCount} 件／失敗 ${failedCount} 件。\n` +
    (remaining > 0
      ? `未完了（あと約 ${remaining} 件）。もう一度この関数（refreshDrawingAttributesForExistingRows）を実行して続きを処理してください。`
      : '全件完了しました。')
  );

  if (remaining === 0) {
    props.deleteProperty(ATTR_REFRESH_PROP_KEY_);
    // 図面台帳のATTRIBUTES / ATTRIBUTES_SUMMARY列を書き換えたため、
    // 検索用索引（図面索引・承認待ち索引）も古い内容のまま残らないよう再構築する
    rebuildSearchIndexes(); // 05_SearchIndex.gs
    console.log('検索用索引（図面索引・承認待ち索引）を再構築しました。');
  }
}

/**
 * 再抽出の進捗（スクリプトプロパティ）をリセットする。
 * 次回 refreshDrawingAttributesForExistingRows() を実行すると、図面台帳の
 * 先頭行からやり直しになる。
 */
function resetAttributesRefreshProgress() {
  PropertiesService.getScriptProperties().deleteProperty(ATTR_REFRESH_PROP_KEY_);
  console.log('進捗をリセットしました。次回実行時は先頭行から処理します。');
}

/**
 * 現在の進捗状況を確認する（デバッグ用）
 */
function checkAttributesRefreshProgress() {
  const props = PropertiesService.getScriptProperties();
  const savedNextRow = props.getProperty(ATTR_REFRESH_PROP_KEY_);
  if (!savedNextRow) {
    console.log('進捗の記録はありません（未実行、または前回で全件完了済みです）。');
    return;
  }
  const ss      = SpreadsheetApp.getActiveSpreadsheet();
  const dbSheet = ss.getSheetByName(SHEET_NAMES.DRAWING_DB);
  const lastRow = dbSheet ? dbSheet.getLastRow() : 0;
  console.log(`次回は ${savedNextRow} 行目から再開します（図面台帳の最終行：${lastRow}）。`);
}