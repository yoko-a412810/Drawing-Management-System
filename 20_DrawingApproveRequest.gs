/**
 * ============================================================
 * 図面承認・採番・出図管理システム
 * ファイル: 20_DrawingApproveRequest.gs  ―  図面一括登録・承認申請モジュール
 * ============================================================
 *
 * 【処理の流れ】
 *   ① カスタムメニュー「① PDF一括読み取り（OCR）」
 *       → runOcrAndFillInputSheet()
 *       → 「読取図面保存」フォルダをスキャン → Gemini API → 入力シートへ転記
 *         → 読み取り済みファイルは即座に「申請中図面」フォルダへ移動
 *       ※ 共有ドライブ運用のため、複数人のPDFが混在すると誤って他人の
 *          図面を読み取ってしまう事故が起こり得る。これを防ぐため、
 *          「読取図面保存」フォルダ→「申請中図面」フォルダ→「承認済み図面」
 *          フォルダの3段階構成にし、OCR読み取り後は即座に次のフォルダへ
 *          移動することで、「読取図面保存」フォルダを常に空（またはこれから
 *          読み取る1人分のみ）に保つ運用とする。
 *          運用ルール：「読取図面保存」フォルダに他の人の図面が既にある場合は、
 *          それがなくなるまで自分の図面を置かない。
 *
 *   ② 担当者が入力シートの内容を確認・修正
 *      （検図者・課長・部長は「② 図面を一括登録・申請する」実行時の
 *        ダイアログで選択するため、シートへの直接入力は不要）
 *      ※ K列「AIチェック」には、①のOCR読取と同時にAIが検出した
 *        「気になる点」が件数で表示される（詳細はセルのノート参照）。
 *        これは正式な検図の代替ではなく、申請者自身が検図者へ回す前に
 *        ケアレスミスへ気づけるようにするための一次チェックであり、
 *        登録可否の判定には使用しない（21_OcrService.gs の
 *        checkDrawingConcerns_ を参照）。
 *
 *   ③ カスタムメニュー「② 図面を一括登録・申請する」
 *       → RegisterDrawings()
 *       → OCRで読み取った図面番号をそのまま使用して図面台帳へ登録
 *       → 検図者へ承認依頼メールを送信
 *
 * 【スクリプトプロパティ（事前設定が必要）】
 *   GEMINI_API_KEY          : Gemini API キー
 *   OCR_FOLDER_ID            : 「読取図面保存」フォルダ（PDFをアップロードするDriveフォルダ）のID
 *   PENDING_DRAWING_FOLDER_ID: 「申請中図面」フォルダ（OCR読み取り後に移動する先）のID
 *   APPROVAL_WEB_APP_URL     : 承認待ち一覧 Web アプリの URL（後で設定）
 *
 * 【入力シート（図面登録）列構成】（2026-07 マイグレーション済み：検図者・課長・部長列を削除）
 *   A=図面番号(DWG.NO.), B=図名(JPN), C=英名(NAME), D=ユニット名(UNIT),
 *   E=機械名(MODEL), F=材質(MATL.), G=縮尺(SCALE), H=図面サイズ(A0~A4),
 *   I=図面ファイルURL, J=AI-OCR読取結果, K=AIチェック（気になる点）, L=登録結果
 *
 * 【図面台帳列構成】（2026-07 マイグレーション済み：差戻し理由・差戻し回数列を削除／
 *                    申請者列を申請日の前に移動／図面サイズ列を追加／
 *                    改訂連番・最新フラグ列を削除／承認済みPDFファイルURL列を
 *                    図面ファイルURL列（U列＝「図面リンク」）に統合）
 *   A=フル図番, B=主図番, C=子図番, D=改訂記号,
 *   E=図名(JPN), F=英名(NAME), G=ユニット名(UNIT), H=機械名(MODEL),
 *   I=材質(MATL.), J=縮尺(SCALE), K=図面サイズ(A0~A4), L=ステータス,
 *   M=申請者, N=申請日, O=検図者, P=検図者承認日時, Q=課長, R=課長承認日時,
 *   S=部長, T=部長承認日時,
 *   U=図面リンク（承認前＝申請中図面／承認後＝承認済み図面）, V=申請バッチID
 */

// ============================================================
// ① OCR 読み取り（カスタムメニューから呼ばれる）
// ============================================================
function runOcrAndFillInputSheet() {
  const ui = SpreadsheetApp.getUi();

  // フォルダ ID の確認
  const folderId = PropertiesService.getScriptProperties().getProperty('OCR_FOLDER_ID');
  if (!folderId) {
    ui.alert(
      '設定エラー',
      'スクリプトプロパティ「OCR_FOLDER_ID」が設定されていません。\n' +
      'Apps Script エディタ →「プロジェクトの設定」→「スクリプトプロパティ」から設定してください。',
      ui.ButtonSet.OK
    );
    return;
  }
  const pendingFolderId = PropertiesService.getScriptProperties().getProperty('PENDING_DRAWING_FOLDER_ID');
  if (!pendingFolderId) {
    ui.alert(
      '設定エラー',
      'スクリプトプロパティ「PENDING_DRAWING_FOLDER_ID」が設定されていません。\n' +
      '「申請中図面」フォルダのIDを、Apps Script エディタ →「プロジェクトの設定」→' +
      '「スクリプトプロパティ」から設定してください。',
      ui.ButtonSet.OK
    );
    return;
  }

  // フォルダ取得
  let folder, pendingFolder;
  try {
    folder = DriveApp.getFolderById(folderId);
    pendingFolder = DriveApp.getFolderById(pendingFolderId);
  } catch (e) {
    ui.alert('エラー',
      `指定フォルダが見つかりません。OCR_FOLDER_ID・PENDING_DRAWING_FOLDER_ID を確認してください。\n${e.message}`,
      ui.ButtonSet.OK);
    return;
  }

  // フォルダ内の PDF 一覧を取得
  // ※ 共有ドライブ運用のため、「読取図面保存」フォルダには常に1人分の図面
  //   （＝これから読み取る対象）だけが置かれている前提で、フォルダ内の
  //   PDFをすべて処理対象とする（他人の図面との混在防止は、読み取り後に
  //   即座に「申請中図面」フォルダへ移動する運用でカバーする。
  //   「読取図面保存」フォルダに別の人の図面が既にある場合は、
  //   それがなくなるまで自分の図面を置かないという運用ルールとセットで使うこと）
  const files = folder.getFilesByType(MimeType.PDF);
  const pdfFiles = [];
  while (files.hasNext()) pdfFiles.push(files.next());

  if (pdfFiles.length === 0) {
    ui.alert('PDFなし', '指定フォルダに PDF ファイルが見つかりませんでした。', ui.ButtonSet.OK);
    return;
  }

  // ※ 以前はここで「入力シートに既に同じファイルURLが記録されている
  //   PDFを処理対象から除外する」重複防止フィルタを設けていたが、これは
  //   「同じPDFファイル（Drive上の同一ファイルID）を修正して同じ場所に
  //   保存し直し、再度読み取らせて上書きする」という運用と相性が悪く、
  //   本来読み取り直したいファイルまでもがスキップされてしまう問題があった。
  //   図面番号が既存の未申請行と一致する場合の上書き判定は、OCR結果を
  //   得たあとの writeOcrResultsToInputSheet_ / findPendingRowByDrawingNo_
  //   側で行うため、ここでの事前フィルタは不要と判断し撤廃した。
  //   フォルダ内の「読取図面保存」フォルダに置かれたPDFは、ボタンを押す
  //   たびに常に全件処理対象とする。
  const ss         = SpreadsheetApp.getActiveSpreadsheet();
  const inputSheet = ss.getSheetByName(SHEET_NAMES.INPUT_DRAWING);
  if (!inputSheet) throw new Error(`"${SHEET_NAMES.INPUT_DRAWING}" シートが見つかりません。`);

  const unprocessed = pdfFiles;

  // OCR 実行 → 完了したファイルから順に「申請中図面」フォルダへ移動する
  // （途中でエラーが起きても、既に読み取り済みのファイルは移動済みの状態を保つ）
  // OCR抽出に成功した場合は、続けてAIによる簡易検図チェック（気になる点の
  // 洗い出し）も実行する（21_OcrService.gs の checkDrawingConcerns_ を参照）。
  // チェックに失敗しても警告ログのみとし、OCR結果の転記自体は継続する。
  const results = [];
  const errors  = [];
  const moveFailures = []; // 「申請中図面」フォルダへの移動に失敗したファイル名
  unprocessed.forEach((file, idx) => {
    try {
      ss.toast(
        `OCR処理中... (${idx + 1}/${unprocessed.length}) ${file.getName()}`,
        'PDF読み取り', 10
      );
      const extracted = extractDrawingInfoByOcr_(file);
      let concerns = null;
      if (extracted) {
        try {
          concerns = checkDrawingConcerns_(file);
        } catch (concernErr) {
          console.warn(`AIチェックに失敗しました: ${file.getName()}`, concernErr);
        }
      }
      results.push({ file, extracted, concerns });
    } catch (e) {
      console.error(`OCR失敗: ${file.getName()}`, e);
      errors.push(file.getName());
      results.push({ file, extracted: null, concerns: null });
    } finally {
      // OCRの成否によらず、入力シートへの転記対象になった時点で
      // 「読取図面保存」フォルダからは取り除き、次の人が使えるようにする
      // ※ 共有ドライブ環境では DriveApp.File.moveTo() が親の付け替えに
      //   失敗することがあるため、より確実な Drive 詳細サービス（v3）の
      //   addParents/removeParents を使用する。移動に失敗した場合、その
      //   ファイルは「読取図面保存」フォルダに残り続け、次回のOCR実行時に
      //   再度処理対象となる（重複防止フィルタは撤廃済みのため）。この際、
      //   図面番号が既存の未申請行と一致すると「上書き」と判定されるが、
      //   trashOldInputFileIfPossible_ 側で自己破棄防止のガードが入って
      //   いるため、正常なファイルが誤ってゴミ箱へ送られることはない。
      try {
        Drive.Files.update({}, file.getId(), null, {
          addParents:        pendingFolder.getId(),
          removeParents:     folder.getId(),
          supportsAllDrives: true,
        });
      } catch (moveErr) {
        console.error(`「申請中図面」フォルダへの移動に失敗しました: ${file.getName()}`, moveErr);
        moveFailures.push(file.getName());
      }
    }
    Utilities.sleep(1000);  // API レート制限対策（1秒待機。OCR・AIチェック2回分のAPI呼び出し後にまとめて待機）
  });

  // 入力シートへ転記（図面番号が既存の未申請行と一致する場合は、その行を
  // 新しいOCR結果で上書きする＝図面を修正しての再読み取りとして扱う。
  // 詳細は writeOcrResultsToInputSheet_ を参照）
  const writeSummary = writeOcrResultsToInputSheet_(inputSheet, results);

  // 入力シートへ移動
  ss.setActiveSheet(inputSheet);

  const errMsg = errors.length > 0
    ? `\n\n以下のファイルは読み取りに失敗しました（手動入力してください）：\n${errors.join('\n')}`
    : '';
  const updatedMsg = writeSummary.updated.length > 0
    ? `\n\n以下は既存の未申請行を新しいOCR結果で上書きしました` +
      `（図面修正後の再読み取りとして扱い、不要になった旧ファイルはゴミ箱へ移動しました）：\n` +
      writeSummary.updated.map(no => `・${no}`).join('\n')
    : '';
  const moveFailMsg = moveFailures.length > 0
    ? `\n\n⚠️ 以下のファイルは「申請中図面」フォルダへの移動に失敗し、` +
      `「読取図面保存」フォルダに残っています（次回のOCR実行時に自動的に再試行されます）：\n` +
      moveFailures.map(name => `・${name}`).join('\n')
    : '';
  ui.alert(
    'OCR完了',
    `${results.length} 件の PDF を読み取りました` +
    `（新規追加 ${writeSummary.added.length} 件／既存行の上書き ${writeSummary.updated.length} 件）。\n` +
    `読み取り済みのファイルは「申請中図面」フォルダへ移動済みです。\n` +
    `黄色のセルを確認・修正のうえ、K列「AIチェック」の内容もあわせてご確認ください。\n` +
    `確認・修正後、「② 図面を一括登録・申請する」を実行してください。${updatedMsg}${errMsg}${moveFailMsg}`,
    ui.ButtonSet.OK
  );
}

// ============================================================
// ② 図面一括登録・承認申請（カスタムメニューから呼ばれる）
//    → 承認者選択ダイアログを開く
// ============================================================
function RegisterDrawings() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const inputSheet = ss.getSheetByName(SHEET_NAMES.INPUT_DRAWING);
  if (!inputSheet) throw new Error(`"${SHEET_NAMES.INPUT_DRAWING}" シートが見つかりません。`);

  // 未登録行の件数確認
  const pendingRows = getPendingRows_(inputSheet);
  if (pendingRows.length === 0) {
    ui.alert('対象なし',
      '登録対象の行がありません。\nL列（登録結果）が空の行が対象です。',
      ui.ButtonSet.OK);
    return;
  }

  // 図面番号の形式チェックのみ先行バリデーション
  const formatErrors = validateDrawingNos_(pendingRows);
  if (formatErrors.length > 0) {
    ui.alert('入力エラー',
      `以下の行に図面番号の形式エラーがあります。修正後に再実行してください。\n\n` +
      formatErrors.join('\n'),
      ui.ButtonSet.OK);
    return;
  }

  // 1回の申請は同一の主図番（機械）の図面をまとめたものである前提のため、
  // 異なる主図番が混在していないかを確認する
  const mainNos = new Set(
    pendingRows.map(row => parseDrawingNo_(String(row.data[COL_INPUT.DRAWING_NO - 1]).trim()).mainNo)
  );
  if (mainNos.size > 1) {
    ui.alert('入力エラー',
      '異なる主図番の図面が混在しています。1回の申請では同一の主図番の図面のみを\n' +
      `まとめて登録してください。\n\n対象の主図番：${Array.from(mainNos).join('、')}`,
      ui.ButtonSet.OK);
    return;
  }
  const mainNo = mainNos.values().next().value;

  // 既に説明文が登録されていれば、ダイアログの入力欄に初期表示する
  // （修正の有無に関わらず、登録処理時に毎回機械台帳へ再書き込みする）
  const mainSheet = ss.getSheetByName(SHEET_NAMES.MAIN_NUMBER);
  const mainInfo  = mainSheet ? getMainNumberInfoFull_(mainSheet, mainNo) : null;
  const existingDescription = mainInfo ? mainInfo.description : '';

  // ユーザーマスタからユーザーリストを取得
  const userList = getUserListForDialog_();
  if (userList.length === 0) {
    ui.alert('エラー',
      'ユーザーマスタに氏名が登録されていません。\n先にユーザーマスタのB列（氏名）を入力してください。',
      ui.ButtonSet.OK);
    return;
  }

  // 承認者選択ダイアログを表示
  const html = HtmlService.createTemplateFromFile('22_ApprovalDialog');
  html.pendingCount = pendingRows.length;
  html.userList     = JSON.stringify(userList);
  html.mainNo       = JSON.stringify(mainNo);
  html.description  = JSON.stringify(existingDescription);

  SpreadsheetApp.getUi().showModalDialog(
    html.evaluate().setWidth(420).setHeight(640),
    '承認者を選択して一括登録'
  );
}

// ============================================================
// ダイアログから呼ばれる：承認者を一括セットして登録処理を実行
// ============================================================
function setApproversAndRegister(form) {
  try {
    if (!form.reviewer || !form.manager || !form.director) {
      return { success: false, message: '承認者の選択が不完全です。' };
    }

    const ss         = SpreadsheetApp.getActiveSpreadsheet();
    const inputSheet = ss.getSheetByName(SHEET_NAMES.INPUT_DRAWING);
    const dbSheet    = ss.getSheetByName(SHEET_NAMES.DRAWING_DB);
    if (!inputSheet) throw new Error(`"${SHEET_NAMES.INPUT_DRAWING}" シートが見つかりません。`);
    if (!dbSheet)    throw new Error(`"${SHEET_NAMES.DRAWING_DB}" シートが見つかりません。`);

    // 未登録行を再取得
    const pendingRows = getPendingRows_(inputSheet);
    if (pendingRows.length === 0) {
      return { success: false, message: '登録対象の行が見つかりませんでした。' };
    }

    // 承認者（検図者・課長・部長）はダイアログで選択された form.reviewer/manager/director を
    // そのまま図面台帳への書き込みに使用する（入力シートには列自体が存在しないため書き戻し不要）

    // データを再取得して全項目バリデーション
    const updatedRows = getPendingRows_(inputSheet);
    const errors = validatePendingRows_(updatedRows);
    if (errors.length > 0) {
      return {
        success: false,
        message: `入力エラーがあります。\n\n${errors.join('\n')}`,
      };
    }

    // ── 排他制御 ──────────────────────────────────
    const lock = LockService.getScriptLock();
    lock.waitLock(15000);

    try {
      const today        = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
      const batchId       = Utilities.getUuid(); // この1回の申請をまとめて識別するID
      const newDbRows    = [];
      const inputUpdates = [];

      updatedRows.forEach(row => {
        const drawingNo = String(row.data[COL_INPUT.DRAWING_NO - 1]).trim();
        const parsed    = parseDrawingNo_(drawingNo);

        // 入力シートのFILE_URL列はHYPERLINK数式なので、getValues()の表示テキストではなく
        // 数式から実際のURLを取り出す
        const fileUrlFormula = inputSheet.getRange(row.rowIndex, COL_INPUT.FILE_URL).getFormula();
        const fileUrl = extractUrlFromCellFormula_(fileUrlFormula) ||
                        String(row.data[COL_INPUT.FILE_URL - 1]).trim();

        // AIチェック（気になる点）はK列のセルノートに詳細を保持しているため、
        // ここで取得して図面台帳へそのまま引き継ぐ（検図者への参考情報。
        // 21_OcrService.gs の checkDrawingConcerns_ を参照）
        const aiCheckNote = inputSheet.getRange(row.rowIndex, COL_INPUT.AI_CHECK).getNote();

        // 特徴属性（類似図面検索用）。入力シートM列にJSON文字列として保持
        // されているものを、そのまま図面台帳へ引き継ぐ（21_OcrService.gs の
        // extractDrawingInfoByOcr_ / normalizeDrawingAttributes_ を参照）。
        // summaryのみ複製し、シートを直接見たときに一覧性を確保する。
        const attributesJson = String(row.data[COL_INPUT.ATTRIBUTES - 1] || '').trim();
        let attributesSummary = '';
        if (attributesJson) {
          try {
            attributesSummary = String(JSON.parse(attributesJson).summary || '');
          } catch (e) {
            console.warn(`特徴属性JSONのパースに失敗しました（${drawingNo}）: ${e.message}`);
          }
        }

        newDbRows.push([
          drawingNo,                                // A: フル図番
          parsed.mainNo,                            // B: 主図番
          parsed.subNo,                             // C: 子図番
          parsed.revMark,                           // D: 改訂記号
          row.data[COL_INPUT.NAME_JP    - 1],       // E: 図名（JPN）
          row.data[COL_INPUT.NAME_EN    - 1],       // F: 英名（NAME）
          row.data[COL_INPUT.UNIT_NAME  - 1],       // G: ユニット名（UNIT）
          row.data[COL_INPUT.MODEL_NAME - 1],       // H: 機械名（MODEL）
          row.data[COL_INPUT.MATERIAL   - 1],       // I: 材質（MATL.）
          row.data[COL_INPUT.SCALE      - 1],       // J: 縮尺（SCALE）
          String(row.data[COL_INPUT.DRAWING_SIZE - 1]).trim().toUpperCase(), // K: 図面サイズ
          '検図中',                                 // L: ステータス
          getFullNameFromMaster_() || '',           // M: 申請者
          today,                                    // N: 申請日
          form.reviewer,                            // O: 検図者
          '',                                       // P: 検図者承認日時
          form.manager,                             // Q: 課長
          '',                                       // R: 課長承認日時
          form.director,                            // S: 部長
          '',                                       // T: 部長承認日時
          buildFileLinkFormula_(fileUrl, '📄図面を開く'), // U: 図面リンク（申請中図面。承認完了時に承認済みリンクへ上書き）
          batchId,                                  // V: 申請バッチID
          aiCheckNote,                              // W: AIチェック（検図者への参考情報）
          attributesJson,                           // X: 特徴属性（JSON。類似図面検索用）
          attributesSummary,                        // Y: 特徴属性の要約
        ]);

        inputUpdates.push({
          rowIndex: row.rowIndex,
          drawingNo,
          reviewer: form.reviewer,
        });
      });

      // 図面台帳へ一括書き込み
      const startRow = dbSheet.getLastRow() + 1;
      dbSheet.getRange(startRow, 1, newDbRows.length, DRAWING_DB_COL_COUNT).setValues(newDbRows);
      formatDbRows_(dbSheet, startRow, newDbRows.length);

      // 検索用索引（図面索引・承認待ち索引）を更新する（05_SearchIndex.gs を参照）。
      // newDbRows は組み立て時点でFILE_URLが数式文字列のままなのでそのまま渡せる。
      // 索引はあくまで検索高速化のための付随データのため、失敗しても登録処理
      // 自体は継続する。
      try {
        newDbRows.forEach(rowValues => {
          upsertDrawingIndex_(rowValues);
          upsertPendingIndex_(rowValues);
        });
      } catch (indexErr) {
        console.warn(`検索用索引の更新に失敗しました: ${indexErr.message}`);
      }

      // 説明文を機械台帳へ保存する（1申請＝同一主図番のため、newDbRowsの先頭行から取得）
      // 修正の有無に関わらず、登録処理のたびに毎回上書きする
      const description = String(form.description || '').trim();
      const mainNo       = newDbRows.length > 0 ? newDbRows[0][COL_DB.MAIN_NO - 1] : '';
      const mainSheet    = ss.getSheetByName(SHEET_NAMES.MAIN_NUMBER);
      if (mainSheet && mainNo) {
        const mainInfo = getMainNumberInfoFull_(mainSheet, mainNo);
        if (mainInfo) {
          mainSheet.getRange(mainInfo.rowIndex, COL_MAIN.DESCRIPTION).setValue(description);
        } else {
          console.warn(`機械台帳に「${mainNo}」が見つからず、説明文を保存できませんでした。`);
        }
      }

      // 入力シートへ書き戻し
      inputUpdates.forEach(u => {
        inputSheet.getRange(u.rowIndex, COL_INPUT.REG_RESULT).setValue('登録済み');
        inputSheet.getRange(u.rowIndex, 1, 1, INPUT_DRAWING_COL_COUNT).setBackground('#e8f5e9');
      });

      // 検図者へ通知メール送信（この申請バッチのIDをURLに付与）
      sendReviewerNotifications_(inputUpdates, batchId, getFullNameFromMaster_() || '', description);

      // 図面台帳へ移動・登録行を選択
      ss.setActiveSheet(dbSheet);
      dbSheet.setActiveRange(
        dbSheet.getRange(startRow, 1, newDbRows.length, DRAWING_DB_COL_COUNT)
      );

      return {
        success: true,
        message: `${newDbRows.length} 件を図面台帳へ登録しました。\n` +
                 `検図者（${form.reviewer}）へ承認依頼メールを送信しました。`,
      };

    } finally {
      lock.releaseLock();
    }

  } catch (e) {
    console.error('setApproversAndRegister error:', e);
    return { success: false, message: `エラーが発生しました：${e.message}` };
  }
}

// ============================================================
// プライベートヘルパー
// ============================================================

/**
 * OCR 結果・AIチェック結果を入力シートへ転記する。
 *
 * 【図面修正 → 再OCR読み取りへの対応】
 *   OCRで読み取った図面番号が、入力シート内の「未申請（登録結果が空）」の
 *   既存行と一致する場合は、その行を新しいOCR結果・AIチェック結果で
 *   丸ごと上書きする（＝申請者が図面を修正してから読み取り直したケースと
 *   して扱う）。これにより、同じ図面のために重複行ができるのを防ぐ。
 *   この際、上書きで不要になった旧PDFファイル（前回の読み取りで
 *   「申請中図面」フォルダへ移動されたもの）はゴミ箱へ移動する
 *   （trashOldInputFileIfPossible_ を参照。完全削除ではなく復元可能な形）。
 *   一致する既存行が見つからない場合（初回読み取り、または今回OCR自体が
 *   失敗し図面番号が取得できなかった場合）は、従来どおり新規行として追加する。
 *
 * @returns {{added: string[], updated: string[]}}
 *   added/updated はそれぞれ図面番号（またはOCR失敗時はファイル名）の配列。
 *   呼び出し元（runOcrAndFillInputSheet）の完了メッセージ表示に使用する。
 */
function writeOcrResultsToInputSheet_(inputSheet, results) {
  const added   = [];
  const updated = [];
  // 同一バッチ内の複数ファイルが同じ図面番号を持っていた場合に、同じ既存行へ
  // 二重にマッチしてしまわないよう、このバッチ内で使用済みの行番号を記録する
  const consumedRowIndexes = new Set();

  results.forEach(({ file, extracted, concerns }) => {
    const fileUrl = file.getUrl();
    const ocrNote = extracted
      ? (extracted.hasWarning ? '⚠️ 要確認（一部読取不可）' : '✅ OCR済み')
      : '❌ 読取失敗（手動入力してください）';

    // AIチェック（気になる点）のステータス文言。詳細はセルのノートに記載する
    // ※ concernsがnullなのは「実施していない」のではなく「実施したが失敗した」
    //   場合（Gemini APIの一時的なエラー等）である点に注意。checkDrawingConcerns_
    //   は失敗時に1回リトライしたうえでnullを返す設計のため、ここでnullなら
    //   一時的な問題である可能性が高く、再度「①図面一括読取」をやり直すことで
    //   解消することが多い。
    let aiCheckNote;
    if (!extracted) {
      aiCheckNote = '―（OCR失敗のためスキップ）';
    } else if (!concerns) {
      aiCheckNote = '❌ チェック失敗（再度読み取りをお試しください）';
    } else if (concerns.points.length === 0) {
      aiCheckNote = '✅ 特になし';
    } else {
      aiCheckNote = `⚠️ ${concerns.points.length}件の確認事項`;
    }

    // 特徴属性（類似図面検索用の下ごしらえデータ）。抽出に失敗していても
    // OCR自体は継続する付随データのため、欠損時は空文字にする
    const attributesJson = (extracted && extracted.attributes)
      ? JSON.stringify(extracted.attributes)
      : '';

    const newRowValues = [
      extracted?.drawingNo   || '',  // A: 図面番号（DWG.NO.）
      extracted?.drawingName || '',  // B: 図名（JPN）
      extracted?.nameEn      || '',  // C: 英名（NAME）
      extracted?.unitName    || '',  // D: ユニット名（UNIT）
      extracted?.modelName   || '',  // E: 機械名（MODEL）
      extracted?.material    || '',  // F: 材質（MATL.）
      extracted?.scale       || '',  // G: 縮尺（SCALE）
      extracted?.drawingSize || '',  // H: 図面サイズ（A0〜A4）
      buildFileLinkFormula_(fileUrl, '📄図面を開く'), // I: 図面ファイルURL（リンク表示）
      ocrNote,                       // J: AI-OCR読取結果
      aiCheckNote,                   // K: AIチェック（要確認点。詳細はノート参照）
      '',                            // L: 登録結果
      attributesJson,                // M: 特徴属性（JSON。類似図面検索用）
    ];

    // 図面番号が既存の未申請行と一致するか探す（OCRが成功した場合のみ）
    const matchedRow = (extracted && extracted.drawingNo)
      ? findPendingRowByDrawingNo_(inputSheet, extracted.drawingNo, consumedRowIndexes)
      : null;

    let targetRow;
    if (matchedRow) {
      targetRow = matchedRow.rowIndex;
      consumedRowIndexes.add(targetRow);

      // 上書きで不要になる旧ファイルをゴミ箱へ移動（復元可能な形で破棄）。
      // ただし、今回処理中のファイルそのものが既存行のファイルと同一の
      // 場合は自己破棄になってしまうため、trashOldInputFileIfPossible_
      // 内でスキップされる（詳細は同関数のコメントを参照）。
      trashOldInputFileIfPossible_(matchedRow.fileUrl, file.getId(), extracted.drawingNo);

      // 前回のOCRで付いた黄色ハイライト等が、今回は不要になった項目にも
      // 残ってしまわないよう、行の背景色を一旦リセットしてから上書きする
      inputSheet.getRange(targetRow, 1, 1, INPUT_DRAWING_COL_COUNT).setBackground(null);
      inputSheet.getRange(targetRow, 1, 1, newRowValues.length).setValues([newRowValues]);

      updated.push(extracted.drawingNo);
    } else {
      inputSheet.appendRow(newRowValues);
      targetRow = inputSheet.getLastRow();
      added.push(extracted?.drawingNo || file.getName());
    }

    if (extracted) {
      // 読み取れなかった項目を黄色でハイライト
      if (!extracted.drawingNo)   highlightCell_(inputSheet, targetRow, COL_INPUT.DRAWING_NO);
      if (!extracted.drawingName) highlightCell_(inputSheet, targetRow, COL_INPUT.NAME_JP);
      if (!extracted.nameEn)      highlightCell_(inputSheet, targetRow, COL_INPUT.NAME_EN);
      if (!extracted.unitName)    highlightCell_(inputSheet, targetRow, COL_INPUT.UNIT_NAME);
      if (!extracted.modelName)   highlightCell_(inputSheet, targetRow, COL_INPUT.MODEL_NAME);
      if (!extracted.material)    highlightCell_(inputSheet, targetRow, COL_INPUT.MATERIAL);
      if (!extracted.scale)       highlightCell_(inputSheet, targetRow, COL_INPUT.SCALE);
      if (!extracted.drawingSize) highlightCell_(inputSheet, targetRow, COL_INPUT.DRAWING_SIZE);
    } else {
      // 読取失敗行全体をハイライト
      inputSheet.getRange(targetRow, 1, 1, COL_INPUT.FILE_URL).setBackground('#fff2cc');
    }

    // AIチェックの結果に応じてK列の背景色を出し分ける：
    //   気になる点あり（水色）／チェック失敗（薄赤・要再実行）／それ以外（無色）
    // 黄色＝必須項目未入力の警告とは意図的に色を分け、それぞれの状態が
    // 視覚的に区別できるようにしている
    const aiCheckCell = inputSheet.getRange(targetRow, COL_INPUT.AI_CHECK);
    if (concerns && concerns.points.length > 0) {
      aiCheckCell
        .setBackground('#e3f2fd')
        .setNote(concerns.points.map((p, i) => `${i + 1}. ${p}`).join('\n'));
    } else if (extracted && !concerns) {
      // 実施したが失敗したケース（一時的なAPIエラー等）。要再実行であることが
      // 分かるよう、必須項目未入力の警告色（黄）とも情報表示色（水色）とも
      // 異なる薄赤で目立たせる
      aiCheckCell
        .setBackground('#fdecea')
        .setNote('AIチェックの実行に失敗しました（一時的なAPIエラーの可能性があります）。\n' +
                 '「① 図面一括読取（OCR）」を再実行すると、この図面のみ再チェックされます。');
    } else {
      aiCheckCell.setBackground(null).clearNote();
    }
  });

  return { added, updated };
}

/**
 * 入力シート内で「未申請（登録結果が空）」かつ図面番号が一致する行を探す
 * （図面修正後の再OCR読み取りで、上書き対象の行を特定するために使用）
 * @param {Sheet} inputSheet
 * @param {string} drawingNo
 * @param {Set<number>} excludeRowIndexes - 同一バッチ内で既にマッチ済みの行番号（除外対象）
 * @returns {{rowIndex: number, fileUrl: string}|null}
 */
function findPendingRowByDrawingNo_(inputSheet, drawingNo, excludeRowIndexes) {
  const lastRow = inputSheet.getLastRow();
  if (lastRow < GLOBAL_ROW.DATA_START) return null;

  const range    = inputSheet.getRange(GLOBAL_ROW.DATA_START, 1, lastRow - GLOBAL_ROW.DATA_START + 1, INPUT_DRAWING_COL_COUNT);
  const values   = range.getValues();
  const formulas = range.getFormulas();

  for (let i = 0; i < values.length; i++) {
    const rowIndex = GLOBAL_ROW.DATA_START + i;
    if (excludeRowIndexes.has(rowIndex)) continue;

    const rowDrawingNo = String(values[i][COL_INPUT.DRAWING_NO - 1]).trim();
    const regResult    = String(values[i][COL_INPUT.REG_RESULT - 1]).trim();
    if (rowDrawingNo === drawingNo && !regResult) {
      const fileUrl = extractUrlFromCellFormula_(formulas[i][COL_INPUT.FILE_URL - 1]) ||
                      String(values[i][COL_INPUT.FILE_URL - 1]).trim();
      return { rowIndex, fileUrl };
    }
  }
  return null;
}

/**
 * 図面修正に伴う再OCRで不要になった旧PDFファイルをゴミ箱へ移動する
 * （完全削除ではなく setTrashed(true) による復元可能な破棄。ファイルが
 *   既に手動で削除・移動されているなど、取得できない場合はエラーにせず
 *   警告ログのみ出して処理を継続する）
 *
 * 【重要：自己破棄の防止】
 *   マッチした既存行（上書き対象）に記録されていたファイルIDが、今まさに
 *   処理中のファイルと同一の場合は、ゴミ箱への移動を行わない。
 *   これは、前回のOCR実行時に「申請中図面」フォルダへの移動が何らかの
 *   理由で失敗し、同じファイルが「読取図面保存」フォルダに残ったまま
 *   次回のOCRで再度読み取られた場合に起こり得る。この場合、図面番号が
 *   一致するため「図面修正後の再読み取り」と判定されるが、実際には
 *   修正後の別ファイルではなく今回処理中の有効なファイルそのものであり、
 *   区別せずに破棄すると、正常に読み取られたファイルが図面台帳への
 *   登録前にゴミ箱へ送られてしまう（＝図面リンクが「ゴミ箱にあります」
 *   という状態になる不具合の原因だった）。
 *
 * @param {string} oldFileUrl - 上書き対象の既存行に記録されていたファイルURL
 * @param {string} currentFileId - 今回処理中のファイルのID（自己破棄防止の比較用）
 * @param {string} drawingNo - ログ表示用の図面番号
 */
function trashOldInputFileIfPossible_(oldFileUrl, currentFileId, drawingNo) {
  if (!oldFileUrl) return;
  try {
    const fileId = extractDriveFileId_(oldFileUrl); // 25_PdfApprovalStamp.gs で定義済み
    if (!fileId) return;

    if (fileId === currentFileId) {
      console.warn(
        `図面番号「${drawingNo}」の既存行は今回処理中のファイルと同一のため、` +
        `ゴミ箱への移動をスキップしました（前回のフォルダ移動が失敗していた可能性があります）。`
      );
      return;
    }

    const file = DriveApp.getFileById(fileId);
    file.setTrashed(true);
  } catch (e) {
    console.warn(`図面修正に伴う旧ファイルのゴミ箱移動に失敗しました（${drawingNo}）: ${e.message}`);
  }
}

/**
 * 入力シートから未登録行（L列＝登録結果が空）を取得する
 * @returns {Array<{rowIndex, data}>}
 */
function getPendingRows_(inputSheet) {
  const lastRow = inputSheet.getLastRow();
  if (lastRow < GLOBAL_ROW.DATA_START) return [];

  const values = inputSheet
    .getRange(GLOBAL_ROW.DATA_START, 1, lastRow - GLOBAL_ROW.DATA_START + 1, INPUT_DRAWING_COL_COUNT)
    .getValues();

  return values
    .map((row, i) => ({ rowIndex: GLOBAL_ROW.DATA_START + i, data: row }))
    .filter(({ data }) => {
      const regResult = String(data[COL_INPUT.REG_RESULT - 1]).trim();
      const drawingNo = String(data[COL_INPUT.DRAWING_NO - 1]).trim();
      return !regResult && drawingNo;  // 登録結果が空 かつ 図面番号がある行
    });
}

/**
 * 未登録行のバリデーション
 * @returns {string[]} エラーメッセージの配列（空なら問題なし）
 */
function validatePendingRows_(pendingRows) {
  const errors        = [];
  const registeredNos = new Set();  // 今回の登録内での重複チェック用

  // 既存台帳のフル図番を取得
  const ss        = SpreadsheetApp.getActiveSpreadsheet();
  const dbSheet   = ss.getSheetByName(SHEET_NAMES.DRAWING_DB);
  const existingNos = getExistingDrawingNos_(dbSheet);

  pendingRows.forEach(row => {
    const r           = row.rowIndex;
    const drawingNo   = String(row.data[COL_INPUT.DRAWING_NO   - 1]).trim();
    const nameJp      = String(row.data[COL_INPUT.NAME_JP      - 1]).trim();
    const drawingSize = String(row.data[COL_INPUT.DRAWING_SIZE - 1]).trim().toUpperCase();
    const fileUrl     = String(row.data[COL_INPUT.FILE_URL     - 1]).trim();

    // 必須項目チェック（検図者・課長・部長はダイアログ側で選択必須のため、
    // ここでは入力シート由来の項目のみをチェックする）
    const missing = [];
    if (!drawingNo)   missing.push('図面番号(DWG.NO.)');
    if (!nameJp)      missing.push('図名(JPN)');
    if (!drawingSize) missing.push('図面サイズ');
    if (!fileUrl)     missing.push('図面ファイルURL');
    if (missing.length > 0) {
      errors.push(`${r}行目：${missing.join('・')} が未入力です`);
      return;
    }

    // 図面サイズの値チェック（A0〜A4のみ許容）
    if (VALID_DRAWING_SIZES.indexOf(drawingSize) === -1) {
      errors.push(
        `${r}行目：図面サイズ「${drawingSize}」が不正です。` +
        `${VALID_DRAWING_SIZES.join('・')} のいずれかを入力してください。`
      );
      return;
    }

    // 図面番号の形式チェック
    const parsed = parseDrawingNo_(drawingNo);
    if (!parsed.isValid) {
      errors.push(`${r}行目：${parsed.errorMsg}`);
      return;
    }

    // 図面台帳との重複チェック
    if (existingNos.has(drawingNo)) {
      errors.push(`${r}行目：図面番号「${drawingNo}」は図面台帳に既に登録されています`);
      return;
    }

    // 今回の登録内での重複チェック
    if (registeredNos.has(drawingNo)) {
      errors.push(`${r}行目：図面番号「${drawingNo}」が入力シート内で重複しています`);
      return;
    }
    registeredNos.add(drawingNo);
  });

  return errors;
}

/**
 * 図面台帳に登録済みのフル図番一覧をSetで返す
 */
function getExistingDrawingNos_(dbSheet) {
  const lastRow = dbSheet.getLastRow();
  if (lastRow < GLOBAL_ROW.DATA_START) return new Set();
  const nos = dbSheet
    .getRange(GLOBAL_ROW.DATA_START, COL_DB.FULL_NO, lastRow - GLOBAL_ROW.DATA_START + 1, 1)
    .getValues().flat().filter(v => v);
  return new Set(nos.map(String));
}

/**
 * 図面台帳の新規追加行に書式を適用する
 */
function formatDbRows_(dbSheet, startRow, count) {
  const range = dbSheet.getRange(startRow, 1, count, DRAWING_DB_COL_COUNT);
  range.setBorder(true, true, true, true, true, true,
    '#cccccc', SpreadsheetApp.BorderStyle.SOLID);

  // 中央揃えの列
  [COL_DB.FULL_NO, COL_DB.MAIN_NO, COL_DB.SUB_NO, COL_DB.REV_MARK,
   COL_DB.STATUS, COL_DB.APP_DATE, COL_DB.SCALE].forEach(col => {
    dbSheet.getRange(startRow, col, count, 1).setHorizontalAlignment('center');
  });
}

/**
 * 検図者ごとに通知メールを送信する
 * @param {Array} inputUpdates
 * @param {string} batchId - この一括申請の識別子（承認画面のURLに付与する）
 * @param {string} [applicantName] - 申請者（この登録操作を行った本人）の氏名。
 *   指定があれば、送信者の表示名に反映する。
 * @param {string} [description] - 説明文（機械台帳に保存した内容と同一）。
 *   指定があればメール本文に含める。
 */
function sendReviewerNotifications_(inputUpdates, batchId, applicantName, description) {
  const baseUrl = PropertiesService.getScriptProperties()
    .getProperty('APPROVAL_WEB_APP_URL') || '（承認WebアプリURLは後で設定予定）';
  const webAppUrl = appendBatchIdToUrl_(baseUrl, batchId);
  const descriptionBlock = description ? `■ 説明文\n${description}\n\n` : '';

  // 検図者ごとにグルーピング
  const grouped = {};
  inputUpdates.forEach(u => {
    if (!grouped[u.reviewer]) grouped[u.reviewer] = [];
    grouped[u.reviewer].push(u);
  });

  // ユーザーマスタから氏名→メールアドレスのマップを取得
  const emailMap = getEmailMapFromUserMaster_();

  const senderName = applicantName
    ? `図面承認・採番・出図管理システム（${applicantName}の操作）`
    : '図面承認・採番・出図管理システム';

  Object.entries(grouped).forEach(([reviewer, updates]) => {
    const email = emailMap[reviewer];
    if (!email) {
      console.warn(
        `検図者「${reviewer}」のメールアドレスがユーザーマスタに見つかりません。スキップします。`
      );
      return;
    }

    const count    = updates.length;
    const itemList = updates.map(u => `  ・${u.drawingNo}`).join('\n');

    const subject = `【図面承認依頼】検図をお願いします（${count}件）`;
    const body    =
      `${reviewer} 様\n\n` +
      `以下の図面について、検図をお願いします。\n\n` +
      descriptionBlock +
      `■ 対象図面（${count}件）\n${itemList}\n\n` +
      `■ 承認待ち一覧（Webアプリ）\n${webAppUrl}\n\n` +
      `上記URLにアクセスし、Googleアカウントでログインして承認・差戻しを行ってください。\n\n` +
      `---\n図面承認・採番・出図管理システム（自動送信）`;

    try {
      GmailApp.sendEmail(email, subject, body, { name: senderName });
    } catch (e) {
      console.error(`メール送信失敗（${reviewer} / ${email}）:`, e);
    }
  });
}

/**
 * ユーザーマスタから 氏名→メールアドレス の Map を返す
 */
function getEmailMapFromUserMaster_() {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAMES.USER_MASTER);
  if (!sheet) return {};
  const lastRow = sheet.getLastRow();
  if (lastRow < GLOBAL_ROW.DATA_START) return {};
  const data = sheet
    .getRange(GLOBAL_ROW.DATA_START, 1, lastRow - GLOBAL_ROW.DATA_START + 1, 2)
    .getValues();
  const map = {};
  data.forEach(row => {
    const email = String(row[0]).trim();
    const name  = String(row[1]).trim();
    if (email && name) map[name] = email;
  });
  return map;
}

/**
 * セルを黄色でハイライトする（要確認の印）
 */
function highlightCell_(sheet, row, col) {
  sheet.getRange(row, col).setBackground('#fff2cc');
}

// ============================================================
// 【デバッグ用】
// ============================================================
function TEST_RegisterDrawings() {
  RegisterDrawings();
}

function TEST_getPendingRows() {
  const ss         = SpreadsheetApp.getActiveSpreadsheet();
  const inputSheet = ss.getSheetByName(SHEET_NAMES.INPUT_DRAWING);
  const rows = getPendingRows_(inputSheet);
  console.log(`未登録行数: ${rows.length}`);
  rows.forEach(r => console.log(`行${r.rowIndex}:`, r.data[0]));
}

/**
 * 図面番号の形式チェックのみ（承認者未入力は許容）
 * ダイアログ表示前の事前チェック用
 * @returns {string[]} エラーメッセージの配列
 */
function validateDrawingNos_(pendingRows) {
  const errors        = [];
  const registeredNos = new Set();
  const ss            = SpreadsheetApp.getActiveSpreadsheet();
  const dbSheet       = ss.getSheetByName(SHEET_NAMES.DRAWING_DB);
  const existingNos   = getExistingDrawingNos_(dbSheet);

  pendingRows.forEach(row => {
    const r         = row.rowIndex;
    const drawingNo = String(row.data[COL_INPUT.DRAWING_NO - 1]).trim();

    if (!drawingNo) {
      errors.push(`${r}行目：図面番号(DWG.NO.)が未入力です`);
      return;
    }
    const parsed = parseDrawingNo_(drawingNo);
    if (!parsed.isValid) {
      errors.push(`${r}行目：${parsed.errorMsg}`);
      return;
    }
    if (existingNos.has(drawingNo)) {
      errors.push(`${r}行目：図面番号「${drawingNo}」は図面台帳に既に登録されています`);
      return;
    }
    if (registeredNos.has(drawingNo)) {
      errors.push(`${r}行目：図面番号「${drawingNo}」が入力シート内で重複しています`);
      return;
    }
    registeredNos.add(drawingNo);
  });
  return errors;
}

/**
 * ユーザーマスタから氏名リストを取得してダイアログ用に返す
 * @returns {Array<{name}>}
 */
function getUserListForDialog_() {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAMES.USER_MASTER);
  if (!sheet) return [];

  const lastRow = sheet.getLastRow();
  if (lastRow < GLOBAL_ROW.DATA_START) return [];

  const data = sheet
    .getRange(GLOBAL_ROW.DATA_START, 1, lastRow - GLOBAL_ROW.DATA_START + 1, 2)
    .getValues();

  return data
    .filter(row => row[0] && row[1])  // メール・氏名の両方がある行のみ
    .map(row => ({ name: String(row[1]).trim() }));
}