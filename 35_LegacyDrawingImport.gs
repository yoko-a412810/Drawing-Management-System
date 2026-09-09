/**
 * ============================================================
 * 図面承認・採番・出図管理システム
 * ファイル: 40_LegacyDrawingImport.gs  ―  過去図面の一括インポートモジュール
 * ============================================================
 *
 * 【役割】
 *   本システムの運用開始以前に、紙の表題欄で検図・承認が完了している過去図面を、
 *   正規の承認フロー（検図→課長→部長）を経ずに図面台帳へ直接登録する。
 *
 * 【通常ルートとの違い】
 *   - 通常ルート（20_DrawingApproveRequest.gs）は「申請者＝ログインユーザー」を
 *     前提とするが、過去図面は表題欄の「設計者」が申請者に相当し、システムへの
 *     登録操作者（ログインユーザー）とは一致しない。そのため、通常ルートでは
 *     読み取らない「設計者・検図者・承認者」の氏名・日付をOCRで追加抽出し、
 *     そのまま図面台帳へ転記する（ユーザーマスタとの照合は行わない。氏名が
 *     ユーザーマスタに存在しない場合や英語表記の場合でも、読み取った文字列を
 *     そのまま記録する）。
 *   - ステータスは「過去図面」（STATUS_LEGACY、00_Config.gs で定義）で登録する。
 *     このステータスは承認Webアプリの一覧取得・承認/差戻し処理のいずれの対象にも
 *     一切含まれない（23_ApprovalWebApp.gs の getMyPendingApprovals 等）。
 *   - 25_PdfApprovalStamp.gs によるPDFスタンプ処理は行わない。読み取った元PDFを
 *     そのまま「承認済み図面」フォルダのサブフォルダへコピーするのみ。
 *   - 課長（MANAGER・MANAGER_DATE）欄は、紙の表題欄に課長の押印欄が存在しない
 *     ため、空欄のままとする。
 *
 * 【抜け道対策：移行期間限定の提供】
 *   このルートは、設計者・検図者・承認者の情報をOCRでそのまま台帳に書き込める
 *   ため、悪用すると正規の承認フローを経ずに新規図面を登録できてしまう。
 *   これを防ぐため、スクリプトプロパティ「LEGACY_IMPORT_DEADLINE」で定めた
 *   期限を過ぎると、メニュー自体が表示されなくなり（01_Menu.gs の onOpen）、
 *   関数を直接実行しても冒頭のガード処理（legacyImportDeadlineError_）で
 *   必ず中断される。
 *
 * 【前提条件】
 *   過去図面をインポートするより先に、対象機械の主図番を通常の主図番発行機能
 *   （11_MainNumberDialog.html）で発行しておくこと。機械台帳に存在しない
 *   主図番の図面は登録時にエラーとして弾かれる。
 *
 * 【スクリプトプロパティ（事前設定が必要）】
 *   LEGACY_OCR_FOLDER_ID     : 過去図面のPDFを置く読み取り対象フォルダのID
 *   LEGACY_IMPORT_DEADLINE   : 本機能の提供期限（例："2027-03-31"）。未設定の場合は利用不可
 *   GEMINI_API_KEY           : 21_OcrService.gs と共通（OCRに使用）
 *   APPROVED_DRAWING_FOLDER_ID: 25_PdfApprovalStamp.gs と共通（PDFの保存先）
 *
 * 【過去図面登録申請（入力）シートの列構成】
 *   A=図面番号(DWG.NO.), B=図名(JPN), C=英名(NAME), D=ユニット名(UNIT),
 *   E=機械名(MODEL), F=材質(MATL.), G=縮尺(SCALE), H=図面サイズ(A0~A4),
 *   I=図面ファイルURL, J=設計者, K=設計日, L=検図者, M=検図日,
 *   N=承認者, O=承認日, P=AI-OCR読取結果, Q=登録結果, R=特徴属性（JSON）
 */

// ============================================================
// 移行期間の判定
// ============================================================

/**
 * 過去図面インポート機能が現在利用できない場合、その理由メッセージを返す
 * （利用可能な場合は null を返す）
 * @returns {string|null}
 */
function legacyImportDeadlineError_() {
  const deadline = PropertiesService.getScriptProperties().getProperty('LEGACY_IMPORT_DEADLINE');
  if (!deadline) {
    return (
      '過去図面インポート機能の利用期限（スクリプトプロパティ「LEGACY_IMPORT_DEADLINE」）が' +
      '設定されていないため、利用できません。管理者に設定を依頼してください。'
    );
  }
  const deadlineDate = new Date(`${deadline}T23:59:59`);
  if (isNaN(deadlineDate.getTime())) {
    return `スクリプトプロパティ「LEGACY_IMPORT_DEADLINE」の値（${deadline}）が日付として解釈できません。`;
  }
  if (new Date() > deadlineDate) {
    return `過去図面インポート機能の移行期間（〜${deadline}）は終了しました。`;
  }
  return null;
}

/**
 * 過去図面インポート機能が現在利用可能かどうか（01_Menu.gs のメニュー表示判定用）
 * @returns {boolean}
 */
function isLegacyImportAvailable_() {
  return legacyImportDeadlineError_() === null;
}

// ============================================================
// ① 過去図面OCR（カスタムメニューから呼ばれる）
// ============================================================
function runLegacyOcrAndFillInputSheet() {
  const ui = SpreadsheetApp.getUi();

  const guardMsg = legacyImportDeadlineError_();
  if (guardMsg) {
    ui.alert('利用期限終了', guardMsg, ui.ButtonSet.OK);
    return;
  }

  const folderId = PropertiesService.getScriptProperties().getProperty('LEGACY_OCR_FOLDER_ID');
  if (!folderId) {
    ui.alert(
      '設定エラー',
      'スクリプトプロパティ「LEGACY_OCR_FOLDER_ID」が設定されていません。\n' +
      'Apps Script エディタ →「プロジェクトの設定」→「スクリプトプロパティ」から設定してください。',
      ui.ButtonSet.OK
    );
    return;
  }

  let folder;
  try {
    folder = DriveApp.getFolderById(folderId);
  } catch (e) {
    ui.alert('エラー',
      `指定フォルダが見つかりません。LEGACY_OCR_FOLDER_ID を確認してください。\n${e.message}`,
      ui.ButtonSet.OK);
    return;
  }

  const files = folder.getFilesByType(MimeType.PDF);
  const pdfFiles = [];
  while (files.hasNext()) pdfFiles.push(files.next());

  if (pdfFiles.length === 0) {
    ui.alert('PDFなし', '指定フォルダに PDF ファイルが見つかりませんでした。', ui.ButtonSet.OK);
    return;
  }

  const ss         = SpreadsheetApp.getActiveSpreadsheet();
  const inputSheet = ensureLegacyInputSheet_(ss);

  // 既に読み取り済みのPDFは対象外（このルートはファイルを移動しないため、
  // 再実行時の重複読み取りを防ぐには入力シート側でのチェックが必須）
  const existingUrls = getExistingLegacyFileUrls_(inputSheet);
  const unprocessed   = pdfFiles.filter(f => !existingUrls.has(f.getUrl()));
  if (unprocessed.length === 0) {
    ui.alert('完了',
      'フォルダ内のすべての PDF は既に入力シートに登録済みです。',
      ui.ButtonSet.OK);
    return;
  }

  const results = [];
  const errors  = [];
  unprocessed.forEach((file, idx) => {
    try {
      ss.toast(
        `過去図面OCR処理中... (${idx + 1}/${unprocessed.length}) ${file.getName()}`,
        '過去図面読み取り', 10
      );
      const extracted = extractLegacyDrawingInfoByOcr_(file);
      results.push({ file, extracted });
    } catch (e) {
      console.error(`過去図面OCR失敗: ${file.getName()}`, e);
      errors.push(file.getName());
      results.push({ file, extracted: null });
    }
    Utilities.sleep(1000); // API レート制限対策
  });

  writeLegacyOcrResultsToInputSheet_(inputSheet, results);
  ss.setActiveSheet(inputSheet);

  const errMsg = errors.length > 0
    ? `\n\n以下のファイルは読み取りに失敗しました（手動入力してください）：\n${errors.join('\n')}`
    : '';
  ui.alert(
    'OCR完了',
    `${results.length} 件の PDF を読み取り、入力シートへ転記しました。\n` +
    `黄色のセルを確認・修正後、「過去図面を一括登録する」を実行してください。\n` +
    `※ 対象の主図番が機械台帳に未発行の場合は、先に主図番を発行してください。${errMsg}`,
    ui.ButtonSet.OK
  );
}

// ============================================================
// ② 過去図面の一括登録（カスタムメニューから呼ばれる）
// ============================================================
function RegisterLegacyDrawings() {
  const ui = SpreadsheetApp.getUi();

  const guardMsg = legacyImportDeadlineError_();
  if (guardMsg) {
    ui.alert('利用期限終了', guardMsg, ui.ButtonSet.OK);
    return;
  }

  const ss         = SpreadsheetApp.getActiveSpreadsheet();
  const inputSheet = ensureLegacyInputSheet_(ss);
  const mainSheet  = ss.getSheetByName(SHEET_NAMES.MAIN_NUMBER);
  const dbSheet    = ss.getSheetByName(SHEET_NAMES.DRAWING_DB);
  if (!mainSheet) throw new Error(`"${SHEET_NAMES.MAIN_NUMBER}" シートが見つかりません。`);
  if (!dbSheet)   throw new Error(`"${SHEET_NAMES.DRAWING_DB}" シートが見つかりません。`);

  const pendingRows = getPendingLegacyRows_(inputSheet);
  if (pendingRows.length === 0) {
    ui.alert('対象なし',
      '登録対象の行がありません。Q列（登録結果）が空の行が対象です。',
      ui.ButtonSet.OK);
    return;
  }

  const errors = validateLegacyRows_(pendingRows, mainSheet, dbSheet);
  if (errors.length > 0) {
    ui.alert('入力エラー',
      `以下の行にエラーがあります。修正後に再実行してください。\n\n${errors.join('\n')}`,
      ui.ButtonSet.OK);
    return;
  }

  const confirmResp = ui.alert(
    '過去図面の一括登録',
    `${pendingRows.length} 件の図面を「過去図面」として図面台帳へ登録します。\n` +
    `正規の承認フロー（検図→課長→部長）は経ないため、承認Webアプリの対象にはなりません。\n` +
    `また、通知メールも送信されません。\n\n実行しますか？`,
    ui.ButtonSet.YES_NO
  );
  if (confirmResp !== ui.Button.YES) return;

  const lock = LockService.getScriptLock();
  lock.waitLock(15000);

  try {
    const newDbRows    = [];
    const rowUpdates   = [];
    const importErrors = [];

    pendingRows.forEach(row => {
      const drawingNo = String(row.data[COL_LEGACY_INPUT.DRAWING_NO - 1]).trim();
      const parsed     = parseDrawingNo_(drawingNo);

      // 入力シートのFILE_URL列はHYPERLINK数式なので、数式から実際のURLを取り出す
      const fileUrlFormula = inputSheet.getRange(row.rowIndex, COL_LEGACY_INPUT.FILE_URL).getFormula();
      const sourceFileUrl = extractUrlFromCellFormula_(fileUrlFormula) ||
                             String(row.data[COL_LEGACY_INPUT.FILE_URL - 1]).trim();

      // 元PDFを「承認済み図面」フォルダのサブフォルダへそのままコピー
      // （25_PdfApprovalStamp.gs と同じ保存先・命名規則を使うが、スタンプ処理は行わない）
      let approvedFileUrl = '';
      try {
        approvedFileUrl = copyLegacyPdfToApprovedFolder_(sourceFileUrl, parsed.mainNo);
      } catch (e) {
        importErrors.push(`${row.rowIndex}行目（${drawingNo}）：PDFの保存に失敗したため、この行はスキップしました：${e.message}`);
        return;
      }

      // 特徴属性（類似図面検索用）。入力シートR列にJSON文字列として保持
      // されているものを、そのまま図面台帳へ引き継ぐ（21_OcrService.gs の
      // normalizeDrawingAttributes_、extractLegacyDrawingInfoByOcr_ を参照）。
      // summaryのみ複製し、シートを直接見たときに一覧性を確保する。
      const attributesJson = String(row.data[COL_LEGACY_INPUT.ATTRIBUTES - 1] || '').trim();
      let attributesSummary = '';
      if (attributesJson) {
        try {
          attributesSummary = String(JSON.parse(attributesJson).summary || '');
        } catch (e) {
          console.warn(`特徴属性JSONのパースに失敗しました（${drawingNo}）: ${e.message}`);
        }
      }

      newDbRows.push([
        drawingNo,                                                       // A: フル図番
        parsed.mainNo,                                                   // B: 主図番
        parsed.subNo,                                                    // C: 子図番
        parsed.revMark,                                                  // D: 改訂記号
        row.data[COL_LEGACY_INPUT.NAME_JP    - 1],                       // E: 図名（JPN）
        row.data[COL_LEGACY_INPUT.NAME_EN    - 1],                       // F: 英名（NAME）
        row.data[COL_LEGACY_INPUT.UNIT_NAME  - 1],                       // G: ユニット名（UNIT）
        row.data[COL_LEGACY_INPUT.MODEL_NAME - 1],                       // H: 機械名（MODEL）
        row.data[COL_LEGACY_INPUT.MATERIAL   - 1],                       // I: 材質（MATL.）
        row.data[COL_LEGACY_INPUT.SCALE      - 1],                       // J: 縮尺（SCALE）
        String(row.data[COL_LEGACY_INPUT.DRAWING_SIZE - 1]).trim().toUpperCase(), // K: 図面サイズ
        STATUS_LEGACY,                                                   // L: ステータス＝「過去図面」
        row.data[COL_LEGACY_INPUT.DESIGNER    - 1],                      // M: 申請者 ← 設計者（OCRのまま）
        row.data[COL_LEGACY_INPUT.DESIGN_DATE - 1],                      // N: 申請日 ← 設計日
        row.data[COL_LEGACY_INPUT.REVIEWER    - 1],                      // O: 検図者
        row.data[COL_LEGACY_INPUT.REVIEW_DATE - 1],                      // P: 検図者承認日時 ← 検図日
        '',                                                               // Q: 課長（過去図面は押印欄自体が存在しないため空欄）
        '',                                                               // R: 課長承認日時
        row.data[COL_LEGACY_INPUT.APPROVER    - 1],                      // S: 部長 ← 承認者
        row.data[COL_LEGACY_INPUT.APPROVE_DATE - 1],                     // T: 部長承認日時 ← 承認日
        buildFileLinkFormula_(approvedFileUrl, '📄図面を開く'),           // U: 図面リンク（承認済み図面フォルダ内のコピー）
        '',                                                               // V: 申請バッチID（過去図面は通知フローを使わないため空欄）
        '',                                                               // W: AIチェック（過去図面インポートでは実施しないため空欄）
        attributesJson,                                                  // X: 特徴属性（JSON。類似図面検索用）
        attributesSummary,                                               // Y: 特徴属性の要約
      ]);

      rowUpdates.push({ rowIndex: row.rowIndex, drawingNo });
    });

    if (newDbRows.length > 0) {
      const startRow = dbSheet.getLastRow() + 1;
      dbSheet.getRange(startRow, 1, newDbRows.length, DRAWING_DB_COL_COUNT).setValues(newDbRows);
      formatDbRows_(dbSheet, startRow, newDbRows.length);

      // 図面索引を更新する（05_SearchIndex.gs を参照）。過去図面は
      // ステータスが STATUS_LEGACY（承認進行中ではない）ため、
      // 承認待ち索引の更新は不要。
      try {
        newDbRows.forEach(rowValues => upsertDrawingIndex_(rowValues));
      } catch (indexErr) {
        console.warn(`検索用索引の更新に失敗しました: ${indexErr.message}`);
      }
    }

    rowUpdates.forEach(u => {
      inputSheet.getRange(u.rowIndex, COL_LEGACY_INPUT.REG_RESULT).setValue('登録済み');
      inputSheet.getRange(u.rowIndex, 1, 1, LEGACY_INPUT_COL_COUNT).setBackground('#e8f5e9');
    });

    if (newDbRows.length > 0) {
      ss.setActiveSheet(dbSheet);
    }

    const errMsg = importErrors.length > 0
      ? `\n\n以下は登録できませんでした：\n${importErrors.join('\n')}`
      : '';
    ui.alert(
      '登録完了',
      `${newDbRows.length} 件を「過去図面」として図面台帳へ登録しました。${errMsg}`,
      ui.ButtonSet.OK
    );

  } finally {
    lock.releaseLock();
  }
}

// ============================================================
// 過去図面用OCR
// ============================================================

const LEGACY_OCR_PROMPT = `
この画像は、過去に紙で運用されていた機械図面です。図面の表題欄（タイトルブロック）を
読み取り、以下の項目をJSON形式で返してください。表題欄にはDESIGNED（設計）・
CHECKED（検図）・APPROVED（承認）の各欄に、手書きまたは印影で氏名と日付の
記入があります。

抽出する項目（括弧内は表題欄上の表示ラベル）：
- drawingNo    : 図面番号（DWG.NO.）英大文字2桁＋数字4桁＋数字3桁＋1桁の計10桁。例：JB2601001-
- drawingName  : 図名（JPN）日本語の図面名称
- nameEn       : 英名（NAME）英語の図面名称
- unitName     : ユニット名（UNIT）
- modelName    : 機械名（MODEL）
- material     : 材質（MATL.）例：SUS304、S45C
- scale        : 縮尺（SCALE）例：1:2、1:5
- drawingSize  : 図面サイズ（用紙サイズ）。「A0」「A1」「A2」「A3」「A4」のいずれか
- designerName : DESIGNED欄の氏名（手書き文字・印影のいずれも可能な限り読み取る）
- designDate   : DESIGNED欄の日付。西暦 yyyy-MM-dd 形式に変換して返す（元号表記の場合は西暦へ変換する）
- reviewerName : CHECKED欄の氏名
- reviewDate   : CHECKED欄の日付。yyyy-MM-dd形式
- approverName : APPROVED欄の氏名
- approveDate  : APPROVED欄の日付。yyyy-MM-dd形式

さらに、図中の部品の形状・特徴について、以下の属性（attributes）も
あわせて抽出してください（将来的に、似た過去図面を検索する機能で使用します。
断定できない項目は無理に埋めず null にしてください）：
- shapeCategory : 部品の形状分類。次のいずれか1つ："板金","軸物","ブロック・切削","ブラケット","その他"
- overallDimensions : 図面から読み取れる部品のおおよその外形寸法（mm、数値）。
    { "length": 全長など最大寸法, "width": 幅方向の寸法, "height": 高さ・厚み方向の寸法 }
    軸物など該当しない軸がある場合は null にしてください。
- features : 図面から読み取れる加工上の特徴
    { "holeCount": 穴の合計個数（読み取れなければnull）,
      "hasThreads": ねじ加工の有無（true/false）,
      "threadSizes": ねじ径の配列（例：["M8","M12"]。なければ空配列）,
      "hasBending": 曲げ加工の有無（true/false）,
      "hasChamfer": 面取り指示の有無（true/false）,
      "hasGeometricTolerance": 幾何公差（データムを伴う公差記入枠）指示の有無（true/false）,
      "geometricToleranceTypes": 図面上に記載されている幾何公差の種類の配列。
          次の用語から該当するものを選んでください（複数可）：
          ["真直度","平面度","真円度","円筒度","線の輪郭度","面の輪郭度",
           "平行度","直角度","傾斜度","位置度","同心度","対称度","円周振れ","全振れ"]。
          該当なし、または読み取れなければ空配列,
      "hasSurfaceTreatment": 表面処理指示の有無（true/false）。表題欄・注記欄いずれも確認,
      "surfaceTreatmentType": 表面処理の具体的な種類（例："三価黒色クロメートメッキ",
          "アルマイト処理(黒)","無電解ニッケルメッキ","黒染め"）。指示はあるが具体的な
          種類名までは読み取れない場合は"表面処理あり（詳細不明）"としてください。
          指示自体がなければ null,
      "hasHeatTreatment": 熱処理指示の有無（true/false）。表題欄・注記欄いずれも確認,
      "heatTreatmentType": 熱処理の具体的な種類（例："焼き入れ HRC50〜55","浸炭焼入れ",
          "高周波焼入れ","焼きなまし"）。指示はあるが具体的な種類名までは読み取れない
          場合は"熱処理あり（詳細不明）"としてください。指示自体がなければ null }
- summary : 部品の形状・特徴を日本語1〜3文で簡潔に要約した説明文。
    幾何公差・表面処理・熱処理の指示がある場合は、その内容にも触れてください
    （例："SS400製の平板ベース。長穴14箇所、C1面取り指定。位置度公差指示あり、
    三価黒色クロメートメッキ仕上げ。"）

返答はJSONオブジェクトのみとし、マークダウンのコードブロック(\`\`\`)は含めないでください。
読み取れなかった項目は空文字（""）としてください。日付が読み取れても西暦への変換に
自信が持てない場合も、無理に変換せず読み取れた通りの文字列を返してください。
例：{
  "drawingNo":"JB2601001-",
  "drawingName":"ブラケット",
  "nameEn":"BRACKET",
  "unitName":"フレームユニット",
  "modelName":"○○装置",
  "material":"SUS304",
  "scale":"1:2",
  "drawingSize":"A3",
  "designerName":"K.Yokoyama",
  "designDate":"2018-04-10",
  "reviewerName":"O.Hamasaka",
  "reviewDate":"2018-04-12",
  "approverName":"M.Fujihara",
  "approveDate":"2018-04-15",
  "attributes": {
    "shapeCategory": "ブラケット",
    "overallDimensions": { "length": 20, "width": 20, "height": 16 },
    "features": {
      "holeCount": 2,
      "hasThreads": false,
      "threadSizes": [],
      "hasBending": true,
      "hasChamfer": true,
      "hasGeometricTolerance": true,
      "geometricToleranceTypes": ["位置度", "平行度"],
      "hasSurfaceTreatment": true,
      "surfaceTreatmentType": "三価黒色クロメートメッキ",
      "hasHeatTreatment": false,
      "heatTreatmentType": null
    },
    "summary": "SPCC製のL字ブラケット。長穴2箇所、C2面取り指定。位置度・平行度公差指示あり、三価黒色クロメートメッキ仕上げ。"
  }
}
`;

/**
 * 過去図面PDFをGemini APIでOCRし、通常項目・部品属性に加えて設計者・検図者・
 * 承認者の氏名・日付を抽出する。21_OcrService.gs の extractDrawingInfoByOcr_ と
 * 同じ GEMINI_API_KEY / GEMINI_MODEL / GEMINI_API_URL / normalizeDrawingAttributes_
 * を流用する。
 *
 * @param {GoogleAppsScript.Drive.File} file
 * @returns {{drawingNo, drawingName, nameEn, unitName, modelName, material, scale,
 *            drawingSize, designerName, designDate, reviewerName, reviewDate,
 *            approverName, approveDate, attributes, hasWarning}|null}
 */
function extractLegacyDrawingInfoByOcr_(file) {
  const apiKey = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');
  if (!apiKey) {
    throw new Error(
      'スクリプトプロパティ「GEMINI_API_KEY」が設定されていません。\n' +
      'Apps Script エディタ →「プロジェクトの設定」→「スクリプトプロパティ」から設定してください。'
    );
  }

  const pdfBlob   = file.getBlob();
  const base64Pdf = Utilities.base64Encode(pdfBlob.getBytes());

  const requestBody = {
    contents: [{
      parts: [
        { text: LEGACY_OCR_PROMPT },
        { inline_data: { mime_type: 'application/pdf', data: base64Pdf } },
      ]
    }],
    generationConfig: {
      temperature:     0,
      maxOutputTokens: 1024, // 項目数が多い（設計者等の氏名・日付＋幾何公差等の属性）ため余裕を持たせる
    }
  };

  const response = UrlFetchApp.fetch(`${GEMINI_API_URL}?key=${apiKey}`, {
    method:             'post',
    contentType:        'application/json',
    payload:            JSON.stringify(requestBody),
    muteHttpExceptions: true,
  });

  const statusCode = response.getResponseCode();
  if (statusCode !== 200) {
    console.error(`Gemini API エラー（過去図面OCR, ${statusCode}）: ${response.getContentText()}`);
    throw new Error(
      `Gemini API がステータス ${statusCode} を返しました。APIキーとモデル名を確認してください。`
    );
  }

  const responseJson = JSON.parse(response.getContentText());
  const text = responseJson?.candidates?.[0]?.content?.parts?.[0]?.text;
  if (!text) {
    console.warn('Gemini API から有効なテキストが返りませんでした。', JSON.stringify(responseJson));
    return null;
  }

  let extracted;
  try {
    const cleaned = text.replace(/```json|```/g, '').trim();
    extracted = JSON.parse(cleaned);
  } catch (e) {
    console.error('JSON パース失敗:', text, e);
    return null;
  }

  const result = {
    drawingNo:    String(extracted.drawingNo    || '').trim(),
    drawingName:  String(extracted.drawingName  || '').trim(),
    nameEn:       String(extracted.nameEn       || '').trim(),
    unitName:     String(extracted.unitName     || '').trim(),
    modelName:    String(extracted.modelName    || '').trim(),
    material:     String(extracted.material     || '').trim(),
    scale:        String(extracted.scale        || '').trim(),
    drawingSize:  String(extracted.drawingSize  || '').trim().toUpperCase(),
    designerName: String(extracted.designerName || '').trim(),
    designDate:   String(extracted.designDate   || '').trim(),
    reviewerName: String(extracted.reviewerName || '').trim(),
    reviewDate:   String(extracted.reviewDate   || '').trim(),
    approverName: String(extracted.approverName || '').trim(),
    approveDate:  String(extracted.approveDate  || '').trim(),
  };

  if (result.drawingSize && VALID_DRAWING_SIZES.indexOf(result.drawingSize) === -1) {
    console.warn(`OCRが想定外の図面サイズを返しました: "${result.drawingSize}"。空欄として扱います。`);
    result.drawingSize = '';
  }

  // 必須項目、および過去図面インポート特有の項目（設計者・検図者・承認者）が
  // 読み取れなかった場合は警告フラグを立て、入力シート上で目立たせる
  result.hasWarning =
    !result.drawingNo || !result.drawingName || !result.drawingSize ||
    !result.designerName || !result.reviewerName || !result.approverName;

  // 部品の形状・特徴属性（類似図面検索用）。normalizeDrawingAttributes_ は
  // 21_OcrService.gs で定義済み（同一プロジェクト内でグローバルに利用可能）。
  // あくまで付随データのため、抽出の成否は hasWarning に影響させない。
  result.attributes = normalizeDrawingAttributes_(extracted.attributes);

  return result;
}

// ============================================================
// 過去図面PDFの保存（スタンプ処理なし）
// ============================================================

/**
 * 過去図面の元PDFを、「承認済み図面」フォルダの主図番サブフォルダへそのまま
 * コピーする（25_PdfApprovalStamp.gs のスタンプ処理は行わない。既にスタンプ・
 * 押印済みの内容をOCRで読み取っているだけのため）。
 * サブフォルダの命名規則・フォルダリンクの記録は 25_PdfApprovalStamp.gs と共通化する
 * （getMainNumberInfo_ / getOrCreateSubfolder_ / recordFolderLinks_ を流用）。
 *
 * @param {string} sourceFileUrl - 過去図面登録申請（入力）シートに登録された元PDFのURL
 * @param {string} mainNo        - 主図番
 * @returns {string} コピー後のPDFのURL
 */
function copyLegacyPdfToApprovedFolder_(sourceFileUrl, mainNo) {
  const approvedFolderId = PropertiesService.getScriptProperties().getProperty('APPROVED_DRAWING_FOLDER_ID');
  if (!approvedFolderId) {
    throw new Error('スクリプトプロパティ「APPROVED_DRAWING_FOLDER_ID」が未設定です。');
  }

  const fileId = extractDriveFileId_(sourceFileUrl);
  if (!fileId) {
    throw new Error(`元PDFのURLからファイルIDを取得できませんでした: ${sourceFileUrl}`);
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const mainInfo = getMainNumberInfo_(ss, mainNo); // 25_PdfApprovalStamp.gs で定義済み
  if (!mainInfo) {
    throw new Error(`機械台帳に「${mainNo}」が見つかりませんでした。`);
  }
  if (!mainInfo.modelNameJp) {
    throw new Error(`機械台帳の「${mainNo}」に機械名（日本語）が未登録です。`);
  }

  const subfolderName = mainInfo.modelNameEn
    ? `${mainNo}_${mainInfo.modelNameJp}_${mainInfo.modelNameEn}`
    : `${mainNo}_${mainInfo.modelNameJp}`;

  const sourceFile = DriveApp.getFileById(fileId);
  const baseFolder = DriveApp.getFolderById(approvedFolderId);
  const subfolder  = getOrCreateSubfolder_(baseFolder, subfolderName); // 25_PdfApprovalStamp.gs で定義済み
  const copiedFile = sourceFile.makeCopy(sourceFile.getName(), subfolder);

  try {
    recordFolderLinks_(ss, mainNo, subfolder); // 25_PdfApprovalStamp.gs で定義済み
  } catch (linkErr) {
    console.warn(`機械台帳へのフォルダリンク記録に失敗しました（${mainNo}）: ${linkErr.message}`);
  }

  return copiedFile.getUrl();
}

// ============================================================
// 過去図面登録申請（入力）シートの管理
// ============================================================

/**
 * 「過去図面登録申請（入力）」シートを取得する。存在しなければヘッダー付きで新規作成する。
 * @param {Spreadsheet} ss
 * @returns {Sheet}
 */
function ensureLegacyInputSheet_(ss) {
  let sheet = ss.getSheetByName(SHEET_NAMES.LEGACY_INPUT);
  if (sheet) return sheet;

  sheet = ss.insertSheet(SHEET_NAMES.LEGACY_INPUT);
  const headers = [
    '図面番号(DWG.NO.)', '図名(JPN)', '英名(NAME)', 'ユニット名(UNIT)', '機械名(MODEL)',
    '材質(MATL.)', '縮尺(SCALE)', '図面サイズ', '図面ファイルURL',
    '設計者', '設計日', '検図者', '検図日', '承認者', '承認日',
    'AI-OCR読取結果', '登録結果', '特徴属性（JSON）',
  ];
  sheet.getRange(GLOBAL_ROW.HEADER, 1, 1, headers.length)
    .setValues([headers])
    .setFontWeight('bold')
    .setBackground('#f5f7fa');
  sheet.setFrozenRows(GLOBAL_ROW.HEADER);
  sheet.setColumnWidths(1, headers.length, 140);

  return sheet;
}

/**
 * 入力シートの既存ファイルURL一覧をSetで返す（重複読み取り防止）
 */
function getExistingLegacyFileUrls_(inputSheet) {
  const lastRow = inputSheet.getLastRow();
  if (lastRow < GLOBAL_ROW.DATA_START) return new Set();
  const formulas = inputSheet
    .getRange(GLOBAL_ROW.DATA_START, COL_LEGACY_INPUT.FILE_URL, lastRow - GLOBAL_ROW.DATA_START + 1, 1)
    .getFormulas().flat();
  const urls = formulas.map(f => extractUrlFromCellFormula_(f)).filter(v => v);
  return new Set(urls);
}

/**
 * 過去図面OCR結果を入力シートへ転記する
 */
function writeLegacyOcrResultsToInputSheet_(inputSheet, results) {
  results.forEach(({ file, extracted }) => {
    const fileUrl = file.getUrl();
    const ocrNote = extracted
      ? (extracted.hasWarning ? '⚠️ 要確認（一部読取不可）' : '✅ OCR済み')
      : '❌ 読取失敗（手動入力してください）';

    // 特徴属性（類似図面検索用の下ごしらえデータ）。抽出に失敗していても
    // OCR自体は継続する付随データのため、欠損時は空文字にする
    const attributesJson = (extracted && extracted.attributes)
      ? JSON.stringify(extracted.attributes)
      : '';

    const newRow = [
      extracted?.drawingNo    || '',
      extracted?.drawingName  || '',
      extracted?.nameEn       || '',
      extracted?.unitName     || '',
      extracted?.modelName    || '',
      extracted?.material     || '',
      extracted?.scale        || '',
      extracted?.drawingSize  || '',
      buildFileLinkFormula_(fileUrl, '📄図面を開く'),
      extracted?.designerName || '',
      extracted?.designDate   || '',
      extracted?.reviewerName || '',
      extracted?.reviewDate   || '',
      extracted?.approverName || '',
      extracted?.approveDate  || '',
      ocrNote,
      '',
      attributesJson,
    ];
    inputSheet.appendRow(newRow);

    const lastRow = inputSheet.getLastRow();

    if (extracted) {
      if (!extracted.drawingNo)    highlightCell_(inputSheet, lastRow, COL_LEGACY_INPUT.DRAWING_NO);
      if (!extracted.drawingName)  highlightCell_(inputSheet, lastRow, COL_LEGACY_INPUT.NAME_JP);
      if (!extracted.nameEn)       highlightCell_(inputSheet, lastRow, COL_LEGACY_INPUT.NAME_EN);
      if (!extracted.unitName)     highlightCell_(inputSheet, lastRow, COL_LEGACY_INPUT.UNIT_NAME);
      if (!extracted.modelName)    highlightCell_(inputSheet, lastRow, COL_LEGACY_INPUT.MODEL_NAME);
      if (!extracted.material)     highlightCell_(inputSheet, lastRow, COL_LEGACY_INPUT.MATERIAL);
      if (!extracted.scale)        highlightCell_(inputSheet, lastRow, COL_LEGACY_INPUT.SCALE);
      if (!extracted.drawingSize)  highlightCell_(inputSheet, lastRow, COL_LEGACY_INPUT.DRAWING_SIZE);
      if (!extracted.designerName) highlightCell_(inputSheet, lastRow, COL_LEGACY_INPUT.DESIGNER);
      if (!extracted.designDate)   highlightCell_(inputSheet, lastRow, COL_LEGACY_INPUT.DESIGN_DATE);
      if (!extracted.reviewerName) highlightCell_(inputSheet, lastRow, COL_LEGACY_INPUT.REVIEWER);
      if (!extracted.reviewDate)   highlightCell_(inputSheet, lastRow, COL_LEGACY_INPUT.REVIEW_DATE);
      if (!extracted.approverName) highlightCell_(inputSheet, lastRow, COL_LEGACY_INPUT.APPROVER);
      if (!extracted.approveDate)  highlightCell_(inputSheet, lastRow, COL_LEGACY_INPUT.APPROVE_DATE);
    } else {
      inputSheet.getRange(lastRow, 1, 1, COL_LEGACY_INPUT.FILE_URL).setBackground('#fff2cc');
    }
  });
}

/**
 * 入力シートから未登録行（Q列が空）を取得する
 * @returns {Array<{rowIndex, data}>}
 */
function getPendingLegacyRows_(inputSheet) {
  const lastRow = inputSheet.getLastRow();
  if (lastRow < GLOBAL_ROW.DATA_START) return [];

  const values = inputSheet
    .getRange(GLOBAL_ROW.DATA_START, 1, lastRow - GLOBAL_ROW.DATA_START + 1, LEGACY_INPUT_COL_COUNT)
    .getValues();

  return values
    .map((row, i) => ({ rowIndex: GLOBAL_ROW.DATA_START + i, data: row }))
    .filter(({ data }) => {
      const regResult = String(data[COL_LEGACY_INPUT.REG_RESULT - 1]).trim();
      const drawingNo = String(data[COL_LEGACY_INPUT.DRAWING_NO - 1]).trim();
      return !regResult && drawingNo;
    });
}

/**
 * 過去図面インポートのバリデーション
 * - 通常ルートと同じ基本項目（図番・図名・図面サイズ・ファイルURL・重複チェック）に加え、
 *   主図番が機械台帳に既に発行済みであることを確認する
 *   （過去図面インポートより先に主図番発行が必要、という運用ルールのガード）
 * - 設計者・検図者・承認者の氏名・日付は、判読不能な古い図面もあり得るため必須にはしない
 *   （空欄のまま登録され、後から図面台帳を直接編集して補完できる）
 * @returns {string[]} エラーメッセージの配列（空なら問題なし）
 */
function validateLegacyRows_(pendingRows, mainSheet, dbSheet) {
  const errors        = [];
  const registeredNos = new Set();
  const existingNos   = getExistingDrawingNos_(dbSheet); // 20_DrawingApproveRequest.gs で定義済み

  pendingRows.forEach(row => {
    const r           = row.rowIndex;
    const drawingNo   = String(row.data[COL_LEGACY_INPUT.DRAWING_NO   - 1]).trim();
    const nameJp      = String(row.data[COL_LEGACY_INPUT.NAME_JP      - 1]).trim();
    const drawingSize = String(row.data[COL_LEGACY_INPUT.DRAWING_SIZE - 1]).trim().toUpperCase();
    const fileUrl      = String(row.data[COL_LEGACY_INPUT.FILE_URL     - 1]).trim();

    const missing = [];
    if (!drawingNo)   missing.push('図面番号(DWG.NO.)');
    if (!nameJp)      missing.push('図名(JPN)');
    if (!drawingSize) missing.push('図面サイズ');
    if (!fileUrl)     missing.push('図面ファイルURL');
    if (missing.length > 0) {
      errors.push(`${r}行目：${missing.join('・')} が未入力です`);
      return;
    }

    if (VALID_DRAWING_SIZES.indexOf(drawingSize) === -1) {
      errors.push(
        `${r}行目：図面サイズ「${drawingSize}」が不正です。` +
        `${VALID_DRAWING_SIZES.join('・')} のいずれかを入力してください。`
      );
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

    // 主図番が既に発行済みであることを確認（getMainNumberInfoFull_ は 30_DrawingSearch.gs で定義済み）
    const mainInfo = getMainNumberInfoFull_(mainSheet, parsed.mainNo);
    if (!mainInfo) {
      errors.push(
        `${r}行目：主図番「${parsed.mainNo}」が機械台帳に見つかりません。` +
        `先に主図番発行機能で主図番を発行してください。`
      );
      return;
    }
  });

  return errors;
}

// ============================================================
// 【デバッグ用】
// ============================================================
function TEST_runLegacyOcrAndFillInputSheet() {
  runLegacyOcrAndFillInputSheet();
}

function TEST_RegisterLegacyDrawings() {
  RegisterLegacyDrawings();
}

function checkLegacyImportStatus() {
  const err = legacyImportDeadlineError_();
  if (err) {
    console.log(`❌ 過去図面インポート機能は利用できません: ${err}`);
  } else {
    const deadline = PropertiesService.getScriptProperties().getProperty('LEGACY_IMPORT_DEADLINE');
    console.log(`✅ 過去図面インポート機能は利用可能です（期限: ${deadline}）`);
  }
}