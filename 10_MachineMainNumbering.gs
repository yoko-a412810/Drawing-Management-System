/**
 * ============================================================
 * 図面承認・採番・出図管理システム
 * ファイル: 10_MachineMainNumbering.gs  ―  主図番採番モジュール
 * ============================================================
 *
 * 【図番フォーマット】
 *   主図番（6桁） = 部署コード2桁 + 着手日の西暦下2桁 + 連番2桁
 *   例：JB2601（2026年・JB部署・1番目）
 *   ※連番は部署コード×年度ごとにリセット（1年間最大99件）
 *
 * 【シート構成（前提）】
 *   - "部署コードマスタ" : A列=部署名, B列=コード(2桁), ヘッダ行=4行目, データ行=5行目〜
 *   - "機械台帳"     : A=主図番, B=部署コード, C=連番, D=機械名（日本語）,
 *                          E=機械名（英語・任意）, F=申請者, G=着手日, H=備考,
 *                          I=承認済み図面フォルダリンク, ヘッダ行=4行目, データ行=5行目〜
 *   - "ユーザーマスタ"   : A=メールアドレス, B=氏名, C=NAME（英語表記）, D=登録日, ヘッダ行=4行目, データ行=5行目〜
 */

// ============================================================
// 定数
// ============================================================
// 部署コードマスタの列インデックス
const COL_DEPT = {
  NAME: 1,  // A: 部署名
  CODE: 2,  // B: コード
};

// ============================================================
// ダイアログを開く
// ============================================================
function openMainNumberDialog() {
  const deptList = getDepartmentList_();
  if (deptList.length === 0) {
    SpreadsheetApp.getUi().alert(
      'エラー',
      '部署コードマスタにデータが登録されていません。\n先にマスタへ部署情報を追加してください。',
      SpreadsheetApp.getUi().ButtonSet.OK
    );
    return;
  }

  const html = HtmlService.createTemplateFromFile('11_MainNumberDialog');
  html.deptList = JSON.stringify(deptList);
  html.today    = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');

  SpreadsheetApp.getUi().showModalDialog(
    html.evaluate().setWidth(460).setHeight(640),
    '主図番の発行'
  );
}

// ============================================================
// 主図番を採番・登録する（ダイアログから呼ばれる）
// ============================================================
/**
 * @param {Object} form  - ダイアログから渡されるフォームデータ
 *   {string} deptCode      - 部署コード（例: "JB"）
 *   {string} deptName      - 部署名
 *   {string} modelNameJp   - 機械名（日本語）
 *   {string} modelNameEn   - 機械名（英語・任意）
 *   {string} startDate     - 着手日（例: "2026-07-01"）
 *   {string} note          - 備考（任意）
 * @returns {Object} result - { success, mainNumber, lastRow, message }
 */
function issueMainNumber(form) {
  try {
    // ── 入力バリデーション ──────────────────────────
    if (!form.deptCode || !form.modelNameJp || !form.startDate || !form.cadFolderUrl) {
      return { success: false, message: '必須項目（部署・機械名（日本語）・着手日・CADフォルダ）を入力してください。' };
    }
    if (!/^[A-Z]{2}$/.test(form.deptCode)) {
      return { success: false, message: '部署コードは英大文字2桁でなければなりません。' };
    }

    // ── CADフォルダの検証 ──────────────────────────
    // 発行後にフォルダを作るのではなく、事前に（中身が空でも）作成済みのフォルダの
    // リンクを発行時点で必須入力してもらう。Drive File Streamの「リンクをクリップボード
    // にコピー」で取得したURLを想定（extractDriveFileId_ は 25_PdfApprovalStamp.gs で定義）。
    const cadFolderId = extractDriveFileId_(form.cadFolderUrl);
    if (!cadFolderId) {
      return {
        success: false,
        message:
          'CADフォルダのURLを認識できませんでした。\n' +
          'エクスプローラーで対象フォルダを右クリック→「Google Drive」→' +
          '「リンクをクリップボードにコピー」で取得したリンクを貼り付けてください。',
      };
    }
    let cadFolder;
    try {
      cadFolder = DriveApp.getFolderById(cadFolderId);
    } catch (e) {
      return { success: false, message: '指定されたCADフォルダが見つかりませんでした。URLを確認してください。' };
    }

    // エクスプローラーパスは発行時点のものを初期値として保存する
    // （表示側では、31_DrawingSearchDialog.htmlの詳細パネルを開くたびにDrive APIで
    //   再計算するため、フォルダが後から移動されても表示上は最新の状態に追従する。
    //   この初期値はあくまで再計算に失敗した場合のフォールバック用）
    let cadFolderPath = '';
    try {
      cadFolderPath = buildExplorerPath_(cadFolder);
    } catch (e) {
      console.warn(`CADフォルダのエクスプローラーパス組み立てに失敗しました: ${e.message}`);
    }

    // ── 申請者氏名の取得（未登録なら中断） ─────────
    const applicant = getFullNameFromMaster_();
    if (!applicant) {
      const ss = SpreadsheetApp.getActiveSpreadsheet();
      ss.setActiveSheet(ss.getSheetByName(SHEET_NAMES.USER_MASTER));
      return {
        success: false,
        message:
          'ユーザーマスタに氏名が登録されていません。\n' +
          '「ユーザーマスタ」シートのB列（氏名）に氏名を入力してから、再度お試しください。',
      };
    }

    const ss    = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(SHEET_NAMES.MAIN_NUMBER);
    if (!sheet) throw new Error(`"${SHEET_NAMES.MAIN_NUMBER}" シートが見つかりません。`);

    // ── 排他制御（同時実行による重複採番を防止） ───
    const lock = LockService.getScriptLock();
    lock.waitLock(10000);

    try {
      // ── 連番の採番（部署コード×年度でリセット） ──
      const year      = form.startDate.slice(2, 4);         // 西暦下2桁（例："26"）
      const nextSeq   = getNextSequence_(sheet, form.deptCode, year);
      const seqPadded = String(nextSeq).padStart(2, '0');   // 2桁ゼロ埋め
      const mainNumber = form.deptCode + year + seqPadded;  // 例：JB2601

      // 連番が99を超える場合はエラー
      if (nextSeq > 99) {
        return {
          success: false,
          message: `${form.deptCode}部署の${year}年度の主図番が上限（99件）に達しています。`,
        };
      }

      // ── 部品リスト（BOM）のGoogleスプレッドシートを自動生成 ──
      // 作成タイミングが担当者によってバラつき、リンクの貼り忘れが発生する問題を
      // 構造的に防ぐため、主図番発行時にテンプレートから空のシートを複製し、
      // あらかじめリンクしておく。付随機能のため、失敗しても主図番の発行自体は
      // 止めない（PDF承認スタンプ機能と同様の設計方針）。
      let bomUrl = '';
      try {
        bomUrl = createBomSpreadsheet_(mainNumber, form.modelNameJp);
      } catch (e) {
        console.warn(`部品リストの自動作成に失敗しました（${mainNumber}）: ${e.message}`);
      }

      // ── 台帳へ書き込み ────────────────────────────
      const newRow = [
        mainNumber,            // A: 主図番
        form.deptCode,         // B: 部署コード
        nextSeq,                // C: 連番（数値）
        form.modelNameJp,       // D: 機械名（日本語）
        form.modelNameEn || '', // E: 機械名（英語・任意）
        applicant,               // F: 申請者（氏名）
        form.startDate,          // G: 着手日（文字列のまま渡してタイムゾーンズレを回避）
        form.note || '',         // H: 備考
        '',                      // I: 承認済み図面フォルダリンク（承認完了時に自動記入）
        '',                      // J: エクスプローラーパス（承認完了時に自動記入）
        buildFileLinkFormula_(cadFolder.getUrl(), '📁CADデータ'), // K: CADフォルダリンク
        cadFolderPath,            // L: CADフォルダパス（初期値。表示時に再計算される）
        buildFileLinkFormula_(bomUrl, '📝部品リストを開く'), // M: 部品リストURL（自動作成に失敗した場合は空欄）
      ];
      sheet.appendRow(newRow);

      // ── 着手日・連番セルの書式を整える ───────────
      const lastRow = sheet.getLastRow();
      sheet.getRange(lastRow, COL_MAIN.DATE).setNumberFormat('yyyy-mm-dd');
      sheet.getRange(lastRow, COL_MAIN.SEQ).setNumberFormat('0');

      // ── 行の書式を整える（罫線・文字寄せ） ───────
      formatNewRow_(sheet, lastRow);

      return {
        success:    true,
        mainNumber: mainNumber,
        lastRow:    lastRow,  // ダイアログ側でシート切り替え・行選択に使用
        message:    `主図番 [ ${mainNumber} ] を発行しました。\n機械名：${form.modelNameJp}`,
      };

    } finally {
      lock.releaseLock();
    }

  } catch (e) {
    console.error('issueMainNumber error:', e);
    return { success: false, message: `エラーが発生しました：${e.message}` };
  }
}

/**
 * 部品リスト（BOM）のテンプレートを複製し、部品リスト保存用フォルダへ新規保存する
 *
 * 【スクリプトプロパティ（事前設定が必要）】
 *   BOM_FOLDER_ID        : 部品リストの保存先フォルダのID（CADフォルダとは独立した専用フォルダ）
 *   BOM_TEMPLATE_FILE_ID : 複製元となる雛形Googleスプレッドシートのファイル ID
 *
 * どちらか一方でも未設定の場合は、自動作成をスキップして空文字を返す
 * （呼び出し元で警告ログのみ出し、主図番の発行自体は継続する）。
 *
 * @param {string} mainNumber - 主図番（例："JB2601"）
 * @param {string} modelNameJp - 機械名（日本語）
 * @returns {string} 複製された部品リストのURL（スキップ・失敗時は空文字）
 */
function createBomSpreadsheet_(mainNumber, modelNameJp) {
  const folderId   = PropertiesService.getScriptProperties().getProperty('BOM_FOLDER_ID');
  const templateId = PropertiesService.getScriptProperties().getProperty('BOM_TEMPLATE_FILE_ID');

  if (!folderId || !templateId) {
    console.warn(
      'スクリプトプロパティ「BOM_FOLDER_ID」または「BOM_TEMPLATE_FILE_ID」が未設定のため、' +
      '部品リストの自動作成をスキップしました。'
    );
    return '';
  }

  const folder   = DriveApp.getFolderById(folderId);
  const template = DriveApp.getFileById(templateId);
  const fileName = `${mainNumber}_${modelNameJp}_部品リスト`;
  const copy     = template.makeCopy(fileName, folder);

  return copy.getUrl();
}

// ============================================================
// ダイアログから呼ばれる：機械台帳に切り替えて登録行を選択する
// ============================================================
function activateMainNumberRow(rowIndex) {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAMES.MAIN_NUMBER);
  ss.setActiveSheet(sheet);
  sheet.setActiveRange(sheet.getRange(rowIndex, 1, 1, Object.keys(COL_MAIN).length));
}

// ============================================================
// ダイアログ用：プレビュー表示のための次の連番を返す
// ============================================================
/**
 * @param {Object} params - { deptCode, year }
 * @returns {number} 次の連番
 */
function getNextSequenceForPreview(params) {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAMES.MAIN_NUMBER);
  if (!sheet) return 1;
  return getNextSequence_(sheet, params.deptCode, params.year);
}

// ============================================================
// プライベートヘルパー
// ============================================================

/**
 * 部署コードマスタからプルダウン用リストを取得する
 * @returns {Array<{name, code}>}
 */
function getDepartmentList_() {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAMES.DEPT_MASTER);
  if (!sheet) return [];

  const lastRow = sheet.getLastRow();
  if (lastRow < GLOBAL_ROW.DATA_START) return [];

  const data = sheet.getRange(GLOBAL_ROW.DATA_START, 1, lastRow - GLOBAL_ROW.DATA_START + 1, 2).getValues();
  return data
    .filter(row => row[COL_DEPT.NAME - 1] && row[COL_DEPT.CODE - 1])
    .map(row => ({
      name: String(row[COL_DEPT.NAME - 1]).trim(),
      code: String(row[COL_DEPT.CODE - 1]).trim().toUpperCase(),
    }));
}

/**
 * 指定した部署コード×年度に対応する次の連番を返す
 * 該当行が存在しない場合は 1 を返す
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet
 * @param {string} deptCode - 部署コード（例："JB"）
 * @param {string} year     - 西暦下2桁（例："26"）
 * @returns {number}
 */
function getNextSequence_(sheet, deptCode, year) {
  const lastRow = sheet.getLastRow();
  if (lastRow < GLOBAL_ROW.DATA_START) return 1;

  // A列(主図番), B列(部署コード), C列(連番) の3列を取得
  const rows = sheet.getRange(
    GLOBAL_ROW.DATA_START,
    COL_MAIN.MAIN_NO,
    lastRow - GLOBAL_ROW.DATA_START + 1,
    3
  ).getValues();

  let maxSeq = 0;
  rows.forEach(row => {
    const mainNo  = String(row[0]).trim();               // 例："JB2601"
    const code    = String(row[1]).trim().toUpperCase();
    const seq     = Number(row[2]);
    const rowYear = mainNo.slice(2, 4);                  // 主図番の3〜4文字目が年度

    // 同じ部署コード かつ 同じ年度の行だけ対象
    if (code === deptCode && rowYear === year && seq > maxSeq) maxSeq = seq;
  });

  return maxSeq + 1;
}

/**
 * ユーザーマスタからログインユーザーの氏名を取得する
 * 未登録の場合はメールアドレスで仮登録し、null を返して発行を中断させる
 * @returns {string|null} 氏名（未登録・氏名空の場合は null）
 */
function getFullNameFromMaster_() {
  const email = Session.getActiveUser().getEmail();
  const ss    = SpreadsheetApp.getActiveSpreadsheet();

  // シートがなければ自動作成
  let sheet = ss.getSheetByName(SHEET_NAMES.USER_MASTER);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAMES.USER_MASTER);
    sheet.getRange(4, 1, 1, 4).setValues([['メールアドレス', '氏名', 'NAME', '登録日']]);
  }

  // 既存データを検索
  const lastRow = sheet.getLastRow();
  if (lastRow >= GLOBAL_ROW.DATA_START) {
    const data  = sheet.getRange(GLOBAL_ROW.DATA_START, 1, lastRow - GLOBAL_ROW.DATA_START + 1, COL_USER.NAME_JP).getValues();
    const match = data.find(row => row[COL_USER.EMAIL - 1] === email);
    if (match && match[COL_USER.NAME_JP - 1]) return String(match[COL_USER.NAME_JP - 1]);  // 氏名登録済み → 返す
    if (match && !match[COL_USER.NAME_JP - 1]) return null;             // 仮登録済みだが氏名未入力
  }

  // 未登録 → メールアドレスで仮登録して null を返す
  registerUserWithEmail_(sheet, email);
  return null;
}

/**
 * ユーザーマスタにメールアドレスで仮登録する
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet
 * @param {string} email
 */
function registerUserWithEmail_(sheet, email) {
  const today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
  sheet.appendRow([email, '', '', today]);

  // 氏名・NAME（英語表記）セルを黄色でハイライトして入力を促す
  const lastRow = sheet.getLastRow();
  sheet.getRange(lastRow, COL_USER.NAME_JP)
    .setBackground('#fff2cc')
    .setNote('氏名を入力してください');
  sheet.getRange(lastRow, COL_USER.NAME_EN)
    .setBackground('#fff2cc')
    .setNote('NAME（英語表記）を入力してください。例：K.Yokoyama\n承認日PDFスタンプ機能で使用します。');
}

/**
 * 新規追加行に書式（罫線・文字寄せ）を適用する
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet
 * @param {number} rowIndex
 */
function formatNewRow_(sheet, rowIndex) {
  const range = sheet.getRange(rowIndex, 1, 1, Object.keys(COL_MAIN).length);

  // 罫線
  range.setBorder(true, true, true, true, true, true,
    '#cccccc', SpreadsheetApp.BorderStyle.SOLID);

  // 中央揃え
  sheet.getRange(rowIndex, COL_MAIN.MAIN_NO).setHorizontalAlignment('center');
  sheet.getRange(rowIndex, COL_MAIN.DEPT_CODE).setHorizontalAlignment('center');
  sheet.getRange(rowIndex, COL_MAIN.SEQ).setHorizontalAlignment('center');
  sheet.getRange(rowIndex, COL_MAIN.DATE).setHorizontalAlignment('center');
}

// ============================================================
// onEdit トリガー：ユーザーマスタのB列（氏名）・C列（NAME英語表記）に
//                  入力されたら、黄色ハイライトとコメントを自動解除する
// ============================================================
function onEdit(e) {
  const sheet = e.range.getSheet();

  // ユーザーマスタのB列・C列以外は何もしない
  if (sheet.getName() !== SHEET_NAMES.USER_MASTER) return;
  const col = e.range.getColumn();
  if (col !== COL_USER.NAME_JP && col !== COL_USER.NAME_EN) return;
  if (e.range.getRow() < GLOBAL_ROW.DATA_START) return;

  const value = e.range.getValue();
  if (value) {
    // 入力されたら黄色とコメントを解除
    e.range.setBackground(null).clearNote();
  }
}

// ============================================================
// 【デバッグ用】スクリプトエディタから直接実行して動作確認できる
// ============================================================
function debugDeptList() {
  console.log(JSON.stringify(getDepartmentList_(), null, 2));
}

function TEST_issueMainNumber() {
  const result = issueMainNumber({
    deptCode:    'JB',
    deptName:    'ブレード技術部 設備開発課',
    modelNameJp: '【テスト】扇形チップ溶接機',
    modelNameEn: 'CTW2021',
    startDate:   '2026-07-01',
    cadFolderUrl: 'ここにテスト用の実在するDriveフォルダURLを入力',
    note:        'テスト実行',
  });
  console.log(JSON.stringify(result, null, 2));
}

// スクリプトエディタから1回だけ実行する
function setupSpreadsheetId() {
  const id = SpreadsheetApp.getActiveSpreadsheet().getId();
  PropertiesService.getScriptProperties().setProperty('SPREADSHEET_ID', id);
  console.log('保存完了:', id);
}