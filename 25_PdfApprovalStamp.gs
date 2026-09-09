/**
 * ============================================================
 * 図面承認・採番・出図管理システム
 * ファイル: 25_PdfApprovalStamp.gs  ―  承認日PDFスタンプ機能
 * ============================================================
 *
 * 【役割】
 *   部長承認完了時に、Cloud Function（Python + PyMuPDF）を呼び出して
 *   図面PDFの表題欄に「申請日・申請者」「検図日・検図者」「承認日・承認者」を
 *   書き込む。
 *
 *   元のPDF（検図・承認プロセス中の作業ファイル）は一切変更しない。
 *   スタンプ済みのPDFは、「承認済み図面」フォルダの下に主図番ごとのサブフォルダ
 *   （[主図番]_[機械名（日本語）]_[機械名（英語・任意）]）を作成し、その中へ
 *   新規ファイルとして保存する。サブフォルダは機械台帳から機械名を引いて
 *   決定し、同じ主図番の図面は毎回同じサブフォルダに集約される。
 *   作成・特定したサブフォルダのURLは、機械台帳の該当行にも記録する。
 *
 *   ※ 当初は元ファイルを Drive.Files.update() で上書きする設計だったが、
 *      実運用で「更新APIは成功しmd5Checksumも一致するのに、閲覧すると
 *      反映されていないように見える」という原因不明の事象が発生したため、
 *      新規ファイルとして保存する方式に変更した。新規ファイルであれば
 *      既存ファイルの更新・キャッシュにまつわる問題が原理的に起こらない。
 *      副次的に、承認済みの原本を書き換えずに済むため、文書管理上の
 *      トレーサビリティの観点でも望ましい。
 *
 * 【スクリプトプロパティ（事前設定が必要）】
 *   PDF_STAMP_FUNCTION_URL   : Cloud FunctionのURL
 *   PDF_STAMP_API_KEY        : Cloud Function呼び出し用の共有シークレット
 *                               （Cloud Function側の環境変数 STAMP_API_KEY と一致させる）
 *   APPROVED_DRAWING_FOLDER_ID : スタンプ済みPDFの保存先「承認済み図面」フォルダ ID
 *                                （この直下に主図番ごとのサブフォルダが作られる）
 *   SHARED_DRIVE_LETTER      : Google Drive for desktopでの共有ドライブのマウント先
 *                               ドライブ文字（例："G"）。未設定時は "G" を既定値として使う。
 *
 * 【拡張サービスの要件】
 *   エクスプローラーパスの組み立てに Drive API (v3) を使用するため、
 *   Apps Script エディタの「サービス」から Drive API (v3) を有効化しておくこと。
 *
 * 【設計方針】
 *   PDFスタンプはあくまで付随機能。処理に失敗しても、承認フロー本体
 *   （ステータス更新・通知メール）は止めない。呼び出し元
 *   （23_ApprovalWebApp.gs）は戻り値の success を見てログに警告を出すのみ。
 */

/**
 * 図面PDFの表題欄（DESIGNED/CHECKED/APPROVED各行）に日付・氏名をスタンプし、
 * 主図番ごとのサブフォルダ（[主図番]_[機械名日本語]_[機械名英語]）へ新規保存する
 *
 * @param {string} fileUrl     - 図面台帳に保存されている元PDFのDriveファイルURL
 * @param {string} drawingSize - 図面サイズ（'A0'〜'A4'）
 * @param {string} mainNo      - 主図番（例："JB2015"）。保存先サブフォルダの決定に使用する
 * @param {Object} stampData   - 書き込む日付・氏名（いずれも空文字なら該当欄は書き込まない）
 *   @param {string} stampData.applicantDate - 申請日時（'yyyy-MM-dd HH:mm:ss'）→ DESIGNED行
 *   @param {string} stampData.applicantName - 申請者氏名                      → DESIGNED行
 *   @param {string} stampData.reviewerDate  - 検図者承認日時                  → CHECKED行
 *   @param {string} stampData.reviewerName  - 検図者氏名                      → CHECKED行
 *   @param {string} stampData.directorDate  - 部長承認日時                    → APPROVED行
 *   @param {string} stampData.directorName  - 部長氏名                        → APPROVED行
 * @returns {{success: boolean, message: string, approvedFileUrl?: string}}
 */
function stampApprovalDatesOnPdf_(fileUrl, drawingSize, mainNo, stampData) {
  try {
    if (!fileUrl) {
      return { success: false, message: '図面ファイルURLが空のためスキップしました。' };
    }
    if (VALID_DRAWING_SIZES.indexOf(drawingSize) === -1) {
      return { success: false, message: `図面サイズ「${drawingSize}」が不正なためスキップしました。` };
    }
    if (!mainNo) {
      return { success: false, message: '主図番が空のためスキップしました（保存先フォルダを決定できません）。' };
    }

    const functionUrl     = PropertiesService.getScriptProperties().getProperty('PDF_STAMP_FUNCTION_URL');
    const apiKey           = PropertiesService.getScriptProperties().getProperty('PDF_STAMP_API_KEY');
    const approvedFolderId = PropertiesService.getScriptProperties().getProperty('APPROVED_DRAWING_FOLDER_ID');
    if (!functionUrl || !apiKey) {
      return {
        success: false,
        message: 'スクリプトプロパティ「PDF_STAMP_FUNCTION_URL」または「PDF_STAMP_API_KEY」が未設定です。',
      };
    }
    if (!approvedFolderId) {
      return {
        success: false,
        message: 'スクリプトプロパティ「APPROVED_DRAWING_FOLDER_ID」が未設定です。',
      };
    }

    // ── 機械台帳から機械名を引き、保存先サブフォルダ名を決定 ──
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const mainInfo = getMainNumberInfo_(ss, mainNo);
    if (!mainInfo) {
      return { success: false, message: `機械台帳に「${mainNo}」が見つかりませんでした。` };
    }
    if (!mainInfo.modelNameJp) {
      return { success: false, message: `機械台帳の「${mainNo}」に機械名（日本語）が未登録です。` };
    }
    const subfolderName = mainInfo.modelNameEn
      ? `${mainNo}_${mainInfo.modelNameJp}_${mainInfo.modelNameEn}`
      : `${mainNo}_${mainInfo.modelNameJp}`;

    const fileId = extractDriveFileId_(fileUrl);
    if (!fileId) {
      return { success: false, message: `URLからDriveファイルIDを抽出できませんでした: ${fileUrl}` };
    }

    // ── 元PDFを取得してbase64化（元ファイルはこの後も一切変更しない） ──
    const file    = DriveApp.getFileById(fileId);
    const pdfBlob = file.getBlob();
    const pdfB64  = Utilities.base64Encode(pdfBlob.getBytes());

    // ── 日付を YYYY/MM/DD に整形（時刻部分はPDF上では不要なため） ──
    const toDateOnly = (v) => (formatSheetDateTime_(v).split(' ')[0] || '').replace(/-/g, '/');

    // ── Cloud Functionへリクエスト ──────────────────
    const payload = {
      apiKey:         apiKey,
      drawingSize:    drawingSize,
      applicantDate:  toDateOnly(stampData.applicantDate),
      applicantName:  String(stampData.applicantName || '').trim(),
      reviewerDate:   toDateOnly(stampData.reviewerDate),
      reviewerName:   String(stampData.reviewerName || '').trim(),
      directorDate:   toDateOnly(stampData.directorDate),
      directorName:   String(stampData.directorName || '').trim(),
      pdfBase64:      pdfB64,
    };

    const response = UrlFetchApp.fetch(functionUrl, {
      method:             'post',
      contentType:        'application/json',
      payload:            JSON.stringify(payload),
      muteHttpExceptions: true,
    });

    const statusCode = response.getResponseCode();
    if (statusCode !== 200) {
      console.error(`PDFスタンプ関数エラー (${statusCode}): ${response.getContentText()}`);
      return { success: false, message: `Cloud Functionがステータス${statusCode}を返しました。` };
    }

    const resultJson = JSON.parse(response.getContentText());
    if (resultJson.debug) {
      console.log('PDFスタンプ診断情報: ' + JSON.stringify(resultJson.debug));
    }
    if (!resultJson.success) {
      return {
        success: false,
        message: resultJson.message || 'Cloud Function側で処理に失敗しました。',
        debug:   resultJson.debug,
      };
    }

    // ── スタンプ済みPDFを「承認済み図面」フォルダ配下のサブフォルダへ新規保存 ──
    const stampedBytes = Utilities.base64Decode(resultJson.pdfBase64);
    const stampedBlob  = Utilities.newBlob(stampedBytes, 'application/pdf', file.getName());

    const baseFolder = DriveApp.getFolderById(approvedFolderId);
    const subfolder   = getOrCreateSubfolder_(baseFolder, subfolderName);
    const approvedFile = subfolder.createFile(stampedBlob);

    // ── 機械台帳にサブフォルダのリンクを記録（未記入の場合のみ） ──
    try {
      recordFolderLinks_(ss, mainNo, subfolder);
    } catch (linkErr) {
      console.warn(`機械台帳へのフォルダリンク記録に失敗しました（${mainNo}）: ${linkErr.message}`);
    }

    return {
      success:        true,
      message:        `PDFへ日付・氏名をスタンプし、「${subfolderName}」フォルダへ保存しました（${approvedFile.getName()}）。`,
      approvedFileUrl: approvedFile.getUrl(),
      debug:           resultJson.debug,
    };

  } catch (e) {
    console.error('stampApprovalDatesOnPdf_ error:', e);
    return { success: false, message: `エラーが発生しました：${e.message}` };
  }
}

/**
 * 機械台帳から、指定した主図番の機械名（日本語・英語）を取得する
 * @param {Spreadsheet} ss
 * @param {string} mainNo - 主図番（例："JB2015"）
 * @returns {{modelNameJp: string, modelNameEn: string, rowIndex: number}|null}
 */
function getMainNumberInfo_(ss, mainNo) {
  const sheet = ss.getSheetByName(SHEET_NAMES.MAIN_NUMBER);
  if (!sheet) return null;
  const lastRow = sheet.getLastRow();
  if (lastRow < GLOBAL_ROW.DATA_START) return null;

  const data = sheet
    .getRange(GLOBAL_ROW.DATA_START, 1, lastRow - GLOBAL_ROW.DATA_START + 1, COL_MAIN.DESCRIPTION)
    .getValues();

  for (let i = 0; i < data.length; i++) {
    if (String(data[i][COL_MAIN.MAIN_NO - 1]).trim() === mainNo) {
      return {
        modelNameJp: String(data[i][COL_MAIN.MODEL_NAME_JP - 1] || '').trim(),
        modelNameEn: String(data[i][COL_MAIN.MODEL_NAME_EN - 1] || '').trim(),
        description: String(data[i][COL_MAIN.DESCRIPTION   - 1] || '').trim(),
        rowIndex:    GLOBAL_ROW.DATA_START + i,
      };
    }
  }
  return null;
}

/**
 * 機械台帳のフォルダリンク欄・エクスプローラーパス欄を、実際に今回
 * 保存に使ったフォルダの内容で常に上書きする。
 *
 * 【設計変更の経緯】
 *   当初は「空欄の場合のみ書き込む」仕様だった（同じ主図番なら通常は
 *   同じサブフォルダになるはずなので、無用な書き込みを避ける意図）。
 *   しかし、保存先サブフォルダ名は機械台帳の機械名（日本語／英語、
 *   D列・E列）から都度組み立てられるため（stampApprovalDatesOnPdf_の
 *   subfolderName参照）、図面を登録し直す過程で機械名が変更されると、
 *   次回承認時には別名の新しいサブフォルダが作成される。この際「空欄
 *   でないから」という理由でI・J列を更新しないと、実際の保存先とは
 *   異なる古いフォルダへのリンクが残り続けてしまう不整合が起きていた。
 *   そのため、既存値の有無に関わらず、常に「今回実際に使ったフォルダ」
 *   の情報で上書きする方式に変更した（同じフォルダであれば同じ値を
 *   再度書き込むだけなので副作用はない）。
 *
 * @param {Spreadsheet} ss
 * @param {string} mainNo
 * @param {Folder} folder - 対象のDriveフォルダ（URL取得・パス組み立ての両方に使用）
 */
function recordFolderLinks_(ss, mainNo, folder) {
  const sheet = ss.getSheetByName(SHEET_NAMES.MAIN_NUMBER);
  if (!sheet) return;
  const info = getMainNumberInfo_(ss, mainNo);
  if (!info) return;

  sheet.getRange(info.rowIndex, COL_MAIN.FOLDER_URL)
    .setFormula(buildFileLinkFormula_(folder.getUrl(), '📁ブラウザで開く'));

  try {
    const explorerPath = buildExplorerPath_(folder);
    sheet.getRange(info.rowIndex, COL_MAIN.EXPLORER_PATH).setValue(explorerPath || '');
  } catch (e) {
    // エクスプローラーパスの組み立てに失敗しても、致命的ではないので警告のみ
    // （この場合はJ列を更新せず、直前までの値を残す）
    console.warn(`エクスプローラーパスの組み立てに失敗しました（${mainNo}）: ${e.message}`);
  }
}

/**
 * 指定したDriveフォルダの、Google Drive for desktop（旧File Stream）における
 * ローカルパス文字列を組み立てる。
 *
 * 【前提】
 *   - 対象フォルダは共有ドライブ上にあること
 *   - Google Drive for desktopが「[ドライブ文字]:\共有ドライブ\[共有ドライブ表示名]\...」
 *     という規定の形式でマウントされていること
 *   - ドライブ文字はスクリプトプロパティ「SHARED_DRIVE_LETTER」で設定する（未設定時は既定値 "G"）
 *
 * 【注意】
 *   この形式はGoogle側の仕様変更や、各PCのマウント設定（ドライブ文字・言語設定等）に
 *   よって変わる可能性がある。ズレる場合は SHARED_DRIVE_LETTER を調整するか、
 *   このロジック自体の見直しが必要になる。
 *
 * @param {Folder} folder
 * @returns {string} 例："G:\共有ドライブ\図面管理\JB2015_扇形チップ溶接機_CTW2021"
 */
function buildExplorerPath_(folder) {
  const driveLetter = PropertiesService.getScriptProperties().getProperty('SHARED_DRIVE_LETTER') || 'G';

  // ── フォルダ階層をたどり、共有ドライブのルートまでの名前を集める ──
  const pathParts = [folder.getName()];
  let currentId = folder.getId();

  for (let i = 0; i < 50; i++) {  // 異常な無限ループ防止のための上限
    const meta = Drive.Files.get(currentId, {
      fields: 'id, name, parents, driveId',
      supportsAllDrives: true,
    });

    if (!meta.parents || meta.parents.length === 0) {
      // 親がない＝共有ドライブのルート自体に到達
      break;
    }

    const parentId = meta.parents[0];
    if (parentId === meta.driveId) {
      // 親IDが共有ドライブID自体 → 共有ドライブのルート直下だった
      break;
    }

    const parentMeta = Drive.Files.get(parentId, {
      fields: 'id, name, parents, driveId',
      supportsAllDrives: true,
    });
    pathParts.unshift(parentMeta.name);
    currentId = parentId;
  }

  // ── 共有ドライブ自体の表示名を取得 ──
  const fileMeta = Drive.Files.get(folder.getId(), { fields: 'driveId', supportsAllDrives: true });
  if (!fileMeta.driveId) {
    // マイドライブ上のフォルダだった場合は非対応（共有ドライブ運用前提のため）
    console.warn('対象フォルダは共有ドライブ上にないため、エクスプローラーパスの生成をスキップしました。');
    return '';
  }
  const driveMeta = Drive.Drives.get(fileMeta.driveId, { fields: 'name' });
  const sharedDriveName = driveMeta.name;

  return `${driveLetter}:\\共有ドライブ\\${sharedDriveName}\\${pathParts.join('\\')}`;
}

/**
 * 指定した親フォルダの直下に、指定した名前のフォルダを取得（なければ作成）する
 *
 * 【注意：ゴミ箱内フォルダの扱い】
 *   DriveApp.Folder.getFoldersByName() は、対象フォルダがゴミ箱に
 *   入っている（削除済みの）場合でも検索結果に含めてしまう（Apps Script /
 *   Drive APIの既知の挙動）。そのため、この関数が単純に「見つかった
 *   最初の1件」を返すと、運用者が「承認済み図面」フォルダ配下の
 *   機械別サブフォルダを一度ゴミ箱へ移動したうえで図面を承認し直した
 *   場合に、ゴミ箱内の古いフォルダをそのまま「既存フォルダ」として
 *   使い続けてしまい、新しく承認したPDFがゴミ箱の中に保存され、
 *   機械台帳のフォルダリンクもゴミ箱内を指したままになる、という
 *   不具合が起きる。
 *   これを防ぐため、検索結果からゴミ箱内のフォルダは除外し、
 *   見つからなければ新規にフォルダを作成する（＝ゴミ箱内の古いフォルダを
 *   自動復元することはしない。復元してしまうと、ユーザーが意図的に
 *   削除した内容が復活してしまい、かえって混乱を招くため）。
 *
 * @param {Folder} parentFolder
 * @param {string} folderName
 * @returns {Folder}
 */
function getOrCreateSubfolder_(parentFolder, folderName) {
  const candidates = parentFolder.getFoldersByName(folderName);
  while (candidates.hasNext()) {
    const candidate = candidates.next();
    if (!candidate.isTrashed()) {
      return candidate;
    }
  }
  return parentFolder.createFolder(folderName);
}

/**
 * DriveのファイルURLからファイルIDを抽出する
 * 対応形式の例：
 *   https://drive.google.com/file/d/FILE_ID/view?usp=drivesdk
 *   https://drive.google.com/open?id=FILE_ID
 * @param {string} url
 * @returns {string|null}
 */
function extractDriveFileId_(url) {
  if (!url) return null;
  const m1 = url.match(/\/d\/([a-zA-Z0-9_-]+)/);
  if (m1) return m1[1];
  const m2 = url.match(/[?&]id=([a-zA-Z0-9_-]+)/);
  if (m2) return m2[1];
  return null;
}

// ============================================================
// 【デバッグ用】
// ============================================================
/**
 * 任意のDriveファイルIDに対してスタンプ処理を単体テストする
 */
function TEST_stampApprovalDatesOnPdf() {
  const TEST_FILE_ID = 'ここにテスト用PDFのDriveファイルIDを入力';
  const result = stampApprovalDatesOnPdf_(
    `https://drive.google.com/file/d/${TEST_FILE_ID}/view`,
    'A3',
    'JB2015',  // 機械台帳に存在する主図番を指定すること
    {
      applicantDate: '2026-07-09 10:00:00',
      applicantName: '鈴木一郎',
      reviewerDate:  '2026-07-10 09:00:00',
      reviewerName:  '田中次郎',
      directorDate:  '2026-07-11 15:30:00',
      directorName:  '山田三郎',
    }
  );
  console.log(JSON.stringify(result, null, 2));
}