/**
 * ============================================================
 * 図面承認・採番・出図管理システム
 * ファイル: 23_ApprovalWebApp.gs  ―  承認待ち一覧 Web アプリ
 * ============================================================
 *
 * 【機能】
 *   - ログインユーザーの承認待ち図面を一覧表示
 *   - 承認・差戻しの処理と次の承認者への通知
 *
 * 【承認フロー】
 *   検図中 → 課長承認待ち → 部長承認待ち → 承認済
 *   ※ どの段階でも差戻し可能（ステータスを「検図中」に戻し申請者へ通知）
 *
 * 【スクリプトプロパティ（事前設定が必要）】
 *   SPREADSHEET_ID       : スプレッドシートのID
 *   APPROVAL_WEB_APP_URL : このWebアプリのデプロイURL（自身のURL）
 *
 * 【Webアプリのデプロイ手順】
 *   1. Apps Script エディタ → 「デプロイ」→「新しいデプロイ」
 *   2. 種類：ウェブアプリ
 *   3. 次のユーザーとして実行：「自分」
 *   4. アクセスできるユーザー：「組織内の全員」
 *   5. デプロイ後のURLをスクリプトプロパティ「APPROVAL_WEB_APP_URL」に登録
 *
 * 【図面台帳の列構成（22列・2026-07マイグレーション済み）】
 *   A=フル図番, B=主図番, C=子図番, D=改訂記号,
 *   E=図名(JPN), F=英名(NAME), G=ユニット名, H=機械名, I=材質, J=縮尺,
 *   K=図面サイズ(A0~A4), L=ステータス, M=申請者, N=申請日, O=検図者,
 *   P=検図者承認日時, Q=課長, R=課長承認日時, S=部長, T=部長承認日時,
 *   U=図面リンク（承認前＝申請中図面／承認後＝承認済み図面）, V=申請バッチID
 */

// ============================================================
// Web アプリのエントリーポイント
// ============================================================
function doGet(e) {
  try {
    const batchId = (e && e.parameter && e.parameter.batchId) ? String(e.parameter.batchId) : '';
    const template = HtmlService.createTemplateFromFile('24_ApprovalWebApp');
    template.batchId = batchId;
    return template
      .evaluate()
      .setTitle('図面承認待ち一覧')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  } catch (err) {
    // ここで例外が発生すると、対処しない場合はブラウザに何も表示されず
    // 「真っ白なページ」になってしまう。必ずエラー内容が見えるHTMLを返す。
    console.error('doGet error:', err);
    return HtmlService.createHtmlOutput(
      '<div style="font-family:sans-serif;padding:24px;color:#721c24;' +
      'background:#f8d7da;border:1px solid #f5c6cb;border-radius:8px;">' +
      '<h3>⚠️ ページの表示中にエラーが発生しました</h3>' +
      '<p>' + (err && err.message ? err.message : String(err)) + '</p>' +
      '<p style="font-size:12px;color:#555;margin-top:10px;">' +
      '管理者の方へ：スクリプトエディタの「実行数」ログでスタックトレースを確認してください。' +
      '再デプロイ（新しいバージョン）が反映されていない可能性もあります。</p>' +
      '</div>'
    ).setTitle('図面承認待ち一覧 - エラー');
  }
}

// ============================================================
// クライアントから呼ばれるサーバー関数
// ============================================================

/**
 * ログインユーザーの承認待ち図面一覧を返す
 * @param {string} [batchId] - 指定があれば、その申請バッチの図面のみに絞り込む
 * @returns {{ userName, items: Array, error?, batchId? }}
 */
function getMyPendingApprovals(batchId) {
  const ss = getSpreadsheet_();

  // 一覧表示は図面台帳を直接スキャンせず、承認進行中の行だけを保持する
  // 「承認待ち索引」を参照する（05_SearchIndex.gs を参照）。図面台帳の
  // 総件数が増えても、この索引は常に「今進行中の承認件数」程度の
  // 小ささを保つため、Web Appを開く・更新するたびに発生するこの
  // 最頻出の読み込みを高速に保てる。
  const pendingSheet = ensureIndexSheet_(SHEET_NAMES.PENDING_INDEX);

  const email    = Session.getActiveUser().getEmail();
  const userName = getNameByEmail_(ss, email);

  if (!userName) {
    return {
      userName: '',
      items: [],
      error: 'ユーザーマスタに氏名が登録されていません。管理者に連絡してください。',
    };
  }

  const lastRow = pendingSheet.getLastRow();
  if (lastRow < GLOBAL_ROW.DATA_START) return { userName, items: [] };

  const range = pendingSheet.getRange(GLOBAL_ROW.DATA_START, 1, lastRow - GLOBAL_ROW.DATA_START + 1, DRAWING_DB_COL_COUNT);
  const rows     = range.getValues();
  const formulas = range.getFormulas(); // FILE_URL列（HYPERLINK数式）から実URLを取り出すために使用

  const items = [];
  rows.forEach((row, i) => {
    const status   = String(row[COL_DB.STATUS   - 1]);
    const reviewer = String(row[COL_DB.REVIEWER - 1]);
    const manager  = String(row[COL_DB.MANAGER  - 1]);
    const director = String(row[COL_DB.DIRECTOR - 1]);
    const rowBatch = String(row[COL_DB.BATCH_ID - 1] || '');

    const isMyTurn =
      (status === STATUS.REVIEWING        && reviewer === userName) ||
      (status === STATUS.WAITING_MANAGER  && manager  === userName) ||
      (status === STATUS.WAITING_DIRECTOR && director === userName);

    if (!isMyTurn) return;

    // batchId が指定されている場合は、そのバッチの図面のみに絞り込む
    if (batchId && rowBatch !== batchId) return;

    const appDate = row[COL_DB.APP_DATE - 1];
    const fileUrl = extractUrlFromCellFormula_(formulas[i][COL_DB.FILE_URL - 1]) ||
                    String(row[COL_DB.FILE_URL - 1]).trim();

    // AIチェック（検図者への参考情報）。改行区切りの箇条書きテキストとして
    // 図面台帳W列に保存されているものを、配列に分解してクライアントへ渡す
    // （20_DrawingApproveRequest.gs の setApproversAndRegister を参照）。
    const aiCheckRaw = String(row[COL_DB.AI_CHECK - 1] || '').trim();
    const aiCheckPoints = aiCheckRaw
      ? aiCheckRaw.split('\n').map(line => line.replace(/^\d+\.\s*/, '').trim()).filter(v => v)
      : [];

    items.push({
      rowIndex:     GLOBAL_ROW.DATA_START + i,
      fullNo:       String(row[COL_DB.FULL_NO      - 1]),
      mainNo:       String(row[COL_DB.MAIN_NO      - 1]),
      nameJp:       String(row[COL_DB.NAME_JP      - 1]),
      modelName:    String(row[COL_DB.MODEL_NAME   - 1]),
      status,
      appDate:      appDate instanceof Date
        ? Utilities.formatDate(appDate, Session.getScriptTimeZone(), 'yyyy-MM-dd')
        : String(appDate),
      applicant:    String(row[COL_DB.APPLICANT    - 1]),
      fileUrl:      fileUrl,
      batchId:      rowBatch,
      aiCheckPoints: aiCheckPoints, // AIチェック（検図者への参考情報。正式な検図結果ではない）
    });
  });

  // 説明文は主図番（機械）単位で保存されている。1申請＝同一主図番のため、
  // 一覧内の最初の項目の主図番から代表して1件だけ取得する
  // （getMainNumberInfo_ は 25_PdfApprovalStamp.gs で定義）。
  let description = '';
  if (items.length > 0) {
    const mainInfo = getMainNumberInfo_(ss, items[0].mainNo);
    if (mainInfo) description = mainInfo.description;
  }

  return { userName, items, batchId: batchId || '', description };
}

/**
 * 図面を承認する
 * @param {string} fullNo - フル図番（10桁）
 * @returns {{ success, message }}
 */
function approveDrawing(fullNo) {
  try {
    const ss       = getSpreadsheet_();
    const dbSheet  = ss.getSheetByName(SHEET_NAMES.DRAWING_DB);
    const email    = Session.getActiveUser().getEmail();
    const userName = getNameByEmail_(ss, email);
    if (!userName) return { success: false, message: 'ユーザーマスタに氏名が登録されていません。' };

    const row = findRow_(dbSheet, fullNo);
    if (!row) return { success: false, message: `図面番号「${fullNo}」が見つかりません。` };

    const status    = String(row.data[COL_DB.STATUS    - 1]);
    const reviewer  = String(row.data[COL_DB.REVIEWER  - 1]);
    const manager   = String(row.data[COL_DB.MANAGER   - 1]);
    const director  = String(row.data[COL_DB.DIRECTOR  - 1]);
    const nameJp    = String(row.data[COL_DB.NAME_JP   - 1]);
    const applicant = String(row.data[COL_DB.APPLICANT - 1]);
    const batchId   = String(row.data[COL_DB.BATCH_ID  - 1] || '');
    const fileUrl       = row.fileUrl;
    const mainNo        = String(row.data[COL_DB.MAIN_NO       - 1] || '');
    const drawingSize   = String(row.data[COL_DB.DRAWING_SIZE  - 1] || '').trim().toUpperCase();
    const reviewerDate  = formatSheetDateTime_(row.data[COL_DB.REVIEWER_DATE - 1], ss.getSpreadsheetTimeZone());
    const applicantDate = formatSheetDateTime_(row.data[COL_DB.APP_DATE      - 1], ss.getSpreadsheetTimeZone());
    const nextUrl   = appendBatchIdToUrl_(getWebAppUrl_(), batchId);
    const now       = Utilities.formatDate(
      new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss'
    );

    if (status === STATUS.REVIEWING && reviewer === userName) {
      // 検図者承認 → 課長承認待ちへ
      dbSheet.getRange(row.rowIndex, COL_DB.REVIEWER_DATE).setValue(now);
      dbSheet.getRange(row.rowIndex, COL_DB.STATUS).setValue(STATUS.WAITING_MANAGER);
      syncIndexesForFullNo_(dbSheet, fullNo); // 05_SearchIndex.gs
      sendNotificationEmail_(ss, manager,
        '【図面承認依頼】課長承認をお願いします',
        `${manager} 様\n\n下記の図面について、課長承認をお願いします。\n\n` +
        `図番：${fullNo}\n図名：${nameJp}\n検図者承認日時：${now}\n\n` +
        `▼ 承認はこちらから\n${nextUrl}\n\n---\n図面承認・採番・出図管理システム`,
        userName
      );
      return { success: true, message: `「${fullNo}」を承認しました。課長（${manager}）へ通知しました。` };

    } else if (status === STATUS.WAITING_MANAGER && manager === userName) {
      // 課長承認 → 部長承認待ちへ
      dbSheet.getRange(row.rowIndex, COL_DB.MANAGER_DATE).setValue(now);
      dbSheet.getRange(row.rowIndex, COL_DB.STATUS).setValue(STATUS.WAITING_DIRECTOR);
      syncIndexesForFullNo_(dbSheet, fullNo); // 05_SearchIndex.gs
      sendNotificationEmail_(ss, director,
        '【図面承認依頼】部長承認をお願いします',
        `${director} 様\n\n下記の図面について、部長承認をお願いします。\n\n` +
        `図番：${fullNo}\n図名：${nameJp}\n課長承認日時：${now}\n\n` +
        `▼ 承認はこちらから\n${nextUrl}\n\n---\n図面承認・採番・出図管理システム`,
        userName
      );
      return { success: true, message: `「${fullNo}」を承認しました。部長（${director}）へ通知しました。` };

    } else if (status === STATUS.WAITING_DIRECTOR && director === userName) {
      // 部長承認 → 承認済
      dbSheet.getRange(row.rowIndex, COL_DB.DIRECTOR_DATE).setValue(now);
      dbSheet.getRange(row.rowIndex, COL_DB.STATUS).setValue(STATUS.APPROVED);

      // PDFへ日付・氏名をスタンプ（失敗しても承認処理自体は継続する）
      const stampResult = stampApprovalDatesOnPdf_(fileUrl, drawingSize, mainNo, {
        applicantDate: applicantDate,
        applicantName: getEnglishNameByJapaneseName_(ss, applicant),
        reviewerDate:  reviewerDate,
        reviewerName:  getEnglishNameByJapaneseName_(ss, reviewer),
        directorDate:  now,
        directorName:  getEnglishNameByJapaneseName_(ss, director),
      });
      if (!stampResult.success) {
        console.warn(`PDFスタンプに失敗しました（${fullNo}）: ${stampResult.message}`);
      } else if (stampResult.approvedFileUrl) {
        // 図面リンク（U列）を、申請中図面のリンクから承認済み図面のリンクへ上書きする
        // （申請中図面フォルダのファイルは定期的に削除される運用のため）
        dbSheet.getRange(row.rowIndex, COL_DB.FILE_URL)
          .setFormula(buildFileLinkFormula_(stampResult.approvedFileUrl, '📄図面を開く'));
      }
      syncIndexesForFullNo_(dbSheet, fullNo); // 05_SearchIndex.gs（承認済みへの遷移で承認待ち索引から自動的に除外される）

      sendNotificationEmail_(ss, applicant,
        '【図面承認完了】承認されました',
        `${applicant} 様\n\n下記の図面が承認されました。\n\n` +
        `図番：${fullNo}\n図名：${nameJp}\n部長承認日時：${now}\n\n---\n図面承認・採番・出図管理システム`,
        userName
      );
      return { success: true, message: `「${fullNo}」が承認済になりました。申請者（${applicant}）へ通知しました。` };

    } else {
      return { success: false, message: `現在のステータス「${status}」ではこの操作はできません。` };
    }

  } catch (e) {
    console.error('approveDrawing error:', e);
    return { success: false, message: `エラーが発生しました：${e.message}` };
  }
}

/**
 * 図面を一括承認する（50件など大量の一括登録に対応）
 * 通知メールは宛先ごとにまとめて1通に集約して送信する。
 * @param {string[]} fullNoList - フル図番の配列
 * @returns {{ success, message, results: Array<{fullNo, success, message}> }}
 */
function approveDrawingsBulk(fullNoList) {
  if (!Array.isArray(fullNoList) || fullNoList.length === 0) {
    return { success: false, message: '承認対象が選択されていません。', results: [] };
  }

  try {
    const ss       = getSpreadsheet_();
    const dbSheet  = ss.getSheetByName(SHEET_NAMES.DRAWING_DB);
    const email    = Session.getActiveUser().getEmail();
    const userName = getNameByEmail_(ss, email);
    if (!userName) {
      return { success: false, message: 'ユーザーマスタに氏名が登録されていません。', results: [] };
    }

    const now = Utilities.formatDate(
      new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss'
    );

    // 通知をまとめるためのバケツ（宛先×次段階ごとに1通へ集約）
    const notifyBucket = {};
    const results      = [];
    let okCount = 0, ngCount = 0;
    const stampFailures = []; // PDFスタンプに失敗した図面（画面へ返して即確認できるようにする）

    fullNoList.forEach(fullNo => {
      try {
        const row = findRow_(dbSheet, fullNo);
        if (!row) {
          results.push({ fullNo, success: false, message: '図面番号が見つかりません。' });
          ngCount++;
          return;
        }

        const status    = String(row.data[COL_DB.STATUS    - 1]);
        const reviewer  = String(row.data[COL_DB.REVIEWER  - 1]);
        const manager   = String(row.data[COL_DB.MANAGER   - 1]);
        const director  = String(row.data[COL_DB.DIRECTOR  - 1]);
        const nameJp    = String(row.data[COL_DB.NAME_JP   - 1]);
        const applicant = String(row.data[COL_DB.APPLICANT - 1]);
        const batchId   = String(row.data[COL_DB.BATCH_ID  - 1] || '');

        let nextRecipient, stageKey, resultMsg;

        if (status === STATUS.REVIEWING && reviewer === userName) {
          dbSheet.getRange(row.rowIndex, COL_DB.REVIEWER_DATE).setValue(now);
          dbSheet.getRange(row.rowIndex, COL_DB.STATUS).setValue(STATUS.WAITING_MANAGER);
          nextRecipient = manager; stageKey = 'manager';
          resultMsg = `承認しました。課長（${manager}）へ通知します。`;

        } else if (status === STATUS.WAITING_MANAGER && manager === userName) {
          dbSheet.getRange(row.rowIndex, COL_DB.MANAGER_DATE).setValue(now);
          dbSheet.getRange(row.rowIndex, COL_DB.STATUS).setValue(STATUS.WAITING_DIRECTOR);
          nextRecipient = director; stageKey = 'director';
          resultMsg = `承認しました。部長（${director}）へ通知します。`;

        } else if (status === STATUS.WAITING_DIRECTOR && director === userName) {
          dbSheet.getRange(row.rowIndex, COL_DB.DIRECTOR_DATE).setValue(now);
          dbSheet.getRange(row.rowIndex, COL_DB.STATUS).setValue(STATUS.APPROVED);
          nextRecipient = applicant; stageKey = 'applicant';
          resultMsg = `承認済になりました。申請者（${applicant}）へ通知します。`;

          // PDFへ日付・氏名をスタンプ（失敗しても承認処理自体は継続する）
          const fileUrl      = row.fileUrl;
          const mainNo       = String(row.data[COL_DB.MAIN_NO       - 1] || '');
          const drawingSize  = String(row.data[COL_DB.DRAWING_SIZE  - 1] || '').trim().toUpperCase();
          const reviewerDate  = formatSheetDateTime_(row.data[COL_DB.REVIEWER_DATE - 1], ss.getSpreadsheetTimeZone());
          const applicantDate = formatSheetDateTime_(row.data[COL_DB.APP_DATE      - 1], ss.getSpreadsheetTimeZone());
          const stampResult  = stampApprovalDatesOnPdf_(fileUrl, drawingSize, mainNo, {
            applicantDate: applicantDate,
            applicantName: getEnglishNameByJapaneseName_(ss, applicant),
            reviewerDate:  reviewerDate,
            reviewerName:  getEnglishNameByJapaneseName_(ss, reviewer),
            directorDate:  now,
            directorName:  getEnglishNameByJapaneseName_(ss, director),
          });
          if (!stampResult.success) {
            console.warn(`PDFスタンプに失敗しました（${fullNo}）: ${stampResult.message}`);
            stampFailures.push(`${fullNo}：${stampResult.message}`);
          } else if (stampResult.approvedFileUrl) {
            // 図面リンク（U列）を、申請中図面のリンクから承認済み図面のリンクへ上書きする
            dbSheet.getRange(row.rowIndex, COL_DB.FILE_URL)
              .setFormula(buildFileLinkFormula_(stampResult.approvedFileUrl, '📄図面を開く'));
          }

        } else {
          results.push({ fullNo, success: false, message: `現在のステータス「${status}」ではこの操作はできません。` });
          ngCount++;
          return;
        }

        syncIndexesForFullNo_(dbSheet, fullNo); // 05_SearchIndex.gs（3分岐いずれの場合も共通で同期する）

        // 宛先×次段階×バッチIDごとに1通へ集約（別バッチの図面は別メールになる）
        const bucketKey = stageKey + '|' + nextRecipient + '|' + batchId;
        if (!notifyBucket[bucketKey]) {
          notifyBucket[bucketKey] = { recipient: nextRecipient, stageKey, batchId, items: [] };
        }
        notifyBucket[bucketKey].items.push({ fullNo, nameJp });

        results.push({ fullNo, success: true, message: resultMsg });
        okCount++;

      } catch (itemErr) {
        console.error('approveDrawingsBulk item error:', fullNo, itemErr);
        results.push({ fullNo, success: false, message: `エラー：${itemErr.message}` });
        ngCount++;
      }
    });

    // 宛先×次段階ごとにまとめて1通のメールを送信
    Object.keys(notifyBucket).forEach(key => {
      const bucket    = notifyBucket[key];
      const listText  = bucket.items.map(it => `・${it.fullNo}　${it.nameJp}`).join('\n');
      let subject, bodyIntro, urlLine;

      if (bucket.stageKey === 'manager') {
        subject   = '【図面承認依頼】課長承認をお願いします（複数件）';
        bodyIntro = `下記 ${bucket.items.length} 件の図面について、課長承認をお願いします。`;
        urlLine   = `\n\n▼ 承認はこちらから\n${appendBatchIdToUrl_(getWebAppUrl_(), bucket.batchId)}`;
      } else if (bucket.stageKey === 'director') {
        subject   = '【図面承認依頼】部長承認をお願いします（複数件）';
        bodyIntro = `下記 ${bucket.items.length} 件の図面について、部長承認をお願いします。`;
        urlLine   = `\n\n▼ 承認はこちらから\n${appendBatchIdToUrl_(getWebAppUrl_(), bucket.batchId)}`;
      } else {
        subject   = '【図面承認完了】承認されました（複数件）';
        bodyIntro = `下記 ${bucket.items.length} 件の図面が承認されました。`;
        urlLine   = '';
      }

      sendNotificationEmail_(ss, bucket.recipient, subject,
        `${bucket.recipient} 様\n\n${bodyIntro}\n\n${listText}${urlLine}\n\n---\n図面承認・採番・出図管理システム`,
        userName
      );
    });

    const stampWarningText = stampFailures.length > 0
      ? `\n［PDFスタンプ失敗：${stampFailures.length}件］\n` + stampFailures.join('\n')
      : '';

    return {
      success: ngCount === 0,
      message: `${okCount} 件を承認しました。` +
               (ngCount > 0 ? `（${ngCount} 件は処理できませんでした。詳細は結果を確認してください）` : '') +
               stampWarningText,
      results,
    };

  } catch (e) {
    console.error('approveDrawingsBulk error:', e);
    return { success: false, message: `エラーが発生しました：${e.message}`, results: [] };
  }
}

/**
 * 図面を一括差戻しする（同一の差戻し理由を複数件へ適用）
 * @param {string[]} fullNoList - フル図番の配列
 * @param {string} reason - 差戻し理由（全件共通）
 * @returns {{ success, message, results: Array<{fullNo, success, message}> }}
 */
function rejectDrawingsBulk(fullNoList, reason) {
  if (!reason || !reason.trim()) {
    return { success: false, message: '差戻し理由を入力してください。', results: [] };
  }
  if (!Array.isArray(fullNoList) || fullNoList.length === 0) {
    return { success: false, message: '差戻し対象が選択されていません。', results: [] };
  }

  try {
    const ss       = getSpreadsheet_();
    const dbSheet  = ss.getSheetByName(SHEET_NAMES.DRAWING_DB);
    const email    = Session.getActiveUser().getEmail();
    const userName = getNameByEmail_(ss, email);
    if (!userName) {
      return { success: false, message: 'ユーザーマスタに氏名が登録されていません。', results: [] };
    }

    const now = Utilities.formatDate(
      new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss'
    );

    const notifyBucket = {}; // applicant -> items[]
    const results       = [];
    let okCount = 0, ngCount = 0;

    fullNoList.forEach(fullNo => {
      try {
        const row = findRow_(dbSheet, fullNo);
        if (!row) {
          results.push({ fullNo, success: false, message: '図面番号が見つかりません。' });
          ngCount++;
          return;
        }

        const status    = String(row.data[COL_DB.STATUS    - 1]);
        const reviewer  = String(row.data[COL_DB.REVIEWER  - 1]);
        const manager   = String(row.data[COL_DB.MANAGER   - 1]);
        const director  = String(row.data[COL_DB.DIRECTOR  - 1]);
        const nameJp    = String(row.data[COL_DB.NAME_JP   - 1]);
        const applicant = String(row.data[COL_DB.APPLICANT - 1]);
        const mainNo    = String(row.data[COL_DB.MAIN_NO   - 1]);
        const subNo     = String(row.data[COL_DB.SUB_NO    - 1]);

        const canReject =
          (status === STATUS.REVIEWING        && reviewer === userName) ||
          (status === STATUS.WAITING_MANAGER  && manager  === userName) ||
          (status === STATUS.WAITING_DIRECTOR && director === userName);

        if (!canReject) {
          results.push({ fullNo, success: false, message: `現在のステータス「${status}」ではこの操作はできません。` });
          ngCount++;
          return;
        }

        // 削除対象の行番号を記録（後でまとめて削除）
        if (!notifyBucket[applicant]) notifyBucket[applicant] = [];
        notifyBucket[applicant].push({ fullNo, nameJp, rowIndex: row.rowIndex, mainNo, subNo });

        results.push({ fullNo, success: true, message: `差戻しました。申請者（${applicant}）へ通知します。` });
        okCount++;

      } catch (itemErr) {
        console.error('rejectDrawingsBulk item error:', fullNo, itemErr);
        results.push({ fullNo, success: false, message: `エラー：${itemErr.message}` });
        ngCount++;
      }
    });

    // 行番号が大きい順に削除（上から削除すると行番号がずれるため）
    const allRows = Object.values(notifyBucket).flat();
    allRows.sort((a, b) => b.rowIndex - a.rowIndex);
    allRows.forEach(item => dbSheet.deleteRow(item.rowIndex));

    // 索引からも削除する（05_SearchIndex.gs）。図面台帳からの行削除が
    // 全て完了した後にまとめて行う（rebuildDrawingIndexEntryIfNeeded_ が
    // 図面台帳を再スキャンするため、削除が済んだ状態で呼ぶ必要がある）。
    try {
      allRows.forEach(item => {
        removeFromPendingIndex_(item.fullNo);
        rebuildDrawingIndexEntryIfNeeded_(item.mainNo, item.subNo, item.fullNo);
      });
    } catch (indexErr) {
      console.warn(`索引の更新に失敗しました: ${indexErr.message}`);
    }

    // 申請者ごとにまとめて通知メールを送信（差戻し理由はメール本文にのみ記載）
    Object.keys(notifyBucket).forEach(applicant => {
      const items    = notifyBucket[applicant];
      const listText = items.map(it => `・${it.fullNo}　${it.nameJp}`).join('\n');

      sendNotificationEmail_(ss, applicant,
        '【図面差戻し】修正をお願いします（複数件）',
        `${applicant} 様\n\n下記 ${items.length} 件の図面が差戻されました。内容を確認・修正後、再申請してください。\n\n` +
        `${listText}\n\n差戻し者：${userName}\n差戻し日時：${now}\n差戻し理由：${reason.trim()}\n\n` +
        `---\n図面承認・採番・出図管理システム`,
        userName
      );
    });

    return {
      success: ngCount === 0,
      message: `${okCount} 件を差戻しました。` + (ngCount > 0 ? `（${ngCount} 件は処理できませんでした。詳細は結果を確認してください）` : ''),
      results,
    };

  } catch (e) {
    console.error('rejectDrawingsBulk error:', e);
    return { success: false, message: `エラーが発生しました：${e.message}`, results: [] };
  }
}

/**
 * 図面を差戻しする
 * @param {string} fullNo  - フル図番（10桁）
 * @param {string} reason  - 差戻し理由
 * @returns {{ success, message }}
 */
function rejectDrawing(fullNo, reason) {
  try {
    if (!reason || !reason.trim()) {
      return { success: false, message: '差戻し理由を入力してください。' };
    }

    const ss       = getSpreadsheet_();
    const dbSheet  = ss.getSheetByName(SHEET_NAMES.DRAWING_DB);
    const email    = Session.getActiveUser().getEmail();
    const userName = getNameByEmail_(ss, email);
    if (!userName) return { success: false, message: 'ユーザーマスタに氏名が登録されていません。' };

    const row = findRow_(dbSheet, fullNo);
    if (!row) return { success: false, message: `図面番号「${fullNo}」が見つかりません。` };

    const status    = String(row.data[COL_DB.STATUS    - 1]);
    const reviewer  = String(row.data[COL_DB.REVIEWER  - 1]);
    const manager   = String(row.data[COL_DB.MANAGER   - 1]);
    const director  = String(row.data[COL_DB.DIRECTOR  - 1]);
    const nameJp    = String(row.data[COL_DB.NAME_JP   - 1]);
    const applicant = String(row.data[COL_DB.APPLICANT - 1]);
    const mainNo    = String(row.data[COL_DB.MAIN_NO   - 1]);
    const subNo     = String(row.data[COL_DB.SUB_NO    - 1]);

    const canReject =
      (status === STATUS.REVIEWING        && reviewer === userName) ||
      (status === STATUS.WAITING_MANAGER  && manager  === userName) ||
      (status === STATUS.WAITING_DIRECTOR && director === userName);

    if (!canReject) {
      return { success: false, message: `現在のステータス「${status}」ではこの操作はできません。` };
    }

    const now = Utilities.formatDate(
      new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss'
    );

    // 図面台帳から該当行を削除
    dbSheet.deleteRow(row.rowIndex);

    // 索引からも削除する（05_SearchIndex.gs）。図面索引側は、削除された行が
    // 「最新」として登録されていた場合のみ、図面台帳を再スキャンして
    // 次点のリビジョンを再計算する（差戻しは低頻度操作のため許容）。
    try {
      removeFromPendingIndex_(fullNo);
      rebuildDrawingIndexEntryIfNeeded_(mainNo, subNo, fullNo);
    } catch (indexErr) {
      console.warn(`索引の更新に失敗しました（${fullNo}）: ${indexErr.message}`);
    }

    // 申請者へ差戻し通知（差戻し理由はメール本文にのみ記載）
    sendNotificationEmail_(ss, applicant,
      '【図面差戻し】修正をお願いします',
      `${applicant} 様\n\n下記の図面が差戻されました。内容を確認・修正後、再申請してください。\n\n` +
      `図番：${fullNo}\n図名：${nameJp}\n差戻し者：${userName}\n差戻し日時：${now}\n` +
      `差戻し理由：${reason.trim()}\n\n` +
      `---\n図面承認・採番・出図管理システム`,
      userName
    );

    return {
      success: true,
      message: `「${fullNo}」を差戻しました。申請者（${applicant}）へ通知しました。`,
    };

  } catch (e) {
    console.error('rejectDrawing error:', e);
    return { success: false, message: `エラーが発生しました：${e.message}` };
  }
}

// ============================================================
// プライベートヘルパー
// ============================================================

function getSpreadsheet_() {
  const id = PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID');
  if (!id) throw new Error(
    'スクリプトプロパティ「SPREADSHEET_ID」が設定されていません。'
  );
  return SpreadsheetApp.openById(id);
}

function getNameByEmail_(ss, email) {
  const sheet = ss.getSheetByName(SHEET_NAMES.USER_MASTER);
  if (!sheet) return '';
  const lastRow = sheet.getLastRow();
  if (lastRow < GLOBAL_ROW.DATA_START) return '';
  const data  = sheet.getRange(GLOBAL_ROW.DATA_START, 1, lastRow - GLOBAL_ROW.DATA_START + 1, 2).getValues();
  const match = data.find(row => String(row[0]).trim() === email);
  return match ? String(match[1]).trim() : '';
}

/**
 * ユーザーマスタの氏名（日本語）から、対応するNAME（英語表記）を取得する。
 * PDF承認スタンプ機能で、図面台帳に保存されている日本語氏名を
 * 英語表記に変換するために使用する。
 * @param {Spreadsheet} ss
 * @param {string} nameJp - 図面台帳の申請者・検図者・部長欄に入っている日本語氏名
 * @returns {string} 英語表記（見つからない場合は空文字）
 */
function getEnglishNameByJapaneseName_(ss, nameJp) {
  if (!nameJp) return '';
  const sheet = ss.getSheetByName(SHEET_NAMES.USER_MASTER);
  if (!sheet) return '';
  const lastRow = sheet.getLastRow();
  if (lastRow < GLOBAL_ROW.DATA_START) return '';
  const data = sheet
    .getRange(GLOBAL_ROW.DATA_START, 1, lastRow - GLOBAL_ROW.DATA_START + 1, COL_USER.NAME_EN)
    .getValues();
  const match = data.find(row => String(row[COL_USER.NAME_JP - 1]).trim() === String(nameJp).trim());
  if (!match) {
    console.warn(`ユーザーマスタに氏名「${nameJp}」が見つかりませんでした（NAME取得不可）。`);
    return '';
  }
  const nameEn = String(match[COL_USER.NAME_EN - 1]).trim();
  if (!nameEn) {
    console.warn(`氏名「${nameJp}」のNAME（英語表記）が未登録です。`);
  }
  return nameEn;
}

function findRow_(dbSheet, fullNo) {
  const lastRow = dbSheet.getLastRow();
  if (lastRow < GLOBAL_ROW.DATA_START) return null;
  const range = dbSheet.getRange(GLOBAL_ROW.DATA_START, 1, lastRow - GLOBAL_ROW.DATA_START + 1, DRAWING_DB_COL_COUNT);
  const rows  = range.getValues();
  const idx   = rows.findIndex(row => String(row[COL_DB.FULL_NO - 1]).trim() === fullNo);
  if (idx === -1) return null;

  // FILE_URL列はHYPERLINK数式のため、実際のURLはgetFormulaで別途取得する
  const fileUrlFormula = dbSheet.getRange(GLOBAL_ROW.DATA_START + idx, COL_DB.FILE_URL).getFormula();
  const fileUrl = extractUrlFromCellFormula_(fileUrlFormula) ||
                  String(rows[idx][COL_DB.FILE_URL - 1]).trim();

  return { rowIndex: GLOBAL_ROW.DATA_START + idx, data: rows[idx], fileUrl: fileUrl };
}

/**
 * ユーザーマスタから宛先メールアドレスを引いて通知メールを送信する
 * @param {Spreadsheet} ss
 * @param {string} recipientName - 宛先の氏名（ユーザーマスタのB列と一致させる）
 * @param {string} subject
 * @param {string} body
 * @param {string} [actorName] - 実際にこの操作を行ったユーザーの氏名。
 *   指定があれば、送信者の表示名に反映する（実際の送信元アドレスは
 *   Webアプリの実行者＝オーナーのままだが、誰の操作によるものかが
 *   受信者から見て分かるようにするため）。
 */
function sendNotificationEmail_(ss, recipientName, subject, body, actorName) {
  const sheet = ss.getSheetByName(SHEET_NAMES.USER_MASTER);
  if (!sheet) return;
  const lastRow = sheet.getLastRow();
  if (lastRow < GLOBAL_ROW.DATA_START) return;
  const data  = sheet.getRange(GLOBAL_ROW.DATA_START, 1, lastRow - GLOBAL_ROW.DATA_START + 1, 2).getValues();
  const match = data.find(row => String(row[1]).trim() === recipientName);
  if (!match) {
    console.warn(`送信先「${recipientName}」がユーザーマスタに見つかりません。`);
    return;
  }
  const senderName = actorName
    ? `図面承認・採番・出図管理システム（${actorName}の操作）`
    : '図面承認・採番・出図管理システム';
  try {
    GmailApp.sendEmail(String(match[0]).trim(), subject, body, { name: senderName });
  } catch (e) {
    console.error(`メール送信失敗（${recipientName}）:`, e);
  }
}

function getWebAppUrl_() {
  return PropertiesService.getScriptProperties()
    .getProperty('APPROVAL_WEB_APP_URL') || '（承認WebアプリのURLは未設定です）';
}