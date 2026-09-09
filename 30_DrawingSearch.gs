/**
 * ============================================================
 * 図面承認・採番・出図管理システム
 * ファイル: 30_DrawingSearch.gs  ―  図面検索ビュー
 * ============================================================
 *
 * 【役割】
 *   図面台帳（フル図番単位）と機械台帳（機械単位）を横断的に
 *   キーワード検索し、モーダルダイアログ内で結果一覧・詳細（承認状況・
 *   改訂履歴・紐づく図面一覧）を確認できるようにする。
 *
 * 【設計方針】
 *   Webアプリとして別デプロイするほど重い機能ではないため、
 *   サイドバーから開くモーダルダイアログ1枚で完結させる
 *   （一覧表示 → 詳細表示は同一ダイアログ内でパネルを切り替える）。
 *   検索・詳細取得はいずれも読み取り専用（シートへの書き込みなし）。
 *
 * 【検索対象】
 *   - 図面台帳：フル図番・主図番・図名(JPN)・英名(NAME)・ユニット名・
 *               機械名・材質・申請者・ステータス・特徴属性の要約
 *               （OCR時にAIが生成した形状・特徴の要約文。図名等に出てこない
 *               形状キーワードでも図面を探せるようにするため対象に含める）
 *   - 機械台帳：主図番・機械名（日本語/英語）・申請者・備考
 *   スペース区切りの複数キーワードはAND検索（大文字小文字を区別しない）。
 *
 * 【絞り込み条件（フィルタ）】
 *   キーワード検索とは別に、以下の条件を追加のAND条件として指定できる。
 *   - 主図番（部分一致）／申請者（部分一致）／期間（From〜To）
 *     …図面・機械の両方に共通の項目（期間は図面＝申請日、機械＝着手日として扱う）
 *   - 図面サイズ（複数選択）／材質（部分一致）／検図者（部分一致）
 *     …図面台帳のみに存在する項目のため、機械台帳の検索には使用しない
 *   検索対象（図面／機械）はチェックボックスで個別にON/OFFでき、OFFにした
 *   対象の検索は行わない（クライアント側で該当フィルタ欄もグレーアウトする）。
 *
 * 【注意：図面台帳の列構成について】
 *   2026-07 マイグレーション④で「改訂連番」「最新フラグ」列は削除され、
 *   「承認済みPDFファイルURL」列は「図面リンク」列（U列＝FILE_URL）に
 *   統合されている（00_Config.gs の COL_DB 参照）。そのため、このファイルでは：
 *     - 改訂順は REV_MARK の文字コード順（'-' → 'A' → 'B' …）で判定する
 *     - 「最新」は同一主図番・子図番内で REV_MARK が最大の行を動的に判定する
 *     - ファイルリンクは FILE_URL 1本のみを使用する
 *       （承認前＝申請中図面／承認後＝承認済み図面のリンクに自動で切り替わる）
 */

// ============================================================
// ダイアログを開く（サイドバーから呼ばれる）
// ============================================================
function openDrawingSearchDialog() {
  const html = HtmlService.createHtmlOutputFromFile('31_DrawingSearchDialog')
    .setWidth(760)
    .setHeight(620);
  SpreadsheetApp.getUi().showModalDialog(html, '🔍 図面検索');
}

// ============================================================
// クライアントから呼ばれる：横断キーワード検索
// ============================================================
/**
 * @param {Object} params
 *   @param {string}   params.keyword          - 検索キーワード（スペース区切りでAND検索、空文字可）
 *   @param {boolean}  [params.searchDrawings=true]    - 図面台帳を検索対象に含めるか
 *   @param {boolean}  [params.searchMainNumbers=true] - 機械台帳を検索対象に含めるか
 *   @param {Object}   [params.filters] - 絞り込み条件（すべて任意、AND条件で適用）
 *     @param {string}   filters.mainNo       - 主図番（部分一致。図面・機械の両方に適用）
 *     @param {string}   filters.applicant    - 申請者（部分一致。図面・機械の両方に適用）
 *     @param {string}   filters.dateFrom     - 期間開始（'yyyy-MM-dd'。図面＝申請日／機械＝着手日）
 *     @param {string}   filters.dateTo       - 期間終了（'yyyy-MM-dd'。同上）
 *     @param {string[]} filters.drawingSizes - 図面サイズ（複数可。図面台帳のみに適用）
 *     @param {string}   filters.material     - 材質（部分一致。図面台帳のみに適用）
 *     @param {string}   filters.reviewer     - 検図者（部分一致。図面台帳のみに適用）
 * @returns {{drawings: Array, mainNumbers: Array}}
 */
function searchDrawingsAndMainNumbers(params) {
  params = params || {};
  const tokens  = String(params.keyword || '').trim().toLowerCase().split(/\s+/).filter(v => v);
  const filters = params.filters || {};

  const searchDrawings   = params.searchDrawings   !== false; // 未指定時はtrue扱い
  const searchMainNums   = params.searchMainNumbers !== false;

  // キーワード・フィルタのいずれも未指定なら、誤って全件を返さないよう空にする
  if (tokens.length === 0 && !hasActiveFilters_(filters)) {
    return { drawings: [], mainNumbers: [] };
  }
  // 検索対象が両方ともOFFの場合も何もしない
  if (!searchDrawings && !searchMainNums) {
    return { drawings: [], mainNumbers: [] };
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  return {
    drawings:    searchDrawings ? searchDrawingDb_(ss, tokens, filters)   : [],
    mainNumbers: searchMainNums ? searchMainNumbers_(ss, tokens, filters) : [],
  };
}

/**
 * フィルタオブジェクトに何か1つでも有効な条件が入っているかを判定する
 * （キーワードが空でも、フィルタだけで検索したいケースに対応するため）
 * @param {Object} filters
 * @returns {boolean}
 */
function hasActiveFilters_(filters) {
  if (!filters) return false;
  return !!(
    String(filters.mainNo    || '').trim() ||
    String(filters.applicant || '').trim() ||
    String(filters.material  || '').trim() ||
    String(filters.reviewer  || '').trim() ||
    String(filters.dateFrom  || '').trim() ||
    String(filters.dateTo    || '').trim() ||
    (Array.isArray(filters.drawingSizes) && filters.drawingSizes.length > 0)
  );
}

/**
 * 日付セルの値（Dateオブジェクトまたは文字列）を 'yyyy-MM-dd' 形式に正規化する。
 * 期間フィルタ（dateFrom/dateTo）との文字列比較に使用する。
 * @param {Date|string|*} value
 * @returns {string} 'yyyy-MM-dd'、変換できない場合は空文字
 */
function normalizeDateForCompare_(value) {
  if (!value) return '';
  if (value instanceof Date) {
    const tz = SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone();
    return Utilities.formatDate(value, tz, 'yyyy-MM-dd');
  }
  const str = String(value).trim();
  // 'yyyy-MM-dd' の先頭10文字のみ使う（日時文字列が来た場合の保険）
  const match = str.match(/^\d{4}-\d{2}-\d{2}/);
  return match ? match[0] : '';
}

/**
 * 日付が期間フィルタ（From〜To、どちらも省略可）の範囲内かを判定する
 * @param {string} dateStr  - normalizeDateForCompare_() 済みの 'yyyy-MM-dd'（空文字なら対象外）
 * @param {string} dateFrom - 'yyyy-MM-dd'（空文字なら下限なし）
 * @param {string} dateTo   - 'yyyy-MM-dd'（空文字なら上限なし）
 * @returns {boolean}
 */
function isWithinDateRange_(dateStr, dateFrom, dateTo) {
  if (!dateFrom && !dateTo) return true;
  if (!dateStr) return false; // 期間指定があるのに日付自体が空の行は対象外
  if (dateFrom && dateStr < dateFrom) return false;
  if (dateTo   && dateStr > dateTo)   return false;
  return true;
}

/**
 * 図面をキーワード検索する（絞り込みフィルタ対応）
 * ※ 図面台帳を直接スキャンせず、「最新リビジョンのみ」を保持する
 *   図面索引（SHEET_NAMES.DRAWING_INDEX）を検索対象にする。図面台帳の
 *   総件数（改訂履歴込み）が増えても、索引はユニークな部品点数以上には
 *   増えないため、検索速度への影響を抑えられる（05_SearchIndex.gs を参照）。
 *   そのため、この検索結果には各部品の最新リビジョンのみが含まれ、
 *   過去の改訂は出てこない（改訂履歴を確認したい場合は詳細パネルの
 *   「改訂履歴」を参照する）。
 * @param {Object} filters - searchDrawingsAndMainNumbers() の filters をそのまま渡す
 * @returns {Array<{fullNo, mainNo, nameJp, modelName, status, applicant, reviewer}>}
 */
function searchDrawingDb_(ss, tokens, filters) {
  filters = filters || {};
  const sheet = ensureIndexSheet_(SHEET_NAMES.DRAWING_INDEX); // 05_SearchIndex.gs
  const lastRow = sheet.getLastRow();
  if (lastRow < GLOBAL_ROW.DATA_START) return [];

  // FILE_URL列（HYPERLINK数式）から実URLを取り出すため、getFormulas()も併せて取得する
  const range    = sheet.getRange(GLOBAL_ROW.DATA_START, 1, lastRow - GLOBAL_ROW.DATA_START + 1, DRAWING_DB_COL_COUNT);
  const rows     = range.getValues();
  const formulas = range.getFormulas();

  // フィルタ条件を事前に正規化しておく
  // ※ 申請者・検図者は normalizeRomanName_（00_Config.gs）で比較する。
  //   過去図面インポートのローマ字氏名は表記ゆれ（大文字小文字・"."後の
  //   スペース有無）が生じやすいため、単純な toLowerCase() の部分一致だけでは
  //   例えば "K. Yokoyama"（スペースあり）と "K.YOKOYAMA"（スペースなし）を
  //   同一人物として拾えない。normalizeLegacyNamesInDrawingDb()（36_
  //   NormalizeLegacyNames.gs）で図面台帳側のデータ自体は正規化済みの想定だが、
  //   万一未実行のデータが残っていても検索側で吸収できるよう、双方を同じ
  //   ルールで正規化してから比較する（日本語氏名に対しては実質無害）。
  const fMainNo    = String(filters.mainNo    || '').trim().toLowerCase();
  const fApplicant = normalizeRomanName_(filters.applicant);
  const fMaterial  = String(filters.material  || '').trim().toLowerCase();
  const fReviewer  = normalizeRomanName_(filters.reviewer);
  const fDateFrom  = String(filters.dateFrom  || '').trim();
  const fDateTo    = String(filters.dateTo    || '').trim();
  const fSizes     = Array.isArray(filters.drawingSizes) ? filters.drawingSizes : [];

  const results = [];
  rows.forEach((row, i) => {
    // 特徴属性の要約（COL_DB.ATTRIBUTES_SUMMARY）もキーワード検索対象に含める。
    // OCR時にAIが生成した形状・特徴の要約文（例："SUS304製のL字ブラケット。
    // 長穴2箇所、C2面取り指定。"）で、図名や機械名には出てこない形状特徴の
    // キーワード（穴の数・材質形状・加工内容など）で図面を探せるようにするため。
    const searchable = [
      row[COL_DB.FULL_NO    - 1], row[COL_DB.MAIN_NO   - 1], row[COL_DB.NAME_JP    - 1],
      row[COL_DB.NAME_EN    - 1], row[COL_DB.UNIT_NAME - 1], row[COL_DB.MODEL_NAME - 1],
      row[COL_DB.MATERIAL   - 1], row[COL_DB.APPLICANT - 1], row[COL_DB.STATUS     - 1],
      row[COL_DB.ATTRIBUTES_SUMMARY - 1],
    ].join(' ').toLowerCase();

    if (!tokens.every(t => searchable.indexOf(t) !== -1)) return;

    const mainNo    = String(row[COL_DB.MAIN_NO    - 1]);
    const applicant = String(row[COL_DB.APPLICANT  - 1]);
    const material  = String(row[COL_DB.MATERIAL   - 1]);
    const reviewer  = String(row[COL_DB.REVIEWER   - 1]);
    const status    = String(row[COL_DB.STATUS     - 1]);
    const size      = String(row[COL_DB.DRAWING_SIZE - 1]).trim().toUpperCase();
    const appDate   = normalizeDateForCompare_(row[COL_DB.APP_DATE - 1]);

    if (fMainNo    && mainNo.toLowerCase().indexOf(fMainNo) === -1)             return;
    if (fApplicant && normalizeRomanName_(applicant).indexOf(fApplicant) === -1) return;
    if (fMaterial  && material.toLowerCase().indexOf(fMaterial) === -1)         return;
    if (fReviewer  && normalizeRomanName_(reviewer).indexOf(fReviewer) === -1)   return;
    if (fSizes.length > 0 && fSizes.indexOf(size) === -1)                       return;
    if (!isWithinDateRange_(appDate, fDateFrom, fDateTo))                       return;

    // FILE_URL は1列のみ（承認前＝申請中図面／承認後＝承認済み図面のリンクに
    // approveDrawing() 側で自動的に上書きされる。ラベルはステータスで出し分ける。
    // getDrawingDetail() と同じ判定ロジック（STATUS.APPROVED）を使用）
    const fileUrl = extractUrlFromCellFormula_(formulas[i][COL_DB.FILE_URL - 1]) ||
                    String(row[COL_DB.FILE_URL - 1]).trim();

    results.push({
      fullNo:        String(row[COL_DB.FULL_NO    - 1]),
      mainNo:        mainNo,
      nameJp:        String(row[COL_DB.NAME_JP    - 1]),
      modelName:     String(row[COL_DB.MODEL_NAME - 1]),
      status:        status,
      applicant:     applicant,
      reviewer:      reviewer,
      fileUrl:       fileUrl,
      fileLinkLabel: status === STATUS.APPROVED ? '承認済みPDFを開く' : '図面ファイルを開く',
    });
  });

  results.sort((a, b) => a.fullNo.localeCompare(b.fullNo));
  return results;
}

/**
 * 機械台帳をキーワード検索する（絞り込みフィルタ対応）
 * ※ 機械台帳には図面サイズ・材質・検図者の概念が存在しないため、
 *   それらのフィルタは無視する（主図番・申請者・期間のみ適用）。
 *   期間フィルタは「着手日」に対して適用する（図面検索での「申請日」に相当）。
 * @param {Object} filters - searchDrawingsAndMainNumbers() の filters をそのまま渡す
 * @returns {Array<{mainNo, modelNameJp, modelNameEn, applicant, date}>}
 */
function searchMainNumbers_(ss, tokens, filters) {
  filters = filters || {};
  const sheet = ss.getSheetByName(SHEET_NAMES.MAIN_NUMBER);
  if (!sheet) return [];
  const lastRow = sheet.getLastRow();
  if (lastRow < GLOBAL_ROW.DATA_START) return [];

  const range    = sheet.getRange(GLOBAL_ROW.DATA_START, 1, lastRow - GLOBAL_ROW.DATA_START + 1, COL_MAIN.EXPLORER_PATH);
  const rows     = range.getValues();
  const formulas = range.getFormulas(); // FOLDER_URL列（HYPERLINK数式）から実URLを取り出すために使用

  const fMainNo    = String(filters.mainNo    || '').trim().toLowerCase();
  const fApplicant = normalizeRomanName_(filters.applicant);
  const fDateFrom  = String(filters.dateFrom  || '').trim();
  const fDateTo    = String(filters.dateTo    || '').trim();

  const results = [];
  rows.forEach((row, i) => {
    const searchable = [
      row[COL_MAIN.MAIN_NO       - 1], row[COL_MAIN.MODEL_NAME_JP - 1],
      row[COL_MAIN.MODEL_NAME_EN - 1], row[COL_MAIN.APPLICANT     - 1],
      row[COL_MAIN.NOTE          - 1],
    ].join(' ').toLowerCase();

    if (!tokens.every(t => searchable.indexOf(t) !== -1)) return;

    const mainNo    = String(row[COL_MAIN.MAIN_NO   - 1]);
    const applicant = String(row[COL_MAIN.APPLICANT - 1]);
    const startDate = normalizeDateForCompare_(row[COL_MAIN.DATE - 1]);

    if (fMainNo    && mainNo.toLowerCase().indexOf(fMainNo) === -1)             return;
    if (fApplicant && normalizeRomanName_(applicant).indexOf(fApplicant) === -1) return;
    if (!isWithinDateRange_(startDate, fDateFrom, fDateTo))                     return;

    // 承認済み図面フォルダのリンク（主図番発行直後はまだ空欄のことがある）
    const folderUrl = extractUrlFromCellFormula_(formulas[i][COL_MAIN.FOLDER_URL - 1]) ||
                       String(row[COL_MAIN.FOLDER_URL - 1]).trim();

    results.push({
      mainNo:      mainNo,
      modelNameJp: String(row[COL_MAIN.MODEL_NAME_JP - 1]),
      modelNameEn: String(row[COL_MAIN.MODEL_NAME_EN - 1]),
      applicant:   applicant,
      date:        formatDateOnly_(row[COL_MAIN.DATE - 1]),
      folderUrl:   folderUrl,
    });
  });

  results.sort((a, b) => a.mainNo.localeCompare(b.mainNo));
  return results;
}

// ============================================================
// クライアントから呼ばれる：検索フィルタ（申請者・検図者）の氏名候補一覧
// ============================================================
/**
 * 検索フィルタの「申請者」「検図者」欄（<input list="..."> のdatalist）で使う
 * 氏名候補一覧を返す。
 *
 * 【候補に含める範囲】
 *   - ユーザーマスタに登録済みの氏名
 *   - 図面索引（最新リビジョンのみ）に実際に記録されている申請者・検図者名
 *     …過去図面インポート機能（35_LegacyDrawingImport.gs）はユーザーマスタとの
 *       照合を行わずOCR読み取りのまま図面台帳へ記録するため、ユーザーマスタに
 *       存在しない氏名（表記ゆれ・退職者・OCR誤読含む）が図面台帳側にだけ
 *       存在するケースがある。そうした氏名も検索フィルタとして使いたいという
 *       要望に対応するため、両方のソースを合わせて候補にする。
 *   あくまで「候補」であり、この一覧にない文字列も自由入力でき、検索側は
 *   部分一致で判定するため、候補にない氏名を入力しても問題なく機能する。
 *
 * @returns {string[]} 重複除去・五十音順ソート済みの氏名配列
 */
function getNameCandidatesForSearchFilter() {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const names = new Set();

  // ユーザーマスタ登録者
  // normalizeRomanName_ は日本語氏名（漢字）には実質影響しないため、
  // ソースを問わず一律で通しておく（ローマ字氏名の表記ゆれを候補上で集約するため）
  const userSheet = ss.getSheetByName(SHEET_NAMES.USER_MASTER);
  if (userSheet) {
    const lastRow = userSheet.getLastRow();
    if (lastRow >= GLOBAL_ROW.DATA_START) {
      userSheet
        .getRange(GLOBAL_ROW.DATA_START, COL_USER.NAME_JP, lastRow - GLOBAL_ROW.DATA_START + 1, 1)
        .getValues()
        .flat()
        .forEach(v => { const n = normalizeRomanName_(v); if (n) names.add(n); });
    }
  }

  // 図面索引（最新リビジョンのみ）に記録されている申請者・検図者
  // （ユーザーマスタに存在しない過去図面の氏名もここから拾える）
  const indexSheet = ensureIndexSheet_(SHEET_NAMES.DRAWING_INDEX); // 05_SearchIndex.gs
  const lastRow = indexSheet.getLastRow();
  if (lastRow >= GLOBAL_ROW.DATA_START) {
    const rows = indexSheet
      .getRange(GLOBAL_ROW.DATA_START, 1, lastRow - GLOBAL_ROW.DATA_START + 1, DRAWING_DB_COL_COUNT)
      .getValues();
    rows.forEach(row => {
      const applicant = normalizeRomanName_(row[COL_DB.APPLICANT - 1]);
      const reviewer  = normalizeRomanName_(row[COL_DB.REVIEWER  - 1]);
      if (applicant) names.add(applicant);
      if (reviewer)  names.add(reviewer);
    });
  }

  return Array.from(names).sort((a, b) => a.localeCompare(b, 'ja'));
}


//
// 【役割】
//   ブラウザからアップロードされた1枚のPDF（未登録の図面でもよい）を
//   その場でAI解析し、図面台帳に蓄積されている「特徴属性」（COL_DB.ATTRIBUTES。
//   OCR読取時に 21_OcrService.gs の extractDrawingInfoByOcr_ が抽出したもの）
//   と比較して、形状・特徴が近い過去図面をランキング表示する。
//
// 【設計方針】
//   - 検索のたびに全図面をAIで再解析することはしない（コスト・時間が
//     図面台帳の件数に比例して増えてしまうため）。AIを使うのはアップロード
//     された「検索したい1枚」の解析のみで、既存図面との比較は図面台帳に
//     保存済みの属性データを使ったローカル計算（ルールベースの重み付け
//     スコアリング）で完結させる。
//   - 属性データは各主図番・子図番ごとに複数リビジョンが存在し得るため、
//     REV_MARKが最大（＝最新）の行のみを比較対象とする（同一部品の
//     旧リビジョンが重複して結果に並ぶのを防ぐため）。
//   - 特徴属性が未登録の行（本機能の導入前に登録された図面、および
//     現時点で属性抽出未対応の過去図面インポートによる行）は、比較対象
//     から自然に除外される（属性JSONが空のため）。
//

/**
 * アップロードされたPDF（base64）を解析し、図面台帳に蓄積された特徴属性と
 * 比較して類似度の高い図面をランキングして返す
 * @param {string} base64Pdf - base64エンコードされたPDFデータ
 * @returns {{success, message?, query?, totalCandidates?, results?}}
 */
function searchSimilarDrawings(base64Pdf) {
  try {
    if (!base64Pdf) return { success: false, message: 'PDFデータが空です。' };

    // extractDrawingInfoFromBase64_ は 21_OcrService.gs で定義済み
    // （extractDrawingInfoByOcr_ と共通の抽出ロジックを、Driveファイルを
    //   介さずbase64データから直接呼び出せるようにしたもの）
    const extracted = extractDrawingInfoFromBase64_(base64Pdf);
    if (!extracted) {
      return {
        success: false,
        message: 'アップロードされたPDFの解析に失敗しました。別のファイルでお試しいただくか、しばらく待って再度お試しください。',
      };
    }
    if (!extracted.attributes || !extracted.attributes.shapeCategory) {
      return {
        success: false,
        message: 'アップロードされたPDFから形状の特徴を十分に読み取れませんでした。図面として認識できる鮮明なPDFかご確認ください。',
      };
    }

    const query = {
      material:   extracted.material,
      attributes: extracted.attributes,
    };

    const ss = SpreadsheetApp.getActiveSpreadsheet();

    const queryPreview = {
      shapeCategory: query.attributes.shapeCategory || '',
      material:      query.material || '',
      summary:       query.attributes.summary || '',
    };

    // 図面台帳を直接スキャンせず、「最新リビジョンのみ」を保持する図面索引
    // （SHEET_NAMES.DRAWING_INDEX）を比較対象にする。索引側で既に
    // 主図番＋子図番ごとの最新化が済んでいるため、ここでのグルーピング
    // 処理は不要になる（05_SearchIndex.gs を参照）。
    const indexSheet = ensureIndexSheet_(SHEET_NAMES.DRAWING_INDEX);
    const lastRow = indexSheet.getLastRow();
    if (lastRow < GLOBAL_ROW.DATA_START) {
      return { success: true, query: queryPreview, totalCandidates: 0, results: [] };
    }

    const range = indexSheet.getRange(GLOBAL_ROW.DATA_START, 1, lastRow - GLOBAL_ROW.DATA_START + 1, DRAWING_DB_COL_COUNT);
    const rows     = range.getValues();
    const formulas = range.getFormulas(); // FILE_URL列（HYPERLINK数式）から実URLを取り出すために使用

    // 特徴属性（JSON）が保存されている行のみを比較対象とする
    const candidates = rows
      .map((row, index) => {
        const attributesJson = String(row[COL_DB.ATTRIBUTES - 1] || '');
        let attributes = null;
        try { attributes = attributesJson ? JSON.parse(attributesJson) : null; } catch (e) { attributes = null; }
        const fileUrl = extractUrlFromCellFormula_(formulas[index][COL_DB.FILE_URL - 1]) ||
                        String(row[COL_DB.FILE_URL - 1]).trim();
        return {
          fullNo:    String(row[COL_DB.FULL_NO    - 1]),
          mainNo:    String(row[COL_DB.MAIN_NO    - 1]),
          nameJp:    String(row[COL_DB.NAME_JP    - 1]),
          modelName: String(row[COL_DB.MODEL_NAME - 1]),
          material:  String(row[COL_DB.MATERIAL   - 1]),
          status:    String(row[COL_DB.STATUS     - 1]),
          summary:   String(row[COL_DB.ATTRIBUTES_SUMMARY - 1] || ''),
          fileUrl:   fileUrl,
          attributes,
        };
      })
      .filter(c => c.attributes && c.attributes.shapeCategory);

    // ルールベースの類似度スコアを算出し、降順に並べる（computeSimilarityScore_
    // は本ファイル末尾で定義。0点＝共通点なしの図面は結果から除外する）
    const scored = candidates
      .map(c => ({ ...c, score: computeSimilarityScore_(query, c) }))
      .filter(c => c.score > 0)
      .sort((a, b) => b.score - a.score)
      .slice(0, 10);

    return {
      success:         true,
      query:           queryPreview,
      totalCandidates: candidates.length,
      results:         scored,
    };

  } catch (e) {
    console.error('searchSimilarDrawings error:', e);
    return { success: false, message: `エラーが発生しました：${e.message}` };
  }
}

/**
 * 2つの図面の特徴属性を比較し、類似度スコア（0〜100点）を算出する
 * （ルールベースの重み付け加点方式。断定できない項目は加点しない
 *   ＝データが不足している場合は無理にスコアを底上げしない設計）
 *
 * 【配点内訳（合計100点）】
 *   形状分類の一致：20点／材質の一致：8点／外形寸法の近さ：最大15点
 *   穴の個数の近さ：最大8点／ねじ加工の有無の一致：4点／ねじ径の重なり：4点
 *   曲げ加工の有無の一致：4点／面取りの有無の一致：4点
 *   幾何公差の有無の一致：4点／幾何公差の種類の重なり：4点
 *   表面処理の有無の一致：3点／表面処理の種類の類似度：最大3点
 *   熱処理の有無の一致：3点／熱処理の種類の類似度：最大3点
 *   要約文の類似度：最大13点
 *
 * 【要約文・種類テキストの類似度について】
 *   summary（OCR時にAIが生成した形状・特徴の要約文、例："SUS304製のL字
 *   ブラケット。長穴2箇所、C2面取り指定。"）や、表面処理・熱処理の具体的な
 *   種類（例："三価黒色クロメートメッキ"）は自由記述の日本語文であり、
 *   スペース区切りの単語分割ができない（形態素解析エンジンをこのプロジェクトの
 *   実行環境（Apps Script）に導入するのは重い）。そのため、文字2-gram
 *   （隣接2文字の組）の集合同士のDice係数（2×共通要素数 ÷ 両集合の要素数の和）
 *   という簡易的な手法でテキスト類似度を近似する（textSimilarityDice_ を参照）。
 *   形態素解析ほどの精度はないが、共通する単語・言い回しが多い文字列同士は
 *   自然と高いスコアになるため、構造化属性だけでは拾いきれない類似性の
 *   補助的な手がかりとして十分に機能する。
 *
 * @param {{material: string, attributes: Object}} query
 * @param {{material: string, attributes: Object}} candidate
 * @returns {number} 0〜100の類似度スコア
 */
function computeSimilarityScore_(query, candidate) {
  const qa = query.attributes     || {};
  const ca = candidate.attributes || {};
  const qf = qa.features || {};
  const cf = ca.features || {};

  let score = 0;

  // 形状分類の一致（20点）
  if (qa.shapeCategory && ca.shapeCategory && qa.shapeCategory === ca.shapeCategory) {
    score += 20;
  }

  // 材質の一致（8点、大文字小文字・前後空白を無視して比較）
  if (query.material && candidate.material &&
      String(query.material).trim().toUpperCase() === String(candidate.material).trim().toUpperCase()) {
    score += 8;
  }

  // 外形寸法の近さ（最大15点）
  score += 15 * dimensionSimilarity_(qa.overallDimensions, ca.overallDimensions);

  // 穴の個数の近さ（最大8点）
  score += 8 * countSimilarity_(qf.holeCount, cf.holeCount);

  // 加工特徴の一致（各4点）
  if (typeof qf.hasThreads === 'boolean' && qf.hasThreads === cf.hasThreads) score += 4;
  if (typeof qf.hasBending === 'boolean' && qf.hasBending === cf.hasBending) score += 4;
  if (typeof qf.hasChamfer === 'boolean' && qf.hasChamfer === cf.hasChamfer) score += 4;

  // ねじ径の重なり（4点。どちらか一方でも共通のねじ径があれば加点）
  if (threadSizesOverlap_(qf.threadSizes, cf.threadSizes)) score += 4;

  // 幾何公差の有無の一致（4点）
  if (typeof qf.hasGeometricTolerance === 'boolean' &&
      qf.hasGeometricTolerance === cf.hasGeometricTolerance) {
    score += 4;
  }
  // 幾何公差の種類の重なり（4点。どちらか一方でも共通の種類があれば加点）
  if (arraysHaveOverlap_(qf.geometricToleranceTypes, cf.geometricToleranceTypes)) score += 4;

  // 表面処理の有無の一致（3点）
  if (typeof qf.hasSurfaceTreatment === 'boolean' &&
      qf.hasSurfaceTreatment === cf.hasSurfaceTreatment) {
    score += 3;
  }
  // 表面処理の種類の類似度（最大3点。自由記述のためテキスト類似度で近似）
  score += 3 * textSimilarityDice_(qf.surfaceTreatmentType, cf.surfaceTreatmentType);

  // 熱処理の有無の一致（3点）
  if (typeof qf.hasHeatTreatment === 'boolean' &&
      qf.hasHeatTreatment === cf.hasHeatTreatment) {
    score += 3;
  }
  // 熱処理の種類の類似度（最大3点。自由記述のためテキスト類似度で近似）
  score += 3 * textSimilarityDice_(qf.heatTreatmentType, cf.heatTreatmentType);

  // 要約文の類似度（最大13点。文字2-gramのDice係数による簡易テキスト類似度）
  score += 13 * textSimilarityDice_(qa.summary, ca.summary);

  return Math.round(score);
}

/**
 * 2つの文字列配列に共通する要素が1つでもあるかを判定する（大文字小文字の
 * 違いは無視しない＝幾何公差の種類名など、プロンプトで提示した用語一覧からの
 * 選択を期待する項目向け。ねじ径のような英数字表記の揺れを吸収したい場合は
 * threadSizesOverlap_ を使うこと）
 * @param {string[]} a
 * @param {string[]} b
 * @returns {boolean}
 */
function arraysHaveOverlap_(a, b) {
  if (!Array.isArray(a) || !Array.isArray(b) || a.length === 0 || b.length === 0) return false;
  const setA = new Set(a.map(s => String(s).trim()));
  return b.some(s => setA.has(String(s).trim()));
}

/**
 * 2つの文字列の類似度を、文字2-gram（隣接2文字の組）集合のDice係数で近似する。
 * 日本語の形態素解析を使わない簡易手法（本ファイル冒頭 computeSimilarityScore_
 * のコメントを参照）。
 * @param {string} a
 * @param {string} b
 * @returns {number} 0〜1（どちらかが2文字未満、または空の場合は0）
 */
function textSimilarityDice_(a, b) {
  const setA = bigramSet_(a);
  const setB = bigramSet_(b);
  if (setA.size === 0 || setB.size === 0) return 0;

  let intersectionCount = 0;
  setA.forEach(gram => { if (setB.has(gram)) intersectionCount++; });

  return (2 * intersectionCount) / (setA.size + setB.size);
}

/**
 * 文字列から2-gram（隣接2文字）の集合を作る（空白は除去してから分割する）
 * @param {string} str
 * @returns {Set<string>}
 */
function bigramSet_(str) {
  const s = String(str || '').replace(/\s+/g, '');
  const grams = new Set();
  for (let i = 0; i < s.length - 1; i++) {
    grams.add(s.substring(i, i + 2));
  }
  return grams;
}

/**
 * 2つの外形寸法（{length,width,height}）オブジェクトの近さを 0〜1 で返す
 * （両方に値がある軸のみを比較し、値がまったく比較できない場合は0を返す。
 *   1軸あたりの近さは 1 - |差| / max(値) で算出し、各軸の平均を取る）
 */
function dimensionSimilarity_(a, b) {
  if (!a || !b) return 0;
  const keys = ['length', 'width', 'height'];
  const ratios = keys
    .map(k => [a[k], b[k]])
    .filter(([x, y]) => typeof x === 'number' && typeof y === 'number' && (x > 0 || y > 0))
    .map(([x, y]) => {
      const maxV = Math.max(x, y, 0.001);
      return Math.max(0, 1 - Math.abs(x - y) / maxV);
    });
  if (ratios.length === 0) return 0;
  return ratios.reduce((sum, v) => sum + v, 0) / ratios.length;
}

/**
 * 2つの個数（穴の数など）の近さを 0〜1 で返す（値が欠けている場合は0）
 */
function countSimilarity_(a, b) {
  if (typeof a !== 'number' || typeof b !== 'number') return 0;
  if (a === 0 && b === 0) return 1;
  const maxV = Math.max(a, b, 1);
  return Math.max(0, 1 - Math.abs(a - b) / maxV);
}

/**
 * 2つのねじ径配列（例：["M8","M12"]）に共通するものがあるかを判定する
 */
function threadSizesOverlap_(a, b) {
  if (!Array.isArray(a) || !Array.isArray(b) || a.length === 0 || b.length === 0) return false;
  const setA = new Set(a.map(s => String(s).trim().toUpperCase()));
  return b.some(s => setA.has(String(s).trim().toUpperCase()));
}

// ============================================================
// クライアントから呼ばれる：図面（フル図番）の詳細を取得
// ============================================================
/**
 * @param {string} fullNo - フル図番（10桁）
 * @returns {{success, message?, detail?}}
 */
function getDrawingDetail(fullNo) {
  try {
    const ss      = SpreadsheetApp.getActiveSpreadsheet();
    const dbSheet = ss.getSheetByName(SHEET_NAMES.DRAWING_DB);
    if (!dbSheet) return { success: false, message: `"${SHEET_NAMES.DRAWING_DB}" シートが見つかりません。` };

    // findRow_ は 23_ApprovalWebApp.gs で定義済み（同一プロジェクト内でグローバルに利用可能）
    const row = findRow_(dbSheet, fullNo);
    if (!row) return { success: false, message: `図面番号「${fullNo}」が見つかりません。` };

    const d      = row.data;
    const mainNo = String(d[COL_DB.MAIN_NO - 1]);
    const subNo  = String(d[COL_DB.SUB_NO  - 1]);
    const status = String(d[COL_DB.STATUS  - 1]);
    const tz     = ss.getSpreadsheetTimeZone();

    // FILE_URL は1列のみ（承認前＝申請中図面／承認後＝承認済み図面のリンクに
    // approveDrawing() 側で自動的に上書きされる。ラベルはステータスで出し分ける）
    // STATUS 定数は 00_Config.gs で定義済み（同一プロジェクト内でグローバルに利用可能）
    // 過去図面（STATUS.LEGACY）は正規の承認フローを経ていないが、表示上は
    // 承認済みと同様に扱う（PDFはすでに「承認済み図面」フォルダへ保存されているため）
    const fileLinkLabel = (status === STATUS.APPROVED || status === STATUS.LEGACY)
      ? '承認済みPDFを開く' : '図面ファイルを開く';

    const detail = {
      fullNo:          String(d[COL_DB.FULL_NO      - 1]),
      mainNo:          mainNo,
      subNo:           subNo,
      revMark:         String(d[COL_DB.REV_MARK     - 1]),
      nameJp:          String(d[COL_DB.NAME_JP      - 1]),
      nameEn:          String(d[COL_DB.NAME_EN      - 1]),
      unitName:        String(d[COL_DB.UNIT_NAME    - 1]),
      modelName:       String(d[COL_DB.MODEL_NAME   - 1]),
      material:        String(d[COL_DB.MATERIAL     - 1]),
      scale:           String(d[COL_DB.SCALE        - 1]),
      drawingSize:     String(d[COL_DB.DRAWING_SIZE - 1]),
      status:          status,
      applicant:       String(d[COL_DB.APPLICANT    - 1]),
      appDate:         formatDateOnly_(d[COL_DB.APP_DATE - 1]),
      reviewer:        String(d[COL_DB.REVIEWER      - 1]),
      reviewerDate:    formatSheetDateTime_(d[COL_DB.REVIEWER_DATE - 1], tz),
      manager:         String(d[COL_DB.MANAGER       - 1]),
      managerDate:     formatSheetDateTime_(d[COL_DB.MANAGER_DATE  - 1], tz),
      director:        String(d[COL_DB.DIRECTOR      - 1]),
      directorDate:    formatSheetDateTime_(d[COL_DB.DIRECTOR_DATE - 1], tz),
      fileUrl:         row.fileUrl,
      fileLinkLabel:   fileLinkLabel,
      revisionHistory: getRevisionHistory_(dbSheet, mainNo, subNo, fullNo),
    };

    return { success: true, detail };

  } catch (e) {
    console.error('getDrawingDetail error:', e);
    return { success: false, message: `エラーが発生しました：${e.message}` };
  }
}

/**
 * 同一主図番・同一子図番の改訂履歴を、改訂記号（REV_MARK）順に取得する。
 * 改訂連番（REV_SEQ）列は存在しないため、REV_MARK の文字コード順
 * （'-'＝初版 → 'A' → 'B' … の順に大きくなる）を代わりに用いる。
 * 同一主図番・子図番内で REV_MARK が最大の行を「最新」として動的に判定する。
 * @returns {Array<{fullNo, revMark, status, appDate, isLatest, isCurrent}>}
 */
function getRevisionHistory_(dbSheet, mainNo, subNo, currentFullNo) {
  const lastRow = dbSheet.getLastRow();
  if (lastRow < GLOBAL_ROW.DATA_START) return [];

  const rows = dbSheet
    .getRange(GLOBAL_ROW.DATA_START, 1, lastRow - GLOBAL_ROW.DATA_START + 1, DRAWING_DB_COL_COUNT)
    .getValues();

  const matched = rows.filter(row =>
    String(row[COL_DB.MAIN_NO - 1]) === mainNo &&
    String(row[COL_DB.SUB_NO  - 1]) === subNo
  );

  // REV_MARK の文字コード順にソート（'-' < 'A' < 'B' … なので初版が自然に先頭に来る）
  matched.sort((a, b) =>
    String(a[COL_DB.REV_MARK - 1]).localeCompare(String(b[COL_DB.REV_MARK - 1]))
  );

  const latestMark = matched.length > 0
    ? String(matched[matched.length - 1][COL_DB.REV_MARK - 1])
    : null;

  return matched.map(row => {
    const revMark = String(row[COL_DB.REV_MARK - 1]);
    return {
      fullNo:    String(row[COL_DB.FULL_NO - 1]),
      revMark:   revMark,
      status:    String(row[COL_DB.STATUS  - 1]),
      appDate:   formatDateOnly_(row[COL_DB.APP_DATE - 1]),
      isLatest:  revMark === latestMark,
      isCurrent: String(row[COL_DB.FULL_NO - 1]) === currentFullNo,
    };
  });
}

// ============================================================
// クライアントから呼ばれる：主図番（機械）の詳細を取得
// ============================================================
/**
 * @param {string} mainNo - 主図番（6桁）
 * @returns {{success, message?, detail?}}
 */
function getMainNumberDetail(mainNo) {
  try {
    const ss    = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(SHEET_NAMES.MAIN_NUMBER);
    if (!sheet) return { success: false, message: `"${SHEET_NAMES.MAIN_NUMBER}" シートが見つかりません。` };

    const info = getMainNumberInfoFull_(sheet, mainNo);
    if (!info) return { success: false, message: `主図番「${mainNo}」が見つかりません。` };

    // 保存済みのパスに頼らず、表示のたびにDrive APIで最新のパスへ計算し直す
    // （CADフォルダが後から移動されても表示がズレないようにするため）
    info.cadFolderPath = resolveLiveCadFolderPath_(info);

    const dbSheet = ss.getSheetByName(SHEET_NAMES.DRAWING_DB);
    info.linkedDrawings = getLinkedDrawings_(dbSheet, mainNo);

    return { success: true, detail: info };

  } catch (e) {
    console.error('getMainNumberDetail error:', e);
    return { success: false, message: `エラーが発生しました：${e.message}` };
  }
}

/**
 * 機械台帳の部品リスト（BOM）URLを更新する（図面検索の詳細パネルから呼ばれる）
 * 主図番発行時の自動作成に失敗した場合の手動リンクや、差し替えに使用する。
 * @param {string} mainNo - 主図番
 * @param {string} newUrl - 新しい部品リストのURL（GoogleスプレッドシートのURL）
 * @returns {{success, message, bomUrl?}}
 */
function updateBomLink(mainNo, newUrl) {
  try {
    const url = String(newUrl || '').trim();
    if (!url) return { success: false, message: 'URLを入力してください。' };

    const fileId = extractDriveFileId_(url);
    if (!fileId) {
      return { success: false, message: 'URLを認識できませんでした。部品リストのスプレッドシートを開いた状態のアドレスバーのURLを貼り付けてください。' };
    }

    let file;
    try {
      file = DriveApp.getFileById(fileId);
    } catch (e) {
      return { success: false, message: '指定されたファイルが見つかりませんでした。URLを確認してください。' };
    }
    if (file.getMimeType() !== MimeType.GOOGLE_SHEETS) {
      return { success: false, message: 'Googleスプレッドシート以外のファイルは部品リストとして登録できません。' };
    }

    const ss    = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(SHEET_NAMES.MAIN_NUMBER);
    if (!sheet) return { success: false, message: `"${SHEET_NAMES.MAIN_NUMBER}" シートが見つかりません。` };

    const info = getMainNumberInfoFull_(sheet, mainNo);
    if (!info) return { success: false, message: `主図番「${mainNo}」が見つかりません。` };

    sheet.getRange(info.rowIndex, COL_MAIN.BOM_URL)
      .setFormula(buildFileLinkFormula_(file.getUrl(), '📊部品リストを開く'));

    return { success: true, message: '部品リストのリンクを更新しました。', bomUrl: file.getUrl() };

  } catch (e) {
    console.error('updateBomLink error:', e);
    return { success: false, message: `エラーが発生しました：${e.message}` };
  }
}

/**
 * 機械台帳のCADフォルダリンクを更新する（図面検索の詳細パネルから呼ばれる）
 * @param {string} mainNo - 主図番
 * @param {string} newUrl - 新しいCADフォルダのURL
 * @returns {{success, message, cadFolderUrl?, cadFolderPath?}}
 */
function updateCadFolderLink(mainNo, newUrl) {
  try {
    const url = String(newUrl || '').trim();
    if (!url) return { success: false, message: 'URLを入力してください。' };

    const fileId = extractDriveFileId_(url);
    if (!fileId) {
      return {
        success: false,
        message:
          'URLを認識できませんでした。エクスプローラーで対象フォルダを右クリック→' +
          '「Google Drive」→「リンクをクリップボードにコピー」で取得したリンクを貼り付けてください。',
      };
    }

    let folder;
    try {
      folder = DriveApp.getFolderById(fileId);
    } catch (e) {
      return { success: false, message: '指定されたフォルダが見つかりませんでした。URLを確認してください。' };
    }

    const ss    = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(SHEET_NAMES.MAIN_NUMBER);
    if (!sheet) return { success: false, message: `"${SHEET_NAMES.MAIN_NUMBER}" シートが見つかりません。` };

    const info = getMainNumberInfoFull_(sheet, mainNo);
    if (!info) return { success: false, message: `主図番「${mainNo}」が見つかりません。` };

    let cadFolderPath = '';
    try {
      cadFolderPath = buildExplorerPath_(folder);
    } catch (e) {
      console.warn(`CADフォルダのエクスプローラーパス組み立てに失敗しました（${mainNo}）: ${e.message}`);
    }

    sheet.getRange(info.rowIndex, COL_MAIN.CAD_FOLDER_URL)
      .setFormula(buildFileLinkFormula_(folder.getUrl(), '📁CADフォルダを開く'));
    sheet.getRange(info.rowIndex, COL_MAIN.CAD_FOLDER_PATH).setValue(cadFolderPath);

    return {
      success:       true,
      message:       'CADフォルダのリンクを更新しました。',
      cadFolderUrl:  folder.getUrl(),
      cadFolderPath: cadFolderPath,
    };

  } catch (e) {
    console.error('updateCadFolderLink error:', e);
    return { success: false, message: `エラーが発生しました：${e.message}` };
  }
}

/**
 * 機械台帳の1行分の情報をフルで取得する
 * （25_PdfApprovalStamp.gs の getMainNumberInfo_ は機械名のみを返す簡易版のため、
 *   検索ビュー用に全項目を返す別関数として用意する）
 */
function getMainNumberInfoFull_(sheet, mainNo) {
  const lastRow = sheet.getLastRow();
  if (lastRow < GLOBAL_ROW.DATA_START) return null;

  const range    = sheet.getRange(GLOBAL_ROW.DATA_START, 1, lastRow - GLOBAL_ROW.DATA_START + 1, COL_MAIN.DESCRIPTION);
  const rows     = range.getValues();
  const formulas = range.getFormulas();

  for (let i = 0; i < rows.length; i++) {
    if (String(rows[i][COL_MAIN.MAIN_NO - 1]).trim() !== mainNo) continue;

    const folderUrl = extractUrlFromCellFormula_(formulas[i][COL_MAIN.FOLDER_URL - 1]) ||
                       String(rows[i][COL_MAIN.FOLDER_URL - 1]).trim();
    const cadFolderUrl = extractUrlFromCellFormula_(formulas[i][COL_MAIN.CAD_FOLDER_URL - 1]) ||
                          String(rows[i][COL_MAIN.CAD_FOLDER_URL - 1]).trim();
    const bomUrl = extractUrlFromCellFormula_(formulas[i][COL_MAIN.BOM_URL - 1]) ||
                    String(rows[i][COL_MAIN.BOM_URL - 1]).trim();

    return {
      rowIndex:     GLOBAL_ROW.DATA_START + i,
      mainNo:       mainNo,
      deptCode:     String(rows[i][COL_MAIN.DEPT_CODE     - 1]),
      modelNameJp:  String(rows[i][COL_MAIN.MODEL_NAME_JP - 1]),
      modelNameEn:  String(rows[i][COL_MAIN.MODEL_NAME_EN - 1]),
      applicant:    String(rows[i][COL_MAIN.APPLICANT     - 1]),
      date:         formatDateOnly_(rows[i][COL_MAIN.DATE - 1]),
      note:         String(rows[i][COL_MAIN.NOTE          - 1]),
      folderUrl:    folderUrl,
      explorerPath: String(rows[i][COL_MAIN.EXPLORER_PATH - 1]),
      cadFolderUrl:  cadFolderUrl,
      cadFolderPath: String(rows[i][COL_MAIN.CAD_FOLDER_PATH - 1]), // フォールバック値（発行時点のもの）
      bomUrl:        bomUrl,
      description:   String(rows[i][COL_MAIN.DESCRIPTION - 1]),
    };
  }
  return null;
}

/**
 * CADフォルダのエクスプローラーパスを、保存済みの値に頼らずDrive APIで都度再計算する。
 * フォルダが後からGoogleドライブ上で移動されても、表示のたびに最新のパスへ追従させるため
 * （25_PdfApprovalStamp.gs の buildExplorerPath_ / extractDriveFileId_ を流用）。
 * @param {{cadFolderUrl: string, cadFolderPath: string}} info - getMainNumberInfoFull_() の戻り値
 * @returns {string} 最新のエクスプローラーパス文字列（取得できなければフォールバック値、それも無ければ空文字）
 */
function resolveLiveCadFolderPath_(info) {
  if (!info || !info.cadFolderUrl) return '';
  const fileId = extractDriveFileId_(info.cadFolderUrl);
  if (!fileId) return info.cadFolderPath || '';
  try {
    const folder = DriveApp.getFolderById(fileId);
    return buildExplorerPath_(folder);
  } catch (e) {
    console.warn(`CADフォルダパスの再計算に失敗しました（${info.mainNo}）: ${e.message}`);
    return info.cadFolderPath || '（フォルダを取得できませんでした）';
  }
}

/**
 * 指定した主図番に紐づく図面台帳の全行（全改訂・全子図番）を取得する
 * @returns {Array<{fullNo, nameJp, status}>}
 */
function getLinkedDrawings_(dbSheet, mainNo) {
  if (!dbSheet) return [];
  const lastRow = dbSheet.getLastRow();
  if (lastRow < GLOBAL_ROW.DATA_START) return [];

  const rows = dbSheet
    .getRange(GLOBAL_ROW.DATA_START, 1, lastRow - GLOBAL_ROW.DATA_START + 1, DRAWING_DB_COL_COUNT)
    .getValues();

  return rows
    .filter(row => String(row[COL_DB.MAIN_NO - 1]) === mainNo)
    .map(row => ({
      fullNo: String(row[COL_DB.FULL_NO - 1]),
      nameJp: String(row[COL_DB.NAME_JP - 1]),
      status: String(row[COL_DB.STATUS  - 1]),
    }))
    .sort((a, b) => a.fullNo.localeCompare(b.fullNo));
}

// ============================================================
// プライベートヘルパー
// ============================================================

/**
 * 日付セルを 'yyyy-MM-dd' 文字列に整形する（時刻は不要なため formatSheetDateTime_ とは別に用意）
 * @param {Date|string|*} value
 * @returns {string}
 */
function formatDateOnly_(value) {
  if (!value) return '';
  if (value instanceof Date) {
    const tz = SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone();
    return Utilities.formatDate(value, tz, 'yyyy-MM-dd');
  }
  return String(value);
}