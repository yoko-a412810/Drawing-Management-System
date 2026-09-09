/**
 * ============================================================
 * 図面承認・採番・出図管理システム
 * ファイル: 00_Config.gs  ―  システム共通設定・定数管理
 * ============================================================
 */

// ─── ① シート名の定義 ───
const SHEET_NAMES = {
  DEPT_MASTER:   '部署コードマスタ',
  MAIN_NUMBER:   '機械台帳',
  USER_MASTER:   'ユーザーマスタ',
  INPUT_DRAWING: '図面登録申請（入力）',
  DRAWING_DB:    '図面台帳',
  LEGACY_INPUT:  '過去図面登録（入力）',
  DRAWING_INDEX: '図面索引',     // 図面台帳の「最新リビジョンのみ」を保持する検索専用索引（05_SearchIndex.gs）
  PENDING_INDEX: '承認待ち索引', // 図面台帳のうち「承認進行中の行のみ」を保持する索引（05_SearchIndex.gs）
};

// ─── ② 共通の行ルール ───
const GLOBAL_ROW = {
  HEADER:     4, // ヘッダは4行目
  DATA_START: 5  // データ開始は5行目
};

// ─── ③ 「機械台帳」の列インデックス（1始まり） ───
// ※ 2026-07 マイグレーション：「組立図名/PJ名」列を「機械名（日本語）」に位置づけ直し、
//    直後に「機械名（英語・任意）」列を追加。末尾に「承認済み図面フォルダリンク」
//    （ブラウザ用URL）・「エクスプローラーパス」（ローカルパス文字列）列を追加。
//    （承認完了時、主図番＋機械名でフォルダを作成・PDFを保存し、両方をここに記録する）
const COL_MAIN = {
  MAIN_NO:        1,  // A: 主図番
  DEPT_CODE:      2,  // B: 部署コード
  SEQ:            3,  // C: 連番
  MODEL_NAME_JP:  4,  // D: 機械名（日本語）
  MODEL_NAME_EN:  5,  // E: 機械名（英語・任意）
  APPLICANT:      6,  // F: 申請者
  DATE:           7,  // G: 着手日
  NOTE:           8,  // H: 備考
  FOLDER_URL:     9,  // I: 承認済み図面フォルダリンク（ブラウザ用URL、承認完了時に自動記入）
  EXPLORER_PATH: 10,  // J: エクスプローラーパス（ローカルパス文字列、承認完了時に自動記入）
  CAD_FOLDER_URL:  11, // K: CADデータフォルダリンク（主図番発行時にユーザーが入力）
  CAD_FOLDER_PATH: 12, // L: CADデータフォルダのエクスプローラーパス（発行時に自動生成。表示のたびに再計算するための初期値／フォールバック用）
  BOM_URL:         13, // M: 部品リスト（BOM）URL。Googleスプレッドシート形式。主図番発行時にテンプレートから自動複製・自動記入
  DESCRIPTION:     14  // N: 説明文（機械の大まかな説明。図面承認申請時に22_ApprovalDialog.htmlで入力・更新）
};

// ─── ③b 「ユーザーマスタ」の列インデックス（1始まり） ───
// ※ 2026-07 マイグレーションで「NAME」（英語表記氏名）列を追加。
//    PDF承認スタンプ機能で、日本語の氏名から英語表記を引くために使用する。
const COL_USER = {
  EMAIL:         1, // A: メールアドレス
  NAME_JP:       2, // B: 氏名（日本語）
  NAME_EN:       3, // C: NAME（英語表記。例："K.Yokoyama"）
  REGISTERED_AT: 4  // D: 登録日
};

// ─── ④ 「図面登録申請（入力）」シートの列インデックス（1始まり） ───
// ※ 2026-07 マイグレーション③：図面サイズ列（H）を追加
// ※ 2026-07 マイグレーション④：検図者・課長・部長列（I〜K）を削除
//    （承認者はダイアログで選択する方式のため、シートへの直接入力欄は不要だった）
// ※ 2026-07 マイグレーション⑤：AIチェック（要確認点）列（K）を追加。
//    OCR読取と同じタイミングでAIによる簡易チェックを行い、申請者自身が
//    検図に回す前にケアレスミスを見直せるようにするため。既存の「登録結果」
//    列はL列へ1列後ろへずれるため、実際のスプレッドシート側でもK列の前に
//    新規列を挿入すること（21_OcrService.gs / 20_DrawingApproveRequest.gs 参照）。
// ※ 2026-07 マイグレーション⑥：特徴属性（JSON）列（M）を末尾に追加。
//    OCR読取時に部品の形状・特徴（shapeCategory/overallDimensions/features/
//    summary）もあわせて抽出し（21_OcrService.gs の extractDrawingInfoByOcr_
//    を参照）、登録申請時にそのまま図面台帳（COL_DB.ATTRIBUTES）へ引き継ぐ。
//    将来の「類似図面検索」機能（30_DrawingSearch.gs、未実装）で、図面台帳に
//    蓄積された属性同士を比較して似た過去図面を探すために使用する予定の
//    データであり、申請者が内容を確認・編集する対象ではないため、AIチェック
//    列とは異なりセル表示・ハイライト等は行わない（JSON文字列をそのまま
//    保持するだけの受け渡し用の列）。
const COL_INPUT = {
  DRAWING_NO:  1,  // A: 図面番号（DWG.NO.）10桁フル図番
  NAME_JP:     2,  // B: 図名（JPN）
  NAME_EN:     3,  // C: 英名（NAME）
  UNIT_NAME:   4,  // D: ユニット名（UNIT）
  MODEL_NAME:  5,  // E: 機械名（MODEL）
  MATERIAL:    6,  // F: 材質（MATL.）
  SCALE:       7,  // G: 縮尺（SCALE）
  DRAWING_SIZE:8,  // H: 図面サイズ（A0〜A4）
  FILE_URL:    9,  // I: 図面ファイルURL（Drive）
  OCR_RESULT:  10, // J: AI-OCR読取結果
  AI_CHECK:    11, // K: AIチェック（要確認点。詳細はセルのノートに記載）
  REG_RESULT:  12, // L: 登録結果
  ATTRIBUTES:  13, // M: 特徴属性（JSON。類似図面検索用の下ごしらえデータ）
};

// ─── ⑤ 「図面台帳」の列インデックス（1始まり） ───
// ※ 2026-07 マイグレーション①：S列（差戻し理由）・T列（差戻し回数）を物理削除
//    （差戻し時は行ごと削除する運用のため、これらの列は不要だった）
// ※ 2026-07 マイグレーション②：申請者列を申請日列の直前に移動
// ※ 2026-07 マイグレーション③：図面サイズ列（K）を縮尺列の直後に追加
//    （承認完了時にPDFへ検図日・承認日をスタンプする機能で、書き込み座標の
//      判定に使用する）
// ※ 2026-07 マイグレーション④：改訂連番・最新フラグ列（未使用の補助列）を削除。
//    承認済みPDFファイルURL列を廃止し、図面ファイルURL列（U列＝「図面リンク」）に統合。
//    承認完了前は申請中図面フォルダのリンク、部長承認完了時にPDFスタンプ処理が
//    承認済み図面フォルダのリンクへ上書きする（申請中図面フォルダのファイルは
//    定期的に削除される運用のため、リンクを2列で持つ必要がなくなったため）
// ※ 2026-07 マイグレーション⑤：AIチェック（検図者への参考情報）列（W）を末尾に追加。
//    入力シート（COL_INPUT.AI_CHECK）でのOCR時AIチェック結果のうち、申請者が
//    「検図者に情報として残しておきたい」と判断したものが、申請登録時に
//    そのままこの列へ引き継がれる（21_OcrService.gs の checkDrawingConcerns_、
//    20_DrawingApproveRequest.gs の setApproversAndRegister を参照）。
//    あくまで検図者への参考情報であり、正式な検図結果や承認可否の判定には
//    使用しない。
// ※ 2026-07 マイグレーション⑥：特徴属性列（X・Y）を末尾に追加。
//    入力シート（COL_INPUT.ATTRIBUTES）でOCR時に抽出した部品の形状・特徴
//    JSONを、申請登録時にそのまま引き継ぐ（20_DrawingApproveRequest.gs の
//    setApproversAndRegister を参照）。ATTRIBUTES_SUMMARY は、そのJSON内の
//    summaryフィールドのみを人間が一覧で確認しやすいよう複製した列
//    （シートを直接眺めたときにJSONをパースせずに内容を把握できるようにする
//    目的のみで、機能的な必須データではない）。
//    類似図面検索機能（30_DrawingSearch.gs）は未実装で、現時点ではこの列は
//    データの蓄積のみを行う（検索機能は今後追加予定）。
const COL_DB = {
  FULL_NO:       1,  // A: フル図番（10桁）
  MAIN_NO:       2,  // B: 主図番
  SUB_NO:        3,  // C: 子図番
  REV_MARK:      4,  // D: 改訂記号
  NAME_JP:       5,  // E: 図名（JPN）
  NAME_EN:       6,  // F: 英名（NAME）
  UNIT_NAME:     7,  // G: ユニット名（UNIT）
  MODEL_NAME:    8,  // H: 機械名（MODEL）
  MATERIAL:      9,  // I: 材質（MATL.）
  SCALE:         10, // J: 縮尺（SCALE）
  DRAWING_SIZE:  11, // K: 図面サイズ（A0〜A4）
  STATUS:        12, // L: ステータス
  APPLICANT:     13, // M: 申請者
  APP_DATE:      14, // N: 申請日
  REVIEWER:      15, // O: 検図者
  REVIEWER_DATE: 16, // P: 検図者承認日時
  MANAGER:       17, // Q: 課長
  MANAGER_DATE:  18, // R: 課長承認日時
  DIRECTOR:      19, // S: 部長
  DIRECTOR_DATE: 20, // T: 部長承認日時
  FILE_URL:      21, // U: 図面リンク（承認前＝申請中図面／承認後＝承認済み図面）
  BATCH_ID:      22, // V: 申請バッチID（1回の一括登録・申請ごとに発行される識別子）
  AI_CHECK:      23, // W: AIチェック（検図者への参考情報。改行区切りの箇条書きテキスト）
  ATTRIBUTES:         24, // X: 特徴属性（JSON。類似図面検索用）
  ATTRIBUTES_SUMMARY: 25, // Y: 特徴属性の要約（JSON内のsummaryを複製した閲覧用テキスト）
};

const DRAWING_DB_COL_COUNT    = 25; // 図面台帳の総列数
const INPUT_DRAWING_COL_COUNT = 13; // 入力シートの総列数

// ─── ⑤b 「過去図面登録申請（入力）」シートの列インデックス（1始まり） ───
// ※ ④ 過去データのインポート機能専用。通常の入力シート（COL_INPUT）とは別に、
//    通常ルートでは読み取らない「設計者・検図者・承認者」の氏名・日付を追加で持つ。
//    移行期間限定の機能のため、スクリプトプロパティ LEGACY_IMPORT_DEADLINE で
//    利用可否を制御する（詳細は 40_LegacyDrawingImport.gs を参照）。
// ※ 2026-07 マイグレーション：特徴属性（JSON）列（R）を末尾に追加。
//    通常フローの図面登録申請（入力）シート（COL_INPUT.ATTRIBUTES）と同様、
//    OCR読取時に部品の形状・特徴も併せて抽出し、登録申請時に図面台帳
//    （COL_DB.ATTRIBUTES）へ引き継ぐ（35_LegacyDrawingImport.gs を参照）。
//    過去図面は登録数が多くなりがちな一括インポート用途のため、通常フロー
//    以上に類似図面検索のコーパスとして価値が高い（過去の類似設計の
//    有無を確認する用途に直結するため）。
const COL_LEGACY_INPUT = {
  DRAWING_NO:   1,  // A: 図面番号（DWG.NO.）10桁フル図番
  NAME_JP:      2,  // B: 図名（JPN）
  NAME_EN:      3,  // C: 英名（NAME）
  UNIT_NAME:    4,  // D: ユニット名（UNIT）
  MODEL_NAME:   5,  // E: 機械名（MODEL）
  MATERIAL:     6,  // F: 材質（MATL.）
  SCALE:        7,  // G: 縮尺（SCALE）
  DRAWING_SIZE: 8,  // H: 図面サイズ（A0〜A4）
  FILE_URL:     9,  // I: 図面ファイルURL（Drive、読み取り元の元PDF）
  DESIGNER:     10, // J: 設計者（表題欄DESIGNED欄の氏名。OCR読み取りのまま、ユーザーマスタとの照合はしない）
  DESIGN_DATE:  11, // K: 設計日
  REVIEWER:     12, // L: 検図者（表題欄CHECKED欄の氏名）
  REVIEW_DATE:  13, // M: 検図日
  APPROVER:     14, // N: 承認者（表題欄APPROVED欄の氏名）
  APPROVE_DATE: 15, // O: 承認日
  OCR_RESULT:   16, // P: AI-OCR読取結果
  REG_RESULT:   17, // Q: 登録結果
  ATTRIBUTES:   18, // R: 特徴属性（JSON。類似図面検索用の下ごしらえデータ）
};
const LEGACY_INPUT_COL_COUNT = 18; // 過去図面入力シートの総列数

// ─── ⑤c ステータス値 ───
// ※ 2026-07 移動：STATUS定数は元々 23_ApprovalWebApp.gs で定義していたが、
//    05_SearchIndex.gs 等、より読み込み順の早いファイルからも参照する
//    必要が生じたため、最も早く読み込まれる本ファイル（00_Config.gs）へ
//    移動した。Apps Scriptはファイルをアルファベット順（＝番号順）に
//    読み込み、トップレベルのconst宣言はその読み込み時点で評価されるため、
//    後から読まれるファイルで定義した定数を、先に読まれるファイルの
//    トップレベルで参照すると「is not defined」エラーになる。共通定数は
//    常に最も早く読まれるこのファイルに置くこと。
const STATUS_LEGACY = '過去図面'; // 過去データ移行専用。承認Webアプリの対象に一切含まれない特別なステータス

const STATUS = {
  REVIEWING:        '検図中',
  WAITING_MANAGER:  '課長承認待ち',
  WAITING_DIRECTOR: '部長承認待ち',
  APPROVED:         '承認済',
  LEGACY:           STATUS_LEGACY, // '過去図面'
};

// ─── ⑥ 図面サイズの許容値（OCR結果の検証・PDFスタンプ座標のキーに使用） ───
const VALID_DRAWING_SIZES = ['A0', 'A1', 'A2', 'A3', 'A4'];

/**
 * WebアプリURLに batchId をクエリパラメータとして付与する共通ヘルパー
 * @param {string} url - ベースとなる /exec URL
 * @param {string} batchId - 申請バッチID（空ならそのまま返す）
 * @returns {string}
 */
function appendBatchIdToUrl_(url, batchId) {
  if (!url || !batchId) return url;
  const sep = url.indexOf('?') === -1 ? '?' : '&';
  return `${url}${sep}batchId=${encodeURIComponent(batchId)}`;
}

/**
 * シートから読み込んだ日時セルの値を 'yyyy-MM-dd HH:mm:ss' 文字列に変換する
 *
 * 【背景】
 *   'yyyy-MM-dd HH:mm:ss' 形式の文字列を setValue() でセルに書き込むと、
 *   スプレッドシートが自動的に「日付時刻型」に変換してしまうことがある。
 *   その状態のセルを getValues() で読み込むと、文字列ではなく
 *   JavaScriptのDateオブジェクトが返る。これをそのまま String() すると
 *   "Wed Jul 10 2026 09:15:23 GMT+0900 (Japan Standard Time)" のような
 *   JS標準フォーマットになってしまい、日付として扱えない
 *   （例：先頭が曜日の "Wed" になる等の不具合につながる）。
 *
 *   この関数は、値がDateオブジェクトであれば正しい文字列に整形し、
 *   すでに文字列であればそのまま返すことで、書き込み時の型に関わらず
 *   一貫した文字列を得られるようにする。
 *
 * 【タイムゾーンについての重要な注意】
 *   Sheetsの日付セルをApps Scriptで読み込むと、内部的には
 *   「スプレッドシート自体のタイムゾーン設定」（ファイル→設定→全般）を
 *   基準にDateオブジェクトが構築される。これは appsscript.json の
 *   timeZone（＝Session.getScriptTimeZone()）とは別の設定であり、
 *   両者が食い違っていると、Session.getScriptTimeZone() で再フォーマット
 *   した際に時刻がズレ、日付が前後にずれることがある（例：日本時間の
 *   夕方の時刻が、スプレッドシートのタイムゾーンとのズレにより翌日の
 *   未明として再フォーマットされてしまう）。
 *   そのため、この関数は呼び出し元から明示的にタイムゾーンを受け取り、
 *   スプレッドシート自体のタイムゾーン（ss.getSpreadsheetTimeZone()）を
 *   渡してもらうことを前提にしている。
 *
 * @param {Date|string|*} value - シートセルから取得した値
 * @param {string} [timeZone] - 変換に使うタイムゾーン（省略時はSession.getScriptTimeZone()）。
 *                              シートから読み込んだDateオブジェクトを扱う場合は、
 *                              呼び出し元で ss.getSpreadsheetTimeZone() を明示的に渡すこと。
 * @returns {string} 'yyyy-MM-dd HH:mm:ss' 形式の文字列（空値なら空文字）
 */
function formatSheetDateTime_(value, timeZone) {
  if (!value) return '';
  if (value instanceof Date) {
    const tz = timeZone || Session.getScriptTimeZone();
    return Utilities.formatDate(value, tz, 'yyyy-MM-dd HH:mm:ss');
  }
  return String(value);
}

/**
 * URLを、セルに表示用テキスト付きリンクとして書き込むための HYPERLINK 数式文字列を作る
 * （例：=HYPERLINK("https://drive.google.com/...","📄図面を開く")）
 * @param {string} url   - リンク先URL
 * @param {string} label - セルに表示するテキスト
 * @returns {string} 数式文字列（urlが空なら空文字を返す＝リンクなし）
 */
function buildFileLinkFormula_(url, label) {
  if (!url) return '';
  const escapedUrl   = String(url).replace(/"/g, '""');
  const escapedLabel = String(label).replace(/"/g, '""');
  return `=HYPERLINK("${escapedUrl}","${escapedLabel}")`;
}

/**
 * セルの数式（またはプレーンな値）から、実際のURLを取り出す。
 * HYPERLINK数式なら中のURLを抽出し、数式でなければそのまま返す
 * （移行期に残る可能性のある「URLがそのまま入っている」古いセルにも対応するため）。
 * @param {string} formulaOrValue - Range.getFormula() または getValue() で取得した内容
 * @returns {string} URL文字列（空なら空文字）
 */
function extractUrlFromCellFormula_(formulaOrValue) {
  if (!formulaOrValue) return '';
  const str = String(formulaOrValue).trim();
  const match = str.match(/^=HYPERLINK\("([^"]+)"/i);
  if (match) return match[1];
  return str;
}

/**
 * ローマ字氏名表記のゆれを正規化する（大文字小文字・"."後のスペース有無など）。
 * 例：'K.Yokoyama' / 'K.YOKOYAMA' / 'K. Yokoyama' → 'K.YOKOYAMA'
 *
 * 【背景】
 *   通常フローの申請者・検図者・部長は氏名（日本語）をプルダウンから選択するが、
 *   過去図面インポート（35_LegacyDrawingImport.gs）はOCRで読み取った表題欄の
 *   英語表記氏名（DESIGNED/CHECKED/APPROVED欄）をそのままユーザーマスタとの
 *   照合なしに図面台帳へ記録するため、大文字小文字や"."後のスペース有無など
 *   表記ゆれが生じやすい。検索フィルタ（30_DrawingSearch.gs）で同一人物を
 *   一つの表記として扱えるよう、この関数で統一する。
 *   全角文字・日本語氏名（漢字）に対しては大文字化・スペース除去のいずれも
 *   実質的に影響しないため、通常フローの氏名にそのまま適用しても問題ない。
 *
 * @param {string} name
 * @returns {string} 大文字・スペースなしに正規化した文字列（空値なら空文字）
 */
function normalizeRomanName_(name) {
  if (!name) return '';
  return String(name).toUpperCase().replace(/\s+/g, '');
}