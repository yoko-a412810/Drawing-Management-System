/**
 * ============================================================
 * 図面承認・採番・出図管理システム
 * ファイル: 21_OcrService.gs  ―  Gemini API OCR モジュール
 * ============================================================
 *
 * 【抽出項目】
 *   - drawingNo   : 図面番号（DWG.NO.）10桁フル図番
 *   - drawingName : 図名（JPN）
 *   - nameEn      : 英名（NAME）
 *   - unitName    : ユニット名（UNIT）
 *   - modelName   : 機械名（MODEL）
 *   - material    : 材質（MATL.）
 *   - scale       : 縮尺（SCALE）
 *
 * 【スクリプトプロパティ（事前設定が必要）】
 *   GEMINI_API_KEY : Google AI Studio で取得した API キー
 *
 * 【スクリプトプロパティの設定手順】
 *   1. Apps Script エディタを開く
 *   2. 左メニュー「プロジェクトの設定」（⚙️）をクリック
 *   3.「スクリプトプロパティ」セクションで「プロパティを追加」
 *   4. プロパティ名: GEMINI_API_KEY、値: 取得したAPIキーを入力して保存
 */

// ============================================================
// 定数
// ============================================================
const GEMINI_MODEL   = 'gemini-3.1-flash-lite';
const GEMINI_API_URL =
  `https://generativelanguage.googleapis.com/v1beta/models/${GEMINI_MODEL}:generateContent`;

// Gemini API へ送るプロンプト
// ※ 2026-07 表題欄の項目抽出に加えて、部品の形状・特徴に関する属性
//    （attributes）も同一コールで抽出するよう拡張した。将来の「類似図面検索」
//    機能（30_DrawingSearch.gs）で、図面台帳に蓄積された属性同士を比較して
//    似た過去図面を探すために使用する。表題欄の構造化項目とは異なり厳密な
//    一致を要求しないやや曖昧な抽出のため、JSONパース失敗等のリスクを避ける
//    べく既存の表題欄項目とスキーマを分けているが、精度への悪影響を避ける
//    ため独立呼び出しにはせず、あえて同一コール・同一温度（0）で処理する
//    （表題欄抽出・属性抽出はどちらも「構造化データの抽出」という同種の
//    タスクであり、AIチェック［checkDrawingConcerns_］のような自由記述の
//    所見出しとは性質が異なるため、混在させても抽出精度への影響は小さいと
//    判断した）。
const OCR_PROMPT = `
この画像は機械図面です。図面の表題欄（タイトルブロック）を読み取り、
以下の項目を JSON 形式で返してください。

抽出する項目（括弧内は表題欄上の表示ラベル）：
- drawingNo   : 図面番号（DWG.NO.）英大文字2桁＋数字4桁＋数字3桁＋1桁の計10桁。例：JB2601001-
- drawingName : 図名（JPN）日本語の図面名称
- nameEn      : 英名（NAME）英語の図面名称
- unitName    : ユニット名（UNIT）
- modelName   : 機械名（MODEL）
- material    : 材質（MATL.）例：SUS304、S45C
- scale       : 縮尺（SCALE）例：1:2、1:5
- drawingSize : 図面サイズ（用紙サイズ）。表題欄付近に記載されている「A0」「A1」「A2」
                「A3」「A4」のいずれか。必ずこの5種類のうちの1つの表記（英大文字+数字1桁）
                のみを返してください。

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
読み取れなかった表題欄項目は空文字（""）としてください。
例：{
  "drawingNo":"JB2601001-",
  "drawingName":"ブラケット",
  "nameEn":"BRACKET",
  "unitName":"フレームユニット",
  "modelName":"○○装置",
  "material":"SUS304",
  "scale":"1:2",
  "drawingSize":"A3",
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

// ============================================================
// メイン関数
// ============================================================

/**
 * PDF ファイルを Gemini API で OCR し、表題欄情報・部品属性を返す
 * （Drive上のファイルを扱う通常のOCRフロー用の薄いラッパー。
 *   実体は extractDrawingInfoFromBase64_ で、類似図面検索機能
 *   （30_DrawingSearch.gs）でアップロードされたPDF（base64）を直接
 *   処理する場合にも共通で使用する）
 *
 * @param {GoogleAppsScript.Drive.File} file - Drive上のPDFファイル
 * @returns {{drawingNo, drawingName, nameEn, unitName, modelName,
 *            material, scale, drawingSize, attributes, hasWarning}|null}
 *   抽出成功: オブジェクト（読取不可項目がある場合 hasWarning=true）
 *   抽出失敗: null
 *   attributes: {shapeCategory, overallDimensions, features, summary}
 *     （類似図面検索機能で使用する部品の形状・特徴データ。断定できない
 *       項目はnullのまま保持する。抽出自体に失敗した場合はnull）
 */
function extractDrawingInfoByOcr_(file) {
  const base64Pdf = Utilities.base64Encode(file.getBlob().getBytes());
  return extractDrawingInfoFromBase64_(base64Pdf);
}

/**
 * base64エンコードされたPDFデータを直接 Gemini API へ送り、表題欄情報・
 * 部品属性を抽出する（extractDrawingInfoByOcr_ の実体）。
 * Driveファイルを介さずに呼び出せるため、類似図面検索機能
 * （30_DrawingSearch.gs の searchSimilarDrawings）で、ブラウザから
 * アップロードされたPDFをその場で解析する用途にも使用する。
 *
 * @param {string} base64Pdf - base64エンコードされたPDFのバイナリデータ
 * @returns {{drawingNo, drawingName, nameEn, unitName, modelName,
 *            material, scale, drawingSize, attributes, hasWarning}|null}
 */
function extractDrawingInfoFromBase64_(base64Pdf) {
  const apiKey = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');
  if (!apiKey) {
    throw new Error(
      'スクリプトプロパティ「GEMINI_API_KEY」が設定されていません。\n' +
      'Apps Script エディタ →「プロジェクトの設定」→「スクリプトプロパティ」から設定してください。'
    );
  }

  const requestBody = {
    contents: [{
      parts: [
        { text: OCR_PROMPT },
        {
          inline_data: {
            mime_type: 'application/pdf',
            data: base64Pdf,
          }
        }
      ]
    }],
    generationConfig: {
      temperature:     0,     // 再現性を高めるため温度は 0
      maxOutputTokens: 1024,  // 属性（幾何公差・表面処理・熱処理を追加）抽出分の余裕を追加
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
    const errBody = response.getContentText();
    console.error(`Gemini API エラー (${statusCode}): ${errBody}`);
    throw new Error(
      `Gemini API がステータス ${statusCode} を返しました。APIキーとモデル名を確認してください。`
    );
  }

  const responseJson = JSON.parse(response.getContentText());

  // レスポンスから text 部分を取得
  const text = responseJson?.candidates?.[0]?.content?.parts?.[0]?.text;
  if (!text) {
    console.warn('Gemini API から有効なテキストが返りませんでした。', JSON.stringify(responseJson));
    return null;
  }

  // JSON パース（コードブロックが残っていれば除去）
  let extracted;
  try {
    const cleaned = text.replace(/```json|```/g, '').trim();
    extracted = JSON.parse(cleaned);
  } catch (e) {
    console.error('JSON パース失敗:', text, e);
    return null;
  }

  const result = {
    drawingNo:   String(extracted.drawingNo   || '').trim(),
    drawingName: String(extracted.drawingName || '').trim(),
    nameEn:      String(extracted.nameEn      || '').trim(),
    unitName:    String(extracted.unitName    || '').trim(),
    modelName:   String(extracted.modelName   || '').trim(),
    material:    String(extracted.material    || '').trim(),
    scale:       String(extracted.scale       || '').trim(),
    drawingSize: String(extracted.drawingSize || '').trim().toUpperCase(),
  };

  // drawingSize は A0〜A4 以外の値が返ってきた場合、無効値として空にする
  // （PDFへの承認日スタンプ機能で座標テーブルのキーとして使うため、
  //   想定外の値のまま入力シートへ転記すると後工程でエラーになる）
  if (result.drawingSize && VALID_DRAWING_SIZES.indexOf(result.drawingSize) === -1) {
    console.warn(`OCRが想定外の図面サイズを返しました: "${result.drawingSize}"。空欄として扱います。`);
    result.drawingSize = '';
  }

  // 必須項目（図面番号・図名）が読み取れなかった場合は警告フラグ
  result.hasWarning = !result.drawingNo || !result.drawingName || !result.drawingSize;

  // 部品の形状・特徴属性（類似図面検索用）。あくまで付随データのため、
  // 想定外の形式で返ってきても正規化して極力活かす（欠けている項目は
  // nullのまま許容し、hasWarning等には影響させない＝この抽出の成否が
  // 表題欄OCRの成否判定を左右しないようにする）
  result.attributes = normalizeDrawingAttributes_(extracted.attributes);

  return result;
}

/**
 * OCRレスポンスの attributes フィールドを正規化する
 * （Geminiの出力が想定外の型・欠損だった場合でも、極力壊れずに使える形にする）
 * @param {*} raw - extracted.attributes（未検証の生データ）
 * @returns {{shapeCategory: string|null, overallDimensions: Object|null,
 *            features: Object, summary: string}}
 */
function normalizeDrawingAttributes_(raw) {
  const a = (raw && typeof raw === 'object') ? raw : {};

  const validShapeCategories = ['板金', '軸物', 'ブロック・切削', 'ブラケット', 'その他'];
  const shapeCategory = validShapeCategories.indexOf(a.shapeCategory) !== -1 ? a.shapeCategory : null;

  const dims = (a.overallDimensions && typeof a.overallDimensions === 'object') ? a.overallDimensions : {};
  const toNumOrNull = (v) => (typeof v === 'number' && isFinite(v)) ? v : null;
  const overallDimensions = {
    length: toNumOrNull(dims.length),
    width:  toNumOrNull(dims.width),
    height: toNumOrNull(dims.height),
  };

  // 表面処理・熱処理の種類は自由記述文字列のため、空文字は null に丸める
  // （"該当なし"を空文字と null のどちらでも表現し得るGemini側の揺れを吸収する）
  const toStrOrNull = (v) => { const s = String(v || '').trim(); return s ? s : null; };

  const feat = (a.features && typeof a.features === 'object') ? a.features : {};
  const features = {
    holeCount:   toNumOrNull(feat.holeCount),
    hasThreads:  feat.hasThreads === true,
    threadSizes: Array.isArray(feat.threadSizes) ? feat.threadSizes.map(String) : [],
    hasBending:  feat.hasBending === true,
    hasChamfer:  feat.hasChamfer === true,
    // 幾何公差：種類はプロンプトで提示した用語一覧からの選択を期待するが、
    // Gemini側の表記ゆれに備えて厳密なバリデーションは行わず、文字列配列として
    // 素通しする（想定外の用語が混ざっても検索・表示上の実害は小さいため）
    hasGeometricTolerance:   feat.hasGeometricTolerance === true,
    geometricToleranceTypes: Array.isArray(feat.geometricToleranceTypes)
      ? feat.geometricToleranceTypes.map(String).map(s => s.trim()).filter(s => s)
      : [],
    hasSurfaceTreatment:  feat.hasSurfaceTreatment === true,
    surfaceTreatmentType: toStrOrNull(feat.surfaceTreatmentType),
    hasHeatTreatment:  feat.hasHeatTreatment === true,
    heatTreatmentType: toStrOrNull(feat.heatTreatmentType),
  };

  return {
    shapeCategory,
    overallDimensions,
    features,
    summary: String(a.summary || '').trim(),
  };
}

/**
 * OCR で読み取った図面番号（10桁）をパースして各部品に分解する
 *
 * @param {string} drawingNo - 例："JB2601001-"
 * @returns {{mainNo, subNo, revMark, isValid, errorMsg}}
 */
function parseDrawingNo_(drawingNo) {
  const no = String(drawingNo).trim();

  // 形式チェック：英大文字2桁 + 数字4桁 + 数字3桁 + （"-" or A-Z）= 10桁
  if (!/^[A-Z]{2}\d{4}\d{3}[-A-Z]$/.test(no)) {
    return {
      mainNo: '', subNo: '', revMark: '',
      isValid: false,
      errorMsg: `図面番号「${no}」の形式が正しくありません。` +
                `英大文字2桁＋数字4桁＋数字3桁＋改訂記号1桁（"-"またはA〜Z）の10桁で入力してください。`,
    };
  }

  return {
    mainNo:  no.slice(0, 6),   // 例: JB2601
    subNo:   no.slice(6, 9),   // 例: 001
    revMark: no.slice(9, 10),  // 例: -
    isValid: true,
    errorMsg: '',
  };
}

// ============================================================
// AIによる簡易検図チェック（① 図面一括読取（OCR）と同時に実行）
// ============================================================
//
// 【位置づけ】
//   これは正式な検図の代替ではない。あくまで、申請者（設計者）自身が
//   検図者へ回す前にケアレスミス（記入漏れ・更新漏れ等）へ気づけるように
//   するための「一次チェック」であり、断定的な合否判定は行わない
//   （気になる点を列挙するのみ）。強度・公差・材質の工学的妥当性判断など、
//   設計判断を要する領域は対象外とする（AIによる誤った太鼓判が、かえって
//   過信につながるリスクの方が大きいため）。
//   そのため、このチェックの結果は登録可否の判定には一切使用しない
//   （＝AIチェックで何か検出されても、申請自体はブロックしない）。
//   OCR抽出（extractDrawingInfoByOcr_）とは目的の異なる別のAPI呼び出しと
//   して分離している（構造化データの抽出と、自由記述の所見出しでは
//   求められるプロンプトの性質が異なり、同一呼び出しに混ぜるとOCR側の
//   抽出精度に悪影響が出る可能性があるため）。
//
//   【チェック項目を絞り込んだ経緯】
//   当初は「気になる点の例」を緩く提示するだけの設計だったが、以下の
//   理由から、確認してほしい項目を具体的に3点へ絞り込む形に変更した：
//     - 日本語注記とその英訳は社内規定のテンプレートで語句が固定されて
//       おり、翻訳の妥当性チェックは不要（申請者に確認済み）
//     - 材質欄と図面内容（形状）の整合性判断は、工学的な設計判断に
//       踏み込んでしまい、AIチェックの本来の目的（ケアレスミス予防）を
//       超えるため対象外とした
//     - 項目を絞った方が、該当項目を実際に確認する挙動になりやすく、
//       抽象的な例示だけを与えるより指摘の密度が上がる

const AI_CHECK_PROMPT = `
この画像は申請前の機械図面です。検図者に提出する前に、申請者（設計者）
自身が見直した方が良さそうな「気になる点」があれば、簡潔に列挙してください。

【確認してほしい観点】
1. 表記誤り・誤字（数字と英字の混同、ねじ呼び記号の誤記など）
2. 図面枠・表題欄の明確な誤記・記載不備（※下記【確認不要】に該当する社内仕様を除く）
3. 加工性・構造上の不自然さ（エンジニアリング視点でのセルフチェック）
- 丸棒曲面への直接穴あけ指示（座ぐり指示の有無）
- 段付き軸におけるコーナーRや逃げ溝指示の有無
- ねじ加工の呼び・深さ指示の不自然さ
- 嵌合（かんごう）が予想される部位への公差指示の抜け

【確認不要・対象外とする項目】
- APPROVED、CHECKED、DESIGNED 各欄の空欄（申請前段階のため未記入で正常です）
- MODEL REVISION 欄の空欄（改訂時のみ記入する仕様のため未記入で正常です）
- DWG. NO.（図面番号）末尾のハイフン「-」（初版図面は末尾がハイフンで終わる仕様です）
- 日本語注記とその英訳の対応関係や英文の内容（社内規定テンプレートのためチェック不要）
- 高度な強度計算、材料選定の妥当性、相手部品との組付け干渉（図面単体で判断不能なため）
- MATL.M.（材料質量）、MATL.SIZE（材料サイズ）欄の値

【注意事項】
- あくまで「念のため確認をお勧めする点」の提示に留め、断定的な誤りの指摘や合格・不合格の判定はしないでください。
- 気になる点が特に見当たらない場合は、空配列を返してください。
- 出力は最大5件程度とし、確信の低い些細な点まで無理に挙げる必要はありません。

返答はJSON形式のみとし、マークダウンのコードブロック(\`\`\`)は含めないで
ください。形式：
{"points": ["気になる点の説明1", "気になる点の説明2"]}
気になる点がなければ {"points": []} を返してください。
`;

/**
 * PDF ファイルに対してAIによる簡易検図チェックを行い、気になる点の一覧を返す
 * （extractDrawingInfoByOcr_ と同じ GEMINI_API_KEY / GEMINI_MODEL を流用するが、
 *   別APIコールとして独立させている）
 *
 * @param {GoogleAppsScript.Drive.File} file - Drive上のPDFファイル
 * @returns {{points: string[]}|null}
 *   チェック成功: { points: string[] }（気になる点がなければ空配列）
 *   チェック失敗: null（呼び出し元はOCR結果への転記自体は継続すること）
 */
function checkDrawingConcerns_(file) {
  const apiKey = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');
  if (!apiKey) {
    console.warn('スクリプトプロパティ「GEMINI_API_KEY」が未設定のため、AIチェックをスキップしました。');
    return null;
  }

  const pdfBlob   = file.getBlob();
  const base64Pdf = Utilities.base64Encode(pdfBlob.getBytes());

  const requestBody = {
    contents: [{
      parts: [
        { text: AI_CHECK_PROMPT },
        { inline_data: { mime_type: 'application/pdf', data: base64Pdf } },
      ]
    }],
    generationConfig: {
      temperature:     0.2, // 所見出しのため、抽出（temperature:0）よりわずかに緩める
      maxOutputTokens: 512,
    }
  };

  // 一時的なAPIエラー（レート制限・タイムアウト等）対策として、失敗時に
  // 1回だけ間隔を空けて再試行する（OCR抽出＋AIチェックの2回分のAPI呼び出しが
  // 短時間に連続するため、瞬間的なレート制限に触れることがある）
  const MAX_ATTEMPTS = 2;

  for (let attempt = 1; attempt <= MAX_ATTEMPTS; attempt++) {
    try {
      const response = UrlFetchApp.fetch(`${GEMINI_API_URL}?key=${apiKey}`, {
        method:             'post',
        contentType:        'application/json',
        payload:            JSON.stringify(requestBody),
        muteHttpExceptions: true,
      });

      const statusCode = response.getResponseCode();
      if (statusCode !== 200) {
        console.warn(
          `Gemini API エラー（AIチェック, 試行${attempt}/${MAX_ATTEMPTS}, ` +
          `${statusCode}）: ${file.getName()} / ${response.getContentText()}`
        );
        if (attempt < MAX_ATTEMPTS) { Utilities.sleep(2000); continue; }
        return null;
      }

      const responseJson = JSON.parse(response.getContentText());
      const text = responseJson?.candidates?.[0]?.content?.parts?.[0]?.text;
      if (!text) {
        console.warn(
          `Gemini API から有効なAIチェック結果が返りませんでした（試行${attempt}/${MAX_ATTEMPTS}）: ` +
          `${file.getName()}`, JSON.stringify(responseJson)
        );
        if (attempt < MAX_ATTEMPTS) { Utilities.sleep(2000); continue; }
        return null;
      }

      const cleaned = text.replace(/```json|```/g, '').trim();
      const parsed  = JSON.parse(cleaned);
      const points  = Array.isArray(parsed.points)
        ? parsed.points.map(p => String(p).trim()).filter(p => p)
        : [];

      return { points };

    } catch (e) {
      // AIチェックはあくまで付随機能のため、最終的に失敗してもOCR処理
      // 全体は継続させる（呼び出し元でnullチェックのうえログのみ出す設計）
      console.warn(`checkDrawingConcerns_ 例外（試行${attempt}/${MAX_ATTEMPTS}）: ${file.getName()}`, e);
      if (attempt < MAX_ATTEMPTS) { Utilities.sleep(2000); continue; }
      console.error('checkDrawingConcerns_ error（最終試行も失敗）:', file.getName(), e);
      return null;
    }
  }
  return null; // ここには到達しないはずだが、念のため
}


// ============================================================
// 説明文のAI添削（③ 図面説明文の追加 で使用）
// ============================================================

const PROOFREAD_PROMPT_PREFIX = `
以下は、社内の図面承認システムに登録する「機械の説明文」です。
誤字脱字・文法の誤りを直し、簡潔で分かりやすい技術文書としての表現に整えてください。
事実関係や内容を勝手に追加・削除・変更しないでください。
返答は添削後の本文のみとし、前置きや説明、マークダウンの記号は一切含めないでください。

【添削対象の説明文】
`;

/**
 * 主図番の説明文をGemini APIで添削する（22_ApprovalDialog.htmlから呼ばれる）
 * OCR機能（extractDrawingInfoByOcr_）と同じ GEMINI_API_KEY / GEMINI_MODEL を流用する。
 * 添削結果は自動反映せず、呼び出し元（クライアント）で人間が内容を確認した上で
 * 採用するかどうかを判断する想定。
 *
 * @param {string} text - 添削対象の説明文
 * @returns {{success: boolean, message?: string, original?: string, proofread?: string}}
 */
function proofreadDescription(text) {
  try {
    const original = String(text || '').trim();
    if (!original) {
      return { success: false, message: '説明文が空です。先に本文を入力してください。' };
    }

    const apiKey = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');
    if (!apiKey) {
      return {
        success: false,
        message: 'スクリプトプロパティ「GEMINI_API_KEY」が設定されていません。',
      };
    }

    const requestBody = {
      contents: [{ parts: [{ text: PROOFREAD_PROMPT_PREFIX + original }] }],
      generationConfig: {
        temperature:     0.2,  // 添削なので多少の言い回し調整は許容しつつ、大きく逸脱させない
        maxOutputTokens: 512,
      },
    };

    const response = UrlFetchApp.fetch(`${GEMINI_API_URL}?key=${apiKey}`, {
      method:             'post',
      contentType:        'application/json',
      payload:            JSON.stringify(requestBody),
      muteHttpExceptions: true,
    });

    const statusCode = response.getResponseCode();
    if (statusCode !== 200) {
      console.error(`Gemini API エラー（説明文添削, ${statusCode}）: ${response.getContentText()}`);
      return { success: false, message: `AI添削中にエラーが発生しました（ステータス${statusCode}）。` };
    }

    const responseJson = JSON.parse(response.getContentText());
    const proofread = responseJson?.candidates?.[0]?.content?.parts?.[0]?.text;
    if (!proofread) {
      console.warn('Gemini API から有効な添削結果が返りませんでした。', JSON.stringify(responseJson));
      return { success: false, message: 'AIから有効な添削結果が返りませんでした。' };
    }

    return { success: true, original: original, proofread: proofread.trim() };

  } catch (e) {
    console.error('proofreadDescription error:', e);
    return { success: false, message: `エラーが発生しました：${e.message}` };
  }
}

// ============================================================
// 【デバッグ用】
// ============================================================

/**
 * 指定した Drive ファイル ID の PDF で OCR をテスト実行する
 * 実行前に FILE_ID を実際の PDF ファイル ID に書き換えてください
 */
function TEST_extractDrawingInfoByOcr() {
  const FILE_ID = '1FQzoH7ujDv9ALzGdFFklOHfrEErBuTM7';
  try {
    const file   = DriveApp.getFileById(FILE_ID);
    const result = extractDrawingInfoByOcr_(file);
    console.log('OCR結果:', JSON.stringify(result, null, 2));
    if (result?.drawingNo) {
      const parsed = parseDrawingNo_(result.drawingNo);
      console.log('図番パース:', JSON.stringify(parsed, null, 2));
    }
  } catch (e) {
    console.error('テスト失敗:', e.message);
  }
}

/**
 * 図面番号パースのユニットテスト
 */
function TEST_parseDrawingNo() {
  const cases = [
    'JB2601001-',  // 正常（初版）
    'JB2601002A',  // 正常（改訂A）
    'JB260100',    // 桁数不足
    'jb2601001-',  // 小文字NG
    'JB2601001Z',  // 正常（改訂Z）
  ];
  cases.forEach(c => {
    const result = parseDrawingNo_(c);
    console.log(`${c} → isValid:${result.isValid}`, result.isValid
      ? `mainNo:${result.mainNo} subNo:${result.subNo} revMark:${result.revMark}`
      : result.errorMsg
    );
  });
}

/**
 * スクリプトプロパティの設定状況を確認する
 */
function checkScriptProperties() {
  const props = PropertiesService.getScriptProperties().getProperties();
  const keys  = ['GEMINI_API_KEY', 'OCR_FOLDER_ID', 'APPROVAL_WEB_APP_URL'];
  keys.forEach(key => {
    const val = props[key];
    if (val) {
      const display = key === 'GEMINI_API_KEY'
        ? `設定済み（末尾: ...${val.slice(-4)}）`
        : `設定済み: ${val}`;
      console.log(`✅ ${key}: ${display}`);
    } else {
      console.warn(`❌ ${key}: 未設定`);
    }
  });
}