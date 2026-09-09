/**
 * ============================================================
 * 図面承認・採番・出図管理システム
 * ファイル: 26_CopyFolderPath.gs  ―  エクスプローラーパスのコピー機能
 * ============================================================
 *
 * 【役割】
 *   機械台帳で選択中の行のエクスプローラーパス（J列）を、
 *   ワンクリックでクリップボードにコピーできるダイアログを提供する。
 *
 * 【技術的な制約】
 *   Googleスプレッドシートのセルは、クリックしただけで任意のJavaScriptを
 *   実行できる仕組みを持たない。また、クリップボードへの書き込みは
 *   ブラウザの仕様上、ユーザーの明示的な操作（ボタンクリック等）を
 *   伴う文脈でしか行えない。そのため「セルをクリックしたら自動でコピー」は
 *   実現できず、「メニューから実行 → ダイアログのボタンでコピー」という
 *   形にしている。
 *
 * 【使い方】
 *   1. 機械台帳で、パスを取得したい行の任意のセルを選択する
 *   2. カスタムメニュー「📐 図面管理」→「📋 エクスプローラーパスをコピー」を実行
 *   3. 開いたダイアログでパスが自動的にクリップボードへコピーされる
 *      （自動コピーが働かない場合は「コピー」ボタンを押す）
 */
function openCopyPathDialog() {
  const ui    = SpreadsheetApp.getUi();
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getActiveSheet();

  if (sheet.getName() !== SHEET_NAMES.MAIN_NUMBER) {
    ui.alert(
      'シートを確認してください',
      '「機械台帳」シートで、パスをコピーしたい行のセルを選択してから実行してください。',
      ui.ButtonSet.OK
    );
    return;
  }

  const activeRow = ss.getActiveRange().getRow();
  if (activeRow < GLOBAL_ROW.DATA_START) {
    ui.alert(
      '行を確認してください',
      'データ行（5行目以降）のセルを選択してから実行してください。',
      ui.ButtonSet.OK
    );
    return;
  }

  const mainNo = String(sheet.getRange(activeRow, COL_MAIN.MAIN_NO).getValue()).trim();
  const path   = String(sheet.getRange(activeRow, COL_MAIN.EXPLORER_PATH).getValue()).trim();

  if (!path) {
    ui.alert(
      'パス未登録',
      `主図番「${mainNo || '（空欄）'}」の行には、まだフォルダパスが記録されていません。\n` +
      `この主図番の図面が一度も承認されていない可能性があります。`,
      ui.ButtonSet.OK
    );
    return;
  }

  const html = HtmlService.createTemplateFromFile('27_CopyPathDialog');
  html.mainNo = mainNo;
  html.path   = path;

  ui.showModalDialog(
    html.evaluate().setWidth(480).setHeight(240),
    'フォルダパスをコピー'
  );
}