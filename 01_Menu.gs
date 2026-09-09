// ============================================================
// カスタムメニュー・サイドバー
// ============================================================

/**
 * シンプルトリガー（スプレッドシートを開くと自動実行される）
 * ※ シンプルトリガーは権限が制限されており、showSidebar()のような
 *   一部の機能を呼び出せない。そのためここではメニュー作成のみを行い、
 *   サイドバーの自動表示は別途「インストール型トリガー」
 *   （onOpenInstallable、フル権限で動作）で行う。
 *   セットアップ手順は本ファイル末尾の setupAutoShowSidebarTrigger() を参照。
 */
function onOpen() {
  const menu = SpreadsheetApp.getUi()
    .createMenu('📐 図面管理')
    .addItem('サイドバーを開く', 'showSidebar');

  // 過去図面インポート機能は移行期間限定（LEGACY_IMPORT_DEADLINE を参照）。
  // 抜け道対策として、期限内のみメニューに表示する（40_LegacyDrawingImport.gs を参照）。
  if (isLegacyImportAvailable_()) {
    menu
      .addSeparator()
      .addItem('【過去図面】① 一括読み取り（OCR）', 'runLegacyOcrAndFillInputSheet')
      .addItem('【過去図面】② 一括登録する', 'RegisterLegacyDrawings');
  }

  menu.addToUi();
}

/**
 * インストール型トリガー用：スプレッドシートを開いたら自動的にサイドバーを表示する
 * （setupAutoShowSidebarTrigger() で1回だけ登録すれば、以降は自動実行される）
 */
function onOpenInstallable() {
  showSidebar();
}

/**
 * 各機能をボタン形式で呼び出せるサイドバーを表示する
 * （カスタムメニュー「サイドバーを開く」からも呼ばれる）
 */
function showSidebar() {
  const template = HtmlService.createTemplateFromFile('02_Sidebar');
  // 過去図面インポートのボタンは移行期間限定で出し分ける（40_LegacyDrawingImport.gs を参照）
  template.legacyImportAvailable = isLegacyImportAvailable_();
  const html = template.evaluate().setTitle('📐 図面管理');
  SpreadsheetApp.getUi().showSidebar(html);
}

// ============================================================
// 【初回セットアップ用】自動サイドバー表示のインストール型トリガーを登録する
// ============================================================
/**
 * この関数を、Apps Script エディタから直接1回だけ手動実行してください。
 * 実行時に権限承認のダイアログが出るので、許可してください。
 * これにより、以降はスプレッドシートを開くたびに自動でサイドバーが
 * 表示されるようになります（onOpenInstallable がフル権限で動くようになる）。
 *
 * 実行方法：
 *   1. Apps Script エディタでこのファイルを開く
 *   2. 関数選択プルダウンで setupAutoShowSidebarTrigger を選択
 *   3. ▶実行 をクリックし、権限承認ダイアログで許可する
 *   4. 完了したら、スプレッドシートを開き直して自動表示を確認する
 *
 * ※ 一度登録すれば、このスプレッドシートに対して恒久的に有効。
 *   再実行すると同じトリガーが重複登録されるのを防ぐため、
 *   実行前に既存の同名トリガーを削除してから登録し直す。
 */
function setupAutoShowSidebarTrigger() {
  // 既存の同名トリガーを削除（重複登録防止）
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(t => {
    if (t.getHandlerFunction() === 'onOpenInstallable') {
      ScriptApp.deleteTrigger(t);
    }
  });

  // インストール型のonOpenトリガーを新規登録
  ScriptApp.newTrigger('onOpenInstallable')
    .forSpreadsheet(SpreadsheetApp.getActiveSpreadsheet())
    .onOpen()
    .create();

  console.log('✅ 自動サイドバー表示のインストール型トリガーを登録しました。');
  console.log('スプレッドシートを開き直して、サイドバーが自動表示されるか確認してください。');
}