function colorAndReorderSheets_RUN() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  var CORE   = '#4285F4'; // 青：中核
  var LOG    = '#34A853'; // 緑：ログ・履歴
  var DATA   = '#9C27B0'; // 紫：業務データ
  var REVIEW = '#EA4335'; // 赤：要確認
  var OTHER  = '#9E9E9E'; // 灰：未分類

  // 左から並べる順（用途別）
  var order = [
    ['在庫同期管理', CORE],
    ['eBayRaw', CORE],
    ['eBay在庫承認', CORE],
    ['eBay出品キュー', CORE],
    ['MCF_SKU_MAP', CORE],

    ['eBay受注ログ', LOG],
    ['MCF監視ログ', LOG],
    ['eBay在庫変更履歴', LOG],
    ['eBay_LISTING_LOG', LOG],
    ['MCF_LINE_NOTIFICATION_LOG', LOG],
    ['実行ログ', LOG],

    ['仕入帳', DATA],

    ['eBay出品リスト', REVIEW],
    ['eBay手残り利益候補', REVIEW],
    ['FBA在庫_API', REVIEW],
    ['手順_在庫同期_SPAPI', REVIEW],
    ['トリガ', REVIEW]
  ];

  var pos = 1;
  var placed = {};
  for (var i = 0; i < order.length; i++) {
    var name = order[i][0];
    var sheet = ss.getSheetByName(name);
    if (!sheet) { Logger.log('注意: シートなし → ' + name); continue; }
    sheet.setTabColor(order[i][1]);
    ss.setActiveSheet(sheet);
    ss.moveActiveSheet(pos);
    placed[name] = true;
    pos++;
  }

  // 一覧に無いシートは灰色（位置はそのまま）
  var all = ss.getSheets();
  var leftover = [];
  for (var j = 0; j < all.length; j++) {
    if (!placed[all[j].getName()]) {
      all[j].setTabColor(OTHER);
      leftover.push(all[j].getName());
    }
  }

  Logger.log('=== 色分け＋並べ替え 完了 ===');
  Logger.log('🔵中核 / 🟢ログ / 🟣データ / 🔴要確認 / ⚪未分類');
  Logger.log('未分類（灰）: ' + (leftover.length ? leftover.join(', ') : 'なし'));
}