/* ===================================================================
 * 作業フロー（図解版）  buildWorkflowVisual_RUN
 * -------------------------------------------------------------------
 * スプレッドシートの「作業フロー」シートを、
 * 色付きボックス＋矢印のフロー図として描画する。
 * さらに下部にシート一覧・トリガー一覧を表で記録。
 *
 * 実行：buildWorkflowVisual_RUN
 * =================================================================== */

function buildWorkflowVisual_RUN() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var SHEET_NAME = '作業フロー';

  var sh = ss.getSheetByName(SHEET_NAME);
  if (sh) { ss.deleteSheet(sh); }       // 作り直し
  sh = ss.insertSheet(SHEET_NAME, 0);   // 先頭に作成
  sh.setTabColor('#FBBC04');

  // ── 全体レイアウト整え ──
  sh.setHiddenGridlines(true);
  // 列幅（A〜H）
  var widths = [30, 150, 150, 60, 150, 150, 150, 30];
  for (var c = 0; c < widths.length; c++) sh.setColumnWidth(c + 1, widths[c]);

  // 色
  var BLUE = '#4285F4', BLUE_BG = '#E8F0FE';
  var AMBER = '#F9AB00', AMBER_BG = '#FEF7E0';
  var TEAL = '#1D9E75', TEAL_BG = '#E1F5EE';
  var GRAY = '#5F6368', GRAY_BG = '#F1F3F4';
  var DARK = '#202124';

  // ── タイトル ──
  _wfvTitle_(sh, 1, 'B', 'G', 'Sky Crew Japan  eBay × Amazon 自動化システム  作業フロー', DARK);
  sh.getRange('B2').setValue('最終更新: ' + Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy/MM/dd HH:mm'))
    .setFontColor('#5f6368').setFontSize(9);
  sh.getRange('B2:G2').merge().setHorizontalAlignment('right');

  // ============================================================
  // セクション見出し：受注→出荷
  // ============================================================
  _wfvSectionBar_(sh, 4, 'B', 'G', '■ 受注 → 出荷の自動化（MCF）', BLUE);

  // ① eBay受注検知
  _wfvBox_(sh, 6, 'B', 'C', '① eBay受注検知', 'checkEbayOrders\n5分ごと', BLUE, BLUE_BG);
  _wfvNote_(sh, 6, 'E', 'G', 'eBayの新規注文をGmail解析し、eBay受注ログに記録');
  _wfvArrowDown_(sh, 8, 'B');

  // ② MCF自動作成
  _wfvBox_(sh, 9, 'B', 'C', '② MCF自動作成', 'autoCreateMcfOrders\n受注12h後・1時間ごと', BLUE, BLUE_BG);
  _wfvNote_(sh, 9, 'E', 'G', '受注から12時間後にAmazon MCFへ出荷依頼を自動作成');
  _wfvArrowDown_(sh, 11, 'B');

  // ③ 追跡番号取得
  _wfvBox_(sh, 12, 'B', 'C', '③ 追跡番号 自動取得', 'checkMcfTrackingNumbersAuto\n1時間ごと', BLUE, BLUE_BG);
  _wfvNote_(sh, 12, 'E', 'G', 'Amazonから追跡番号を取得し、eBayへ提出');
  _wfvArrowDown_(sh, 14, 'B');

  // ④ 異常監視（黄）
  _wfvBox_(sh, 15, 'B', 'C', '④ MCF異常監視 v3', 'runMcfAnomalyMonitor\n要トリガー設定', AMBER, AMBER_BG);
  _wfvNote_(sh, 15, 'E', 'G', 'MCF未作成・追跡未取得を検知→LINE通知。種別とレベルを分けて重複を防止（②③を監視）');

  // ============================================================
  // セクション見出し：在庫
  // ============================================================
  _wfvSectionBar_(sh, 18, 'B', 'G', '■ 在庫の自動化（eBay 0 ↔ 1）', TEAL);

  // ⑤ 在庫同期
  _wfvBox_(sh, 20, 'B', 'C', '⑤ 在庫同期', 'runInventorySyncAllAuto\n毎日17:10', TEAL, TEAL_BG);
  _wfvNote_(sh, 20, 'E', 'G', 'FBA貼付在庫を在庫同期管理シートのD列へ反映');
  _wfvArrowDown_(sh, 22, 'B');

  // ⑥ 在庫承認
  _wfvBox_(sh, 23, 'B', 'C', '⑥ 在庫承認 検出＋通知', 'dailyInventoryApprovalNotify\n毎日8:00', TEAL, TEAL_BG);
  _wfvNote_(sh, 23, 'E', 'G', 'eBay数量0↔1の候補を検出→LINE通知。「進めてください」でeBayを更新');

  // ============================================================
  // シート一覧（表）
  // ============================================================
  var r = 26;
  _wfvSectionBar_(sh, r, 'B', 'G', '■ 各シートの意味', GRAY); r += 2;

  var sheetRows = [
    ['🔵中核', '在庫同期管理', 'SKU別 Amazon/eBay在庫・基準数の同期'],
    ['🔵中核', 'eBayRaw', 'eBay出品の生データ（SKU・ItemID）'],
    ['🔵中核', 'eBay在庫承認', '数量0↔1の承認待ち/承認済リスト'],
    ['🔵中核', 'eBay出品キュー', 'eBay自動出品の待ち行列'],
    ['🔵中核', 'eBay出品リスト', 'eBay出品対象リスト'],
    ['🔵中核', 'MCF_SKU_MAP', 'SKU→Amazon商品の対応表（707行）'],
    ['🔵中核', 'FBA在庫_API', 'SP-API取得のFBA在庫（SKU検索の第2参照）'],
    ['🟢ログ', 'eBay受注ログ', '受注〜MCF〜追跡の全記録＋監視列V〜AC'],
    ['🟢ログ', 'MCF監視ログ', 'MCF異常監視の実行履歴'],
    ['🟢ログ', 'eBay在庫変更履歴', 'eBay在庫を0/1に変えた履歴'],
    ['🟢ログ', 'eBay_LISTING_LOG', 'eBay出品処理のログ'],
    ['🟢ログ', 'MCF_LINE_NOTIFICATION_LOG', 'LINE通知の重複防止ログ'],
    ['🟢ログ', '実行ログ', '各種処理の実行ログ'],
    ['🟣データ', '仕入帳', '仕入記録（手動・業務帳簿）'],
    ['🔴削除候補', 'eBay手残り利益候補', 'コード未参照（分析用？要確認）'],
    ['🔴削除候補', '手順_在庫同期_SPAPI', 'コード未参照（手順メモ）'],
    ['🔴削除候補', 'トリガ', 'コード未参照（メモ？要確認）']
  ];
  _wfvTable_(sh, r, ['分類', 'シート名', '意味・役割'], sheetRows, ['B','C','E']);
  r += sheetRows.length + 2;

  // ============================================================
  // トリガー一覧（表）
  // ============================================================
  _wfvSectionBar_(sh, r, 'B', 'G', '■ トリガーの意味', GRAY); r += 2;

  var trigRows = [
    ['checkEbayOrders', '5分ごと', 'eBay新規受注の検知'],
    ['autoCreateMcfOrders', '1時間ごと', 'MCF自動作成（受注12h後）'],
    ['checkMcfTrackingNumbersAuto', '1時間ごと', '追跡番号の取得・eBay提出'],
    ['runInventorySyncAllAuto', '毎日17:10', '在庫同期'],
    ['dailyInventoryApprovalNotify', '毎日8:00', '在庫承認 候補検出＋LINE通知'],
    ['refreshFbaStockListFromSpApi', '毎日3:00(設定時)', 'FBA在庫リスト更新'],
    ['runMcfAnomalyMonitor', '未設定（手動）', 'MCF異常監視（要設定）']
  ];
  _wfvTable_(sh, r, ['トリガー関数', '頻度', '意味'], trigRows, ['B','C','E']);
  r += trigRows.length + 2;

  sh.getRange(r, 2).setValue('※ 実際の設定中トリガーは「トリガー」画面（時計アイコン）で確認できます。')
    .setFontColor('#5f6368').setFontSize(9).setFontStyle('italic');
  r++;
  sh.getRange(r, 2).setValue('※ runMcfAnomalyMonitor はトリガー未設定。自動監視には別途設定が必要。')
    .setFontColor('#c5221f').setFontSize(9).setFontStyle('italic');

  ss.setActiveSheet(sh);
  sh.setActiveSelection('A1');
  Logger.log('=== 作業フロー（図解版）作成完了 ===');
}


// ── タイトルバー ──
function _wfvTitle_(sh, row, colStart, colEnd, text, color) {
  var rng = sh.getRange(colStart + row + ':' + colEnd + row);
  rng.merge().setValue(text)
    .setBackground(color).setFontColor('#ffffff')
    .setFontWeight('bold').setFontSize(13)
    .setHorizontalAlignment('center').setVerticalAlignment('middle');
  sh.setRowHeight(row, 40);
}

// ── セクションバー ──
function _wfvSectionBar_(sh, row, colStart, colEnd, text, color) {
  var rng = sh.getRange(colStart + row + ':' + colEnd + row);
  rng.merge().setValue(text)
    .setBackground(color).setFontColor('#ffffff')
    .setFontWeight('bold').setFontSize(11)
    .setHorizontalAlignment('left').setVerticalAlignment('middle');
  sh.setRowHeight(row, 28);
}

// ── 処理ボックス（2列ぶち抜き・3行分の高さ）──
function _wfvBox_(sh, row, colStart, colEnd, title, sub, borderColor, bgColor) {
  var rng = sh.getRange(colStart + row + ':' + colEnd + (row + 1));
  rng.merge()
    .setBackground(bgColor)
    .setBorder(true, true, true, true, false, false, borderColor, SpreadsheetApp.BorderStyle.SOLID_THICK)
    .setHorizontalAlignment('center').setVerticalAlignment('middle')
    .setWrap(true);
  // タイトル＋サブを1セルに（リッチテキスト風に改行）
  rng.setValue(title + '\n' + sub);
  rng.setFontColor('#202124');
  sh.setRowHeight(row, 24);
  sh.setRowHeight(row + 1, 26);
}

// ── 補足ノート ──
function _wfvNote_(sh, row, colStart, colEnd, text) {
  var rng = sh.getRange(colStart + row + ':' + colEnd + (row + 1));
  rng.merge().setValue(text)
    .setFontColor('#3c4043').setFontSize(10)
    .setHorizontalAlignment('left').setVerticalAlignment('middle')
    .setWrap(true);
}

// ── 下向き矢印 ──
function _wfvArrowDown_(sh, row, col) {
  sh.getRange(col + row).setValue('▼')
    .setFontColor('#9aa0a6').setFontSize(12)
    .setHorizontalAlignment('center').setVerticalAlignment('middle');
  sh.setRowHeight(row, 20);
}

// ── 表 ──
function _wfvTable_(sh, startRow, headers, rows, cols) {
  // ヘッダー
  for (var h = 0; h < headers.length; h++) {
    var cell = (h < cols.length - 1)
      ? sh.getRange(cols[h] + startRow)
      : sh.getRange(cols[h] + startRow + ':G' + startRow);
    if (h === cols.length - 1) cell.merge();
    cell.setValue(headers[h])
      .setBackground('#1a73e8').setFontColor('#ffffff')
      .setFontWeight('bold').setFontSize(10);
  }
  // 本体
  for (var i = 0; i < rows.length; i++) {
    var rr = startRow + 1 + i;
    sh.getRange(cols[0] + rr).setValue(rows[i][0]).setFontSize(10);
    sh.getRange(cols[1] + rr).setValue(rows[i][1]).setFontSize(10);
    var noteCell = sh.getRange(cols[2] + rr + ':G' + rr).merge();
    noteCell.setValue(rows[i][2]).setFontSize(10).setFontColor('#3c4043').setWrap(true);
    var bg = (i % 2 === 0) ? '#ffffff' : '#f8f9fa';
    sh.getRange(cols[0] + rr + ':G' + rr).setBackground(bg);
  }
}