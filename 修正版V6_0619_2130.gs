// ============================================================
// onOpen() 関数 - メニュー作成
// ============================================================
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('eBay 自動化')
    .addItem('📋 受注チェック実行', 'checkEbayOrders')
    .addItem('📬 MCF 追跡番号チェック', 'checkMcfTrackingNumbers')
    .addItem('📤 eBay に追跡番号提出', 'autoSubmitTrackingToEbay')
    .addSeparator()
    .addItem('📖 ヘルプ', 'showHelp')
    .addToUi();
}

function testMcfMailPermission() {
  MailApp.sendEmail({
    to: 'shinicchee@gmail.com',
    subject: 'MCF監視メール権限テスト',
    body: 'MailApp.sendEmail の権限確認テストです。'
  });
}/**
 * corrected_category_audit_v4.gs
 * eBay Taxonomy APIによるカテゴリ候補・Item Specificsの読み取り専用監査。
 * シート書き込み・出品・トリガー操作は行わない。
 */

var EBAY_CATEGORY_AUDIT_CONFIG_ = {
  taxonomyBase: 'https://api.ebay.com/commerce/taxonomy/v1',
  marketplaceId: 'EBAY_US',
  applicationScope: 'https://api.ebay.com/oauth/api_scope',
  searchQueries: [
    'Shimano 23 Vanquish C2000SHG spinning reel',
    'Shimano Vanquish spinning reel'
  ],
  maxCategoriesToInspect: 8
};

function auditCategoryAndItemSpecifics_ReadOnly() {
  Logger.log('=== eBayカテゴリ・Item Specifics読み取り監査 開始 ===');
  Logger.log('シート書き込み・出品・更新処理は行いません。');

  var accessToken = ebayCategoryAudit_GetApplicationToken_();
  if (!accessToken) return;

  var categoryTreeId = ebayCategoryAudit_GetDefaultTreeId_(accessToken);
  if (!categoryTreeId) return;
  Logger.log('Category Tree ID: ' + categoryTreeId);

  var seen = {};
  var categories = [];

  EBAY_CATEGORY_AUDIT_CONFIG_.searchQueries.forEach(function(query) {
    var suggestions = ebayCategoryAudit_GetSuggestions_(accessToken, categoryTreeId, query);
    Logger.log('--- Query: ' + query + ' / suggestions=' + suggestions.length + ' ---');

    suggestions.forEach(function(suggestion) {
      if (!seen[suggestion.categoryId]) {
        seen[suggestion.categoryId] = true;
        categories.push(suggestion);
      }
      Logger.log('CategoryID=' + suggestion.categoryId
        + ' / Name=' + suggestion.categoryName
        + ' / Path=' + suggestion.parentPath);
    });
  });

  if (categories.length === 0) {
    Logger.log('❌ カテゴリ候補を取得できませんでした。');
    return;
  }

  categories = categories.slice(0, EBAY_CATEGORY_AUDIT_CONFIG_.maxCategoriesToInspect);
  Logger.log('Item Specifics調査対象: ' + categories.length + 'カテゴリ');

  categories.forEach(function(category) {
    Logger.log('========================================');
    Logger.log('CategoryID=' + category.categoryId + ' / ' + category.categoryName);

    var aspects = ebayCategoryAudit_GetAspects_(accessToken, categoryTreeId, category.categoryId);
    if (aspects.length === 0) {
      Logger.log('Item Specificsなし、取得失敗、または非リーフカテゴリの可能性があります。');
      return;
    }

    var required = [];
    var recommended = [];
    var optional = [];

    aspects.forEach(function(aspect) {
      var constraint = aspect.aspectConstraint || {};
      var isRequired = constraint.aspectRequired === true
        || String(constraint.aspectRequired).toLowerCase() === 'true';
      var usage = constraint.aspectUsage || 'OPTIONAL';
      var values = (aspect.aspectValues || []).slice(0, 10).map(function(value) {
        return value.localizedValue;
      });

      var item = {
        name: aspect.localizedAspectName || '',
        required: isRequired,
        usage: usage,
        mode: constraint.aspectMode || '',
        cardinality: constraint.itemToAspectCardinality || '',
        enabledForVariations: constraint.aspectEnabledForVariations === true,
        expectedRequiredByDate: constraint.expectedRequiredByDate || '',
        sampleValues: values
      };

      if (isRequired) required.push(item);
      else if (usage === 'RECOMMENDED') recommended.push(item);
      else optional.push(item);
    });

    Logger.log('★ Required (' + required.length + ')');
    required.forEach(function(item) { Logger.log(JSON.stringify(item)); });

    Logger.log('◎ Recommended (' + recommended.length + ')');
    recommended.forEach(function(item) { Logger.log(JSON.stringify(item)); });

    Logger.log('Optional (' + optional.length + '): '
      + optional.map(function(item) { return item.name; }).join(', '));
  });

  Logger.log('=== eBayカテゴリ・Item Specifics読み取り監査 終了 ===');
}

function ebayCategoryAudit_GetApplicationToken_() {
  try {
    var properties = PropertiesService.getScriptProperties();
    var clientId = properties.getProperty('EBAY_CLIENT_ID');
    var clientSecret = properties.getProperty('EBAY_CLIENT_SECRET');

    if (!clientId || !clientSecret) {
      Logger.log('❌ Script Properties不足: EBAY_CLIENT_ID / EBAY_CLIENT_SECRET');
      return null;
    }

    var response = UrlFetchApp.fetch('https://api.ebay.com/identity/v1/oauth2/token', {
      method: 'post',
      headers: {
        Authorization: 'Basic ' + Utilities.base64Encode(clientId + ':' + clientSecret),
        'Content-Type': 'application/x-www-form-urlencoded'
      },
      payload: {
        grant_type: 'client_credentials',
        scope: EBAY_CATEGORY_AUDIT_CONFIG_.applicationScope
      },
      muteHttpExceptions: true
    });

    var httpCode = response.getResponseCode();
    var body = response.getContentText();
    if (httpCode !== 200) {
      Logger.log('❌ Application Token取得失敗 HTTP ' + httpCode + ': ' + body.substring(0, 800));
      return null;
    }

    return JSON.parse(body).access_token || null;
  } catch (error) {
    Logger.log('❌ Application Token取得例外: ' + error.message);
    return null;
  }
}

function ebayCategoryAudit_GetDefaultTreeId_(accessToken) {
  var url = EBAY_CATEGORY_AUDIT_CONFIG_.taxonomyBase
    + '/get_default_category_tree_id?marketplace_id='
    + encodeURIComponent(EBAY_CATEGORY_AUDIT_CONFIG_.marketplaceId);

  var json = ebayCategoryAudit_GetJson_(url, accessToken, 'getDefaultCategoryTreeId');
  return json ? json.categoryTreeId || null : null;
}

function ebayCategoryAudit_GetSuggestions_(accessToken, categoryTreeId, query) {
  var url = EBAY_CATEGORY_AUDIT_CONFIG_.taxonomyBase
    + '/category_tree/' + encodeURIComponent(categoryTreeId)
    + '/get_category_suggestions?q=' + encodeURIComponent(query);

  var json = ebayCategoryAudit_GetJson_(url, accessToken, 'getCategorySuggestions');
  if (!json) return [];

  return (json.categorySuggestions || []).map(function(suggestion) {
    var ancestors = suggestion.categoryTreeNodeAncestors || [];
    var pathParts = ancestors.map(function(ancestor) {
      return ancestor.category && ancestor.category.categoryName
        ? ancestor.category.categoryName
        : '';
    }).filter(function(name) { return name !== ''; });

    return {
      categoryId: suggestion.category ? suggestion.category.categoryId || '' : '',
      categoryName: suggestion.category ? suggestion.category.categoryName || '' : '',
      parentPath: pathParts.join(' > '),
      categoryTreeNodeLevel: suggestion.categoryTreeNodeLevel
    };
  }).filter(function(item) { return item.categoryId !== ''; });
}

function ebayCategoryAudit_GetAspects_(accessToken, categoryTreeId, categoryId) {
  var url = EBAY_CATEGORY_AUDIT_CONFIG_.taxonomyBase
    + '/category_tree/' + encodeURIComponent(categoryTreeId)
    + '/get_item_aspects_for_category?category_id=' + encodeURIComponent(categoryId);

  var json = ebayCategoryAudit_GetJson_(url, accessToken, 'getItemAspectsForCategory:' + categoryId);
  return json ? json.aspects || [] : [];
}

function ebayCategoryAudit_GetJson_(url, accessToken, label) {
  try {
    var response = UrlFetchApp.fetch(url, {
      method: 'get',
      headers: ebayCategoryAudit_BuildRestHeaders_(accessToken),
      muteHttpExceptions: true
    });

    var httpCode = response.getResponseCode();
    var body = response.getContentText();
    if (httpCode !== 200) {
      Logger.log('❌ ' + label + ' HTTP ' + httpCode + ': ' + body.substring(0, 800));
      return null;
    }

    return JSON.parse(body);
  } catch (error) {
    Logger.log('❌ ' + label + ' 例外: ' + error.message);
    return null;
  }
}

function ebayCategoryAudit_BuildRestHeaders_(accessToken) {
  return {
    Authorization: 'Bearer ' + accessToken,
    Accept: 'application/json'
  };
}
/**
 * corrected_policy_audit_v4.gs
 * eBay配送ポリシーの読み取り専用監査。
 * 外部変更・シート書き込み・トリガー操作・出品操作は行わない。
 */

var EBAY_POLICY_AUDIT_CONFIG_ = {
  marketplaceId: 'EBAY_US',
  accountApiBase: 'https://api.ebay.com/sell/account/v1'
};

/**
 * 既存のFulfillment Policyを読み取り、実行ログへ表示する。
 * Item Locationの適法性・正確性は別途確認が必要。
 */
function auditFulfillmentPolicies_ReadOnly() {
  Logger.log('=== eBay配送ポリシー読み取り監査 開始 ===');
  Logger.log('変更処理・シート書き込み・出品処理は行いません。');

  var accessToken = ebayPolicyAudit_GetUserAccessToken_();
  if (!accessToken) {
    Logger.log('❌ User Access Tokenを取得できないため終了します。');
    return;
  }

  var url = EBAY_POLICY_AUDIT_CONFIG_.accountApiBase
    + '/fulfillment_policy?marketplace_id='
    + encodeURIComponent(EBAY_POLICY_AUDIT_CONFIG_.marketplaceId);

  var response;
  try {
    response = UrlFetchApp.fetch(url, {
      method: 'get',
      headers: ebayPolicyAudit_BuildRestHeaders_(accessToken),
      muteHttpExceptions: true
    });
  } catch (error) {
    Logger.log('❌ API呼び出し例外: ' + error.message);
    return;
  }

  var httpCode = response.getResponseCode();
  var body = response.getContentText();
  Logger.log('HTTP Status: ' + httpCode);

  if (httpCode !== 200) {
    Logger.log('❌ getFulfillmentPolicies失敗: ' + body.substring(0, 1000));
    if (httpCode === 401 || httpCode === 403) {
      Logger.log('確認事項: 現在のRefresh Tokenに sell.account または sell.account.readonly の同意スコープが含まれているか確認してください。');
      Logger.log('含まれていない場合は、追加スコープ付きでOAuth同意をやり直し、新しいRefresh Tokenが必要です。');
    }
    return;
  }

  var json;
  try {
    json = JSON.parse(body);
  } catch (error) {
    Logger.log('❌ JSON解析失敗: ' + error.message);
    return;
  }

  var policies = json.fulfillmentPolicies || [];
  Logger.log('取得件数: ' + policies.length + '件');

  policies.forEach(function(policy, index) {
    var handling = policy.handlingTime || {};
    var options = policy.shippingOptions || [];
    var domestic = [];
    var international = [];

    options.forEach(function(option) {
      var services = (option.shippingServices || []).map(function(service) {
        var costText = service.freeShipping
          ? 'Free'
          : (service.shippingCost && service.shippingCost.value != null
            ? service.shippingCost.value + ' ' + (service.shippingCost.currency || '')
            : 'Cost not returned');

        return {
          code: service.shippingServiceCode || '',
          carrier: service.shippingCarrierCode || '',
          cost: costText
        };
      });

      if (option.optionType === 'DOMESTIC') domestic = services;
      if (option.optionType === 'INTERNATIONAL') international = services;
    });

    Logger.log('----------------------------------------');
    Logger.log('[' + (index + 1) + '] Policy ID: ' + (policy.fulfillmentPolicyId || ''));
    Logger.log('Name: ' + (policy.name || ''));
    Logger.log('Marketplace: ' + (policy.marketplaceId || ''));
    Logger.log('Handling Time: ' + (handling.value != null ? handling.value : '') + ' ' + (handling.unit || ''));
    Logger.log('Domestic Services: ' + JSON.stringify(domestic));
    Logger.log('International Services: ' + JSON.stringify(international));
    Logger.log('Ship-to Included: ' + JSON.stringify((policy.shipToLocations || {}).regionIncluded || []));
    Logger.log('Ship-to Excluded: ' + JSON.stringify((policy.shipToLocations || {}).regionExcluded || []));
    Logger.log('注意: この情報だけではMCF用ポリシーとして適合するとは判定しません。');
  });

  Logger.log('=== eBay配送ポリシー読み取り監査 終了 ===');
}

/**
 * Refresh TokenからUser Access Tokenを取得する。
 * scopeを省略し、元の同意時に付与されたスコープを継承する。
 */
function ebayPolicyAudit_GetUserAccessToken_() {
  try {
    var properties = PropertiesService.getScriptProperties();
    var clientId = properties.getProperty('EBAY_CLIENT_ID');
    var clientSecret = properties.getProperty('EBAY_CLIENT_SECRET');
    var refreshToken = properties.getProperty('EBAY_REFRESH_TOKEN');

    if (!clientId || !clientSecret || !refreshToken) {
      Logger.log('❌ Script Properties不足: EBAY_CLIENT_ID / EBAY_CLIENT_SECRET / EBAY_REFRESH_TOKEN');
      return null;
    }

    var response = UrlFetchApp.fetch('https://api.ebay.com/identity/v1/oauth2/token', {
      method: 'post',
      headers: {
        Authorization: 'Basic ' + Utilities.base64Encode(clientId + ':' + clientSecret),
        'Content-Type': 'application/x-www-form-urlencoded'
      },
      payload: {
        grant_type: 'refresh_token',
        refresh_token: refreshToken
      },
      muteHttpExceptions: true
    });

    var httpCode = response.getResponseCode();
    var body = response.getContentText();
    if (httpCode !== 200) {
      Logger.log('❌ OAuth Token取得失敗 HTTP ' + httpCode + ': ' + body.substring(0, 800));
      return null;
    }

    var json = JSON.parse(body);
    return json.access_token || null;
  } catch (error) {
    Logger.log('❌ User Access Token取得例外: ' + error.message);
    return null;
  }
}

function ebayPolicyAudit_BuildRestHeaders_(accessToken) {
  return {
    Authorization: 'Bearer ' + accessToken,
    Accept: 'application/json'
  };
}
/**
 * EbayListingQueue.gs
 * ─────────────────────────────────────────────────────────────────
 * スプレッドシート : eBay shop Management
 * 対象シート      : eBay出品キュー
 *
 * 公開関数（3本）
 *   previewEbayListingQueue()    ← 最初に実行可能。シート記録あり。
 *   verifyApprovedEbayListing()  ← 人間承認「Verify承認」後のみ。
 *   publishApprovedEbayListing() ← 人間承認「本番承認」後のみ。現時点で実行禁止。
 *
 * 認証
 *   getEbayAccessTokenFromRefreshToken() 共通関数を使用。
 *   Trading API ヘッダー : X-EBAY-API-IAF-TOKEN: <Access Token>
 *   ※ X-EBAY-C-IAF-TOKEN / X-EBAY-IAF-TOKEN は使用禁止
 *
 * 安全ガード（全関数共通）
 *   ・画像URLが空 → Verify・本番出品を停止
 *   ・価格が空または 0 → 停止
 *   ・SKU がeBayRaw・eBay出品リストに既存 → 停止
 *   ・ItemID が既にある → 停止
 *   ・一度に処理するのは最大 1 行
 *   ・本番関数は人間承認「本番承認」かつ全条件を満たす行のみ対象
 *   ・トリガーは作成しない
 *   ・配送ポリシーは変更しない
 *   ・在庫同期コードは変更しない
 * ─────────────────────────────────────────────────────────────────
 */

// ═══════════════════════════════════════════════════════════
// 定数：シート名・列インデックス（1始まり）
// ═══════════════════════════════════════════════════════════

var ELQ = {
  // ── シート名 ──
  SHEET_QUEUE    : 'eBay出品キュー',
  SHEET_RAW      : 'eBayRaw',
  SHEET_LIST     : 'eBay出品リスト',
  SHEET_LOG      : 'eBay_LISTING_LOG',

  // ── eBay出品キュー 列番号（1始まり、ヘッダー行=1）──
  // A: 出品タスクID   B: 人間承認       C: 処理状態
  // D: SKU           E: ASIN           F: 商品名
  // G: eBayタイトル   H: CategoryID     I: カテゴリ名
  // J: ConditionID   K: 数量           L: 価格USD
  // M: eBay画像URL   N: ShippingProfileID  O: ReturnProfileID
  // P: PaymentProfileID  Q: City        R: State
  // S: Country       T: 商品説明        U: プレビュー結果
  // V: Verify状態    W: Verify Ack     X: Verifyメッセージ
  // Y: ItemID        Z: 出品URL        AA: エラーコード
  // AB: エラー内容   AC: 作成日         AD: 更新日
  // AE: 備考
  COL: {
    TASK_ID         :  1,   // A
    HUMAN_APPROVAL  :  2,   // B
    STATUS          :  3,   // C
    SKU             :  4,   // D
    ASIN            :  5,   // E
    PRODUCT_NAME    :  6,   // F
    EBAY_TITLE      :  7,   // G
    CATEGORY_ID     :  8,   // H
    CATEGORY_NAME   :  9,   // I
    CONDITION_ID    : 10,   // J
    QUANTITY        : 11,   // K
    PRICE_USD       : 12,   // L
    PICTURE_URL     : 13,   // M
    SHIPPING_ID     : 14,   // N
    RETURN_ID       : 15,   // O
    PAYMENT_ID      : 16,   // P
    CITY            : 17,   // Q
    STATE           : 18,   // R
    COUNTRY         : 19,   // S
    DESCRIPTION     : 20,   // T
    PREVIEW_RESULT  : 21,   // U
    VERIFY_STATUS   : 22,   // V
    VERIFY_ACK      : 23,   // W
    VERIFY_MESSAGE  : 24,   // X
    ITEM_ID         : 25,   // Y
    LISTING_URL     : 26,   // Z
    ERROR_CODE      : 27,   // AA
    ERROR_DETAIL    : 28,   // AB
    CREATED_DATE    : 29,   // AC
    UPDATED_DATE    : 30,   // AD
    NOTES           : 31,   // AE
  },

  // ── 人間承認の値 ──
  APPROVAL: {
    UNAPPROVED  : '未承認',
    VERIFY_OK   : 'Verify承認',
    PUBLISH_OK  : '本番承認',
  },

  // ── 処理状態の値 ──
  STATE: {
    DRAFT          : '下書き',
    PREVIEW_DONE   : 'プレビュー済み',
    VERIFY_SUCCESS : 'Verify成功',
    VERIFY_FAILED  : 'Verify失敗',
    PUBLISHED      : '出品済み',
  },

  // ── Verify状態の値 ──
  VERIFY: {
    NOT_YET  : '未実行',
    SUCCESS  : 'Verify成功',
    WARNING  : 'Verify警告',
    FAILED   : 'Verify失敗',
  },

  // ── Trading API ──
  TRADING_ENDPOINT   : 'https://api.ebay.com/ws/api.dll',
  COMPATIBILITY_LEVEL: '1225',
  SITE_ID            : '0',
};






// ═══════════════════════════════════════════════════════════
// 内部: データ行取得
// ═══════════════════════════════════════════════════════════

/**
 * シートのデータ行（ヘッダーを除く）を配列で返す。
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet
 * @returns {Array<{rowNum: number, data: Array}>}
 */
function _getDataRows_(sheet) {
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) { return []; }
  var lastCol = Math.max(sheet.getLastColumn(), ELQ.COL.NOTES);
  var values  = sheet.getRange(2, 1, lastRow - 1, lastCol).getValues();
  var result  = [];
  values.forEach(function(row, idx) {
    // 出品タスクIDが空の行はスキップ
    if (String(row[ELQ.COL.TASK_ID - 1] || '').trim() === '') { return; }
    result.push({ rowNum: idx + 2, data: row });
  });
  return result;
}

// ═══════════════════════════════════════════════════════════
// 内部: 必須項目チェック
// ═══════════════════════════════════════════════════════════

/**
 * 行データの必須項目を検査してメッセージ配列を返す。
 * @param {Array} data
 * @returns {string[]}
 */
function _validateRow_(data) {
  var results = [];
  var checks = [
    { col: ELQ.COL.SKU,         label: 'D列 SKU' },
    { col: ELQ.COL.EBAY_TITLE,  label: 'G列 eBayタイトル' },
    { col: ELQ.COL.CATEGORY_ID, label: 'H列 CategoryID' },
    { col: ELQ.COL.CONDITION_ID,label: 'J列 ConditionID' },
    { col: ELQ.COL.QUANTITY,    label: 'K列 数量' },
    { col: ELQ.COL.PRICE_USD,   label: 'L列 価格USD' },
    { col: ELQ.COL.SHIPPING_ID, label: 'N列 ShippingProfileID' },
    { col: ELQ.COL.RETURN_ID,   label: 'O列 ReturnProfileID' },
    { col: ELQ.COL.PAYMENT_ID,  label: 'P列 PaymentProfileID' },
    { col: ELQ.COL.CITY,        label: 'Q列 City' },
    { col: ELQ.COL.COUNTRY,     label: 'S列 Country' },
  ];

  checks.forEach(function(c) {
    var val = String(data[c.col - 1] || '').trim();
    if (val === '') {
      results.push('❌ ' + c.label + ' が空です。');
    } else {
      results.push('✅ ' + c.label + ' = ' + val.substring(0, 40));
    }
  });

  // 価格が 0 でないか
  var price = parseFloat(data[ELQ.COL.PRICE_USD - 1]);
  if (!isNaN(price) && price <= 0) {
    results.push('❌ L列 価格USD が 0 以下です（値: ' + price + '）。');
  }

  // タイトル文字数チェック
  var title = String(data[ELQ.COL.EBAY_TITLE - 1] || '');
  if (title.length > 80) {
    results.push('❌ G列 eBayタイトルが80文字超です（' + title.length + '文字）。');
  } else if (title.length > 0) {
    results.push('✅ G列 eBayタイトル文字数: ' + title.length + '文字（80文字以内）');
  }

  // 画像URL（Verify時は警告、本番時は停止）
  var picture = String(data[ELQ.COL.PICTURE_URL - 1] || '').trim();
  if (picture === '') {
    results.push('⚠️  M列 eBay画像URL が空です（Verify・本番出品は停止されます）。');
  } else if (picture.indexOf('media-amazon') !== -1 || picture.indexOf('images-amazon') !== -1) {
    results.push('❌ M列 eBay画像URL に Amazon 画像 URL が設定されています（使用禁止）。');
  } else {
    results.push('✅ M列 eBay画像URL = ' + picture.substring(0, 60));
  }

  return results;
}

// ═══════════════════════════════════════════════════════════
// 内部: 安全ガード
// ═══════════════════════════════════════════════════════════

/**
 * Verify・本番出品前の安全ガードを実行する。
 * @param {Array} data - 行データ
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} ss
 * @param {boolean} isPublish - true=本番出品用（より厳しいチェック）
 * @returns {{ok: boolean, messages: string[]}}
 */
function _runGuards_(data, ss, isPublish) {
  var messages = [];

  var sku      = String(data[ELQ.COL.SKU         - 1] || '').trim();
  var price    = parseFloat(data[ELQ.COL.PRICE_USD  - 1]) || 0;
  var picture  = String(data[ELQ.COL.PICTURE_URL - 1] || '').trim();
  var itemId   = String(data[ELQ.COL.ITEM_ID     - 1] || '').trim();
  var title    = String(data[ELQ.COL.EBAY_TITLE  - 1] || '');

  // ─ 価格チェック ─
  if (price <= 0) {
    messages.push('価格が空または0以下です（L列: ' + price + '）。');
  }

  // ─ 画像URLチェック ─
  if (picture === '') {
    messages.push('eBay画像URLが空です（M列）。eBay EPSアップロード後のURLを設定してください。');
  } else if (picture.indexOf('media-amazon') !== -1 || picture.indexOf('images-amazon') !== -1) {
    messages.push('Amazon画像URLは使用禁止です（M列）。');
  }

  // ─ ItemID重複チェック（既にItemIDがある = 既出品）─
  if (itemId !== '') {
    messages.push('Y列にItemIDが既に存在します（' + itemId + '）。重複出品を防止するため停止します。');
  }

  // ─ タイトル文字数 ─
  if (title.length > 80) {
    messages.push('eBayタイトルが80文字超です（' + title.length + '文字）。');
  }

  // ─ SKU重複チェック（eBayRaw・eBay出品リスト）─
  if (sku !== '') {
    var dupResult = _checkSkuDuplicate_(sku, ss);
    if (dupResult.found) {
      messages.push('SKU「' + sku + '」が既にeBayに出品されています（' + dupResult.source + '）。重複出品を防止するため停止します。');
    }
  }

  return { ok: messages.length === 0, messages: messages };
}

/**
 * SKU がeBayRaw・eBay出品リストに既存かチェックする。
 * @param {string} sku
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} ss
 * @returns {{found: boolean, source: string}}
 */
function _checkSkuDuplicate_(sku, ss) {
  // ── eBayRaw チェック ──
  var rawSheet = ss.getSheetByName(ELQ.SHEET_RAW);
  if (rawSheet) {
    var rawData = rawSheet.getDataRange().getValues();
    for (var i = 1; i < rawData.length; i++) {
      // eBayRaw: B列(index 1) = SKU
      if (String(rawData[i][1] || '').trim() === sku) {
        return { found: true, source: 'eBayRaw B列 行' + (i + 1) };
      }
    }
  }

  // ── eBay出品リスト チェック ──
  var listSheet = ss.getSheetByName(ELQ.SHEET_LIST);
  if (listSheet) {
    var listData = listSheet.getDataRange().getValues();
    for (var j = 1; j < listData.length; j++) {
      // eBay出品リスト: D列(index 3) = Custom label (SKU)
      if (String(listData[j][3] || '').trim() === sku) {
        return { found: true, source: 'eBay出品リスト D列 行' + (j + 1) };
      }
    }
  }

  return { found: false, source: '' };
}

// ═══════════════════════════════════════════════════════════
// 内部: XML 生成
// ═══════════════════════════════════════════════════════════

/**
 * VerifyAddFixedPriceItemRequest XML を生成する。
 * publishApprovedEbayListing() 内で AddFixedPriceItemRequest へ置換して使用する。
 * RequesterCredentials は使用しない（IAFトークン方式）。
 * PostalCode フィールドは含めない（既存ShippingProfile継承）。
 *
 * @param {Array}  data       - 行データ
 * @param {string} priceStr   - 価格文字列
 * @param {string} pictureUrl - 画像URL文字列
 * @returns {string} XML文字列
 */
function _buildXml_(data, priceStr, pictureUrl) {
  var title       = _esc_(String(data[ELQ.COL.EBAY_TITLE  - 1] || ''));
  var sku         = _esc_(String(data[ELQ.COL.SKU         - 1] || ''));
  var categoryId  = _esc_(String(data[ELQ.COL.CATEGORY_ID - 1] || ''));
  var conditionId = _esc_(String(data[ELQ.COL.CONDITION_ID- 1] || ''));
  var quantity    = _esc_(String(data[ELQ.COL.QUANTITY    - 1] || '1'));
  var shippingId  = _esc_(String(data[ELQ.COL.SHIPPING_ID - 1] || ''));
  var returnId    = _esc_(String(data[ELQ.COL.RETURN_ID   - 1] || ''));
  var paymentId   = _esc_(String(data[ELQ.COL.PAYMENT_ID  - 1] || ''));
  var city        = _esc_(String(data[ELQ.COL.CITY        - 1] || ''));
  var country     = _esc_(String(data[ELQ.COL.COUNTRY     - 1] || ''));
  var description = String(data[ELQ.COL.DESCRIPTION - 1] || '');

  // ShippingProfileName は読み取り専用（シートに列がないため固定値）
  var shippingName = '３日出荷_wide';

  return '<?xml version="1.0" encoding="utf-8"?>\n'
    + '<VerifyAddFixedPriceItemRequest xmlns="urn:ebay:apis:eBLBaseComponents">\n'
    + '  <!-- 認証: IAFトークン方式。RequesterCredentials未使用 -->\n'
    + '  <!-- X-EBAY-API-IAF-TOKEN ヘッダーでAccess Tokenを渡す -->\n'
    + '\n'
    + '  <Item>\n'
    + '    <Title>' + title + '</Title>\n'
    + '    <SKU>' + sku + '</SKU>\n'
    + '\n'
    + '    <PrimaryCategory>\n'
    + '      <CategoryID>' + categoryId + '</CategoryID>\n'
    + '    </PrimaryCategory>\n'
    + '\n'
    + '    <StartPrice currencyID="USD">' + _esc_(priceStr) + '</StartPrice>\n'
    + '    <Quantity>' + quantity + '</Quantity>\n'
    + '    <ListingType>FixedPriceItem</ListingType>\n'
    + '    <ListingDuration>GTC</ListingDuration>\n'
    + '    <OutOfStockControl>true</OutOfStockControl>\n'
    + '\n'
    + '    <ConditionID>' + conditionId + '</ConditionID>\n'
    + '\n'
    + '    <!-- 画像: eBay EPSアップロード済みURLのみ使用。Amazon画像URL禁止 -->\n'
    + '    <PictureDetails>\n'
    + '      <PictureURL>' + _esc_(pictureUrl) + '</PictureURL>\n'
    + '    </PictureDetails>\n'
    + '\n'
    + '    <!-- PostalCode: 既存ShippingProfileを継承。フィールドに含めない -->\n'
    + '    <Location>' + city + '</Location>\n'
    + '    <Currency>USD</Currency>\n'
    + '    <Country>' + country + '</Country>\n'
    + '\n'
    + '    <SellerProfiles>\n'
    + '      <SellerShippingProfile>\n'
    + '        <ShippingProfileID>' + shippingId + '</ShippingProfileID>\n'
    + '        <ShippingProfileName>' + _esc_(shippingName) + '</ShippingProfileName>\n'
    + '      </SellerShippingProfile>\n'
    + '      <SellerReturnProfile>\n'
    + '        <ReturnProfileID>' + returnId + '</ReturnProfileID>\n'
    + '      </SellerReturnProfile>\n'
    + '      <SellerPaymentProfile>\n'
    + '        <PaymentProfileID>' + paymentId + '</PaymentProfileID>\n'
    + '      </SellerPaymentProfile>\n'
    + '    </SellerProfiles>\n'
    + '\n'
    + '    <ItemSpecifics>\n'
    + '      <NameValueList>\n'
    + '        <Name>Brand</Name>\n'
    + '        <Value>Shimano</Value>\n'
    + '      </NameValueList>\n'
    + '      <NameValueList>\n'
    + '        <Name>Type</Name>\n'
    + '        <Value>Spinning</Value>\n'
    + '      </NameValueList>\n'
    + '    </ItemSpecifics>\n'
    + '\n'
    + '    <Description><![CDATA[\n'
    + description + '\n'
    + '    ]]></Description>\n'
    + '\n'
    + '  </Item>\n'
    + '</VerifyAddFixedPriceItemRequest>';
}

// ═══════════════════════════════════════════════════════════
// 内部: Access Token 取得
// ═══════════════════════════════════════════════════════════

/**
 * User Access Token を取得する。
 * 既存の getEbayAccessTokenFromRefreshToken() を優先使用する。
 * 存在しない場合はフォールバック処理を使用する。
 * @returns {string|null}
 */
function _getAccessToken_() {
  // ── 既存共通関数を優先 ──
  // 同一プロジェクトに getEbayAccessTokenFromRefreshToken() があれば
  // 以下のコメントアウトを解除してください。
  // try {
  //   var t = getEbayAccessTokenFromRefreshToken();
  //   if (t) return t;
  // } catch (ignore) {}

  // ── フォールバック ──
  try {
    var props     = PropertiesService.getScriptProperties();
    var clientId  = sampleCsvProps_().getProperty('EBAY_CLIENT_ID');
    var clientSec = sampleCsvProps_().getProperty('EBAY_CLIENT_SECRET');
    var refreshTk = sampleCsvProps_().getProperty('EBAY_REFRESH_TOKEN');

    if (!clientId || !clientSec || !refreshTk) {
      Logger.log('❌ Script Properties未設定: EBAY_CLIENT_ID / EBAY_CLIENT_SECRET / EBAY_REFRESH_TOKEN');
      return null;
    }

    var resp = UrlFetchApp.fetch('https://api.ebay.com/identity/v1/oauth2/token', {
      method  : 'post',
      headers : {
        'Authorization' : 'Basic ' + Utilities.base64Encode(clientId + ':' + clientSec),
        'Content-Type'  : 'application/x-www-form-urlencoded',
      },
      payload : {
        grant_type    : 'refresh_token',
        refresh_token : refreshTk,
      },
      muteHttpExceptions: true,
    });

    var json = JSON.parse(resp.getContentText());
    if (!json.access_token) {
      Logger.log('❌ Token取得失敗: ' + JSON.stringify(json).substring(0, 300));
      return null;
    }
    return json.access_token;

  } catch (e) {
    Logger.log('❌ _getAccessToken_ 例外: ' + e.message);
    return null;
  }
}

// ═══════════════════════════════════════════════════════════
// 内部: ユーティリティ
// ═══════════════════════════════════════════════════════════

/** XML特殊文字エスケープ */
function _esc_(str) {
  return String(str)
    .replace(/&/g,  '&amp;')
    .replace(/</g,  '&lt;')
    .replace(/>/g,  '&gt;')
    .replace(/"/g,  '&quot;')
    .replace(/'/g,  '&apos;');
}

/** XMLタグ値の単純抽出（最初の一致のみ）*/
function _extractXmlValue_(xml, tag) {
  var re    = new RegExp('<' + tag + '>([^<]*)</' + tag + '>');
  var match = xml.match(re);
  return match ? match[1].trim() : '';
}

/** JST現在日時文字列 */
function _nowJst_() {
  return Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy/MM/dd HH:mm:ss');
}function updateEbayRawFromTradingApi() {

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('eBayRaw');

  if (!sheet) {
    throw new Error('eBayRaw シートが見つかりません');
  }

  const token = getEbayAccessTokenFromRefreshToken();

  const xml =
    '<?xml version="1.0" encoding="utf-8"?>' +
    '<GetMyeBaySellingRequest xmlns="urn:ebay:apis:eBLBaseComponents">' +
    '<ActiveList>' +
    '<Include>true</Include>' +
    '<Pagination>' +
    '<EntriesPerPage>200</EntriesPerPage>' +
    '<PageNumber>1</PageNumber>' +
    '</Pagination>' +
    '</ActiveList>' +
    '</GetMyeBaySellingRequest>';

  const res = UrlFetchApp.fetch('https://api.ebay.com/ws/api.dll', {
    method: 'post',
    headers: {
      'X-EBAY-API-CALL-NAME': 'GetMyeBaySelling',
      'X-EBAY-API-COMPATIBILITY-LEVEL': '1199',
      'X-EBAY-API-SITEID': '0',
      'X-EBAY-API-IAF-TOKEN': token,
      'Content-Type': 'text/xml'
    },
    payload: xml,
    muteHttpExceptions: true
  });

  const body = res.getContentText();
  const items = body.match(/<Item>[\s\S]*?<\/Item>/g) || [];

  const output = [[
    'ItemID',
    'SKU',
    'Title',
    'Quantity',
    'QuantitySold',
    'AvailableQuantity'
  ]];

  items.forEach(function(item) {

    const itemId = (item.match(/<ItemID>([^<]+)<\/ItemID>/) || [])[1] || '';
    const sku = (item.match(/<SKU>([^<]+)<\/SKU>/) || [])[1] || '';
    const title = (item.match(/<Title>([^<]+)<\/Title>/) || [])[1] || '';

    const quantity = Number((item.match(/<Quantity>([^<]+)<\/Quantity>/) || [])[1] || 0);
    const quantitySold = Number((item.match(/<QuantitySold>([^<]+)<\/QuantitySold>/) || [])[1] || 0);

    const availableQuantity = Math.max(quantity - quantitySold, 0);

    output.push([
      itemId,
      sku,
      title,
      quantity,
      quantitySold,
      availableQuantity
    ]);
  });

  sheet.clearContents();
  sheet.getRange(1, 1, output.length, output[0].length).setValues(output);

  Logger.log('✅ eBayRaw 更新完了: ' + (output.length - 1) + '件');
}

// ============================================================
// eBay Access Token取得
// ============================================================

function getEbayAccessTokenFromRefreshToken() {

  const props =
    PropertiesService
      .getScriptProperties();

  const clientId =
    sampleCsvProps_().getProperty(
      'EBAY_CLIENT_ID'
    );

  const clientSecret =
    sampleCsvProps_().getProperty(
      'EBAY_CLIENT_SECRET'
    );

  const refreshToken =
    sampleCsvProps_().getProperty(
      'EBAY_REFRESH_TOKEN'
    );

  const basicAuth =
    Utilities.base64Encode(
      clientId + ':' + clientSecret
    );

  const res =
    UrlFetchApp.fetch(
      'https://api.ebay.com/identity/v1/oauth2/token',
      {

        method: 'post',

        headers: {

          Authorization:
            'Basic ' + basicAuth,

          'Content-Type':
            'application/x-www-form-urlencoded'
        },

        payload: {

          grant_type:
            'refresh_token',

          refresh_token:
            refreshToken,

          scope:
            'https://api.ebay.com/oauth/api_scope'
        },

        muteHttpExceptions: true
      }
    );

  const text =
    res.getContentText();

  const data =
    JSON.parse(text);

  if (!data.access_token) {

    throw new Error(
      'Access Token取得失敗: ' +
      text
    );
  }

  return data.access_token;
}

/**
 * eBayRaw → 在庫同期管理 更新
 *
 * eBayRaw シートのSKU / ItemID / Title / AvailableQuantity を読み取り、
 * 在庫同期管理シートへ反映する。
 *
 * 既存の「通知基準数」はSKU単位で保持する。
 */
function updateStockSyncFromEbayRaw() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const ebaySheet = ss.getSheetByName('eBayRaw');
  const syncSheet = ss.getSheetByName('在庫同期管理');

  if (!ebaySheet || !syncSheet) {
    Logger.log('❌ eBayRaw または 在庫同期管理 が見つかりません');
    return;
  }

  // 既存の通知基準数を保存
  const existingData = syncSheet.getDataRange().getValues();
  const warningMap = {};

  for (let i = 1; i < existingData.length; i++) {
    const existingSku = normalizeSku_(existingData[i][0]);
    const warningQty = existingData[i][5];

    if (existingSku) {
      warningMap[existingSku] = warningQty;
    }
  }

  const ebayData = ebaySheet.getDataRange().getValues();

  if (ebayData.length < 2) {
    Logger.log('⚠️ eBayRaw にデータ行がありません');
    return;
  }

  const headers = ebayData[0].map(h => String(h).trim());

  const itemIdIndex = headers.indexOf('ItemID');
  const skuIndex = headers.indexOf('SKU');
  const titleIndex = headers.indexOf('Title');
  const availableQtyIndex = headers.indexOf('AvailableQuantity');

  if (skuIndex === -1) {
    Logger.log('❌ eBayRaw に SKU ヘッダーが見つかりません');
    Logger.log('ヘッダー一覧: ' + headers.join(' | '));
    return;
  }

  const skipSkus = [
    'WAREHOUSES',
    'DELETED',
    'SKU',
    'ITEMID',
    'TITLE',
    'AVAILABLEQUANTITY'
  ];

  const output = [];

  output.push([
    'SKU',
    'ItemID',
    'タイトル',
    'Amazon在庫',
    'eBay設定在庫',
    '通知基準数',
    '通知状態',
    '最終確認日時',
    '確認結果'
  ]);

  for (let i = 1; i < ebayData.length; i++) {
    const row = ebayData[i];

    const sku = normalizeSku_(row[skuIndex]);

    if (!sku) continue;
    if (skipSkus.includes(sku)) continue;

    const itemId =
      itemIdIndex !== -1
        ? row[itemIdIndex]
        : '';

    const title =
      titleIndex !== -1
        ? row[titleIndex]
        : '';

    const availableQty =
      availableQtyIndex !== -1
        ? row[availableQtyIndex]
        : '';

    const warningQty =
      warningMap[sku] !== undefined &&
      warningMap[sku] !== ''
        ? warningMap[sku]
        : 3;

    output.push([
      sku,
      itemId,
      title,
      '',
      availableQty,
      warningQty,
      '',
      '',
      'eBayRaw取込済'
    ]);
  }

  syncSheet.clearContents();

  syncSheet
    .getRange(1, 1, output.length, output[0].length)
    .setValues(output);

  Logger.log('✅ eBayRaw → 在庫同期管理 更新完了');
  Logger.log('反映行数: ' + (output.length - 1));
}

/**
 * SKU正規化 共通関数
 *
 * - 前後空白削除
 * - 大文字化
 * - 末尾 FBA 表記除去
 * - 全角ハイフン類を半角ハイフンへ統一
 * - 空白除去
 * - 連続ハイフン整理
 */
function normalizeSku_(value) {
  return String(value || '')
    .trim()
    .toUpperCase()
    .replace(/\s*-?\s*FBA$/i, '')
    .replace(/[－ー―]/g, '-')
    .replace(/\s+/g, '')
    .replace(/-+/g, '-')
    .trim();
}



/**
 * 在庫数・基準数を安全に数値化する
 */
function toInventoryNumberForZeroCheck_(value) {
  if (
    value === '' ||
    value === null ||
    value === undefined
  ) {
    return null;
  }

  const number = Number(value);

  return Number.isFinite(number)
    ? number
    : null;
}// ============================================================
// eBay受注検知 → LINE即時通知 → 12時間後MCF確認通知 → MCF自動作成
// 統合スクリプト（重複チェック機能付き）
// ============================================================

const SS_ID = '1exGBAEx99-2Qc9d0DLRiZbygvgIo4FY_NoBG3Nkan-A';
const ORDER_SHEET = 'eBay受注ログ';



// ============================================================
// 【新規追加】注文IDが既に Sheets に記録されているかチェック
// ============================================================
function isOrderAlreadyLogged(orderId) {
  if (!orderId) return false;

  try {
    const ss = SpreadsheetApp.openById(SS_ID);
    const sheet = ss.getSheetByName(ORDER_SHEET);
    
    if (!sheet) return false;

    const data = sheet.getDataRange().getValues();

    // B列（注文ID）をチェック
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][1]).trim() === String(orderId).trim()) {
        Logger.log('✅ 既出注文を検知: ' + orderId);
        return true;  // 既に記録されている
      }
    }

    return false;  // 新規注文

  } catch (e) {
    Logger.log('❌ 重複チェックエラー: ' + e.message);
    return false;
  }
}



// ============================================================
// 現在のeBay出品SKUをItem numberから取得
// Trading API GetItem
// ============================================================
function getCurrentListingSkuByItemId_(itemId, accessToken) {
  try {
    const xml =
      '<?xml version="1.0" encoding="UTF-8"?>' +
      '<GetItemRequest xmlns="urn:ebay:apis:eBLBaseComponents">' +
      '<ItemID>' + itemId + '</ItemID>' +
      '<DetailLevel>ReturnAll</DetailLevel>' +
      '</GetItemRequest>';

    const res = UrlFetchApp.fetch(
      'https://api.ebay.com/ws/api.dll',
      {
        method: 'post',
        headers: {
          'X-EBAY-API-IAF-TOKEN': accessToken,
          'X-EBAY-API-CALL-NAME': 'GetItem',
          'X-EBAY-API-SITEID': '0',
          'X-EBAY-API-COMPATIBILITY-LEVEL': '1199',
          'Content-Type': 'text/xml'
        },
        payload: xml,
        muteHttpExceptions: true
      }
    );

    const body = res.getContentText();

    const ack =
      (body.match(/<Ack>([^<]+)<\/Ack>/) || [])[1] || '';

    if (ack !== 'Success' && ack !== 'Warning') {
      Logger.log(
        '⚠️ GetItem失敗 ItemID=' +
        itemId +
        ' body=' +
        body.substring(0, 500)
      );

      return '';
    }

    const sku =
      (body.match(/<SKU>([\s\S]*?)<\/SKU>/) || [])[1] || '';

    return decodeXml_(sku).trim();

  } catch (e) {
    Logger.log('⚠️ GetItem例外: ' + e.message);
    return '';
  }
}

// ============================================================
// REST itemId から数値Item number抽出
// 例: v1|366282457520|0 → 366282457520
// ============================================================
function extractEbayItemId_(value) {
  const text = String(value || '');

  const match = text.match(/\d{10,}/);

  return match ? match[0] : '';
}

// ============================================================
// XMLエスケープ解除
// ============================================================
function decodeXml_(value) {
  return String(value || '')
    .replace(/&amp;/g, '&')
    .replace(/&lt;/g, '<')
    .replace(/&gt;/g, '>')
    .replace(/&quot;/g, '"')
    .replace(/&#39;/g, "'");
}

// ============================================================
// eBayメール解析
// ============================================================
function parseEbayEmail(msg) {
  const subject = msg.getSubject();
  const body = msg.getPlainBody();

  if (!subject.match(/you made the sale|order confirmed|you sold/i)) {
    return null;
  }

  const order = {
    msgId: msg.getId(),
    receivedAt: msg.getDate(),
    subject: subject,
    orderId: '',
    itemId: '',
    itemTitle: subject
      .replace(/you made the sale for\s*/i, '')
      .trim()
      .substring(0, 80),
    sku: '',
    buyerName: '',
    phone: '',
    phone2: '',
    email: '',
    street: '',
    city: '',
    state: '',
    zip: '',
    address: '',
    shipBy: '',
    quantity: '',
    amount: ''
  };

  const m1 = body.match(/(\d{2}-\d{5}-\d{5})/);
  if (m1) order.orderId = m1[1].trim();

  const m2 = body.match(/Your buyer's shipping details:\s*\n\s*([^\n]+)/);
  if (m2) order.buyerName = m2[1].trim().substring(0, 40);

  const m3 = body.match(
    /Your buyer's shipping details:\s*\n\s*[^\n]+\n([\s\S]*?United States)/
  );

  if (m3) {
    const raw =
      m3[1].replace(/\n/g, ' ').trim();

    order.address =
      raw.substring(0, 120);

    const p =
      raw.match(/^(.+?)\s{2,}([^,]+),\s*([A-Z]{2})\s+([\d-]+)/);

    if (p) {
      order.street = p[1].trim();
      order.city = p[2].trim();
      order.state = p[3].trim();
      order.zip = p[4].trim();
    }
  }

  const m4 = body.match(/Ship by:\s*\n?\s*([A-Za-z]+ \d+, \d+)/);
  if (m4) order.shipBy = m4[1].trim();

  const m5 = body.match(/Qty[:\s]*(\d+)/i);
  if (m5) order.quantity = m5[1];

  const m6 = body.match(/(?:Total|Amount)[:\s]*\$?([\d,.]+)/i);
  if (m6) order.amount = '$' + m6[1];

  Logger.log(
    '受注解析: ' +
    order.orderId +
    ' / ' +
    order.itemTitle.substring(0, 40)
  );

  return order;
}

// ============================================================
// LINE 即時通知
// ============================================================
function notifyOrderReceived(order) {
  const jst =
    new Date(order.receivedAt.getTime() + 9 * 60 * 60 * 1000);

  const timeStr =
    Utilities.formatDate(jst, 'Asia/Tokyo', 'MM/dd HH:mm');

  const msg =
    '🛒 eBay 受注通知\n━━━━━━━━━━━━━━━\n' +
    '🕐 ' + timeStr + ' JST\n' +
    '📦 ' + (order.itemTitle || '(商品名)') + '\n' +
    (order.orderId ? '🔖 注文ID: ' + order.orderId + '_eBay\n' : '') +
    (order.itemId ? '🆔 Item ID: ' + order.itemId + '\n' : '') +
    (order.sku ? '🏷 SKU: ' + order.sku + '\n' : '') +
    (order.buyerName ? '👤 購入者: ' + order.buyerName + '\n' : '') +
    (order.address ? '📍 住所: ' + order.address + '\n' : '') +
    (order.phone ? '📞 電話: ' + order.phone + '\n' : '') +
    (order.email ? '✉️ メール: ' + order.email + '\n' : '') +
    (order.shipBy ? '📅 発送期限: ' + order.shipBy + '\n' : '') +
    (order.quantity ? '📊 数量: ' + order.quantity + '\n' : '') +
    (order.amount ? '💰 金額: ' + order.amount + '\n' : '') +
    '\n⏳ 12時間後にMCF自動作成します';

  sendLine(msg);
}

// ============================================================
// 12時間後トリガーをセット
// ============================================================
function scheduleMcfReminder(order) {
  const props = PropertiesService.getScriptProperties();

  props.setProperty(
    'MCF_ORDER_' + (order.orderId || order.msgId),
    JSON.stringify({
      orderId: order.orderId,
      itemId: order.itemId,
      itemTitle: order.itemTitle,
      sku: order.sku,
      quantity: order.quantity,
      buyerName: order.buyerName,
      phone: order.phone,
      phone2: order.phone2 || '',
      email: order.email,
      street: order.street,
      city: order.city,
      state: order.state,
      zip: order.zip,
      address: order.address,
      shipBy: order.shipBy,
      amount: order.amount
    })
  );

  ScriptApp
    .newTrigger('sendMcfReminder')
    .timeBased()
    .at(new Date(Date.now() + 12 * 60 * 60 * 1000))
    .create();

  Logger.log('✅ MCFリマインダートリガーセット');
}

// ============================================================
// 12時間後 MCF確認通知 + MCF自動作成
// ============================================================
function sendMcfReminder() {
  const props = PropertiesService.getScriptProperties();

  const mcfOrders =
    Object.entries(props.getProperties())
      .filter(function(kv) {
        return kv[0].indexOf('MCF_ORDER_') === 0;
      });

  if (mcfOrders.length === 0) {
    Logger.log('MCF対象注文なし');
    return;
  }

  mcfOrders.forEach(function(kv) {
    const o = JSON.parse(kv[1]);

    const msg =
      '📬 Amazon MCF 作成確認\n━━━━━━━━━━━━━━━\n' +
      '📦 ' + (o.itemTitle || '(商品名)') + '\n' +
      (o.orderId ? '🔖 Order ID: ' + o.orderId + '_eBay\n' : '') +
      (o.itemId ? '🆔 Item ID: ' + o.itemId + '\n' : '') +
      (o.sku ? '🏷 SKU: ' + o.sku + '\n' : '') +
      (o.buyerName ? '👤 購入者: ' + o.buyerName + '\n' : '') +
      (o.street ? '🏠 住所: ' + o.street + '\n' : '') +
      (o.city ? '🏙 市: ' + o.city + '\n' : '') +
      (o.state ? '📌 州: ' + o.state + '\n' : '') +
      (o.zip ? '📮 ZIP: ' + o.zip + '\n' : '') +
      (o.phone ? '📞 電話: ' + o.phone + '\n' : '') +
      (o.email ? '✉️ メール: ' + o.email + '\n' : '') +
      (o.shipBy ? '📅 発送期限: ' + o.shipBy + '\n' : '') +
      (o.quantity ? '📊 数量: ' + o.quantity + '\n' : '') +
      '\n🚀 MCF自動作成を実行します';

    sendLine(msg);

    Logger.log(
      '✅ MCFリマインダー送信: ' +
      (o.orderId || o.itemTitle)
    );

    props.deleteProperty(kv[0]);

    updateMcfStatus(o.orderId);
  });

  Logger.log('🚀 MCF自動作成を開始します');

  autoCreateMcfOrders();

  ScriptApp.getProjectTriggers().forEach(function(t) {
    if (t.getHandlerFunction() === 'sendMcfReminder') {
      ScriptApp.deleteTrigger(t);
    }
  });
}



// ============================================================
// MCFステータス更新
// ============================================================
function updateMcfStatus(orderId) {
  if (!orderId) return;

  try {
    const ss =
      SpreadsheetApp.openById(SS_ID);

    const sheet =
      ss.getSheetByName(ORDER_SHEET);

    if (!sheet) return;

    const data =
      sheet.getDataRange().getValues();

    for (let i = 1; i < data.length; i++) {
      if (data[i][1] === orderId) {
        sheet.getRange(i + 1, 17).setValue('確認通知済み');
        break;
      }
    }

  } catch (e) {
    Logger.log('❌ MCFステータス更新エラー: ' + e.message);
  }
}

// ============================================================
// LINE送信
// ============================================================
function sendLine(message) {
  const props = PropertiesService.getScriptProperties();

  const token =
    sampleCsvProps_().getProperty('LINE_CHANNEL_TOKEN');

  const userId =
    sampleCsvProps_().getProperty('LINE_USER_ID');

  if (!token || !userId) {
    Logger.log('❌ LINE設定なし');
    return;
  }

  UrlFetchApp.fetch(
    'https://api.line.me/v2/bot/message/push',
    {
      method: 'post',
      headers: {
        'Authorization': 'Bearer ' + token,
        'Content-Type': 'application/json'
      },
      payload: JSON.stringify({
        to: userId,
        messages: [{
          type: 'text',
          text: message
        }]
      }),
      muteHttpExceptions: true
    }
  );
}

// ============================================================
// SP-API アクセストークン取得
// ============================================================
function getSpApiAccessToken() {

  const props = PropertiesService.getScriptProperties();

  try {

    const response = UrlFetchApp.fetch(
      'https://api.amazon.com/auth/o2/token',
      {
        method: 'post',
        contentType: 'application/x-www-form-urlencoded',
        payload: {
          grant_type: 'refresh_token',
          refresh_token: sampleCsvProps_().getProperty('SP_REFRESH_TOKEN'),
          client_id: sampleCsvProps_().getProperty('AMAZON_CLIENT_ID'),
          client_secret: sampleCsvProps_().getProperty('AMAZON_CLIENT_SECRET')
        },
        muteHttpExceptions: true
      }
    );

    const json = JSON.parse(response.getContentText());

    if (json.access_token) {

      Logger.log('✅ SP-APIトークン取得成功');
      return json.access_token;
    }

    Logger.log('❌ トークンエラー: ' + response.getContentText());

    return null;

  } catch (e) {

    Logger.log('❌ トークン取得例外: ' + e.message);

    return null;
  }
}

// ============================================================
// MCF追跡番号チェック
// ============================================================
function checkMcfTrackingNumbers() {

  const ss = SpreadsheetApp.openById(SS_ID);
  const sheet = ss.getSheetByName(ORDER_SHEET);
  const data = sheet.getDataRange().getValues();

  const accessToken = getMcfSpApiAccessToken_();

  if (!accessToken) {
    Logger.log('❌ アクセストークン取得失敗');
    return;
  }

  let checkedCount = 0;
  let foundCount = 0;

  for (let i = 1; i < data.length; i++) {

    const row = data[i];

    const mcfValue = row[15];           // P列 MCF作成
    const existingTracking = row[18];   // S列 Tracking
    const trackingStatus = row[20];     // U列 Tracking確認ステータス

    if (!mcfValue || !String(mcfValue).startsWith('MCF:')) continue;

    // U列「チェック不要」はスキップ
    if (trackingStatus === 'チェック不要') continue;

    // 既にTracking取得済みならスキップ
    if (existingTracking) continue;

    const fulfillmentOrderId =
      String(mcfValue).replace('MCF:', '').trim();

    Logger.log('🔍 MCF追跡確認: ' + fulfillmentOrderId);

    const result = getMcfTrackingByFulfillmentOrderId(
      accessToken,
      fulfillmentOrderId
    );

    checkedCount++;

    if (result.success && result.trackingNumber) {

      // R列 Carrier
      sheet.getRange(i + 1, 18)
        .setValue(result.carrierCode || '');

      // S列 Tracking
      sheet.getRange(i + 1, 19)
        .setValue(result.trackingNumber);

      // T列 Tracking取得日時
      sheet.getRange(i + 1, 20)
        .setValue(new Date());

      // U列 Tracking確認ステータス
      sheet.getRange(i + 1, 21)
        .setValue('追跡番号取得済');

      foundCount++;

      sendMcfTrackingNotification(
        row[1],
        fulfillmentOrderId,
        result.carrierCode,
        result.trackingNumber
      );

      Logger.log('✅ 追跡番号取得: ' + result.trackingNumber);

    } else if (result.success) {

      sheet.getRange(i + 1, 21)
        .setValue(result.status || '追跡待ち');

      Logger.log('⏳ ' + (result.status || '追跡番号未発行'));

    } else {

      sheet.getRange(i + 1, 21)
        .setValue('確認エラー:' + result.error);

      Logger.log('❌ 確認エラー: ' + result.error);
    }

    Utilities.sleep(1000);
  }

  Logger.log(
    '=== MCF追跡確認完了 checked:' +
    checkedCount +
    ' found:' +
    foundCount +
    ' ==='
  );
}

// ============================================================
// MCF Order → packageNumber → Tracking取得
// ============================================================
function getMcfTrackingByFulfillmentOrderId(
  accessToken,
  fulfillmentOrderId
) {

  const orderResult =
    getMcfFulfillmentOrderRaw(
      accessToken,
      fulfillmentOrderId
    );

  if (!orderResult.success) {
    return orderResult;
  }

  const json = orderResult.json;

  // trackingNumberを直接探す
  const directTracking = findTrackingInfo_(json);

  if (directTracking && directTracking.trackingNumber) {

    return {
      success: true,
      carrierCode:
        directTracking.carrierCode ||
        directTracking.carrierName ||
        '',
      trackingNumber:
        directTracking.trackingNumber
    };
  }

  // packageNumber を探す
  const packageNumbers =
    findPackageNumbers_(json);

  if (!packageNumbers.length) {

    return {
      success: true,
      status: '追跡待ち（packageNumber未発行）'
    };
  }

  Logger.log(
    '📦 packageNumber検出: ' +
    packageNumbers.join(', ')
  );

  for (let i = 0; i < packageNumbers.length; i++) {

    const packageNumber =
      packageNumbers[i];

    const trackingResult =
      getPackageTrackingDetails(
        accessToken,
        packageNumber
      );

    if (
      trackingResult.success &&
      trackingResult.trackingNumber
    ) {

      return trackingResult;
    }

    Utilities.sleep(500);
  }

  return {
    success: true,
    status:
      '追跡待ち（packageNumberあり・tracking未発行）'
  };
}

// ============================================================
// Amazon SP-API: getFulfillmentOrder 生データ取得
// ============================================================
function getMcfFulfillmentOrderRaw(
  accessToken,
  fulfillmentOrderId
) {

  const url =
    'https://sellingpartnerapi-na.amazon.com/fba/outbound/2020-07-01/fulfillmentOrders/' +
    encodeURIComponent(
      fulfillmentOrderId
    );

  try {

    const response = UrlFetchApp.fetch(
      url,
      {
        method: 'get',
        headers: {
          'Authorization':
            'Bearer ' + accessToken,
          'x-amz-access-token':
            accessToken,
          'Content-Type':
            'application/json'
        },
        muteHttpExceptions: true
      }
    );

    const status =
      response.getResponseCode();

    const text =
      response.getContentText();

    Logger.log(
      'MCF GET [' +
      status +
      ']: ' +
      text.substring(0, 1500)
    );

    if (status !== 200) {

      return {
        success: false,
        error:
          status +
          ':' +
          text.substring(0, 300)
      };
    }

    return {
      success: true,
      json: JSON.parse(text)
    };

  } catch (e) {

    return {
      success: false,
      error: e.message
    };
  }
}

// ============================================================
// Amazon SP-API: getPackageTrackingDetails
// ============================================================
function getPackageTrackingDetails(
  accessToken,
  packageNumber
) {

  const url =
    'https://sellingpartnerapi-na.amazon.com/fba/outbound/2020-07-01/tracking' +
    '?packageNumber=' +
    encodeURIComponent(packageNumber);

  try {

    const response = UrlFetchApp.fetch(
      url,
      {
        method: 'get',
        headers: {
          'Authorization':
            'Bearer ' + accessToken,
          'x-amz-access-token':
            accessToken,
          'Content-Type':
            'application/json'
        },
        muteHttpExceptions: true
      }
    );

    const status =
      response.getResponseCode();

    const text =
      response.getContentText();

    Logger.log(
      'TRACKING GET [' +
      status +
      '] packageNumber=' +
      packageNumber +
      ': ' +
      text.substring(0, 1500)
    );

    if (status !== 200) {

      return {
        success: false,
        error:
          status +
          ':' +
          text.substring(0, 300)
      };
    }

    const json = JSON.parse(text);

    const found =
      findTrackingInfo_(json);

    if (
      found &&
      found.trackingNumber
    ) {

      return {
        success: true,
        carrierCode:
          found.carrierCode ||
          found.carrierName ||
          '',
        trackingNumber:
          found.trackingNumber
      };
    }

    return {
      success: true,
      trackingNumber: ''
    };

  } catch (e) {

    return {
      success: false,
      error: e.message
    };
  }
}

// ============================================================
// JSON全体から trackingNumber を探す
// ============================================================
function findTrackingInfo_(obj) {

  if (!obj || typeof obj !== 'object') return null;

  if (obj.trackingNumber) {

    return {
      trackingNumber:
        obj.trackingNumber,
      carrierCode:
        obj.carrierCode || '',
      carrierName:
        obj.carrierName || ''
    };
  }

  for (const key in obj) {

    const result =
      findTrackingInfo_(obj[key]);

    if (result) return result;
  }

  return null;
}

// ============================================================
// JSON全体から packageNumber を探す
// ============================================================
function findPackageNumbers_(obj) {

  const results = [];

  function walk(value) {

    if (!value || typeof value !== 'object') return;

    if (value.packageNumber) {
      results.push(
        String(value.packageNumber)
      );
    }

    for (const key in value) {
      walk(value[key]);
    }
  }

  walk(obj);

  return [...new Set(results)];
}

// ============================================================
// LINE通知: MCF追跡番号取得
// ============================================================
function sendMcfTrackingNotification(
  ebayOrderId,
  fulfillmentOrderId,
  carrier,
  trackingNumber
) {

  const props =
    PropertiesService.getScriptProperties();

  const lineToken =
    sampleCsvProps_().getProperty(
      'LINE_CHANNEL_TOKEN'
    );

  const lineUserId =
    sampleCsvProps_().getProperty(
      'LINE_USER_ID'
    );

  if (!lineToken || !lineUserId) return;

  UrlFetchApp.fetch(
    'https://api.line.me/v2/bot/message/push',
    {
      method: 'post',
      headers: {
        'Authorization':
          'Bearer ' + lineToken,
        'Content-Type':
          'application/json'
      },

      payload: JSON.stringify({

        to: lineUserId,

        messages: [{

          type: 'text',

          text:
            '📦 MCF追跡番号取得\n' +
            '━━━━━━━━━━━━\n' +

            'eBay注文: ' +
            ebayOrderId +
            '\n' +

            'MCF ID: ' +
            fulfillmentOrderId +
            '\n' +

            '配送会社: ' +
            (carrier || '-') +
            '\n' +

            '追跡番号: ' +
            trackingNumber +
            '\n\n' +

            'eBay Orders:\n' +
            'https://www.ebay.com/sh/ord/?filter=status%3AALL_ORDERS'

        }]
      }),

      muteHttpExceptions: true
    }
  );
}

// ============================================================
// MCF自動作成
// eBay受注ログのP列が空白または「未作成」の行だけ作成
// 追加仕様：受注から12時間未満は WAIT_12H として作成しない
// ============================================================
function autoCreateMcfOrders() {

  const ss = SpreadsheetApp.openById(SS_ID);
  const orderSheet = ss.getSheetByName(ORDER_SHEET);

  if (!orderSheet) {
    Logger.log('❌ eBay受注ログ シート未発見');
    return;
  }

  const data = orderSheet.getDataRange().getValues();

  const accessToken = getMcfSpApiAccessToken_();

  if (!accessToken) {
    Logger.log('❌ アクセストークン取得失敗');
    return;
  }

  let createdCount = 0;
  let errorCount = 0;
  let waitCount = 0;
  let skipCount = 0;

  const skipStatuses = [
    'MCF作成済',
    '手動対応済み',
    '対象外',
    '対応不要',
    'キャンセル',
    '返金済み',
    '手動確認',
    '確認通知済み'
  ];

  for (let i = 1; i < data.length; i++) {

    const row = data[i];

    const rowNumber = i + 1;
    const mcfStatus = String(row[15] || '').trim(); // P列 MCF作成
    const qStatus = String(row[16] || '').trim();   // Q列 ステータス
    const ebayOrderId = String(row[1] || '').trim(); // B列

    // P列が空白 or 未作成 のみ対象
    if (mcfStatus && mcfStatus !== '未作成') {
      Logger.log('⏭ [SKIP_P] 行' + rowNumber + ' ' + ebayOrderId + ' P列=' + mcfStatus);
      skipCount++;
      continue;
    }

    // Q列が処理済み・手動対応系ならスキップ
    if (skipStatuses.some(function(status) { return qStatus.indexOf(status) !== -1; })) {
      Logger.log('⏭ [SKIP_Q] 行' + rowNumber + ' ' + ebayOrderId + ' Q列=' + qStatus);
      skipCount++;
      continue;
    }

    // 受注から12時間未満はMCF作成しない
    const delayCheck = shouldCreateMcfNow_(row[0]); // A列：受注日時

    if (!delayCheck.ready) {
      if (String(delayCheck.reason || '').indexOf('ERROR') === 0) {
        Logger.log('❌ [ORDER_DATE_ERROR] 行' + rowNumber + ' ' + ebayOrderId + ' ' + delayCheck.reason);
        errorCount++;
      } else {
        Logger.log('⏳ [WAIT_12H] 行' + rowNumber + ' ' + ebayOrderId + ' ' + delayCheck.reason);
        waitCount++;
      }
      // WAIT_12H / 日時エラーでは、P列/Q列は書き換えない
      continue;
    }

    const sku = String(row[2] || '').trim();         // C列
    const qty = Number(row[4]) || 1;                 // E列

    const buyerName = String(row[6] || '').trim();   // G列
    const street = String(row[7] || '').trim();      // H列
    const city = String(row[8] || '').trim();        // I列
    const state = String(row[9] || '').trim();       // J列
    const zip = String(row[10] || '').trim();        // K列
    const phone = String(row[11] || '').trim();      // L列
    const phone2 = String(row[12] || '').trim();     // M列
    const email = String(row[13] || '').trim();      // N列

    if (!ebayOrderId || !sku || !street || !city || !state || !zip) {

      orderSheet
        .getRange(rowNumber, 17) // Q列
        .setValue('情報不足');

      Logger.log(
        '⚠️ 行' +
        rowNumber +
        ' 必須情報不足'
      );

      errorCount++;
      continue;
    }

    const skuInfo =
      getSkuInfoBySku(
        ss,
        sku
      );

    if (!skuInfo) {

      orderSheet
        .getRange(rowNumber, 16) // P列
        .setValue('SKU未発見');

      orderSheet
        .getRange(rowNumber, 17) // Q列
        .setValue('MCF未作成');

      Logger.log(
        '⚠️ SKU未発見: ' +
        sku
      );

      errorCount++;
      continue;
    }

    const sellerOrderId =
      'EBAY-' +
      ebayOrderId +
      '-' +
      new Date()
        .getTime()
        .toString()
        .slice(-4);

    Logger.log(
      '🚀 MCF作成開始: ' +
      sellerOrderId +
      ' SKU=' +
      skuInfo.sellerSku
    );

    const result =
      createMcfOrder(
        accessToken,
        {
          sellerOrderId: sellerOrderId,
          sellerSku: skuInfo.sellerSku,
          asin: skuInfo.asin,
          quantity: qty,
          buyerName: buyerName || 'eBay Customer',
          street: street,
          city: city,
          state: state,
          zip: zip,
          phone: phone || phone2 || '',
          email: email,
          shippingSpeed: 'Standard'
        }
      );

    if (result.success) {

      orderSheet
        .getRange(rowNumber, 16) // P列
        .setValue(
          'MCF:' +
          result.fulfillmentOrderId
        );

      orderSheet
        .getRange(rowNumber, 17) // Q列
        .setValue('MCF作成済');

      // MCF作成成功のみのLINE通知は不要のため停止
      // sendMcfCreatedNotification(
      //   ebayOrderId,
      //   skuInfo.asin,
      //   result.fulfillmentOrderId
      // );

      Logger.log(
        '✅ MCF作成成功: ' +
        result.fulfillmentOrderId
      );

      createdCount++;

    } else {

      orderSheet
        .getRange(rowNumber, 17) // Q列
        .setValue(
          '作成エラー:' +
          result.error
        );

      Logger.log(
        '❌ MCF作成失敗: ' +
        result.error
      );

      errorCount++;
    }

    Utilities.sleep(1000);
  }

  Logger.log(
    '=== MCF作成完了 成功:' +
    createdCount +
    ' エラー:' +
    errorCount +
    ' 待機:' +
    waitCount +
    ' スキップ:' +
    skipCount +
    ' ==='
  );
}

// ============================================================
// FBA在庫_DL からSKU情報取得
// ============================================================
function getSkuInfoBySku(
  ss,
  sku
) {

  const stockSheet =
    ss.getSheetByName(
      'FBA在庫_DL'
    );

  if (!stockSheet) {

    Logger.log(
      '❌ FBA在庫_DL シート未発見'
    );

    return null;
  }

  const data =
    stockSheet
      .getDataRange()
      .getValues();

  const headers =
    data[0];

  const skuIndex =
    headers.indexOf('sku');

  const asinIndex =
    headers.indexOf('asin');

  if (
    skuIndex === -1 ||
    asinIndex === -1
  ) {

    Logger.log(
      '❌ FBA在庫_DL ヘッダー不一致'
    );

    return null;
  }

  const targetSku =
    normalizeMcfSku_(
      sku
    );

  for (let i = 1; i < data.length; i++) {

    const fbaSku =
      normalizeMcfSku_(
        data[i][skuIndex]
      );

    if (fbaSku === targetSku) {

      return {
        sellerSku:
          String(
            data[i][skuIndex] || ''
          ).trim(),

        asin:
          String(
            data[i][asinIndex] || ''
          ).trim()
      };
    }
  }

  return null;
}

// ============================================================
// SKU正規化
// ============================================================
function normalizeMcfSku_(value) {

  return String(value || '')
    .trim()
    .replace(
      /\s*-?\s*FBA$/i,
      ''
    )
    .trim();
}

// ============================================================
// Amazon SP-API: MCFオーダー作成
// Blank Box / Amazon Logistics除外 対応版
// POST版（正常動作版）
// ============================================================
function createMcfOrder(
  accessToken,
  params
) {

  const fulfillmentOrderId =
    params.sellerOrderId;

  const url =
    'https://sellingpartnerapi-na.amazon.com' +
    '/fba/outbound/2020-07-01/fulfillmentOrders';

  const body = {

    marketplaceId:
      'ATVPDKIKX0DER',

    sellerFulfillmentOrderId:
      fulfillmentOrderId,

    displayableOrderId:
      fulfillmentOrderId,

    displayableOrderDate:
      new Date().toISOString(),

    displayableOrderComment:
      'eBay order auto-fulfilled via MCF',

    shippingSpeedCategory:
      params.shippingSpeed || 'Standard',

    // ========================================================
    // Blank Box / Amazon Logistics除外
    // ========================================================
    featureConstraints: [

      {
        featureName:
          'BLANK_BOX',

        featureFulfillmentPolicy:
          'Required'
      },

      {
        featureName:
          'BLOCK_AMZL',

        featureFulfillmentPolicy:
          'Required'
      }
    ],

    notificationEmails:
      params.email
        ? [params.email]
        : [],

    destinationAddress: {

      name:
        params.buyerName,

      addressLine1:
        params.street,

      addressLine2:
        params.address2 || '',

      city:
        params.city,

      stateOrRegion:
        params.state,

      postalCode:
        params.zip,

      countryCode:
        'US',

      phone:
        params.phone || '',

      email:
        params.email || ''
    },

    items: [{

      sellerSku:
        params.sellerSku,

      sellerFulfillmentOrderItemId:
        '1',

      quantity:
        Number(params.quantity) || 1,

      displayableComment:
        'eBay MCF Auto'
    }]
  };

  Logger.log(
    'MCF CREATE BODY:\n' +
    JSON.stringify(body, null, 2)
  );

  try {

    const response =
      UrlFetchApp.fetch(
        url,
        {

          method: 'post',

          headers: {

            'Authorization':
              'Bearer ' + accessToken,

            'x-amz-access-token':
              accessToken,

            'Content-Type':
              'application/json'
          },

          payload:
            JSON.stringify(body),

          muteHttpExceptions:
            true
        }
      );

    const status =
      response.getResponseCode();

    const text =
      response.getContentText();

    Logger.log(
      'MCF CREATE [' +
      status +
      ']: ' +
      text.substring(0, 1200)
    );

    if (
      status === 200 ||
      status === 201 ||
      status === 202
    ) {

      return {
        success: true,
        fulfillmentOrderId:
          fulfillmentOrderId
      };
    }

    let message = text;

    try {

      const err =
        JSON.parse(text);

      message =
        err.errors
          ? err.errors[0].message
          : text;

    } catch (e) {}

    return {
      success: false,
      error:
        status +
        ':' +
        message
    };

  } catch (e) {

    return {
      success: false,
      error:
        e.message
    };
  }
}

// ============================================================
// LINE通知: MCF作成完了
// ============================================================
function sendMcfCreatedNotification(
  ebayOrderId,
  asin,
  fulfillmentOrderId
) {

  const props =
    PropertiesService
      .getScriptProperties();

  const lineToken =
    sampleCsvProps_().getProperty(
      'LINE_CHANNEL_TOKEN'
    );

  const lineUserId =
    sampleCsvProps_().getProperty(
      'LINE_USER_ID'
    );

  if (
    !lineToken ||
    !lineUserId
  ) return;

  UrlFetchApp.fetch(
    'https://api.line.me/v2/bot/message/push',
    {
      method:
        'post',

      headers: {

        'Authorization':
          'Bearer ' +
          lineToken,

        'Content-Type':
          'application/json'
      },

      payload:
        JSON.stringify({
          to:
            lineUserId,

          messages: [{
            type:
              'text',

            text:
              '🚚 MCF自動作成完了\n' +
              '━━━━━━━━━━━━\n' +
              'eBay注文: ' +
              ebayOrderId +
              '\n' +
              'ASIN: ' +
              asin +
              '\n' +
              'MCF ID: ' +
              fulfillmentOrderId +
              '\n' +
              '配送: Standard'
          }]
        }),

      muteHttpExceptions:
        true
    }
  );
}

// ============================================================
// テスト・デバッグ用
// ============================================================
function testLineNotification() {
  sendLine(
    '🧪 テスト通知\n━━━━━━━━━━━━━━━\neBay受注 → LINE通知\n動作確認OK ✅'
  );
}

function debugOrderMailDetail() {
  const threads =
    GmailApp.search(
      'from:ebay.com subject:"You made the sale"',
      0,
      3
    );

  Logger.log('受注メール件数: ' + threads.length);

  threads.forEach(function(t, i) {
    const msg = t.getMessages()[0];

    Logger.log(
      '[' +
      i +
      '] ' +
      msg.getSubject() +
      ' / 未読: ' +
      msg.isUnread()
    );
  });
}

function debugBuyerInfo() {
  const threads =
    GmailApp.search(
      'from:ebay.com subject:"You made the sale"',
      0,
      1
    );

  if (threads.length === 0) {
    Logger.log('メールなし');
    return;
  }

  const body =
    threads[0].getMessages()[0].getPlainBody();

  const m =
    body.match(/(\d{2}-\d{5}-\d{5})/);

  if (!m) {
    Logger.log('注文番号なし');
    return;
  }

  const detail =
    getOrderDetail(m[1]);

  Logger.log(
    '取得結果: ' +
    JSON.stringify(detail)
  );
}

function testAutoCreateMcfOrdersPrecheck() {

  const ss = SpreadsheetApp.openById(SS_ID);
  const orderSheet = ss.getSheetByName(ORDER_SHEET);

  if (!orderSheet) {
    Logger.log('❌ eBay受注ログ シート未発見');
    return;
  }

  const data = orderSheet.getDataRange().getValues();

  let targetCount = 0;
  let okCount = 0;
  let errorCount = 0;

  for (let i = 1; i < data.length; i++) {

    const row = data[i];

    const mcfStatus = String(row[15] || '').trim();

    if (mcfStatus && mcfStatus !== '未作成') continue;

    const ebayOrderId = String(row[1] || '').trim();
    const sku = String(row[2] || '').trim();
    const qty = Number(row[4]) || 1;

    const buyerName = String(row[6] || '').trim();
    const street = String(row[7] || '').trim();
    const city = String(row[8] || '').trim();
    const state = String(row[9] || '').trim();
    const zip = String(row[10] || '').trim();

    targetCount++;

    Logger.log('━━━━━━━━━━━━━━━━━━━━');
    Logger.log('対象行: ' + (i + 1));
    Logger.log('eBay注文: ' + ebayOrderId);
    Logger.log('SKU: ' + sku);
    Logger.log('数量: ' + qty);
    Logger.log('購入者: ' + buyerName);
    Logger.log('住所: ' + street + ', ' + city + ', ' + state + ' ' + zip);

    if (!ebayOrderId || !sku || !street || !city || !state || !zip) {
      Logger.log('❌ 必須情報不足');
      errorCount++;
      continue;
    }

    const skuInfo = getSkuInfoBySku(ss, sku);

    if (!skuInfo) {
      Logger.log('❌ SKU未発見: FBA在庫_DL に一致なし');
      errorCount++;
      continue;
    }

    Logger.log('✅ 作成可能');
    Logger.log('Amazon sellerSku: ' + skuInfo.sellerSku);
    Logger.log('ASIN: ' + skuInfo.asin);

    okCount++;
  }

  Logger.log('━━━━━━━━━━━━━━━━━━━━');
  Logger.log(
    '事前チェック完了 対象:' +
    targetCount +
    ' OK:' +
    okCount +
    ' エラー:' +
    errorCount
  );
}

// ============================================================
// トリガー設定
// ============================================================
function setupOrderCheckTrigger() {
  ScriptApp.getProjectTriggers().forEach(function(t) {
    if (t.getHandlerFunction() === 'checkEbayOrders') {
      ScriptApp.deleteTrigger(t);
    }
  });

  ScriptApp
    .newTrigger('checkEbayOrders')
    .timeBased()
    .everyMinutes(5)
    .create();

  Logger.log('✅ 受注チェックトリガー設定完了（5分おき）');
}

function showHelp() {
  const html = HtmlService.createHtmlOutput(`
    <h2>eBay 自動化 - ヘルプ</h2>
    <h3>メニュー説明</h3>
    <ul>
      <li><b>📋 受注チェック実行</b>: eBay 受注メールをチェック（1 時間ごと自動実行）</li>
      <li><b>📬 MCF 追跡番号チェック</b>: Amazon MCF の追跡番号を取得</li>
      <li><b>📤 eBay に追跡番号提出</b>: 取得した追跡番号を eBay に提出（手動実行）</li>
    </ul>
    <h3>使い方</h3>
    <ol>
      <li>Amazon MCF で追跡番号が取得されると Sheets に自動入力</li>
      <li>「📬 MCF 追跡番号チェック」を実行して追跡番号を確認</li>
      <li>「📤 eBay に追跡番号提出」を実行して eBay に提出</li>
      <li>eBay セラーセンターで「Shipped」ステータスに更新される</li>
    </ol>
  `);
  SpreadsheetApp.getUi().showModelessDialog(html, 'ヘルプ');
}

// ============================================================
// 追跡番号を eBay に提出（手動実行版）
// ============================================================
function autoSubmitTrackingToEbay() {
  try {
    const ss = SpreadsheetApp.openById(SS_ID);
    const sheet = ss.getSheetByName(ORDER_SHEET);
    const data = sheet.getDataRange().getValues();

    let submittedCount = 0;
    let skipCount = 0;
    let errorCount = 0;

    for (let i = 1; i < data.length; i++) {
      const row = data[i];

      const orderId = String(row[1] || '').trim();         // B列：注文ID
      const trackingNumber = String(row[18] || '').trim(); // S列：追跡番号
      const carrierCode = String(row[17] || '').trim();    // R列：キャリア
      const submitStatus = String(row[22] || '').trim();   // W列：提出ステータス

      // 必須条件チェック
      if (!orderId) {
        Logger.log('⏭️  行' + (i + 1) + '：注文ID なし');
        skipCount++;
        continue;
      }

      if (!trackingNumber) {
        Logger.log('⏭️  行' + (i + 1) + '：追跡番号 なし');
        skipCount++;
        continue;
      }

      if (submitStatus === 'eBay提出済み') {
        Logger.log('⏭️  行' + (i + 1) + '：既に提出済み');
        skipCount++;
        continue;
      }

      Logger.log('🚀 eBay提出開始: ' + orderId + ' / ' + trackingNumber);

      const result = submitTrackingToEbayApi_(orderId, trackingNumber, carrierCode);

      if (result.success) {
        sheet.getRange(i + 1, 23).setValue('eBay提出済み');      // W列
        sheet.getRange(i + 1, 24).setValue(new Date());          // X列：提出日時

        sendLine(
          '✅ eBay追跡番号提出成功\n' +
          '━━━━━━━━━━━━\n' +
          '注文ID: ' + orderId + '\n' +
          '追跡番号: ' + trackingNumber + '\n' +
          '配送会社: ' + (carrierCode || 'USPS')
        );

        Logger.log('✅ eBay提出成功: ' + trackingNumber);
        submittedCount++;

      } else {
        sheet.getRange(i + 1, 23).setValue('提出エラー:' + result.error.substring(0, 30));

        sendLine(
          '❌ eBay追跡番号提出失敗\n' +
          '━━━━━━━━━━━━\n' +
          '注文ID: ' + orderId + '\n' +
          '理由: ' + result.error.substring(0, 50)
        );

        Logger.log('❌ eBay提出失敗: ' + result.error);
        errorCount++;
      }

      Utilities.sleep(1500);  // API レート制限対策
    }

    const summary =
      '✅ eBay追跡番号提出完了\n' +
      '━━━━━━━━━━━━\n' +
      '✅ 成功: ' + submittedCount + '件\n' +
      '⏭️  スキップ: ' + skipCount + '件\n' +
      '❌ エラー: ' + errorCount + '件';

    Logger.log(summary);
    sendLine(summary);

    SpreadsheetApp.getUi().alert(summary);

  } catch (e) {
    Logger.log('❌ eBay提出処理エラー: ' + e.message);
    sendLine('❌ eBay提出処理エラー: ' + e.message);
    SpreadsheetApp.getUi().alert('❌ エラーが発生しました: ' + e.message);
  }
}

// ============================================================
// eBay REST API shipOrder 実行
// ============================================================
function submitTrackingToEbayApi_(orderId, trackingNumber, carrierCode) {
  try {
    const props = PropertiesService.getScriptProperties();

    // OAuth トークン取得
    const creds = Utilities.base64Encode(
      sampleCsvProps_().getProperty('EBAY_CLIENT_ID') +
      ':' +
      sampleCsvProps_().getProperty('EBAY_CLIENT_SECRET')
    );

    const tokenRes = UrlFetchApp.fetch(
      'https://api.ebay.com/identity/v1/oauth2/token',
      {
        method: 'post',
        headers: {
          'Authorization': 'Basic ' + creds,
          'Content-Type': 'application/x-www-form-urlencoded'
        },
        payload:
          'grant_type=refresh_token&refresh_token=' +
          encodeURIComponent(sampleCsvProps_().getProperty('EBAY_REFRESH_TOKEN')),
        muteHttpExceptions: true
      }
    );

    const tokenJson = JSON.parse(tokenRes.getContentText());
    const accessToken = tokenJson.access_token;

    if (!accessToken) {
      return {
        success: false,
        error: 'eBayトークン取得失敗: ' + tokenRes.getContentText()
      };
    }

    // shipOrder リクエスト
    const shipBody = {
      lineItems: [
        {
          lineItemId: orderId
        }
      ],
      shipmentCarrier: carrierCode || 'USPS',
      trackingNumber: trackingNumber
    };

    Logger.log('eBay shipOrder リクエスト: ' + JSON.stringify(shipBody));

    const shipRes = UrlFetchApp.fetch(
      'https://api.ebay.com/sell/fulfillment/v1/order/' + orderId + '/shipping/fulfillment',
      {
        method: 'post',
        headers: {
          'Authorization': 'Bearer ' + accessToken,
          'Content-Type': 'application/json'
        },
        payload: JSON.stringify(shipBody),
        muteHttpExceptions: true
      }
    );

    const shipStatus = shipRes.getResponseCode();
    const shipText = shipRes.getContentText();

    Logger.log('eBay shipOrder レスポンス [' + shipStatus + ']: ' + shipText.substring(0, 500));

    if (shipStatus === 200 || shipStatus === 201 || shipStatus === 204) {
      return {
        success: true,
        orderId: orderId,
        trackingNumber: trackingNumber
      };
    } else {
      return {
        success: false,
        error: shipStatus + ': ' + shipText.substring(0, 200)
      };
    }

  } catch (e) {
    return {
      success: false,
      error: 'eBay API エラー: ' + e.message
    };
  }
}

/**
 * 【修正版】SKU情報を複数シートから取得
 * 優先順：MCF_SKU_MAP → FBA在庫_API → FBA在庫_DL
 */
function getSkuInfoBySku(ss, sku) {
  if (!sku) return null;

  const skuNorm = String(sku).toUpperCase().trim();

  // ■１ MCF_SKU_MAP を優先参照
  const mcfSheet = ss.getSheetByName('MCF_SKU_MAP');
  if (mcfSheet) {
    const data = mcfSheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      const ebaySkuCell = String(data[i][0] || '').toUpperCase().trim();
      if (ebaySkuCell === skuNorm) {
        return {
          sellerSku: String(data[i][1] || ''),
          asin: String(data[i][2] || ''),
          source: 'MCF_SKU_MAP'
        };
      }
    }
  }

  // ■２ FBA在庫_API を参照（後で実装）
  // const fbaApiSheet = ss.getSheetByName('FBA在庫_API');
  // ...

  // ■３ FBA在庫_DL を参照（互換用）
  const fbaSheet = ss.getSheetByName('FBA在庫_DL');
  if (fbaSheet) {
    const data = fbaSheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      const cellSku = String(data[i][0] || '').toUpperCase().trim();
      if (cellSku === skuNorm) {
        return {
          sellerSku: String(data[i][0] || ''),
          asin: String(data[i][2] || ''),
          source: 'FBA在庫_DL'
        };
      }
    }
  }

  // ■４ 見つからない場合
  Logger.log('❌ SKU未発見：' + sku);
  return null;
}

function logEbayOrderDirect_(order) {
  try {
    if (!order || !order.orderId) {
      Logger.log('❌ logEbayOrderDirect_: orderIdなし');
      return false;
    }

    const ss = SpreadsheetApp.openById(SS_ID);
    let sheet = ss.getSheetByName(ORDER_SHEET);

    if (!sheet) {
      sheet = ss.insertSheet(ORDER_SHEET);

      const h = [
        '受注日時(JST)',
        'eBay注文番号',
        'SKU',
        '商品名',
        '数量',
        '金額',
        '購入者名',
        'Street',
        'City',
        'State',
        'ZIP',
        '電話番号',
        '電話番号2',
        'メールアドレス',
        '発送期限',
        'MCF作成',
        'ステータス',
        'Carrier',
        'Tracking',
        'Tracking取得日時',
        'Tracking確認ステータス'
      ];

      sheet.appendRow(h);
      sheet.getRange(1, 1, 1, h.length)
        .setFontWeight('bold')
        .setBackground('#1F4E79')
        .setFontColor('#FFFFFF');
      sheet.setFrozenRows(1);
    }

    const receivedAt = order.receivedAt instanceof Date
      ? order.receivedAt
      : new Date();

    Utilities.formatDate(receivedAt, 'Asia/Tokyo', 'yyyy-MM-dd HH:mm')

sheet.appendRow([
  Utilities.formatDate(receivedAt, 'Asia/Tokyo', 'yyyy-MM-dd HH:mm'),
      order.orderId || '',
      order.sku || '',
      order.itemTitle || '',
      order.quantity || '1',
      order.amount || '',
      order.buyerName || '',
      order.street || '',
      order.city || '',
      order.state || '',
      order.zip || '',
      order.phone || '',
      order.phone2 || '',
      order.email || '',
      order.shipBy || '',
      '未作成',
      '受注済',
      '',
      '',
      '',
      ''
    ]);

    Logger.log('✅ logEbayOrderDirect_: 受注ログ記録完了 ' + order.orderId);
    return true;

  } catch (e) {
    Logger.log('❌ logEbayOrderDirect_ エラー: ' + e.message);
    return false;
  }
}


function notifyEbayOrderDirect_(order) {
  try {
    if (!order || !order.orderId) {
      Logger.log('❌ notifyEbayOrderDirect_: orderIdなし');
      return false;
    }

    const receivedAt = order.receivedAt instanceof Date
      ? order.receivedAt
      : new Date();

    const timeStr = Utilities.formatDate(
      receivedAt,
      'Asia/Tokyo',
      'MM/dd HH:mm'
    );

    const msg =
      '🛒 eBay 受注通知\n' +
      '━━━━━━━━━━━━━━━\n' +
      '🕐 ' + timeStr + ' JST\n' +
      '📦 ' + (order.itemTitle || '(商品名)') + '\n' +
      '🔖 注文ID: ' + order.orderId + '_eBay\n' +
      (order.itemId ? '🆔 Item ID: ' + order.itemId + '\n' : '') +
      (order.sku ? '🏷 SKU: ' + order.sku + '\n' : '') +
      (order.buyerName ? '👤 購入者: ' + order.buyerName + '\n' : '') +
      (order.address ? '📍 住所: ' + order.address + '\n' : '') +
      (order.phone ? '📞 電話: ' + order.phone + '\n' : '') +
      (order.email ? '✉️ メール: ' + order.email + '\n' : '') +
      (order.shipBy ? '📅 発送期限: ' + order.shipBy + '\n' : '') +
      (order.quantity ? '📊 数量: ' + order.quantity + '\n' : '') +
      (order.amount ? '💰 金額: ' + order.amount + '\n' : '') +
      '\n⏳ 12時間経過後、autoCreateMcfOrders がMCF作成対象にします';

    sendLine(msg);

    Logger.log('✅ notifyEbayOrderDirect_: LINE通知送信 ' + order.orderId);
    return true;

  } catch (e) {
    Logger.log('❌ notifyEbayOrderDirect_ エラー: ' + e.message);
    return false;
  }
}/**
 * SpApiInventory.gs  （このファイルだけを差し替える）
 * ===============================================================
 * SP-API FBA在庫取得まわりの関数を「1ファイルに集約」したもの。
 *
 * 【安全な更新手順 / なぜ1ファイルに集約するか】
 *   - この8関数だけをこのファイルに置けば、差し替えは「このファイルの中身を
 *     まるごと貼り替える」だけで済みます。
 *   - updateEbayRawFromTradingApi() / sendLine() / runInventorySyncAllAuto()
 *     などは別ファイルにあるため、このファイルを貼り替えても一切影響しません。
 *   - 同名関数の重複（特に updateFbaInventoryFromSpApi が2つある状態）を防ぐため、
 *     旧SP-APIコードが別ファイルや同ファイルに残っていないか必ず確認してください。
 *
 * 含まれる関数（差し替え対象）：
 *   updateFbaInventoryFromSpApi()
 *   _fetchInventoryForSkus_()
 *   _fetchInventoryForSingleSku_()
 *   _buildSummariesUrl_()
 *   _spApiGetRaw_()
 *   getSpApiAccessToken_()   ← Script Properties名を新仕様に変更
 *   _normSku_()
 *   chunkArray_()
 *   debugSpApiSingleSku_HD_BB9M_IXV5()
 *
 * Script Properties（必須）：
 *   AMAZON_CLIENT_ID
 *   AMAZON_CLIENT_SECRET
 *   SP_REFRESH_TOKEN
 *
 * eBay在庫変更処理は一切含みません（取得・シート反映のみ）。
 * ===============================================================
 */

// ===== 設定 =====
var SHEET_STOCK_SYNC = '在庫同期管理';
var SP_API_ENDPOINT  = 'https://sellingpartnerapi-na.amazon.com';
var MARKETPLACE_ID   = 'ATVPDKIKX0DER';   // Amazon.com
var SKU_CHUNK_SIZE   = 50;
var SLEEP_BETWEEN_MS = 700;   // 一括チャンク間
var SLEEP_SINGLE_MS  = 600;   // 単体補完の各呼び出し間（2rps対策）
var MAX_RETRY        = 3;

var SKU_EXCLUDE = ['SKU', 'WAREHOUSES', 'DELETED', 'ITEMID', 'TITLE', 'AVAILABLEQUANTITY'];
var DEBUG_TRACK_SKUS = ['HD-BB9M-IXV5'];

// テスト用SKU（SP-APIに問い合わせず監視除外する）。
// 完全一致のほか、SAMPLE-001 / TEST_SKU / DUMMY123 のような部分一致も対象。
var TEST_SKU_TOKENS = ['SAMPLE', 'TEST', 'DUMMY'];

/*
 * ※注意：定数 SHEET_STOCK_SYNC / SKU_EXCLUDE は AutoOps 側ファイルでも
 *   宣言している場合があります。GASは同名のグローバル var が複数ファイルに
 *   あると後勝ちになり実行は通りますが、混乱を避けるため、どちらか一方に
 *   まとめることを推奨します（重複してもエラーにはなりません）。
 */


/**
 * 【メイン】一括取得 → 未取得SKUを単体補完 → 在庫同期管理D列へ反映
 */
function updateFbaInventoryFromSpApi() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_STOCK_SYNC);
  if (!sheet) { Logger.log('シートが見つかりません: ' + SHEET_STOCK_SYNC); return; }

  var lastRow = sheet.getLastRow();
  if (lastRow < 2) { Logger.log('データ行がありません'); return; }
  var numRows = lastRow - 1;

  var trackNorm = {};
  for (var t = 0; t < DEBUG_TRACK_SKUS.length; t++) {
    trackNorm[_normSku_(DEBUG_TRACK_SKUS[t])] = DEBUG_TRACK_SKUS[t];
  }

  // A列SKU読み取り（両側 _normSku_ で対称化）
  var skuColumn = sheet.getRange(2, 1, numRows, 1).getValues();
  var rowSkus = [];
  var uniqueSkus = {};
  var excludedCount = 0;
  for (var i = 0; i < skuColumn.length; i++) {
    var raw = skuColumn[i][0];
    var norm = _normSku_(raw);
    if (!norm) continue;
    if (SKU_EXCLUDE.indexOf(norm.toUpperCase()) !== -1) continue;
    if (_isTestSku_(norm)) {
      // テストSKUはSP-APIへ問い合わせない。行は保持し「監視除外」を入れる。
      rowSkus.push({ row0: i, raw: raw, norm: norm, excluded: true });
      excludedCount++;
      continue;
    }
    rowSkus.push({ row0: i, raw: raw, norm: norm, excluded: false });
    uniqueSkus[norm] = raw;
  }
  var skus = Object.keys(uniqueSkus);
  Logger.log('対象SKU数: ' + skus.length + '（監視除外: ' + excludedCount + '件）');
  if (skus.length === 0 && excludedCount === 0) { Logger.log('有効なSKUがありません'); return; }

  var fbaMap = {};     // norm → fulfillableQuantity（0含む）
  var fbaSource = {};  // norm → 'batch' | 'single'
  var batchFoundCount = 0;
  var missingSkus = [];
  var singleSuccess = 0, singleFail = 0;

  // 実在SKUがある場合のみSP-APIへ問い合わせる（除外のみの時はトークン取得もしない）
  if (skus.length > 0) {
    var token = getSpApiAccessToken_();

    // STEP1：一括取得
    var chunks = chunkArray_(skus, SKU_CHUNK_SIZE);
    for (var c = 0; c < chunks.length; c++) {
      var chunk = chunks[c];
      var summaries = _fetchInventoryForSkus_(chunk, token);
      Logger.log('[CHUNK ' + (c + 1) + '/' + chunks.length + '] '
        + '送信SKU件数=' + chunk.length + ' / レスポンスsummary件数=' + summaries.length);
      for (var s = 0; s < summaries.length; s++) {
        var sum = summaries[s];
        var nsku = _normSku_(sum.sellerSku);
        if (!nsku) continue;
        var q = (sum.inventoryDetails || {}).fulfillableQuantity;
        if (typeof q === 'number') { fbaMap[nsku] = q; fbaSource[nsku] = 'batch'; }
      }
      if (c < chunks.length - 1) Utilities.sleep(SLEEP_BETWEEN_MS);
    }
    batchFoundCount = Object.keys(fbaMap).length;

    // STEP2：未取得抽出
    for (var m = 0; m < skus.length; m++) {
      if (!fbaMap.hasOwnProperty(skus[m])) missingSkus.push(skus[m]);
    }
    Logger.log('一括取得件数: ' + batchFoundCount + ' / 一括未発見件数: ' + missingSkus.length);

    // STEP3-4：単体補完
    for (var k = 0; k < missingSkus.length; k++) {
      var msku = missingSkus[k];
      var got = _fetchInventoryForSingleSku_(msku, token);
      if (got.found && typeof got.qty === 'number') {
        fbaMap[msku] = got.qty; fbaSource[msku] = 'single'; singleSuccess++;
        Logger.log('  [単体補完OK] ' + msku + ' fulfillableQuantity=' + got.qty);
      } else {
        singleFail++;
        Logger.log('  [単体補完NG] ' + msku + (got.note ? ' (' + got.note + ')' : ''));
      }
      if (k < missingSkus.length - 1) Utilities.sleep(SLEEP_SINGLE_MS);
    }

    for (var nk in trackNorm) {
      if (fbaMap.hasOwnProperty(nk)) {
        Logger.log('[追跡] ' + trackNorm[nk] + ' → 取得成功 source=' + fbaSource[nk] + ' 値=' + fbaMap[nk]);
      } else {
        Logger.log('[追跡] ' + trackNorm[nk] + ' → 一括・単体とも取得失敗');
      }
    }
  } else {
    Logger.log('実在SKUが0件のためSP-API呼び出しをスキップ（監視除外のみ）');
  }

  // STEP5-7：シート反映
  var now = new Date();
  var dCol  = sheet.getRange(2, 4, numRows, 1).getValues();
  var hiCol = sheet.getRange(2, 8, numRows, 2).getValues();
  var excludedWritten = 0;
  for (var r = 0; r < rowSkus.length; r++) {
    var idx = rowSkus[r].row0;
    var norm = rowSkus[r].norm;

    // 監視除外SKU：D列は更新せず、I列に「監視除外」を入れる
    if (rowSkus[r].excluded) {
      hiCol[idx][1] = '監視除外';
      excludedWritten++;
      continue;
    }

    if (fbaMap.hasOwnProperty(norm)) {
      dCol[idx][0]  = fbaMap[norm];
      hiCol[idx][0] = now;
      hiCol[idx][1] = (fbaSource[norm] === 'single') ? 'SP-API取得済（単体補完）' : 'SP-API取得済';
    } else {
      dCol[idx][0]  = '';   // 0は入れない
      hiCol[idx][0] = now;
      hiCol[idx][1] = 'SP-API未発見';
    }
  }
  sheet.getRange(2, 4, numRows, 1).setValues(dCol);
  sheet.getRange(2, 8, numRows, 2).setValues(hiCol);

  var finalMiss = missingSkus.length - singleSuccess;
  Logger.log('===== 反映完了 =====');
  Logger.log('一括取得件数        : ' + batchFoundCount);
  Logger.log('一括未発見件数      : ' + missingSkus.length);
  Logger.log('単体補完成功件数    : ' + singleSuccess);
  Logger.log('単体補完失敗件数    : ' + singleFail);
  Logger.log('監視除外件数        : ' + excludedWritten);
  Logger.log('最終取得済(合計)    : ' + (batchFoundCount + singleSuccess) + ' / ' + skus.length);
  Logger.log('最終未発見          : ' + finalMiss);
}


/** 複数SKU一括取得（sellerSkus 繰り返し付与） */
function _fetchInventoryForSkus_(skuChunk, token) {
  var results = [];
  var nextToken = null;
  var page = 0;
  do {
    var url = _buildSummariesUrl_(skuChunk, nextToken, false);
    var res = _spApiGetRaw_(url, token);
    if (res.code !== 200) {
      Logger.log('一括取得エラー code=' + res.code + ' body=' + res.body.substring(0, 800));
      throw new Error('SP-API batch request failed: ' + res.code);
    }
    var json = JSON.parse(res.body);
    var arr = (json.payload && json.payload.inventorySummaries) || [];
    results = results.concat(arr);
    page++;
    nextToken = (json.pagination && json.pagination.nextToken) || null;
  } while (nextToken && page < 20);
  return results;
}


/** 単体SKU取得（補完用）。例外を投げず {found, qty, note} を返す */
function _fetchInventoryForSingleSku_(sku, token) {
  var url = _buildSummariesUrl_([sku], null, true);
  var res = _spApiGetRaw_(url, token);
  if (res.code !== 200) return { found: false, qty: null, note: 'HTTP ' + res.code };
  var json;
  try { json = JSON.parse(res.body); }
  catch (e) { return { found: false, qty: null, note: 'JSON parse error' }; }
  var summaries = (json.payload && json.payload.inventorySummaries) || [];
  if (summaries.length === 0) return { found: false, qty: null, note: 'summary 0件' };
  var nsku = _normSku_(sku);
  var picked = null;
  for (var i = 0; i < summaries.length; i++) {
    if (_normSku_(summaries[i].sellerSku) === nsku) { picked = summaries[i]; break; }
  }
  if (!picked) picked = summaries[0];
  var q = (picked.inventoryDetails || {}).fulfillableQuantity;
  if (typeof q === 'number') return { found: true, qty: q, note: '' };
  return { found: false, qty: null, note: 'fulfillableQuantity欠落' };
}


/** 単体・一括 共通URL構築 */
function _buildSummariesUrl_(skuList, nextToken, singular) {
  var url = SP_API_ENDPOINT + '/fba/inventory/v1/summaries'
    + '?details=true'
    + '&granularityType=Marketplace'
    + '&granularityId=' + MARKETPLACE_ID
    + '&marketplaceIds=' + MARKETPLACE_ID;
  if (nextToken) {
    url += '&nextToken=' + encodeURIComponent(nextToken);
  } else if (singular) {
    url += '&sellerSku=' + encodeURIComponent(skuList[0]);
  } else {
    for (var i = 0; i < skuList.length; i++) {
      url += '&sellerSkus=' + encodeURIComponent(skuList[i]);
    }
  }
  return url;
}


/** SP-API GET（raw）。429/5xxは指数バックオフ */
function _spApiGetRaw_(url, token) {
  var attempt = 0;
  while (true) {
    var res = UrlFetchApp.fetch(url, {
      method: 'get',
      headers: { 'x-amz-access-token': token },
      muteHttpExceptions: true
    });
    var code = res.getResponseCode();
    var body = res.getContentText();
    if (code === 200) return { code: code, body: body, url: url };
    if ((code === 429 || code >= 500) && attempt < MAX_RETRY) {
      var wait = Math.pow(2, attempt) * 1000;
      Logger.log('HTTP ' + code + ' リトライ ' + (attempt + 1) + '回目 / wait ' + wait + 'ms');
      Utilities.sleep(wait);
      attempt++;
      continue;
    }
    return { code: code, body: body, url: url };
  }
}


/**
 * LWAアクセストークン取得（SP-APIはSigV4不要・LWA3項目のみ）
 * Script Properties: AMAZON_CLIENT_ID / AMAZON_CLIENT_SECRET / SP_REFRESH_TOKEN
 */
function getSpApiAccessToken_() {
  var props = PropertiesService.getScriptProperties();
  var clientId     = sampleCsvProps_().getProperty('AMAZON_CLIENT_ID');
  var clientSecret = sampleCsvProps_().getProperty('AMAZON_CLIENT_SECRET');
  var refreshToken = sampleCsvProps_().getProperty('SP_REFRESH_TOKEN');
  if (!clientId || !clientSecret || !refreshToken) {
    throw new Error('SP-API認証情報(Script Properties)が不足: '
      + 'AMAZON_CLIENT_ID / AMAZON_CLIENT_SECRET / SP_REFRESH_TOKEN を確認してください');
  }
  var res = UrlFetchApp.fetch('https://api.amazon.com/auth/o2/token', {
    method: 'post',
    contentType: 'application/x-www-form-urlencoded',
    payload: {
      grant_type: 'refresh_token',
      refresh_token: refreshToken,
      client_id: clientId,
      client_secret: clientSecret
    },
    muteHttpExceptions: true
  });
  var code = res.getResponseCode();
  var data = JSON.parse(res.getContentText());
  if (code !== 200 || !data.access_token) {
    throw new Error('LWAトークン取得失敗 code=' + code + ' body=' + res.getContentText());
  }
  return data.access_token;
}


/** SKU正規化（既存 normalizeSku_ があれば使用） */
function _normSku_(sku) {
  if (typeof normalizeSku_ === 'function') {
    try { return normalizeSku_(sku); } catch (e) {}
  }
  return String(sku == null ? '' : sku).trim();
}


/** SAMPLE / TEST / DUMMY 等のテストSKUか判定（部分一致・大文字小文字無視） */
function _isTestSku_(sku) {
  var u = String(sku == null ? '' : sku).toUpperCase();
  for (var i = 0; i < TEST_SKU_TOKENS.length; i++) {
    if (u.indexOf(TEST_SKU_TOKENS[i]) !== -1) return true;
  }
  return false;
}


/** 配列をsize件ずつ分割 */
function chunkArray_(arr, size) {
  var out = [];
  for (var i = 0; i < arr.length; i += size) out.push(arr.slice(i, i + size));
  return out;
}


/** 【デバッグ・単数指定】HD-BB9M-IXV5 */
function debugSpApiSingleSku_HD_BB9M_IXV5() {
  var targetSku = 'HD-BB9M-IXV5';
  var token = getSpApiAccessToken_();
  var url = _buildSummariesUrl_([targetSku], null, true);
  var r = _spApiGetRaw_(url, token);
  Logger.log('URL:\n' + r.url);
  Logger.log('HTTP: ' + r.code);
  Logger.log('本文(先頭3000):\n' + r.body.substring(0, 3000));
  if (r.code !== 200) return;
  var json = JSON.parse(r.body);
  var summaries = (json.payload && json.payload.inventorySummaries) || [];
  Logger.log('summaries件数: ' + summaries.length);
  for (var i = 0; i < summaries.length; i++) {
    var d = summaries[i].inventoryDetails || {};
    Logger.log('raw="' + summaries[i].sellerSku + '" norm="' + _normSku_(summaries[i].sellerSku)
      + '" fulfillableQuantity=' + d.fulfillableQuantity
      + ' totalQuantity=' + summaries[i].totalQuantity);
  }
}
/**
 * 在庫同期管理を使ったeBay在庫0候補の確認
 *
 * ルールB：
 * Amazon在庫 <= SKUごとの通知基準数
 * かつ eBay設定在庫 > 0
 *
 * 読み取り専用：
 * - eBay API更新なし
 * - シート書き込みなし
 * - LINE通知なし
 */
function dryRunFbaToEbayZeroCandidates() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('在庫同期管理');

  if (!sheet) {
    throw new Error('在庫同期管理シートが見つかりません');
  }

  const data = sheet.getDataRange().getValues();

  if (data.length < 2) {
    Logger.log('対象データがありません');
    return [];
  }

  const targets = [];

  Logger.log('');
  Logger.log('━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━');
  Logger.log('【eBay在庫0変更候補・DRY_RUN】');
  Logger.log('判定：Amazon在庫 <= SKU別通知基準数');
  Logger.log('★ eBay在庫の変更は行いません');
  Logger.log('━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━');

  for (let i = 1; i < data.length; i++) {
    const row = data[i];

    const sku = String(row[0] || '').trim();       // A列
    const itemId = String(row[1] || '').trim();    // B列
    const title = String(row[2] || '').trim();     // C列
    const amazonQty = toInventoryNumber_(row[3]);  // D列
    const ebayQty = toInventoryNumber_(row[4]);    // E列
    const warningQty = toInventoryNumber_(row[5]); // F列
    const checkedAt = row[7];                      // H列
    const resultText = String(row[8] || '').trim();// I列

    if (!sku || !itemId) {
      continue;
    }

    if (
      amazonQty === null ||
      ebayQty === null ||
      warningQty === null
    ) {
      Logger.log(
        '⏭ 行' + (i + 1) +
        ' 数値不正: ' + sku
      );
      continue;
    }

    const isTarget =
      amazonQty <= warningQty &&
      ebayQty > 0;

    if (!isTarget) {
      continue;
    }

    const target = {
      rowNumber: i + 1,
      sku: sku,
      itemId: itemId,
      title: title,
      amazonQty: amazonQty,
      ebayQty: ebayQty,
      warningQty: warningQty,
      checkedAt: checkedAt,
      resultText: resultText
    };

    targets.push(target);

    Logger.log('');
    Logger.log('⚠️ 在庫0変更候補');
    Logger.log('行番号       : ' + target.rowNumber);
    Logger.log('SKU          : ' + target.sku);
    Logger.log('ItemID       : ' + target.itemId);
    Logger.log('Amazon在庫   : ' + target.amazonQty);
    Logger.log('eBay在庫     : ' + target.ebayQty);
    Logger.log('通知基準数   : ' + target.warningQty);
    Logger.log('最終確認日時 : ' + target.checkedAt);
    Logger.log('確認結果     : ' + target.resultText);
    Logger.log(
      '商品名       : ' +
      target.title.substring(0, 100)
    );
  }

  Logger.log('');
  Logger.log('━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━');
  Logger.log('在庫0変更候補: ' + targets.length + '件');

  if (targets.length === 0) {
    Logger.log('✅ 現在、変更候補はありません');
  } else {
    Logger.log('★ DRY_RUNのためeBayへの変更はしていません');
  }

  Logger.log('━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━');

  return targets;
}

function toInventoryNumber_(value) {
  if (
    value === '' ||
    value === null ||
    value === undefined
  ) {
    return null;
  }

  const number = Number(value);

  return Number.isFinite(number)
    ? number
    : null;
}

/**
 * 旧CSV方式のFBA在庫チェックトリガーだけを削除する
 *
 * 削除対象：
 * checkFbaInventory の時間トリガー
 *
 * 削除しないもの：
 * runInventorySyncAllAuto
 * checkEbayOrders
 * autoCreateMcfOrders
 * その他のトリガー
 */
function deleteLegacyCheckFbaInventoryTrigger() {
  const triggers = ScriptApp.getProjectTriggers();
  let deletedCount = 0;

  triggers.forEach(function(trigger) {
    if (trigger.getHandlerFunction() === 'checkFbaInventory') {
      ScriptApp.deleteTrigger(trigger);
      deletedCount++;
      Logger.log('削除: checkFbaInventory');
    }
  });

  Logger.log('====================================');
  Logger.log(
    '旧 checkFbaInventory トリガー削除数: ' +
    deletedCount
  );

  if (deletedCount === 0) {
    Logger.log('対象トリガーはありませんでした');
  }

  Logger.log('====================================');
}

/**
 * ============================================================
 * E在庫0処理_本番.gs
 *
 * Amazon FBA在庫に応じて、eBay販売可能在庫を0に変更する
 *
 * 【採用ルールB】
 * Amazon在庫 <= SKUごとの通知基準数
 * かつ
 * eBay設定在庫 > 0
 *
 * 【安全仕様】
 * - 手動実行専用
 * - 実行前に確認画面を表示
 * - 在庫データが3時間より古ければ停止
 * - SP-API取得成功行のみ対象
 * - eBayRawと在庫同期管理の数量が一致しなければ停止
 * - LockServiceで二重実行防止
 * - 最大20件を超えた場合は停止
 *
 * 【依存する既存関数】
 * - getNewAccessToken()
 *   または
 * - getEbayAccessTokenFromRefreshToken()
 *
 * 任意：
 * - sendLine()
 * ============================================================
 */


/**
 * 本番実行関数
 *
 * 注意：
 * - 時間トリガーには登録しない
 * - まず runInventorySyncAllAuto
 * - 次に dryRunFbaToEbayZeroCandidates
 * - 最後に本関数を手動実行
 */
function applyFbaToEbayZero_CONFIRMED() {
  const lock = LockService.getScriptLock();

  if (!lock.tryLock(10000)) {
    throw new Error(
      '別の在庫処理が実行中です。時間を置いて再実行してください。'
    );
  }

  try {
    const MAX_DATA_AGE_MINUTES = 180;
    const MAX_TARGET_COUNT = 20;

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const syncSheet = ss.getSheetByName('在庫同期管理');
    const ebayRawSheet = ss.getSheetByName('eBayRaw');

    if (!syncSheet) {
      throw new Error('在庫同期管理シートが見つかりません');
    }

    if (!ebayRawSheet) {
      throw new Error('eBayRawシートが見つかりません');
    }

    const syncData = syncSheet.getDataRange().getValues();
    const rawData = ebayRawSheet.getDataRange().getValues();

    if (syncData.length < 2) {
      throw new Error('在庫同期管理にデータがありません');
    }

    if (rawData.length < 2) {
      throw new Error('eBayRawにデータがありません');
    }

    const syncHeaders = syncData[0].map(function(value) {
      return String(value || '').trim();
    });

    const rawHeaders = rawData[0].map(function(value) {
      return String(value || '').trim();
    });

    const syncCol = {
      sku: syncHeaders.indexOf('SKU'),
      itemId: syncHeaders.indexOf('ItemID'),
      title: syncHeaders.indexOf('タイトル'),
      amazonQty: syncHeaders.indexOf('Amazon在庫'),
      ebayQty: syncHeaders.indexOf('eBay設定在庫'),
      warningQty: syncHeaders.indexOf('通知基準数'),
      notificationStatus: syncHeaders.indexOf('通知状態'),
      checkedAt: syncHeaders.indexOf('最終確認日時'),
      result: syncHeaders.indexOf('確認結果')
    };

    const rawCol = {
      itemId: rawHeaders.indexOf('ItemID'),
      sku: rawHeaders.indexOf('SKU'),
      availableQuantity: rawHeaders.indexOf('AvailableQuantity')
    };

    Object.keys(syncCol).forEach(function(key) {
      if (syncCol[key] === -1) {
        throw new Error(
          '在庫同期管理に必要な列がありません: ' + key
        );
      }
    });

    Object.keys(rawCol).forEach(function(key) {
      if (rawCol[key] === -1) {
        throw new Error(
          'eBayRawに必要な列がありません: ' + key
        );
      }
    });

    // eBayRawをItemID単位で検索できるようにする
    const rawMap = {};

    for (let i = 1; i < rawData.length; i++) {
      const row = rawData[i];

      const itemId = String(
        row[rawCol.itemId] || ''
      ).trim();

      if (!itemId) {
        continue;
      }

      rawMap[itemId] = {
        sku: String(row[rawCol.sku] || '').trim(),
        availableQuantity: toEbayZeroNumber_(
          row[rawCol.availableQuantity]
        )
      };
    }

    const now = new Date();
    const targets = [];

    for (let i = 1; i < syncData.length; i++) {
      const row = syncData[i];

      const sku = String(
        row[syncCol.sku] || ''
      ).trim();

      const itemId = String(
        row[syncCol.itemId] || ''
      ).trim();

      const title = String(
        row[syncCol.title] || ''
      ).trim();

      const amazonQty = toEbayZeroNumber_(
        row[syncCol.amazonQty]
      );

      const ebayQty = toEbayZeroNumber_(
        row[syncCol.ebayQty]
      );

      const warningQty = toEbayZeroNumber_(
        row[syncCol.warningQty]
      );

      const resultText = String(
        row[syncCol.result] || ''
      ).trim();

      if (!sku || !itemId) {
        continue;
      }

      if (
        amazonQty === null ||
        ebayQty === null ||
        warningQty === null
      ) {
        continue;
      }

      // ルールB
      const isTarget =
        amazonQty <= warningQty &&
        ebayQty > 0;

      if (!isTarget) {
        continue;
      }

      // SP-APIで正常取得できた行だけ対象
      if (resultText.indexOf('SP-API取得済') !== 0) {
        throw new Error(
          '安全停止：SP-API取得済ではありません。' +
          ' SKU=' + sku +
          ' / 確認結果=' + resultText
        );
      }

      const checkedAt = parseEbayZeroDate_(
        row[syncCol.checkedAt]
      );

      if (!checkedAt) {
        throw new Error(
          '安全停止：最終確認日時を判定できません。' +
          ' SKU=' + sku
        );
      }

      const ageMinutes =
        (now.getTime() - checkedAt.getTime()) / 60000;

      if (ageMinutes < -5) {
        throw new Error(
          '安全停止：最終確認日時が未来になっています。' +
          ' SKU=' + sku
        );
      }

      if (ageMinutes > MAX_DATA_AGE_MINUTES) {
        throw new Error(
          '安全停止：在庫データが古いです。' +
          ' SKU=' + sku +
          ' / 経過=' +
          Math.floor(ageMinutes) +
          '分'
        );
      }

      if (!/^\d{10,20}$/.test(itemId)) {
        throw new Error(
          '安全停止：ItemID形式が不正です。' +
          ' SKU=' + sku +
          ' / ItemID=' + itemId
        );
      }

      const raw = rawMap[itemId];

      if (!raw) {
        throw new Error(
          '安全停止：eBayRawにItemIDがありません。' +
          ' SKU=' + sku +
          ' / ItemID=' + itemId
        );
      }

      if (raw.availableQuantity === null) {
        throw new Error(
          '安全停止：eBayRawの在庫数を取得できません。' +
          ' SKU=' + sku
        );
      }

      if (raw.availableQuantity !== ebayQty) {
        throw new Error(
          '安全停止：eBay在庫数が一致しません。' +
          ' SKU=' + sku +
          ' / 在庫同期管理=' + ebayQty +
          ' / eBayRaw=' + raw.availableQuantity
        );
      }

      targets.push({
        rowNumber: i + 1,
        sku: sku,
        itemId: itemId,
        title: title,
        amazonQty: amazonQty,
        ebayQty: ebayQty,
        warningQty: warningQty
      });
    }

    if (targets.length === 0) {
      Logger.log('✅ eBay在庫0変更対象はありません');
      return;
    }

    if (targets.length > MAX_TARGET_COUNT) {
      throw new Error(
        '安全停止：対象件数が多すぎます。' +
        ' 対象=' + targets.length + '件'
      );
    }

    Logger.log('');
    Logger.log('━━━━━━━━━━━━━━━━━━━━━━━━━━━━');
    Logger.log('【eBay在庫0・本番変更予定】');

    targets.forEach(function(target) {
      Logger.log(
        target.sku +
        ' / ItemID=' + target.itemId +
        ' / Amazon=' + target.amazonQty +
        ' / eBay=' + target.ebayQty +
        ' / 基準=' + target.warningQty
      );
    });

    Logger.log('━━━━━━━━━━━━━━━━━━━━━━━━━━━━');

// 事前に発行した「1回限りの承認」を確認
const approvalProps =
  PropertiesService.getScriptProperties();

const approvedUntil = Number(
  approvalsampleCsvProps_().getProperty(
    'FBA_EBAY_ZERO_APPROVED_UNTIL'
  ) || 0
);

if (
  !approvedUntil ||
  Date.now() > approvedUntil
) {
  throw new Error(
    '安全停止：本番実行の事前承認がありません。' +
    '先に approveFbaToEbayZeroOnce を実行し、' +
    '10分以内に再実行してください。'
  );
}

// 承認は1回だけ使用可能
approvalProps.deleteProperty(
  'FBA_EBAY_ZERO_APPROVED_UNTIL'
);

Logger.log(
  '✅ 1回限りの本番実行承認を確認しました'
);

    // 既存のeBayアクセストークン取得関数を利用
    let accessToken = '';

    if (typeof getNewAccessToken === 'function') {
      accessToken = getNewAccessToken();

    } else if (
      typeof getEbayAccessTokenFromRefreshToken === 'function'
    ) {
      accessToken =
        getEbayAccessTokenFromRefreshToken();

    } else {
      throw new Error(
        'eBayアクセストークン取得関数が見つかりません。' +
        ' getNewAccessToken または' +
        ' getEbayAccessTokenFromRefreshToken が必要です。'
      );
    }

    if (!accessToken) {
      throw new Error('eBayアクセストークンを取得できません');
    }

    const results = [];

    targets.forEach(function(target) {
      const result = reviseEbayAvailableToZero_(
        target.itemId,
        accessToken
      );

      if (result.success) {
        // E列：eBay設定在庫
        syncSheet
          .getRange(
            target.rowNumber,
            syncCol.ebayQty + 1
          )
          .setValue(0);

        // G列：通知状態
        syncSheet
          .getRange(
            target.rowNumber,
            syncCol.notificationStatus + 1
          )
          .setValue('eBay在庫0変更済');

        // I列：確認結果
        syncSheet
          .getRange(
            target.rowNumber,
            syncCol.result + 1
          )
          .setValue(
            'eBay在庫0変更成功' +
            '（Amazon=' + target.amazonQty +
            ' / 基準=' + target.warningQty + '）'
          );

        Logger.log(
          '✅ eBay在庫0変更成功: ' +
          target.sku +
          ' / ItemID=' +
          target.itemId
        );

      } else {
        syncSheet
          .getRange(
            target.rowNumber,
            syncCol.notificationStatus + 1
          )
          .setValue('eBay在庫0変更エラー');

        syncSheet
          .getRange(
            target.rowNumber,
            syncCol.result + 1
          )
          .setValue(
            'eBay更新失敗: ' +
            String(result.error || '')
              .substring(0, 300)
          );

        Logger.log(
          '❌ eBay在庫0変更失敗: ' +
          target.sku +
          ' / ' +
          result.error
        );
      }

      results.push({
        target: target,
        success: result.success,
        error: result.error || '',
        ack: result.ack || ''
      });

      Utilities.sleep(1000);
    });

    SpreadsheetApp.flush();

    const successCount = results.filter(function(result) {
      return result.success;
    }).length;

    const failureCount =
      results.length - successCount;

    let message =
      '📦 eBay在庫0変更結果\n' +
      '成功：' + successCount + '件\n' +
      '失敗：' + failureCount + '件\n\n';

    results.forEach(function(result) {
      message +=
        (result.success ? '✅ ' : '❌ ') +
        result.target.sku +
        ' / Amazon=' +
        result.target.amazonQty +
        ' / 基準=' +
        result.target.warningQty +
        '\n';
    });

    Logger.log(message);

    if (typeof sendLine === 'function') {
      try {
        sendLine(message);
      } catch (lineError) {
        Logger.log(
          'LINE通知失敗: ' +
          lineError.message
        );
      }
    }

  } finally {
    lock.releaseLock();
  }
}


/**
 * eBayの販売可能在庫を0に変更する
 */
function reviseEbayAvailableToZero_(
  itemId,
  accessToken
) {
  const xml =
    '<?xml version="1.0" encoding="UTF-8"?>' +
    '<ReviseFixedPriceItemRequest ' +
    'xmlns="urn:ebay:apis:eBLBaseComponents">' +
      '<Item>' +
        '<ItemID>' + itemId + '</ItemID>' +
        '<Quantity>0</Quantity>' +
      '</Item>' +
    '</ReviseFixedPriceItemRequest>';

  try {
    const response = UrlFetchApp.fetch(
      'https://api.ebay.com/ws/api.dll',
      {
        method: 'post',
        headers: {
          'X-EBAY-API-IAF-TOKEN': accessToken,
          'X-EBAY-API-CALL-NAME':
            'ReviseFixedPriceItem',
          'X-EBAY-API-SITEID': '0',
          'X-EBAY-API-COMPATIBILITY-LEVEL': '1199',
          'Content-Type': 'text/xml'
        },
        payload: xml,
        muteHttpExceptions: true
      }
    );

    const httpCode =
      response.getResponseCode();

    const body =
      response.getContentText();

    const ack =
      (body.match(
        /<Ack>([^<]+)<\/Ack>/
      ) || [])[1] || '';

    const shortMessage =
      (body.match(
        /<ShortMessage>([\s\S]*?)<\/ShortMessage>/
      ) || [])[1] || '';

    Logger.log(
      'ReviseFixedPriceItem [' +
      httpCode +
      '] ItemID=' +
      itemId +
      ' Ack=' +
      ack +
      '\n' +
      body.substring(0, 1000)
    );

    if (
      httpCode === 200 &&
      (
        ack === 'Success' ||
        ack === 'Warning'
      )
    ) {
      return {
        success: true,
        ack: ack,
        error: ''
      };
    }

    return {
      success: false,
      ack: ack,
      error:
        'HTTP ' + httpCode +
        ' / Ack=' + ack +
        ' / ' + shortMessage
    };

  } catch (error) {
    return {
      success: false,
      ack: '',
      error: error.message
    };
  }
}


/**
 * 数値を安全に変換する
 */
function toEbayZeroNumber_(value) {
  if (
    value === '' ||
    value === null ||
    value === undefined
  ) {
    return null;
  }

  const number = Number(value);

  return Number.isFinite(number)
    ? number
    : null;
}


/**
 * 日付を安全に変換する
 */
function parseEbayZeroDate_(value) {
  if (
    value instanceof Date &&
    !isNaN(value.getTime())
  ) {
    return value;
  }

  const parsed = new Date(value);

  return isNaN(parsed.getTime())
    ? null
    : parsed;
}

/**
 * eBayのOut-of-Stock設定を確認する
 *
 * 読み取り専用：
 * - 出品変更なし
 * - シート変更なし
 * - LINE通知なし
 */
function checkEbayOutOfStockControlPreference() {
  let accessToken = '';

  if (typeof getNewAccessToken === 'function') {
    accessToken = getNewAccessToken();
  } else if (
    typeof getEbayAccessTokenFromRefreshToken === 'function'
  ) {
    accessToken = getEbayAccessTokenFromRefreshToken();
  } else {
    throw new Error(
      'eBayアクセストークン取得関数が見つかりません'
    );
  }

  const xml =
    '<?xml version="1.0" encoding="UTF-8"?>' +
    '<GetUserPreferencesRequest ' +
    'xmlns="urn:ebay:apis:eBLBaseComponents">' +
      '<ShowOutOfStockControlPreference>true' +
      '</ShowOutOfStockControlPreference>' +
    '</GetUserPreferencesRequest>';

  const response = UrlFetchApp.fetch(
    'https://api.ebay.com/ws/api.dll',
    {
      method: 'post',
      headers: {
        'X-EBAY-API-IAF-TOKEN': accessToken,
        'X-EBAY-API-CALL-NAME': 'GetUserPreferences',
        'X-EBAY-API-SITEID': '0',
        'X-EBAY-API-COMPATIBILITY-LEVEL': '1199',
        'Content-Type': 'text/xml'
      },
      payload: xml,
      muteHttpExceptions: true
    }
  );

  const httpCode = response.getResponseCode();
  const body = response.getContentText();

  const ack =
    (body.match(/<Ack>([^<]+)<\/Ack>/) || [])[1] || '';

  const preference =
    (
      body.match(
        /<OutOfStockControlPreference>([^<]+)<\/OutOfStockControlPreference>/
      ) || []
    )[1] || '';

  Logger.log('━━━━━━━━━━━━━━━━━━━━━━━━━━━━');
  Logger.log('【eBay Out-of-Stock設定確認】');
  Logger.log('HTTP: ' + httpCode);
  Logger.log('Ack: ' + ack);
  Logger.log('OutOfStockControlPreference: ' + preference);

  if (String(preference).toLowerCase() === 'true') {
    Logger.log('✅ 有効：数量0で出品を維持しつつ非表示にできます');
  } else {
    Logger.log('❌ 無効または取得失敗：本番処理は実行しないでください');
  }

  Logger.log('━━━━━━━━━━━━━━━━━━━━━━━━━━━━');
}

/**
 * eBay在庫0変更を1回だけ承認する
 *
 * 承認有効時間：10分
 * この関数自体はeBayやシートを変更しない
 */
function approveFbaToEbayZeroOnce() {
  const expiresAt =
    Date.now() + 10 * 60 * 1000;

  PropertiesService
    .getScriptProperties()
    .setProperty(
      'FBA_EBAY_ZERO_APPROVED_UNTIL',
      String(expiresAt)
    );

  Logger.log('━━━━━━━━━━━━━━━━━━━━━━━━━━━━');
  Logger.log('【eBay在庫0変更・1回承認】');
  Logger.log('承認有効時間：10分');
  Logger.log(
    '承認期限：' +
    Utilities.formatDate(
      new Date(expiresAt),
      'Asia/Tokyo',
      'yyyy/MM/dd HH:mm:ss'
    )
  );
  Logger.log(
    '次に applyFbaToEbayZero_CONFIRMED を実行してください'
  );
  Logger.log('━━━━━━━━━━━━━━━━━━━━━━━━━━━━');
}

/**
 * ============================================================
 * FBA在庫連動 eBay在庫0・自動実行
 *
 * フロー：
 * runInventorySyncAllAuto
 * ↓
 * 候補判定
 * ↓
 * eBay在庫0変更
 * ↓
 * eBay在庫変更履歴へ記録
 * ↓
 * 呼出元でLINE通知
 *
 * 有効化プロパティ：
 * FBA_EBAY_ZERO_AUTO_ENABLED = YES
 * ============================================================
 */

var FBA_ZERO_AUTO_CONFIG_ = {
  spreadsheetId: '1exGBAEx99-2Qc9d0DLRiZbygvgIo4FY_NoBG3Nkan-A',
  syncSheetName: '在庫同期管理',
  rawSheetName: 'eBayRaw',
  historySheetName: 'eBay在庫変更履歴',
  enableProperty: 'FBA_EBAY_ZERO_AUTO_ENABLED',

  // runInventorySyncAllAuto直後なので30分以内を要求
  maxDataAgeMinutes: 30,

  // 異常な大量変更を防止
  maxTargetCount: 5
};


/**
 * 自動在庫0処理を有効化する
 */
function enableFbaToEbayZeroAuto() {
  PropertiesService
    .getScriptProperties()
    .setProperty(
      FBA_ZERO_AUTO_CONFIG_.enableProperty,
      'YES'
    );

  Logger.log('✅ FBA→eBay在庫0自動処理を有効化しました');
}


/**
 * 自動在庫0処理を停止する
 */
function disableFbaToEbayZeroAuto() {
  PropertiesService
    .getScriptProperties()
    .deleteProperty(
      FBA_ZERO_AUTO_CONFIG_.enableProperty
    );

  Logger.log('⏸ FBA→eBay在庫0自動処理を停止しました');
}


/**
 * 自動処理の設定状態を確認する
 */
function checkFbaToEbayZeroAutoStatus() {
  const enabled =
    PropertiesService
      .getScriptProperties()
      .getProperty(
        FBA_ZERO_AUTO_CONFIG_.enableProperty
      ) === 'YES';

  Logger.log(
    'FBA→eBay在庫0自動処理: ' +
    (enabled ? '✅ 有効' : '⏸ 無効')
  );
}


/**
 * runInventorySyncAllAutoから呼び出す自動本番処理
 *
 * この関数単体をトリガーへ登録しない。
 */
function applyFbaToEbayZeroAuto_() {
  const enabled =
    PropertiesService
      .getScriptProperties()
      .getProperty(
        FBA_ZERO_AUTO_CONFIG_.enableProperty
      ) === 'YES';

  const emptyResult = {
    enabled: enabled,
    candidateCount: 0,
    successCount: 0,
    failureCount: 0,
    results: []
  };

  if (!enabled) {
    Logger.log(
      '⏸ eBay在庫0自動処理は無効です'
    );
    return emptyResult;
  }

  const ss = SpreadsheetApp.openById(
    FBA_ZERO_AUTO_CONFIG_.spreadsheetId
  );

  const syncSheet = ss.getSheetByName(
    FBA_ZERO_AUTO_CONFIG_.syncSheetName
  );

  const rawSheet = ss.getSheetByName(
    FBA_ZERO_AUTO_CONFIG_.rawSheetName
  );

  if (!syncSheet || !rawSheet) {
    throw new Error(
      '在庫同期管理またはeBayRawが見つかりません'
    );
  }

  const syncData =
    syncSheet.getDataRange().getValues();

  const rawData =
    rawSheet.getDataRange().getValues();

  if (syncData.length < 2) {
    throw new Error(
      '在庫同期管理にデータがありません'
    );
  }

  if (rawData.length < 2) {
    throw new Error(
      'eBayRawにデータがありません'
    );
  }

  const syncHeaders =
    syncData[0].map(function(value) {
      return String(value || '').trim();
    });

  const rawHeaders =
    rawData[0].map(function(value) {
      return String(value || '').trim();
    });

  const syncCol = {
    sku:
      syncHeaders.indexOf('SKU'),

    itemId:
      syncHeaders.indexOf('ItemID'),

    title:
      syncHeaders.indexOf('タイトル'),

    amazonQty:
      syncHeaders.indexOf('Amazon在庫'),

    ebayQty:
      syncHeaders.indexOf('eBay設定在庫'),

    warningQty:
      syncHeaders.indexOf('通知基準数'),

    notificationStatus:
      syncHeaders.indexOf('通知状態'),

    checkedAt:
      syncHeaders.indexOf('最終確認日時'),

    result:
      syncHeaders.indexOf('確認結果')
  };

  const rawCol = {
    itemId:
      rawHeaders.indexOf('ItemID'),

    sku:
      rawHeaders.indexOf('SKU'),

    availableQuantity:
      rawHeaders.indexOf('AvailableQuantity')
  };

  Object.keys(syncCol).forEach(function(key) {
    if (syncCol[key] === -1) {
      throw new Error(
        '在庫同期管理に必要な列がありません: ' +
        key
      );
    }
  });

  Object.keys(rawCol).forEach(function(key) {
    if (rawCol[key] === -1) {
      throw new Error(
        'eBayRawに必要な列がありません: ' +
        key
      );
    }
  });

  // eBayRawをItemIDで照合する
  const rawMap = {};

  for (let i = 1; i < rawData.length; i++) {
    const row = rawData[i];

    const itemId = String(
      row[rawCol.itemId] || ''
    ).trim();

    if (!itemId) {
      continue;
    }

    rawMap[itemId] = {
      sku: String(
        row[rawCol.sku] || ''
      ).trim(),

      availableQuantity:
        toEbayZeroNumber_(
          row[rawCol.availableQuantity]
        )
    };
  }

  const now = new Date();
  const targets = [];

  for (let i = 1; i < syncData.length; i++) {
    const row = syncData[i];

    const sku = String(
      row[syncCol.sku] || ''
    ).trim();

    const itemId = String(
      row[syncCol.itemId] || ''
    ).trim();

    const title = String(
      row[syncCol.title] || ''
    ).trim();

    const amazonQty =
      toEbayZeroNumber_(
        row[syncCol.amazonQty]
      );

    const ebayQty =
      toEbayZeroNumber_(
        row[syncCol.ebayQty]
      );

    const warningQty =
      toEbayZeroNumber_(
        row[syncCol.warningQty]
      );

    const resultText = String(
      row[syncCol.result] || ''
    ).trim();

    if (!sku || !itemId) {
      continue;
    }

    if (
      amazonQty === null ||
      ebayQty === null ||
      warningQty === null
    ) {
      continue;
    }

    // ルールB
    if (!(
      amazonQty <= warningQty &&
      ebayQty > 0
    )) {
      continue;
    }

    // SP-API取得成功行だけ処理
    if (
      resultText.indexOf('SP-API取得済') !== 0
    ) {
      throw new Error(
        '安全停止：SP-API取得失敗行です。' +
        ' SKU=' + sku +
        ' / 結果=' + resultText
      );
    }

    const checkedAt =
      parseEbayZeroDate_(
        row[syncCol.checkedAt]
      );

    if (!checkedAt) {
      throw new Error(
        '安全停止：最終確認日時不明。' +
        ' SKU=' + sku
      );
    }

    const ageMinutes =
      (
        now.getTime() -
        checkedAt.getTime()
      ) / 60000;

    if (
      ageMinutes < -5 ||
      ageMinutes >
        FBA_ZERO_AUTO_CONFIG_
          .maxDataAgeMinutes
    ) {
      throw new Error(
        '安全停止：在庫データが古い、' +
        'または未来日時です。' +
        ' SKU=' + sku +
        ' / 経過=' +
        Math.floor(ageMinutes) +
        '分'
      );
    }

    if (!/^\d{10,20}$/.test(itemId)) {
      throw new Error(
        '安全停止：ItemID不正。' +
        ' SKU=' + sku +
        ' / ItemID=' + itemId
      );
    }

    const raw = rawMap[itemId];

    if (!raw) {
      throw new Error(
        '安全停止：eBayRawにItemIDなし。' +
        ' SKU=' + sku
      );
    }

    if (
      normalizeSku_(raw.sku) !==
      normalizeSku_(sku)
    ) {
      throw new Error(
        '安全停止：SKU不一致。' +
        ' 在庫同期管理=' + sku +
        ' / eBayRaw=' + raw.sku
      );
    }

    if (
      raw.availableQuantity === null ||
      raw.availableQuantity !== ebayQty
    ) {
      throw new Error(
        '安全停止：eBay在庫数不一致。' +
        ' SKU=' + sku +
        ' / 在庫同期管理=' + ebayQty +
        ' / eBayRaw=' +
        raw.availableQuantity
      );
    }

    targets.push({
      rowNumber: i + 1,
      sku: sku,
      itemId: itemId,
      title: title,
      amazonQty: amazonQty,
      ebayQtyBefore: ebayQty,
      warningQty: warningQty
    });
  }

  emptyResult.candidateCount =
    targets.length;

  if (targets.length === 0) {
    Logger.log(
      '✅ eBay在庫0自動変更対象なし'
    );
    return emptyResult;
  }

  if (
    targets.length >
    FBA_ZERO_AUTO_CONFIG_.maxTargetCount
  ) {
    throw new Error(
      '安全停止：自動変更対象が多すぎます。' +
      ' 対象=' + targets.length + '件'
    );
  }

  const accessToken =
    getEbayAccessTokenForFbaZeroAuto_();

  if (
    !isEbayOutOfStockControlEnabled_(
      accessToken
    )
  ) {
    throw new Error(
      '安全停止：eBayのOut-of-Stock設定が無効です'
    );
  }

  const runId =
    Utilities.getUuid();

  const results = [];

  targets.forEach(function(target) {
    const apiResult =
      reviseEbayAvailableToZero_(
        target.itemId,
        accessToken
      );

    const result = {
      runId: runId,
      target: target,
      success: apiResult.success === true,
      ack: apiResult.ack || '',
      error: apiResult.error || ''
    };

    if (result.success) {
      // E列：eBay設定在庫
      syncSheet
        .getRange(
          target.rowNumber,
          syncCol.ebayQty + 1
        )
        .setValue(0);

      // G列：通知状態
      syncSheet
        .getRange(
          target.rowNumber,
          syncCol.notificationStatus + 1
        )
        .setValue(
          'eBay在庫0自動変更済'
        );

      // I列：確認結果
      syncSheet
        .getRange(
          target.rowNumber,
          syncCol.result + 1
        )
        .setValue(
          'eBay在庫0自動変更成功' +
          '（Amazon=' +
          target.amazonQty +
          ' / 基準=' +
          target.warningQty +
          '）'
        );

      Logger.log(
        '✅ eBay在庫0自動変更成功: ' +
        target.sku
      );

    } else {
      syncSheet
        .getRange(
          target.rowNumber,
          syncCol.notificationStatus + 1
        )
        .setValue(
          'eBay在庫0自動変更エラー'
        );

      syncSheet
        .getRange(
          target.rowNumber,
          syncCol.result + 1
        )
        .setValue(
          'eBay在庫0自動変更失敗: ' +
          String(result.error)
            .substring(0, 300)
        );

      Logger.log(
        '❌ eBay在庫0自動変更失敗: ' +
        target.sku +
        ' / ' +
        result.error
      );
    }

    results.push(result);

    Utilities.sleep(1000);
  });

  SpreadsheetApp.flush();

  appendFbaZeroAutoHistory_(
    ss,
    results
  );

  return {
    enabled: true,
    candidateCount: targets.length,

    successCount:
      results.filter(function(result) {
        return result.success;
      }).length,

    failureCount:
      results.filter(function(result) {
        return !result.success;
      }).length,

    results: results
  };
}


/**
 * eBayトークン取得
 */
function getEbayAccessTokenForFbaZeroAuto_() {
  let accessToken = '';

  if (
    typeof getNewAccessToken ===
    'function'
  ) {
    accessToken =
      getNewAccessToken();

  } else if (
    typeof
      getEbayAccessTokenFromRefreshToken ===
    'function'
  ) {
    accessToken =
      getEbayAccessTokenFromRefreshToken();
  }

  if (!accessToken) {
    throw new Error(
      'eBayアクセストークンを取得できません'
    );
  }

  return accessToken;
}


/**
 * Out-of-Stock設定確認
 */
function isEbayOutOfStockControlEnabled_(
  accessToken
) {
  const xml =
    '<?xml version="1.0" encoding="UTF-8"?>' +
    '<GetUserPreferencesRequest ' +
    'xmlns="urn:ebay:apis:eBLBaseComponents">' +
      '<ShowOutOfStockControlPreference>true' +
      '</ShowOutOfStockControlPreference>' +
    '</GetUserPreferencesRequest>';

  const response =
    UrlFetchApp.fetch(
      'https://api.ebay.com/ws/api.dll',
      {
        method: 'post',
        headers: {
          'X-EBAY-API-IAF-TOKEN':
            accessToken,

          'X-EBAY-API-CALL-NAME':
            'GetUserPreferences',

          'X-EBAY-API-SITEID':
            '0',

          'X-EBAY-API-COMPATIBILITY-LEVEL':
            '1199',

          'Content-Type':
            'text/xml'
        },
        payload: xml,
        muteHttpExceptions: true
      }
    );

  const body =
    response.getContentText();

  const ack =
    (
      body.match(
        /<Ack>([^<]+)<\/Ack>/
      ) || []
    )[1] || '';

  const preference =
    (
      body.match(
        /<OutOfStockControlPreference>([^<]+)<\/OutOfStockControlPreference>/
      ) || []
    )[1] || '';

  if (
    ack !== 'Success' &&
    ack !== 'Warning'
  ) {
    throw new Error(
      'GetUserPreferences失敗: ' +
      body.substring(0, 500)
    );
  }

  return (
    String(preference)
      .toLowerCase() === 'true'
  );
}


/**
 * 在庫変更履歴へ追記
 */
function appendFbaZeroAutoHistory_(
  ss,
  results
) {
  if (!results || results.length === 0) {
    return;
  }

  let sheet = ss.getSheetByName(
    FBA_ZERO_AUTO_CONFIG_.historySheetName
  );

  if (!sheet) {
    sheet = ss.insertSheet(
      FBA_ZERO_AUTO_CONFIG_.historySheetName
    );

    sheet.appendRow([
      'run_id',
      '処理日時',
      '結果',
      'SKU',
      'ItemID',
      '商品名',
      'Amazon在庫',
      '通知基準数',
      '変更前eBay在庫',
      '変更後eBay在庫',
      'Ack',
      'エラー',
      '実行元'
    ]);

    sheet.setFrozenRows(1);
  }

  const now = new Date();

  const rows = results.map(function(result) {
    const target = result.target;

    return [
      result.runId,
      now,
      result.success ? '成功' : '失敗',
      target.sku,
      target.itemId,
      target.title,
      target.amazonQty,
      target.warningQty,
      target.ebayQtyBefore,
      result.success
        ? 0
        : target.ebayQtyBefore,
      result.ack,
      result.error,
      'runInventorySyncAllAuto'
    ];
  });

  sheet
    .getRange(
      sheet.getLastRow() + 1,
      1,
      rows.length,
      rows[0].length
    )
    .setValues(rows);
}


/**
 * 在庫0自動変更結果のLINE本文
 */
function buildFbaZeroAutoLineMessage_(
  result
) {
  let message =
    '📦 FBA連動 eBay在庫0処理\n' +
    '━━━━━━━━━━━━━━━\n' +
    '候補：' +
    result.candidateCount +
    '件\n' +
    '成功：' +
    result.successCount +
    '件\n' +
    '失敗：' +
    result.failureCount +
    '件\n\n';

  result.results.forEach(function(item) {
    const target = item.target;

    message +=
      (item.success ? '✅ ' : '❌ ') +
      target.sku +
      '\n' +
      'Amazon：' +
      target.amazonQty +
      ' / 基準：' +
      target.warningQty +
      ' / eBay：' +
      target.ebayQtyBefore +
      ' → ' +
      (item.success
        ? '0'
        : target.ebayQtyBefore) +
      '\n';

    if (!item.success) {
      message +=
        'エラー：' +
        String(item.error)
          .substring(0, 150) +
        '\n';
    }

    message += '\n';
  });

  message +=
    '時刻：' +
    Utilities.formatDate(
      new Date(),
      'Asia/Tokyo',
      'MM/dd HH:mm'
    );

  return message;
}/**
 * InventoryAutoRun.gs  （自動運用v6・追加ファイル）
 * ===============================================================
 * 在庫同期システムの「完全自動運用」関数群。
 * SP-API在庫取得v4（updateFbaInventoryFromSpApi 等）とは別ファイル。
 * v4側には一切触れず、ここで呼び出すだけ。
 *
 * 含まれる関数：
 *   runInventorySyncAllAuto()       … 自動一括実行＋自己診断＋成否LINE
 *   validateInventorySyncResult_()  … 在庫同期管理シートの自己診断
 *   setupInventorySyncAutoTrigger() … 毎日17:10頃JSTのトリガー設定
 *   listInventoryRelatedTriggers()  … 旧方式トリガーの残存確認（削除はしない）
 *
 * 既存関数を呼び出すだけ（再定義しない）：
 *   updateEbayRawFromTradingApi() / updateStockSyncFromEbayRaw()
 *   updateFbaInventoryFromSpApi() / checkFbaLowStockAndNotify1() / sendLine()
 *
 * 【定数の衝突回避】
 *   v4ファイルが SHEET_STOCK_SYNC / SKU_EXCLUDE 等を var 宣言しているため、
 *   ここで同名を再宣言すると「already been declared」で保存不可になる。
 *   そのため、このファイル独自の AUTO_ 接頭辞付き定数を使用する。
 *
 * 禁止：setEbayQuantityZero / ReviseFixedPriceItem / checkFbaInventory
 *   → eBay在庫は一切変更しない（取得・判定・LINE通知のみ）
 * ===============================================================
 */

// ===== このファイル専用の定数（v4と名前が衝突しないよう AUTO_ 接頭辞）=====
var AUTO_SHEET_STOCK_SYNC = '在庫同期管理';

var AUTO_COL_SKU       = 1; // A
var AUTO_COL_AMZ_QTY   = 4; // D Amazon在庫
var AUTO_COL_EBAY_QTY  = 5; // E eBay設定在庫
var AUTO_COL_THRESHOLD = 6; // F 通知基準数
var AUTO_COL_LASTCHECK = 8; // H 最終確認日時
var AUTO_COL_RESULT    = 9; // I 確認結果

var AUTO_SKU_EXCLUDE = ['SKU', 'WAREHOUSES', 'DELETED', 'ITEMID', 'TITLE', 'AVAILABLEQUANTITY'];

// テスト用SKU（在庫同期の対象から除外する）。
// 完全一致のほか、これらを含むSKU（例: SAMPLE-001, TEST_AAA）も除外する。
var AUTO_TEST_SKU_TOKENS = ['SAMPLE', 'TEST', 'DUMMY'];

// 旧方式トリガー（残存していたら警告）
var AUTO_LEGACY_TRIGGER_FUNCS = [
  'checkFbaInventory',
  'setupInventoryTrigger',
  'setupInventorySyncTrigger',
  'setEbayQuantityZero'
];


/**
 * 自動一括実行
 *
 * eBayRaw取得
 * ↓
 * 在庫同期管理更新
 * ↓
 * SP-API在庫取得
 * ↓
 * eBay在庫0自動変更
 * ↓
 * 履歴記録
 * ↓
 * LINE通知
 */
function runInventorySyncAllAuto() {
  var lock =
    LockService.getScriptLock();

  if (!lock.tryLock(10 * 1000)) {
    Logger.log(
      '別の実行が進行中のためスキップ'
    );
    return;
  }

  try {
    _autoRunStep_(
      'eBayRaw取得失敗',
      updateEbayRawFromTradingApi
    );

    _autoRunStep_(
      '在庫同期更新失敗',
      updateStockSyncFromEbayRaw
    );

    _autoRunStep_(
      'FBA貼付在庫取得失敗',
      updateStockSyncDFromFbaPaste_
    );

        // 新規：eBay在庫0自動変更
    var zeroResult =
      _autoRunStep_(
        'eBay在庫0自動変更失敗',
        applyFbaToEbayZeroAuto_
      );

    var v =
      validateInventorySyncResult_();

if (v.sysAbnormal) {
      sendLine(
        _autoBuildSystemErrorMessage_(v)
      );

    } else if (
      zeroResult &&
      (
        zeroResult.successCount > 0 ||
        zeroResult.failureCount > 0
      )
    ) {
      sendLine(
        buildFbaZeroAutoLineMessage_(
          zeroResult
        )
      );

    
} else if (v.hasStockAlert) {
      sendLine(
        _autoBuildStockAlertMessage_(v)
      );

    } else {

_autoClearNotify_('STOCK');

      // ✅ 変更：LINE通知を止め、ログ出力のみにする
      Logger.log(
        'ℹ️ 在庫同期はすべて正常（変更・アラートなし）のため、LINE通知をスキップしました。'
      );
      Logger.log(
        _autoBuildOkMessage_(v)
      );
    }

  } catch (e) {
    var stage =
      e && e._stage
        ? e._stage
        : 'GAS実行エラー';

    var emsg =
      e && e.message
        ? e.message
        : String(e);

    var msg =
      '❌ 在庫同期 システム異常\n\n' +
      '【異常内容】\n' +
      stage +
      '\nエラー: ' +
      emsg +
      '\n\n時刻: ' +
      _autoNowStr_();

    try {
      sendLine(msg);
    } catch (e2) {
      Logger.log(
        'LINE送信失敗: ' + e2
      );
    }

    Logger.log(
      'runInventorySyncAllAuto エラー[' +
      stage +
      ']: ' +
      e
    );

  } finally {
    lock.releaseLock();
  }
}


/**
 * ステップ実行ラッパー
 *
 * 戻り値を呼出元へ返す。
 */
function _autoRunStep_(
  stageLabel,
  fn
) {
  try {
    return fn();

  } catch (e) {
    var wrapped =
      new Error(
        e && e.message
          ? e.message
          : String(e)
      );

    wrapped._stage =
      stageLabel;

    throw wrapped;
  }
}


/**
 * 【2】在庫同期管理シートの自己診断
 * 対象：eBay設定在庫(E) >= 1 のSKU
 *
 * システム異常: 未発見 / Amazon在庫空欄 / 確認結果空欄 / 最終確認日時空欄
 * 在庫アラート: Amazon在庫0 / 通知基準以下 / eBay設定在庫 > Amazon在庫
 * （在庫0・在庫少・eBay超過は NG扱いにしない＝システムは正常）
 */
function validateInventorySyncResult_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(AUTO_SHEET_STOCK_SYNC);

  var v = {
    sysAbnormal: false,
    hasStockAlert: false,
    fatal: null,
    counts: {
      target: 0, got: 0, single: 0,
      missing: 0, amzBlank: 0, resultBlank: 0, lastCheckBlank: 0,
      amzZero: 0, lowStock: 0, ebayOver: 0
    },
    sysIssues: [],
    stockIssues: []
  };

  if (!sheet) {
    v.sysAbnormal = true;
    v.fatal = 'シートが見つかりません: ' + AUTO_SHEET_STOCK_SYNC;
    return v;
  }
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return v;

  var numRows = lastRow - 1;
  var data = sheet.getRange(2, 1, numRows, 9).getValues();

  for (var i = 0; i < data.length; i++) {
    var row = data[i];
    var sku = String(row[AUTO_COL_SKU - 1] == null ? '' : row[AUTO_COL_SKU - 1]).trim();
    if (!sku) continue;
    if (AUTO_SKU_EXCLUDE.indexOf(sku.toUpperCase()) !== -1) continue;
    if (_autoIsTestSku_(sku)) continue;   // SAMPLE / TEST / DUMMY 等は対象外

    var ebayQty = _autoToNum_(row[AUTO_COL_EBAY_QTY - 1]);
    if (!(ebayQty >= 1)) continue;

    v.counts.target++;

    var amzRaw    = row[AUTO_COL_AMZ_QTY - 1];
    var amzBlank  = (amzRaw === '' || amzRaw === null || amzRaw === undefined);
    var amzQty    = _autoToNum_(amzRaw);
    var threshold = _autoToNum_(row[AUTO_COL_THRESHOLD - 1]);
    var checkRes  = String(row[AUTO_COL_RESULT - 1] == null ? '' : row[AUTO_COL_RESULT - 1]).trim();
    var lastCheck = row[AUTO_COL_LASTCHECK - 1];

    if (checkRes === 'SP-API取得済') v.counts.got++;
    else if (checkRes === 'SP-API取得済（単体補完）') { v.counts.got++; v.counts.single++; }

    // ===== システム異常系 =====
    if (checkRes === 'SP-API未発見') {
      v.counts.missing++;
      _autoPush_(v.sysIssues, sku + ' / SP-API未発見 / eBay在庫=' + ebayQty);
    }
    if (checkRes === '') {
      v.counts.resultBlank++;
      _autoPush_(v.sysIssues, sku + ' / 確認結果が空欄');
    }
    if (lastCheck === '' || lastCheck === null) {
      v.counts.lastCheckBlank++;
      _autoPush_(v.sysIssues, sku + ' / 最終確認日時が空欄');
    }
    if (amzBlank) {
      v.counts.amzBlank++;
      _autoPush_(v.sysIssues, sku + ' / Amazon在庫が空欄');
    }

    // ===== 在庫アラート系（Amazon在庫が数値の場合のみ）=====
    if (!amzBlank && !isNaN(amzQty)) {
      if (amzQty === 0) {
        v.counts.amzZero++;
        _autoPush_(v.stockIssues, sku + ' / Amazon在庫0 / eBay在庫=' + ebayQty);
      } else if (threshold >= 1 && amzQty <= threshold) {
        v.counts.lowStock++;
        _autoPush_(v.stockIssues, sku + ' / 在庫少 Amazon=' + amzQty + ' / 基準=' + threshold);
      }
      if (ebayQty > amzQty) {
        v.counts.ebayOver++;
        _autoPush_(v.stockIssues, sku + ' / eBay在庫超過 eBay=' + ebayQty + ' > Amazon=' + amzQty);
      }
    }
  }

  var c = v.counts;
  v.sysAbnormal   = (c.missing > 0 || c.amzBlank > 0 || c.resultBlank > 0 || c.lastCheckBlank > 0);
  v.hasStockAlert = (c.amzZero > 0 || c.lowStock > 0 || c.ebayOver > 0);

  return v;
}


/**
 * 【3】毎日17:10頃JSTのトリガー作成（既存同名トリガーは削除してから）
 */
function setupInventorySyncAutoTrigger() {
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === 'runInventorySyncAllAuto') {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }

  ScriptApp.newTrigger('runInventorySyncAllAuto')
    .timeBased()
    .everyDays(1)
    .atHour(17)
    .nearMinute(10)
    .inTimezone('Asia/Tokyo')
    .create();

  sendLine('✅ 在庫同期 自動トリガー設定完了\n'
    + '毎日17:10頃に自動実行します。\n\n'
    + '今後、人間がスプレッドシートを開いて確認する必要はありません。');

  Logger.log('トリガー設定完了: runInventorySyncAllAuto 毎日17:10頃 JST');
}


/**
 * 【4】旧方式トリガーの残存確認（削除はしない・警告のみ）
 */
function listInventoryRelatedTriggers() {
  var triggers = ScriptApp.getProjectTriggers();
  var all = [];
  var legacyFound = [];

  for (var i = 0; i < triggers.length; i++) {
    var fn = triggers[i].getHandlerFunction();
    all.push(fn);
    if (AUTO_LEGACY_TRIGGER_FUNCS.indexOf(fn) !== -1) {
      legacyFound.push(fn);
    }
  }

  Logger.log('現在のトリガー一覧: ' + (all.length ? all.join(', ') : '（なし）'));

  if (legacyFound.length > 0) {
    var msg = '⚠️ 旧方式トリガー残存警告\n\n'
      + '以下の旧トリガーが残っています。\n'
      + '意図せずeBay在庫が変更される恐れがあります。\n\n'
      + legacyFound.join('\n') + '\n\n'
      + 'GASエディタ「トリガー」画面から手動で削除してください。\n'
      + '（このスクリプトは自動削除しません）';
    try { sendLine(msg); } catch (e) { Logger.log('LINE送信失敗: ' + e); }
    Logger.log('旧トリガー検出: ' + legacyFound.join(', '));
  } else {
    Logger.log('旧方式トリガーは見つかりませんでした（クリーン）');
  }

  return { all: all, legacy: legacyFound };
}


// ===== LINE文面（このファイル専用・_auto 接頭辞）=====

function _autoBuildOkMessage_(v) {
  var c = v.counts;
  return '✅ 在庫同期 自動確認OK\n'
    + 'システム状態: 正常\n\n'
    + '対象SKU: ' + c.target + '件\n'
    + 'SP-API取得済: ' + c.got + '件（単体補完 ' + c.single + '件）\n'
    + '在庫アラート: なし\n\n'
    + '時刻: ' + _autoNowStr_();
}

function _autoBuildStockAlertMessage_(v) {
  var c = v.counts;
  return '📦 在庫アラートあり\n'
    + 'システム状態: 正常\n\n'
    + 'Amazon在庫0: ' + c.amzZero + '件\n'
    + '在庫少: ' + c.lowStock + '件\n'
    + 'eBay在庫超過: ' + c.ebayOver + '件\n\n'
    + '【対象SKU 上位10件】\n'
    + _autoTopList_(v.stockIssues, 10) + '\n\n'
    + '時刻: ' + _autoNowStr_();
}

function _autoBuildSystemErrorMessage_(v) {
  if (v.fatal) {
    return '❌ 在庫同期 システム異常\n\n【異常内容】\n' + v.fatal + '\n\n時刻: ' + _autoNowStr_();
  }
  var c = v.counts;
  var lines = [];
  if (c.missing > 0)        lines.push('SP-API未発見: ' + c.missing + '件');
  if (c.amzBlank > 0)       lines.push('Amazon在庫空欄: ' + c.amzBlank + '件');
  if (c.resultBlank > 0)    lines.push('確認結果空欄: ' + c.resultBlank + '件');
  if (c.lastCheckBlank > 0) lines.push('最終確認日時空欄: ' + c.lastCheckBlank + '件');

  return '❌ 在庫同期 システム異常\n\n'
    + '【異常内容】\n'
    + lines.join('\n') + '\n\n'
    + '【対象SKU 上位10件】\n'
    + _autoTopList_(v.sysIssues, 10) + '\n\n'
    + '時刻: ' + _autoNowStr_();
}

function _autoTopList_(arr, n) { return arr.slice(0, n).join('\n'); }
function _autoPush_(arr, line) { if (arr.indexOf(line) === -1) arr.push(line); }

/** SAMPLE / TEST / DUMMY 等のテストSKUか判定（部分一致・大文字小文字無視） */
function _autoIsTestSku_(sku) {
  var u = String(sku == null ? '' : sku).toUpperCase();
  for (var i = 0; i < AUTO_TEST_SKU_TOKENS.length; i++) {
    if (u.indexOf(AUTO_TEST_SKU_TOKENS[i]) !== -1) return true;
  }
  return false;
}

function _autoToNum_(v) {
  if (v === '' || v === null || v === undefined) return NaN;
  var n = Number(v);
  return isNaN(n) ? NaN : n;
}

function _autoNowStr_() {
  return Utilities.formatDate(new Date(), 'Asia/Tokyo', 'MM/dd HH:mm');
}
/**
 * MCF_LineNotification.gs  v3
 *
 * MCF自動化まわりの LINE通知管理システム
 *
 * ★ 既存の sendLine / pushLine 系関数は削除しない
 * ★ 不要な通知呼び出しは削除・抑制
 * ★ 通知は「人が判断・対応すべきタイミング」のみ
 * ★ MCF_LINE_NOTIFICATION_LOG で重複通知を防止
 * ★ dryRunCheckMcfLineNotifications() で実際の送信前に動作確認可能
 *
 * v3 変更点（v2からの差分）：
 *   - MCF_NOT_CREATED_THRESHOLD_HOURS_ を 12 → 14 に変更
 *
 *   【変更理由】
 *     MCF作成自体を受注から12時間後に行うため（MCF_CreationDelay12h.gs と連動）、
 *     「MCF未作成アラート」の通知タイミングも合わせて遅らせる。
 *
 *   【タイムライン】
 *     受注
 *      ├─  0〜12h : WAIT_12H（MCF作成待機 - キャンセル余地期間）
 *      ├─ 12h〜   : MCF作成タイミング（autoCreateMcfOrders が実行される）
 *      └─ 14h〜   : P列が空欄なら MCF_NOT_CREATED LINE通知
 *                  （MCF作成後2時間以上経過してもP列未記入のケースをアラート）
 *
 * 【通知する種別】
 *   EBAY_ORDER_NEW       新規eBay受注検知
 *   MCF_CREATE_FAILED    MCF作成失敗
 *   MCF_NOT_CREATED      MCF未作成アラート（受注から14時間以上経過・P列空欄）
 *   DEADLINE_RISK        発送期限24時間以内・Tracking未取得（1日1回）
 *   TRACKING_OBTAINED    追跡番号取得（新規発行時）
 *   EBAY_TRACKING_FAILED eBay追跡番号提出失敗
 *
 * 【通知しない種別】
 *   dry-run成功 / MCF作成成功のみ / 追跡番号未発行（PENDING/Processing）
 *   ログ記録成功 / エラー0件 / 読み取り専用チェック成功
 *   受注から14時間未満（WAIT_12H期間 + MCF作成処理中）
 */

// ============================================================
// ■ 定数
// ============================================================

const MCF_LINE_LOG_SHEET_NAME_ = 'MCF_LINE_NOTIFICATION_LOG';

// 通知対象イベント種別
const MCF_NOTIFY_EVENT_TYPES_ = {
  EBAY_ORDER_NEW       : 'EBAY_ORDER_NEW',
  MCF_CREATE_FAILED    : 'MCF_CREATE_FAILED',
  MCF_NOT_CREATED      : 'MCF_NOT_CREATED',
  DEADLINE_RISK        : 'DEADLINE_RISK',
  TRACKING_OBTAINED    : 'TRACKING_OBTAINED',
  EBAY_TRACKING_FAILED : 'EBAY_TRACKING_FAILED'
};

// eBay受注ログ 列インデックス（0始まり）
const MCF_NOTIFY_COL_ = {
  RECEIVED_AT  :  0,  // A: 受注日時
  ORDER_ID     :  1,  // B: eBay注文番号
  SKU          :  2,  // C: SKU
  PRODUCT_NAME :  3,  // D: 商品名
  QUANTITY     :  4,  // E: 数量
  SHIP_BY      : 14,  // O: 発送期限
  MCF_ID       : 15,  // P: MCF作成
  STATUS       : 16,  // Q: ステータス
  CARRIER      : 17,  // R: Carrier
  TRACKING     : 18   // S: Tracking
};

// ★ v3変更点：12 → 14
// MCF作成が受注12h後のため、アラートは受注14h後（MCF作成から2h後）に送信
// 0〜12h：WAIT_12H（MCF作成待機）
// 12h〜 ：MCF作成タイミング
// 14h〜P列空欄：このLINE通知の対象
const MCF_NOT_CREATED_THRESHOLD_HOURS_ = 14;

// 発送期限リスクのしきい値（期限までの残り時間）
const DEADLINE_RISK_THRESHOLD_HOURS_ = 24;

// MCF_LINE_NOTIFICATION_LOG 列インデックス（1始まり・getRange用）
const MCF_LOG_COL_ = {
  NOTIFIED_AT  : 1,  // A: 通知日時
  EVENT_TYPE   : 2,  // B: 通知種別
  ORDER_ID     : 3,  // C: eBay注文番号
  SKU          : 4,  // D: SKU
  NOTIFY_KEY   : 5,  // E: 通知キー
  MESSAGE      : 6,  // F: 本文
  RESULT       : 7   // G: 送信結果
};

// ============================================================
// ■ 0. ensureMcfLineLogSheet_（共通）
//       MCF_LINE_NOTIFICATION_LOG シートを取得または新規作成する
// ============================================================

function ensureMcfLineLogSheet_(ss) {
  let logSheet = ss.getSheetByName(MCF_LINE_LOG_SHEET_NAME_);

  if (!logSheet) {
    logSheet = ss.insertSheet(MCF_LINE_LOG_SHEET_NAME_);
    logSheet.setTabColor('#4FC3F7');

    logSheet.appendRow([
      '通知日時',     // A
      '通知種別',     // B
      'eBay注文番号', // C
      'SKU',          // D
      '通知キー',     // E
      '本文',         // F
      '送信結果'      // G
    ]);

    const hdr = logSheet.getRange(1, 1, 1, 7);
    hdr.setBackground('#1a73e8');
    hdr.setFontColor('#ffffff');
    hdr.setFontWeight('bold');
    logSheet.setFrozenRows(1);
    logSheet.setColumnWidth(1, 160);
    logSheet.setColumnWidth(2, 160);
    logSheet.setColumnWidth(3, 160);
    logSheet.setColumnWidth(4, 140);
    logSheet.setColumnWidth(5, 280);
    logSheet.setColumnWidth(6, 400);
    logSheet.setColumnWidth(7, 100);
  }

  return logSheet;
}


// ============================================================
// ■ 2. shouldNotifyMcfEvent_
//       通知対象かどうかを判定（重複チェック含む）
// ============================================================

function shouldNotifyMcfEvent_(eventType, orderId, notifyKey) {
  const today = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyyMMdd');

  const allowedTypes = Object.values(MCF_NOTIFY_EVENT_TYPES_);
  if (!allowedTypes.includes(eventType)) {
    return { shouldNotify: false, notifyKey: '', reason: '通知対象外のイベント種別: ' + eventType };
  }

  if (!notifyKey) {
    switch (eventType) {
      case MCF_NOTIFY_EVENT_TYPES_.EBAY_ORDER_NEW:
        notifyKey = 'EBAY_ORDER_NEW_' + orderId; break;
      case MCF_NOTIFY_EVENT_TYPES_.MCF_CREATE_FAILED:
        notifyKey = 'MCF_CREATE_FAILED_' + orderId; break;
      case MCF_NOTIFY_EVENT_TYPES_.MCF_NOT_CREATED:
        notifyKey = 'MCF_NOT_CREATED_' + orderId + '_' + today; break;
      case MCF_NOTIFY_EVENT_TYPES_.DEADLINE_RISK:
        notifyKey = 'DEADLINE_RISK_' + orderId + '_' + today; break;
      case MCF_NOTIFY_EVENT_TYPES_.TRACKING_OBTAINED:
        notifyKey = 'TRACKING_OBTAINED_' + orderId; break;
      case MCF_NOTIFY_EVENT_TYPES_.EBAY_TRACKING_FAILED:
        notifyKey = 'EBAY_TRACKING_FAILED_' + orderId; break;
      default:
        notifyKey = eventType + '_' + orderId + '_' + today;
    }
  }

  try {
    const ss       = SpreadsheetApp.getActiveSpreadsheet();
    const logSheet = ss.getSheetByName(MCF_LINE_LOG_SHEET_NAME_);
    if (logSheet && logSheet.getLastRow() > 1) {
      const logData = logSheet.getDataRange().getValues();
      for (let i = 1; i < logData.length; i++) {
        const existingKey    = String(logData[i][MCF_LOG_COL_.NOTIFY_KEY - 1] || '').trim();
        const existingResult = String(logData[i][MCF_LOG_COL_.RESULT - 1]    || '').trim();
        if (existingKey === notifyKey && existingResult === 'SENT') {
          return {
            shouldNotify: false,
            notifyKey   : notifyKey,
            reason      : '重複: 同じ通知キーが既にログに存在します（' + existingResult + '）'
          };
        }
      }
    }
  } catch (e) {
    Logger.log('⚠️ shouldNotifyMcfEvent_: ログシート検索エラー: ' + e.toString());
  }

  return { shouldNotify: true, notifyKey: notifyKey, reason: '通知対象です' };
}

// ============================================================
// ■ 3. logMcfLineNotification_
//       MCF_LINE_NOTIFICATION_LOG に個別通知を記録
// ============================================================

function logMcfLineNotification_(eventType, orderId, sku, notifyKey, message, result) {
  try {
    const ss       = SpreadsheetApp.getActiveSpreadsheet();
    const logSheet = ensureMcfLineLogSheet_(ss);

    const now = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy/MM/dd HH:mm:ss');
    logSheet.appendRow([now, eventType, orderId, sku, notifyKey, message, result]);

    const lastRow    = logSheet.getLastRow();
    const resultCell = logSheet.getRange(lastRow, MCF_LOG_COL_.RESULT);
    if (result === 'SENT') {
      resultCell.setBackground('#e6f4ea').setFontColor('#137333');
    } else if (result.startsWith('ERROR')) {
      resultCell.setBackground('#fce8e6').setFontColor('#c5221f');
    } else if (result === 'DRY_RUN') {
      resultCell.setBackground('#fef7e0').setFontColor('#e37400');
    } else if (result === 'SKIPPED') {
      resultCell.setBackground('#f1f3f4').setFontColor('#5f6368');
    }
  } catch (e) {
    Logger.log('⚠️ logMcfLineNotification_: ログ記録エラー: ' + e.toString());
  }
}

// ============================================================
// ■ 3b. logMcfDryRunSummary_
//        実行後のサマリー行を必ず1行記録する
// ============================================================

function logMcfDryRunSummary_(checkedCount, notifyCount, skipCount) {
  try {
    const ss       = SpreadsheetApp.getActiveSpreadsheet();
    const logSheet = ensureMcfLineLogSheet_(ss);

    const now       = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy/MM/dd HH:mm:ss');
    const resultMsg = (notifyCount === 0)
      ? '結果：通知対象なし'
      : '結果：通知対象 ' + notifyCount + ' 件';

    logSheet.appendRow([
      now,
      'DRY_RUN_SUMMARY',
      String(checkedCount),
      String(notifyCount),
      String(skipCount),
      resultMsg,
      'DRY_RUN'
    ]);

    const lastRow = logSheet.getLastRow();
    logSheet.getRange(lastRow, 1, 1, 7).setBackground('#e8f0fe');
    logSheet.getRange(lastRow, MCF_LOG_COL_.RESULT)
      .setBackground('#fef7e0').setFontColor('#e37400');

    const resultCell = logSheet.getRange(lastRow, MCF_LOG_COL_.MESSAGE);
    if (notifyCount === 0) {
      resultCell.setFontColor('#5f6368');
    } else {
      resultCell.setFontColor('#1a73e8').setFontWeight('bold');
    }

  } catch (e) {
    Logger.log('⚠️ logMcfDryRunSummary_: ログ記録エラー: ' + e.toString());
  }
}

// ============================================================
// ■ 4. pushLineMcf_
//       LINE Messaging API push（内部ラッパー）
//       ★ 既存の sendLine 系関数があれば置き換え可能
// ============================================================

function pushLineMcf_(message) {
  try {
    const props  = PropertiesService.getScriptProperties();
    const token =
  sampleCsvProps_().getProperty('LINE_CHANNEL_ACCESS_TOKEN') ||
  sampleCsvProps_().getProperty('LINE_CHANNEL_TOKEN');
    const userId = sampleCsvProps_().getProperty('LINE_USER_ID');

    if (!token || !userId) {
      throw new Error('LINE_CHANNEL_ACCESS_TOKEN または LINE_USER_ID が Script Properties に設定されていません');
    }

    const payload = { to: userId, messages: [{ type: 'text', text: message }] };

    const response = UrlFetchApp.fetch('https://api.line.me/v2/bot/message/push', {
      method            : 'POST',
      headers           : { 'Content-Type': 'application/json', 'Authorization': 'Bearer ' + token },
      payload           : JSON.stringify(payload),
      muteHttpExceptions: true
    });

    const code = response.getResponseCode();
    if (code === 200) { return true; }
    Logger.log('❌ LINE push 失敗（HTTP ' + code + '）: ' + response.getContentText());
    return false;

  } catch (e) {
    Logger.log('❌ pushLineMcf_ エラー: ' + e.toString());
    return false;
  }
}

// ============================================================
// ■ 5. buildMcfLineMessage_
//       イベント種別に応じたLINEメッセージ本文を生成
// ============================================================

function buildMcfLineMessage_(eventType, order) {
  const orderId     = order.orderId                          || '';
  const sku         = order.sku                              || '';
  const productName = (order.productName || '').substring(0, 40);
  const quantity    = order.quantity                         || '';
  const shipByDate  = order.shipByDate                       || '';
  const mcfId       = order.mcfId                            || '';
  const carrier     = order.carrier                          || '';
  const tracking    = order.tracking                         || '';
  const errorInfo   = order.errorInfo                        || '';

  switch (eventType) {
    case MCF_NOTIFY_EVENT_TYPES_.EBAY_ORDER_NEW:
      return ['【eBay受注】新規受注','注文番号: '+orderId,'SKU: '+sku,'商品: '+productName,'数量: '+quantity,'発送期限: '+shipByDate].join('\n');
    case MCF_NOTIFY_EVENT_TYPES_.MCF_CREATE_FAILED:
      return ['【MCF作成失敗】','注文番号: '+orderId,'SKU: '+sku,'エラー: '+errorInfo,'→ Seller Central で確認してください'].join('\n');
    case MCF_NOTIFY_EVENT_TYPES_.MCF_NOT_CREATED:
      return [
        '【MCF確認】MCF未作成アラート',
        '受注から ' + MCF_NOT_CREATED_THRESHOLD_HOURS_ + ' 時間以上経過',
        '注文番号: ' + orderId,
        'SKU: ' + sku,
        '発送期限: ' + shipByDate,
        '→ P列が空欄のため確認が必要です'
      ].join('\n');
    case MCF_NOTIFY_EVENT_TYPES_.DEADLINE_RISK:
      return ['【発送期限注意】発送期限まで24時間以内','注文番号: '+orderId,'発送期限: '+shipByDate,'MCF状態: '+(mcfId||'未作成'),'Tracking: '+(tracking||'未取得'),'→ Amazon MCFの出荷状況を確認してください'].join('\n');
    case MCF_NOTIFY_EVENT_TYPES_.TRACKING_OBTAINED:
      return ['【追跡番号取得】','注文番号: '+orderId,'SKU: '+sku,'Carrier: '+carrier,'Tracking: '+tracking,'→ eBayへの追跡番号提出をお願いします'].join('\n');
    case MCF_NOTIFY_EVENT_TYPES_.EBAY_TRACKING_FAILED:
      return ['【eBay追跡番号提出失敗】','注文番号: '+orderId,'Carrier: '+carrier,'Tracking: '+tracking,'エラー: '+errorInfo,'→ eBay Seller Hub で手動提出してください'].join('\n');
    default:
      return '【MCF通知】' + eventType + '\n注文番号: ' + orderId;
  }
}

// ============================================================
// ■ 6. sendMcfLineNotification_
//       重複チェック後、必要な場合だけLINE送信
// ============================================================

function sendMcfLineNotification_(eventType, order, opts) {
  opts       = opts       || {};
  const isDryRun  = opts.dryRun    === true;
  const customKey = opts.notifyKey || null;
  const orderId   = order.orderId  || '';
  const sku       = order.sku      || '';

  Logger.log('');
  Logger.log('[sendMcfLineNotification_]');
  Logger.log('  イベント種別: ' + eventType);
  Logger.log('  注文番号    : ' + orderId);
  Logger.log('  DRY_RUN     : ' + isDryRun);

  const check = shouldNotifyMcfEvent_(eventType, orderId, customKey);

  if (!check.shouldNotify) {
    Logger.log('  → スキップ: ' + check.reason);
    logMcfLineNotification_(eventType, orderId, sku, check.notifyKey, '（スキップ）', 'SKIPPED');
    return 'SKIPPED';
  }

  const message = buildMcfLineMessage_(eventType, order);

  if (isDryRun) {
    Logger.log('  → [DRY_RUN] 送信予定（実際には送信しません）');
    Logger.log('  通知キー: ' + check.notifyKey);
    Logger.log('  送信予定本文:');
    Logger.log('  ─────────────────────────────────');
    Logger.log(message.split('\n').map(function(l){ return '  ' + l; }).join('\n'));
    Logger.log('  ─────────────────────────────────');
    logMcfLineNotification_(eventType, orderId, sku, check.notifyKey, message, 'DRY_RUN');
    return 'DRY_RUN';
  }

  const success = pushLineMcf_(message);
  const result  = success ? 'SENT' : 'ERROR: push失敗';
  logMcfLineNotification_(eventType, orderId, sku, check.notifyKey, message, result);

  if (success) {
    Logger.log('  ✅ LINE通知送信完了: ' + eventType + ' / ' + orderId);
  } else {
    Logger.log('  ❌ LINE通知送信失敗: ' + eventType + ' / ' + orderId);
  }
  return result;
}

// ============================================================
// ■ 7. dryRunCheckMcfLineNotifications
//       LINE送信せず、現在の通知対象を確認する（メイン確認関数）
// ============================================================

/**
 * dryRunCheckMcfLineNotifications
 *
 * v3 のMCF_NOT_CREATED 判定について：
 *   受注から MCF_NOT_CREATED_THRESHOLD_HOURS_（14時間）以上経過 & P列空欄 → 通知
 *   14時間未満は通知しない（0〜12h: WAIT_12H, 12〜14h: MCF作成処理バッファ）
 */
function dryRunCheckMcfLineNotifications() {
  const NOW_JST = new Date();

  Logger.log('');
  Logger.log('━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━');
  Logger.log('【dryRunCheckMcfLineNotifications v3】通知対象確認（DRY_RUN）');
  Logger.log('  MCF_NOT_CREATED しきい値: ' + MCF_NOT_CREATED_THRESHOLD_HOURS_ + ' 時間');
  Logger.log('  ★ LINE送信しません / eBay受注ログ書き込みなし');
  Logger.log('  実行日時: ' + Utilities.formatDate(NOW_JST, 'Asia/Tokyo', 'yyyy/MM/dd HH:mm:ss'));
  Logger.log('━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━');
  Logger.log('');

  const ss      = SpreadsheetApp.getActiveSpreadsheet();
  const eBayLog = ss.getSheetByName('eBay受注ログ');
  if (!eBayLog) { Logger.log('❌ 「eBay受注ログ」シートが見つかりません'); return; }

  const allData    = eBayLog.getDataRange().getValues();
  let notifyCount  = 0;
  let skipCount    = 0;
  let checkedCount = 0;

  Logger.log('[eBay受注ログ スキャン開始]');
  Logger.log('  総行数（ヘッダー含む）: ' + allData.length);
  Logger.log('');

  for (let i = 1; i < allData.length; i++) {
    const row     = allData[i];
    const orderId = String(row[MCF_NOTIFY_COL_.ORDER_ID] || '').trim();
    if (!orderId) continue;
    checkedCount++;

    const receivedAtRaw = row[MCF_NOTIFY_COL_.RECEIVED_AT];
    const sku           = String(row[MCF_NOTIFY_COL_.SKU]          || '').trim();
    const productName   = String(row[MCF_NOTIFY_COL_.PRODUCT_NAME] || '').trim();
    const quantity      = row[MCF_NOTIFY_COL_.QUANTITY]             || 0;
    const shipByRaw     = row[MCF_NOTIFY_COL_.SHIP_BY];
    const mcfId         = String(row[MCF_NOTIFY_COL_.MCF_ID]       || '').trim();
    const tracking      = String(row[MCF_NOTIFY_COL_.TRACKING]      || '').trim();
    const shipByDate    = shipByRaw instanceof Date
      ? Utilities.formatDate(shipByRaw, 'Asia/Tokyo', 'yyyy/MM/dd HH:mm')
      : String(shipByRaw || '');

    Logger.log('  ─── 行' + (i + 1) + ': ' + orderId + ' ──────────────────────');

    // ── 判定A：MCF未作成アラート（受注14h以上 & P列空欄）──
    let receivedAtDate = null;
    try {
      receivedAtDate = receivedAtRaw instanceof Date
        ? receivedAtRaw
        : Utilities.parseDate(String(receivedAtRaw).trim(), 'Asia/Tokyo', 'yyyy-MM-dd HH:mm');
    } catch (e) { receivedAtDate = null; }

    const hoursFromOrder = receivedAtDate
      ? (NOW_JST - receivedAtDate) / (1000 * 60 * 60) : null;

    const isMcfNotCreated = !mcfId && hoursFromOrder !== null
      && hoursFromOrder >= MCF_NOT_CREATED_THRESHOLD_HOURS_;

    if (isMcfNotCreated) {
      const order = { orderId, sku, productName, quantity, shipByDate };
      const check = shouldNotifyMcfEvent_(MCF_NOTIFY_EVENT_TYPES_.MCF_NOT_CREATED, orderId);
      Logger.log('  ⚠️ MCF未作成アラート対象');
      Logger.log('     受注から: ' + hoursFromOrder.toFixed(1) + ' 時間経過 / P列: 空欄');
      Logger.log('     通知キー: ' + check.notifyKey);
      Logger.log('     通知すべきか: ' + (check.shouldNotify ? '✅ YES' : '⛔ NO（' + check.reason + '）'));
      if (check.shouldNotify) {
        const msg = buildMcfLineMessage_(MCF_NOTIFY_EVENT_TYPES_.MCF_NOT_CREATED, order);
        Logger.log('     [DRY_RUN] 送信予定メッセージ:');
        msg.split('\n').forEach(function(l){ Logger.log('       ' + l); });
        logMcfLineNotification_(MCF_NOTIFY_EVENT_TYPES_.MCF_NOT_CREATED, orderId, sku, check.notifyKey, msg, 'DRY_RUN');
        notifyCount++;
      } else { skipCount++; }
    }

    // ── 判定B：発送期限リスク（24h以内 & Tracking空欄）──
    let shipByDateObj = null;
    try {
      if (shipByRaw instanceof Date) { shipByDateObj = shipByRaw; }
      else if (shipByRaw) {
        shipByDateObj = new Date(String(shipByRaw));
        if (isNaN(shipByDateObj.getTime())) { shipByDateObj = null; }
      }
    } catch (e) { shipByDateObj = null; }

    const hoursToDeadline = shipByDateObj
      ? (shipByDateObj - NOW_JST) / (1000 * 60 * 60) : null;

    const isDeadlineRisk = !tracking && hoursToDeadline !== null
      && hoursToDeadline >= 0 && hoursToDeadline <= DEADLINE_RISK_THRESHOLD_HOURS_;

    if (isDeadlineRisk) {
      const order = { orderId, sku, productName, quantity, shipByDate, mcfId, tracking };
      const check = shouldNotifyMcfEvent_(MCF_NOTIFY_EVENT_TYPES_.DEADLINE_RISK, orderId);
      Logger.log('  🚨 発送期限リスク対象');
      Logger.log('     発送期限まで: ' + hoursToDeadline.toFixed(1) + ' 時間 / Tracking: 未取得');
      Logger.log('     通知キー: ' + check.notifyKey);
      Logger.log('     通知すべきか: ' + (check.shouldNotify ? '✅ YES' : '⛔ NO（' + check.reason + '）'));
      if (check.shouldNotify) {
        const msg = buildMcfLineMessage_(MCF_NOTIFY_EVENT_TYPES_.DEADLINE_RISK, order);
        Logger.log('     [DRY_RUN] 送信予定メッセージ:');
        msg.split('\n').forEach(function(l){ Logger.log('       ' + l); });
        logMcfLineNotification_(MCF_NOTIFY_EVENT_TYPES_.DEADLINE_RISK, orderId, sku, check.notifyKey, msg, 'DRY_RUN');
        notifyCount++;
      } else { skipCount++; }
    }

    if (!isMcfNotCreated && !isDeadlineRisk) {
      Logger.log('  ✅ 通知不要（正常 or 処理済み）');
      Logger.log('     MCF作成: ' + (mcfId || '空欄') + ' / Tracking: ' + (tracking || '空欄'));
      if (hoursFromOrder !== null && hoursFromOrder < MCF_NOT_CREATED_THRESHOLD_HOURS_ && !mcfId) {
        Logger.log('     ℹ️ 受注から ' + hoursFromOrder.toFixed(1) + 'h（14h未満のため通知対象外）');
      }
    }
    Logger.log('');
  }

  // サマリー行を必ず記録
  logMcfDryRunSummary_(checkedCount, notifyCount, skipCount);

  Logger.log('━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━');
  Logger.log('【DRY_RUN サマリー】');
  Logger.log('  確認件数        : ' + checkedCount + ' 件');
  Logger.log('  通知対象（DRY） : ' + notifyCount  + ' 件');
  Logger.log('  重複スキップ    : ' + skipCount    + ' 件');
  Logger.log('');
  Logger.log('  ★ LINE送信は行いませんでした');
  if (notifyCount > 0) {
    Logger.log('  ★ MCF_LINE_NOTIFICATION_LOG に DRY_RUN として記録しました');
  }
  Logger.log('  ★ MCF_LINE_NOTIFICATION_LOG に DRY_RUN_SUMMARY 行を記録しました');
  Logger.log('');
  Logger.log('【次のステップ】');
  if (notifyCount > 0) {
    Logger.log('  1. MCF_LINE_NOTIFICATION_LOG の DRY_RUN 行で内容を確認');
    Logger.log('  2. 問題なければ sendMcfLineNotification_() で本番送信');
    Logger.log('  3. 重複防止のため DRY_RUN 行はそのまま残してOK');
  } else {
    Logger.log('  現在通知対象の注文はありません。');
    Logger.log('  定期的に dryRunCheckMcfLineNotifications() を実行して確認してください。');
  }
  Logger.log('━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━');
  Logger.log('');
}

// ============================================================
// ■ 使い方ガイド
// ============================================================
//
// 【v3での変更点まとめ】
//   MCF_NOT_CREATED_THRESHOLD_HOURS_ = 14（v2: 12 → v3: 14）
//
//   受注タイムライン：
//     0〜12h  : WAIT_12H（MCF作成待機 - MCF_CreationDelay12h.gs が制御）
//     12h〜   : MCF作成タイミング（autoCreateMcfOrders が処理）
//     14h〜P列空欄: このファイルのMCF_NOT_CREATEDアラート対象
//
// 【DRY_RUN で通知対象を確認する】
//   GASエディタ → dryRunCheckMcfLineNotifications を選択 → ▶ 実行
//
// 【MCF_LINE_NOTIFICATION_LOG サマリー行の読み方】
//   B列: DRY_RUN_SUMMARY  C列: 確認件数  D列: 通知対象件数
//   E列: 重複スキップ件数  F列: 結果  G列: DRY_RUN
//
// 【pushLineMcf_ を既存関数に置き換える場合】
//   function pushLineMcf_(message) { return sendLine(message); }

/**
 * MCF_CreationDelay12h.gs
 *
 * MCF作成タイミング「受注から12時間後」制御モジュール
 *
 * 【目的】
 *   eBay受注直後のキャンセル・住所修正・支払い確認の余地を残すため、
 *   受注日時から12時間経過した注文だけをMCF作成対象にする。
 *
 * 【タイムライン】
 *   受注
 *    │
 *    ├─ 0〜12h ：WAIT_12H（MCF作成待機 - キャンセル余地期間）
 *    │
 *    ├─ 12h〜  ：READY_FOR_MCF（MCF作成対象）
 *    │
 *    └─ 14h〜P列空欄：MCF_NOT_CREATED LINE通知
 *       （MCF_LineNotification.gs v3 の MCF_NOT_CREATED_THRESHOLD_HOURS_ = 14 と連動）
 *
 * 【提供する関数】
 *   dryRunCheckMcfCreationDelay12h() - 全件スキャンして分類確認（dry-run・本番API呼ばない）
 *   shouldCreateMcfNow_(receivedAtRaw) - autoCreateMcfOrders() から呼ぶ判定ヘルパー
 *   getOrderDelayClassification_(row, rowNum) - 行データを分類する内部関数
 *
 * 【autoCreateMcfOrders() への組み込み方法】
 *   既存のP列/Q列スキップチェックの直後に以下を追加：
 *
 *   const delayCheck = shouldCreateMcfNow_(row[0]); // A列: 受注日時
 *   if (!delayCheck.ready) {
 *     Logger.log('⏳ [WAIT_12H] ' + eBayOrderId + ' ' + delayCheck.reason);
 *     continue; // MCF作成をスキップ
 *   }
 *   // ↑ ここを通過した注文だけ既存のMCF作成処理へ進む
 *
 * 【禁止事項】
 *   ★ dryRunCheckMcfCreationDelay12h() は本番MCF APIを呼ばない
 *   ★ eBay受注ログへの書き込みなし
 *   ★ LINE通知なし
 *   ★ トリガー変更なし
 *   ★ 既存のMCF作成済み注文（P列あり）は再作成しない
 *
 * 【現在の注文 27-14706-83123 について】
 *   MCF作成済み（P列: EBAY-27-14706-83123）のため、
 *   今回の12時間ルール変更の対象外。再作成しない。
 */

// ============================================================
// ■ 定数
// ============================================================

// MCF作成開始まで待機する時間（受注から）
const MCF_CREATION_DELAY_HOURS_ = 12;

// eBay受注ログ 列インデックス（0始まり）
const DELAY_COL_ = {
  RECEIVED_AT  :  0,  // A: 受注日時(JST)
  ORDER_ID     :  1,  // B: eBay注文番号
  SKU          :  2,  // C: SKU
  PRODUCT_NAME :  3,  // D: 商品名
  SHIP_BY      : 14,  // O: 発送期限
  MCF_ID       : 15,  // P: MCF作成
  STATUS       : 16   // Q: ステータス
};

// Q列で「処理済み」とみなしてSKIP_PROCESSEDにするステータス
const DELAY_SKIP_PROCESSED_ = [
  'MCF作成済', '手動対応済み', '対象外', '対応不要', 'キャンセル', '返金済み'
];

// Q列で「手動確認」とみなしてSKIP_MANUALにするステータス
const DELAY_SKIP_MANUAL_ = [
  '手動確認', '確認通知済み'
];

// ============================================================
// ■ 1. shouldCreateMcfNow_
//       autoCreateMcfOrders() から呼ぶ12時間判定ヘルパー
// ============================================================

/**
 * shouldCreateMcfNow_
 *
 * 受注日時から MCF_CREATION_DELAY_HOURS_（12時間）経過しているか判定する。
 * autoCreateMcfOrders() の MCF作成判定ループ内から呼ぶ。
 *
 * 【使用例】
 *   const delayCheck = shouldCreateMcfNow_(row[0]);
 *   if (!delayCheck.ready) {
 *     Logger.log('⏳ [WAIT_12H] ' + eBayOrderId + ' ' + delayCheck.reason);
 *     continue;
 *   }
 *
 * @param {Date|string} receivedAtRaw - A列の受注日時（Date型または文字列型）
 * @returns {{
 *   ready         : boolean,  // true=作成可能 / false=待機中またはエラー
 *   hoursElapsed  : number,   // 経過時間（時間）
 *   hoursRemaining: number,   // 残り待機時間（時間）- readyの場合は0
 *   reason        : string    // ログ出力用の説明
 * }}
 */
function shouldCreateMcfNow_(receivedAtRaw) {
  const now = new Date();

  // 受注日時をパース
  let receivedAt = null;
  try {
    if (receivedAtRaw instanceof Date) {
      receivedAt = receivedAtRaw;
    } else if (receivedAtRaw) {
      // "2026-06-06 21:07" 形式を想定
      receivedAt = Utilities.parseDate(
        String(receivedAtRaw).trim(), 'Asia/Tokyo', 'yyyy-MM-dd HH:mm'
      );
    }
  } catch (e) {
    return {
      ready          : false,
      hoursElapsed   : 0,
      hoursRemaining : MCF_CREATION_DELAY_HOURS_,
      reason         : 'ERROR: 受注日時のパース失敗（' + e.toString() + '）'
    };
  }

  if (!receivedAt || isNaN(receivedAt.getTime())) {
    return {
      ready          : false,
      hoursElapsed   : 0,
      hoursRemaining : MCF_CREATION_DELAY_HOURS_,
      reason         : 'ERROR: 受注日時が空または無効値'
    };
  }

  const hoursElapsed    = (now - receivedAt) / (1000 * 60 * 60);
  const hoursRemaining  = Math.max(0, MCF_CREATION_DELAY_HOURS_ - hoursElapsed);

  // 12時間未満 → 待機
  if (hoursElapsed < MCF_CREATION_DELAY_HOURS_) {
    return {
      ready          : false,
      hoursElapsed   : hoursElapsed,
      hoursRemaining : hoursRemaining,
      reason         : 'WAIT_12H: 受注から ' + hoursElapsed.toFixed(1)
                       + 'h経過（残り ' + hoursRemaining.toFixed(1) + 'h）'
    };
  }

  // 12時間以上 → 作成可能
  return {
    ready          : true,
    hoursElapsed   : hoursElapsed,
    hoursRemaining : 0,
    reason         : 'READY: 受注から ' + hoursElapsed.toFixed(1) + 'h経過'
  };
}

// ============================================================
// ■ 2. getOrderDelayClassification_（内部関数）
//       1行の受注データを分類して返す
// ============================================================

/**
 * getOrderDelayClassification_
 *
 * @param {Array}  row    - eBay受注ログの1行データ（0始まり配列）
 * @param {number} rowNum - シート上の行番号（1始まり、Logger出力用）
 * @returns {{
 *   classification: string,   // WAIT_12H / READY_FOR_MCF / SKIP_PROCESSED / SKIP_MANUAL / ERROR
 *   orderId       : string,
 *   receivedAt    : string,   // 表示用文字列
 *   hoursElapsed  : string,   // 表示用文字列
 *   shipByDate    : string,
 *   mcfId         : string,
 *   status        : string,
 *   reason        : string
 * }}
 */
function getOrderDelayClassification_(row, rowNum) {
  const orderId     = String(row[DELAY_COL_.ORDER_ID]     || '').trim();
  const receivedRaw = row[DELAY_COL_.RECEIVED_AT];
  const shipByRaw   = row[DELAY_COL_.SHIP_BY];
  const mcfId       = String(row[DELAY_COL_.MCF_ID]       || '').trim();
  const status      = String(row[DELAY_COL_.STATUS]        || '').trim();

  // 表示用文字列
  const receivedAtStr = receivedRaw instanceof Date
    ? Utilities.formatDate(receivedRaw, 'Asia/Tokyo', 'yyyy/MM/dd HH:mm')
    : String(receivedRaw || '');
  const shipByStr = shipByRaw instanceof Date
    ? Utilities.formatDate(shipByRaw, 'Asia/Tokyo', 'yyyy/MM/dd HH:mm')
    : String(shipByRaw || '');

  // ① P列チェック（MCF作成済み → SKIP_PROCESSED）
  if (mcfId) {
    return {
      classification: 'SKIP_PROCESSED',
      orderId, receivedAt: receivedAtStr, hoursElapsed: '-',
      shipByDate: shipByStr, mcfId, status,
      reason: 'P列に値あり（MCF作成済み）: ' + mcfId
    };
  }

  // ② Q列チェック（処理済みステータス → SKIP_PROCESSED）
  if (DELAY_SKIP_PROCESSED_.some(function(s) { return status.includes(s); })) {
    return {
      classification: 'SKIP_PROCESSED',
      orderId, receivedAt: receivedAtStr, hoursElapsed: '-',
      shipByDate: shipByStr, mcfId, status,
      reason: 'Q列が処理済みステータス: ' + status
    };
  }

  // ③ Q列チェック（手動確認系 → SKIP_MANUAL）
  if (DELAY_SKIP_MANUAL_.some(function(s) { return status.includes(s); })) {
    return {
      classification: 'SKIP_MANUAL',
      orderId, receivedAt: receivedAtStr, hoursElapsed: '-',
      shipByDate: shipByStr, mcfId, status,
      reason: 'Q列が手動確認ステータス: ' + status
    };
  }

  // ④ 12時間判定
  const check = shouldCreateMcfNow_(receivedRaw);

  if (check.reason.startsWith('ERROR')) {
    return {
      classification: 'ERROR',
      orderId, receivedAt: receivedAtStr, hoursElapsed: 'N/A',
      shipByDate: shipByStr, mcfId, status,
      reason: check.reason
    };
  }

  if (!check.ready) {
    return {
      classification: 'WAIT_12H',
      orderId, receivedAt: receivedAtStr,
      hoursElapsed: check.hoursElapsed.toFixed(1) + 'h（残り ' + check.hoursRemaining.toFixed(1) + 'h）',
      shipByDate: shipByStr, mcfId, status,
      reason: check.reason
    };
  }

  return {
    classification: 'READY_FOR_MCF',
    orderId, receivedAt: receivedAtStr,
    hoursElapsed: check.hoursElapsed.toFixed(1) + 'h',
    shipByDate: shipByStr, mcfId, status,
    reason: check.reason
  };
}

// ============================================================
// ■ 3. dryRunCheckMcfCreationDelay12h
//       全件スキャンして12時間判定の分類を確認する（dry-run）
// ============================================================

/**
 * dryRunCheckMcfCreationDelay12h
 *
 * eBay受注ログを全件スキャンし、各注文を以下に分類してLogger出力する。
 *
 * 分類：
 *   ⏳ WAIT_12H      : 受注から12時間未満（MCF作成待機中 - エラーではない）
 *   ✅ READY_FOR_MCF : 受注から12時間以上・MCF未作成（MCF作成対象）
 *   ⏭  SKIP_PROCESSED: P列またはQ列で処理済み（スキップ）
 *   👤 SKIP_MANUAL   : 手動確認・確認通知済みなど（スキップ）
 *   ❌ ERROR         : 受注日時が読めない等
 *
 * 出力項目：
 *   行番号 / eBay注文番号 / 受注日時 / 経過時間 / 発送期限 / MCF作成状態（P列）/ 判定結果
 *
 * ★ 本番MCF APIを呼ばない
 * ★ eBay受注ログへの書き込みなし
 * ★ LINE通知なし
 * ★ トリガー変更なし
 */
function dryRunCheckMcfCreationDelay12h() {
  const NOW_JST = new Date();

  Logger.log('');
  Logger.log('━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━');
  Logger.log('【dryRunCheckMcfCreationDelay12h】MCF作成 12時間待機確認');
  Logger.log('  MCF作成待機時間 : ' + MCF_CREATION_DELAY_HOURS_ + ' 時間');
  Logger.log('  ★ 本番MCF APIを呼びません / eBay受注ログ書き込みなし');
  Logger.log('  実行日時: ' + Utilities.formatDate(NOW_JST, 'Asia/Tokyo', 'yyyy/MM/dd HH:mm:ss'));
  Logger.log('━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━');
  Logger.log('');

  const ss      = SpreadsheetApp.getActiveSpreadsheet();
  const eBayLog = ss.getSheetByName('eBay受注ログ');
  if (!eBayLog) {
    Logger.log('❌ 「eBay受注ログ」シートが見つかりません');
    return;
  }

  const allData = eBayLog.getDataRange().getValues();
  Logger.log('[eBay受注ログ スキャン]');
  Logger.log('  総行数（ヘッダー含む）: ' + allData.length);
  Logger.log('');

  // 分類カウンター
  const counts = {
    WAIT_12H      : 0,
    READY_FOR_MCF : 0,
    SKIP_PROCESSED: 0,
    SKIP_MANUAL   : 0,
    ERROR         : 0
  };

  // サマリー用リスト（WAIT/READY/ERRORのみ詳細出力）
  const lists = {
    WAIT_12H      : [],
    READY_FOR_MCF : [],
    ERROR         : []
  };

  for (let i = 1; i < allData.length; i++) {
    const row     = allData[i];
    const orderId = String(row[DELAY_COL_.ORDER_ID] || '').trim();
    if (!orderId) continue;

    const result = getOrderDelayClassification_(row, i + 1);
    counts[result.classification]++;

    // ── Logger 詳細出力 ──────────────────────────────────
    let symbol = '';
    switch (result.classification) {
      case 'WAIT_12H'      : symbol = '⏳'; break;
      case 'READY_FOR_MCF' : symbol = '✅'; break;
      case 'SKIP_PROCESSED': symbol = '⏭ '; break;
      case 'SKIP_MANUAL'   : symbol = '👤'; break;
      case 'ERROR'         : symbol = '❌'; break;
    }

    Logger.log('  ' + symbol + ' 行' + (i + 1) + ' [' + result.classification + ']');
    Logger.log('     注文番号  : ' + result.orderId);
    Logger.log('     受注日時  : ' + result.receivedAt);
    Logger.log('     経過時間  : ' + result.hoursElapsed);
    Logger.log('     発送期限  : ' + result.shipByDate);
    Logger.log('     P列(MCF)  : ' + (result.mcfId  || '（空欄）'));
    Logger.log('     Q列(状態) : ' + (result.status  || '（空欄）'));
    Logger.log('     判定      : ' + result.reason);
    Logger.log('');

    // サマリー用リスト更新
    if (result.classification === 'WAIT_12H') {
      lists.WAIT_12H.push(result.orderId + '（' + result.hoursElapsed + '）');
    } else if (result.classification === 'READY_FOR_MCF') {
      lists.READY_FOR_MCF.push(result.orderId + '（' + result.hoursElapsed + '）');
    } else if (result.classification === 'ERROR') {
      lists.ERROR.push(result.orderId + '（' + result.reason + '）');
    }
  }

  // ── 分類サマリー ─────────────────────────────────────
  Logger.log('━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━');
  Logger.log('【分類サマリー】');
  Logger.log('');
  Logger.log('  ⏳ WAIT_12H      : ' + counts.WAIT_12H       + ' 件  （受注12h未満 - MCF作成待機中）');
  Logger.log('  ✅ READY_FOR_MCF : ' + counts.READY_FOR_MCF  + ' 件  （12h以上経過・未作成 - MCF作成対象）');
  Logger.log('  ⏭  SKIP_PROCESSED: ' + counts.SKIP_PROCESSED + ' 件  （処理済み - スキップ）');
  Logger.log('  👤 SKIP_MANUAL   : ' + counts.SKIP_MANUAL    + ' 件  （手動確認 - スキップ）');
  Logger.log('  ❌ ERROR         : ' + counts.ERROR           + ' 件  （受注日時読み取り不可）');
  Logger.log('');

  if (lists.WAIT_12H.length > 0) {
    Logger.log('  【⏳ WAIT_12H 一覧】（キャンセル余地期間中）');
    lists.WAIT_12H.forEach(function(s) { Logger.log('     - ' + s); });
    Logger.log('');
  }

  if (lists.READY_FOR_MCF.length > 0) {
    Logger.log('  【✅ READY_FOR_MCF 一覧】（MCF作成対象）');
    lists.READY_FOR_MCF.forEach(function(s) { Logger.log('     - ' + s); });
    Logger.log('');
  }

  if (lists.ERROR.length > 0) {
    Logger.log('  【❌ ERROR 一覧】（受注日時確認が必要）');
    lists.ERROR.forEach(function(s) { Logger.log('     - ' + s); });
    Logger.log('');
  }

  // ── 次のステップ ─────────────────────────────────────
  Logger.log('━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━');
  Logger.log('【次のステップ】');
  Logger.log('');

  if (counts.READY_FOR_MCF > 0) {
    Logger.log('  READY_FOR_MCF が ' + counts.READY_FOR_MCF + ' 件あります。');
    Logger.log('  autoCreateMcfOrders() に shouldCreateMcfNow_() を組み込んで実行してください。');
    Logger.log('');
    Logger.log('  【autoCreateMcfOrders() 組み込みコード】');
    Logger.log('  ─────────────────────────────────────────────');
    Logger.log('  // P列/Q列スキップチェックの直後に追加');
    Logger.log('  const delayCheck = shouldCreateMcfNow_(row[0]); // A列: 受注日時');
    Logger.log('  if (!delayCheck.ready) {');
    Logger.log('    Logger.log("⏳ [WAIT_12H] " + eBayOrderId + " " + delayCheck.reason);');
    Logger.log('    continue; // MCF作成をスキップ');
    Logger.log('  }');
    Logger.log('  // ↑ここを通過した注文だけ既存のMCF作成処理へ進む');
    Logger.log('  ─────────────────────────────────────────────');
  } else if (counts.WAIT_12H > 0) {
    Logger.log('  WAIT_12H が ' + counts.WAIT_12H + ' 件あります。');
    Logger.log('  12時間経過後に再度 dryRunCheckMcfCreationDelay12h() を実行して確認してください。');
  } else if (counts.ERROR > 0) {
    Logger.log('  ERROR が ' + counts.ERROR + ' 件あります。');
    Logger.log('  A列（受注日時）の形式を確認してください。期待形式: yyyy-MM-dd HH:mm または Date型');
  } else {
    Logger.log('  現在MCF作成対象の注文はありません。');
  }

  Logger.log('');
  Logger.log('  ★ 27-14706-83123 はMCF作成済みのため SKIP_PROCESSED になります。');
  Logger.log('  ★ 新規受注から12時間は WAIT_12H → 12時間後に READY_FOR_MCF になります。');
  Logger.log('━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━');
  Logger.log('');
}

// ============================================================
// ■ 参考：autoCreateMcfOrders() への組み込みパターン（コメント）
// ============================================================
//
// 【現在の autoCreateMcfOrders() の想定フロー】
//
//   for (let i = 1; i < data.length; i++) {
//     const row = data[i];
//     const eBayOrderId = String(row[1] || '').trim();
//     if (!eBayOrderId) continue;
//
//     // ── 既存のスキップチェック（変更不要）──
//     const mcfStatus = String(row[15] || '').trim();   // P列
//     const statusCol = String(row[16] || '').trim();   // Q列
//     if (mcfStatus) { skipCount++; continue; }
//     if (['MCF作成済','手動対応済み','キャンセル'...].some(...)) { skipCount++; continue; }
//
//     // ── ★ 追加：12時間待機チェック ──────────────────────
//     const delayCheck = shouldCreateMcfNow_(row[0]); // A列: 受注日時
//     if (!delayCheck.ready) {
//       Logger.log('⏳ [WAIT_12H] ' + eBayOrderId + ' ' + delayCheck.reason);
//       // エラー扱いしない - 単にスキップ
//       continue;
//     }
//     // ─────────────────────────────────────────────────────
//
//     // ← この先が既存のMCF作成処理（変更不要）
//     processedCount++;
//     const skuInfo = getSkuInfoBySku(ss, sku);
//     // ...MCF作成ロジック...
//   }
//
// 【dryRunValidateMcfOrders() への追加】
//   dryRunValidateMcfOrders() にも同じ shouldCreateMcfNow_() チェックを追加すると
//   dry-run段階でも WAIT_12H の注文が除外されるため、整合性が取れます。
//
//   // 既存のスキップチェック後に追加：
//   const dc = shouldCreateMcfNow_(row[0]);
//   if (!dc.ready) {
//     Logger.log('⏳ [WAIT_12H] 行' + (i+1) + ': ' + eBayOrder + ' ' + dc.reason);
//     skipCount++;
//     continue;
//   }

/**
 * MCF_TrackingAuto.gs
 *
 * MCF追跡番号自動取得システム（全件処理版）
 *
 * 【関数一覧】
 *   checkMcfTrackingNumbersAuto()       - 本番用（追跡番号取得・書き込み・LINE通知）
 *   dryRunCheckMcfTrackingNumbersAuto() - 確認用（書き込みなし・LINE通知なし）
 *
 * 【処理対象条件（すべて満たす行のみ処理）】
 *   1. P列（MCF作成）に有効なMCF IDがある
 *      有効：EBAY- で始まる、または MCF:EBAY- で始まる
 *      無効：手動確認 / SKU未発見 / 未作成 / 情報不足 / 作成エラー / 空欄
 *   2. Q列（ステータス）が「MCF作成済」
 *   3. S列（Tracking）が空欄
 *   4. B列（eBay注文番号）に値がある
 *
 * 【書き込み先】（追跡番号が取得できた行のみ書き込む）
 *   R列（18）: Carrier
 *   S列（19）: Tracking
 *   T列（20）: Tracking取得日時
 *   U列（21）: Tracking確認ステータス（「追跡番号取得済」）
 *
 * 【禁止事項】
 *   ★ eBayへの追跡番号提出APIは呼ばない
 *   ★ MCF作成APIは呼ばない
 *   ★ 既存Trackingがある行は上書きしない
 *   ★ P列/Q列は変更しない
 *   ★ トリガーはこのファイルでは設定しない
 *   ★ dryRun関数ではR/S/T/U列に書き込まない
 *   ★ dryRun関数ではLINE通知しない
 */

// ============================================================
// ■ 定数
// ============================================================

// eBay受注ログ 列インデックス（0始まり）
const MCF_TRACKING_COL_ = {
  RECEIVED_AT  :  0,  // A: 受注日時
  ORDER_ID     :  1,  // B: eBay注文番号
  SKU          :  2,  // C: SKU
  PRODUCT_NAME :  3,  // D: 商品名
  SHIP_BY      : 14,  // O: 発送期限
  MCF_ID       : 15,  // P: MCF作成
  STATUS       : 16,  // Q: ステータス
  CARRIER      : 17,  // R: Carrier
  TRACKING     : 18,  // S: Tracking
  TRACKING_AT  : 19,  // T: Tracking取得日時
  TRACKING_STS : 20   // U: Tracking確認ステータス
};

// シート書き込み列番号（1始まり・getRange用）
const MCF_TRACKING_WRITE_COL_ = {
  CARRIER      : 18,  // R列
  TRACKING     : 19,  // S列
  TRACKING_AT  : 20,  // T列
  TRACKING_STS : 21   // U列
};

// GetFulfillmentOrder エンドポイント
const MCF_TRACKING_BASE_URL_ = 'https://sellingpartnerapi-na.amazon.com';

// ============================================================
// ■ 0. normalizeMcfId_
//       P列のMCF IDから先頭の「MCF:」を除去し、
//       有効な EBAY- ID だけを返す
// ============================================================

function normalizeMcfId_(rawMcfId) {
  const value = String(rawMcfId || '').trim().replace(/^MCF:/i, '');

  // Amazon MCF の sellerFulfillmentOrderId だけ有効にする
  // 有効例：
  //   EBAY-27-14706-83123
  //   MCF:EBAY-27-14706-83123 → EBAY-27-14706-83123
  // 無効例：
  //   手動確認 / SKU未発見 / 未作成 / 情報不足 / 作成エラー / 空欄
  if (!/^EBAY-/.test(value)) {
    return '';
  }

  return value;
}

// ============================================================
// ■ 1. fetchMcfTracking_
//       GetFulfillmentOrder API を呼び出して追跡番号を取得する内部関数
// ============================================================

function fetchMcfTracking_(sellerFulfillmentOrderId) {

  let accessToken = '';
  try {
    accessToken = getMcfLwaAccessToken_();
  } catch (e) {
    return {
      success: false,
      hasTracking: false,
      error: 'LWAトークン取得失敗: ' + e.toString()
    };
  }

  const endpoint = MCF_TRACKING_BASE_URL_
    + '/fba/outbound/2020-07-01/fulfillmentOrders/'
    + encodeURIComponent(sellerFulfillmentOrderId);

  let httpResponse;
  let responseCode;
  let responseBody;

  try {
    httpResponse = UrlFetchApp.fetch(endpoint, {
      method: 'GET',
      headers: {
        'x-amz-access-token': accessToken,
        'Content-Type': 'application/json'
      },
      muteHttpExceptions: true
    });

    responseCode = httpResponse.getResponseCode();
    responseBody = httpResponse.getContentText();

  } catch (e) {
    return {
      success: false,
      hasTracking: false,
      error: 'ネットワークエラー: ' + e.toString()
    };
  }

  if (responseCode !== 200) {
    return {
      success: false,
      hasTracking: false,
      error: 'HTTP ' + responseCode + ': ' + responseBody.substring(0, 300)
    };
  }

  let responseJson;
  try {
    responseJson = JSON.parse(responseBody);
  } catch (e) {
    return {
      success: false,
      hasTracking: false,
      error: 'JSONパース失敗: ' + e.toString()
    };
  }

  // payload 層対応
  const payload          = responseJson.payload || responseJson;
  const fulfillmentOrder = payload.fulfillmentOrder || {};
  const shipments        = payload.fulfillmentShipments || [];

  const fulfillmentOrderStatus = fulfillmentOrder.fulfillmentOrderStatus || '';

  let trackingNumber = '';
  let carrierCode    = '';
  let packageStatus  = '';

  for (let s = 0; s < shipments.length; s++) {
    const shipment = shipments[s];
    packageStatus = shipment.fulfillmentShipmentStatus || '';

    const packages = shipment.fulfillmentShipmentPackage || [];

    for (let p = 0; p < packages.length; p++) {
      const pkg = packages[p];

      if (pkg.trackingNumber) {
        trackingNumber = pkg.trackingNumber;
        carrierCode = pkg.carrierCode || '';
        break;
      }
    }

    if (trackingNumber) break;
  }

  return {
    success: true,
    hasTracking: !!trackingNumber,
    trackingNumber: trackingNumber,
    carrierCode: carrierCode,
    packageStatus: packageStatus,
    fulfillmentOrderStatus: fulfillmentOrderStatus,
    shipmentsCount: shipments.length,
    error: null
  };
}

// ============================================================
// ■ 2. checkMcfTrackingNumbersAuto
//       本番用：追跡番号取得・シート書き込み・LINE通知
// ============================================================

function checkMcfTrackingNumbersAuto() {
  const NOW_JST = new Date();

  Logger.log('');
  Logger.log('━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━');
  Logger.log('【checkMcfTrackingNumbersAuto】MCF追跡番号自動取得（本番）');
  Logger.log('  ★ R/S/T/U列に書き込みます / LINE通知あり');
  Logger.log('  ★ eBayへの追跡番号提出は行いません');
  Logger.log('  実行日時: ' + Utilities.formatDate(NOW_JST, 'Asia/Tokyo', 'yyyy/MM/dd HH:mm:ss'));
  Logger.log('━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━');
  Logger.log('');

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const eBayLog = ss.getSheetByName('eBay受注ログ');

  if (!eBayLog) {
    Logger.log('❌ 「eBay受注ログ」シートが見つかりません');
    return;
  }

  const allData = eBayLog.getDataRange().getValues();

  Logger.log('[eBay受注ログ スキャン]');
  Logger.log('  総行数（ヘッダー含む）: ' + allData.length);
  Logger.log('');

  let checkedCount  = 0;
  let obtainedCount = 0;
  let pendingCount  = 0;
  let errorCount    = 0;
  let skipCount     = 0;

  for (let i = 1; i < allData.length; i++) {
    const row      = allData[i];
    const orderId  = String(row[MCF_TRACKING_COL_.ORDER_ID] || '').trim();
    const rawMcfId = String(row[MCF_TRACKING_COL_.MCF_ID] || '').trim();
    const status   = String(row[MCF_TRACKING_COL_.STATUS] || '').trim();
    const tracking = String(row[MCF_TRACKING_COL_.TRACKING] || '').trim();
    const sku      = String(row[MCF_TRACKING_COL_.SKU] || '').trim();

    if (!orderId) continue;

    // 1. P列に値がある
    if (!rawMcfId) {
      skipCount++;
      continue;
    }

    // 2. Q列が「MCF作成済」
    if (status !== 'MCF作成済') {
      skipCount++;
      continue;
    }

    // 3. S列Trackingが空欄
    if (tracking) {
      skipCount++;
      continue;
    }

    // 4. P列の値が有効なMCF IDか確認
    const mcfId = normalizeMcfId_(rawMcfId);

    if (!mcfId) {
      skipCount++;
      Logger.log(
        '⏭  行' +
        (i + 1) +
        ': ' +
        orderId +
        ' → スキップ（P列が有効なMCF IDではありません: ' +
        rawMcfId +
        '）'
      );
      continue;
    }

    checkedCount++;

    Logger.log('  ─── 行' + (i + 1) + ': ' + orderId);
    Logger.log('     MCF ID（正規化後）: ' + mcfId);

    const result = fetchMcfTracking_(mcfId);

    if (!result.success) {
      Logger.log('     ❌ APIエラー: ' + result.error);
      errorCount++;
      Logger.log('');
      continue;
    }

    Logger.log('     fulfillmentOrderStatus: ' + result.fulfillmentOrderStatus);
    Logger.log('     shipments数           : ' + result.shipmentsCount);
    Logger.log('     packageStatus         : ' + (result.packageStatus || '（なし）'));

    if (!result.hasTracking) {
      Logger.log('     ℹ️ 追跡番号未発行 → LINE通知なし / 書き込みなし');
      pendingCount++;
      Logger.log('');
      continue;
    }

    Logger.log('     ✅ 追跡番号取得');
    Logger.log('        Carrier : ' + result.carrierCode);
    Logger.log('        Tracking: ' + result.trackingNumber);

    const writtenAt = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy/MM/dd HH:mm:ss');
    const sheetRow = i + 1;

    try {
      eBayLog.getRange(sheetRow, MCF_TRACKING_WRITE_COL_.CARRIER).setValue(result.carrierCode);
      eBayLog.getRange(sheetRow, MCF_TRACKING_WRITE_COL_.TRACKING).setValue(result.trackingNumber);
      eBayLog.getRange(sheetRow, MCF_TRACKING_WRITE_COL_.TRACKING_AT).setValue(writtenAt);
      eBayLog.getRange(sheetRow, MCF_TRACKING_WRITE_COL_.TRACKING_STS).setValue('追跡番号取得済');

      Logger.log('        R/S/T/U列への書き込み完了');

    } catch (e) {
      Logger.log('     ❌ シート書き込みエラー: ' + e.toString());
      errorCount++;
      Logger.log('');
      continue;
    }

    // LINE通知：追跡番号取得時だけ
    try {
      const order = {
        orderId: orderId,
        sku: sku,
        carrier: result.carrierCode,
        tracking: result.trackingNumber
      };

      const lineResult = sendMcfLineNotification_(
        MCF_NOTIFY_EVENT_TYPES_.TRACKING_OBTAINED,
        order,
        { dryRun: false }
      );

      Logger.log('        LINE通知結果: ' + lineResult);

    } catch (e) {
      Logger.log('     ⚠️ LINE通知エラー（書き込みは完了済み）: ' + e.toString());
    }

    obtainedCount++;
    Logger.log('');
  }

  Logger.log('━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━');
  Logger.log('【サマリー】');
  Logger.log('  確認対象件数    : ' + checkedCount + ' 件');
  Logger.log('  追跡番号取得件数: ' + obtainedCount + ' 件');
  Logger.log('  未発行件数      : ' + pendingCount + ' 件');
  Logger.log('  エラー件数      : ' + errorCount + ' 件');
  Logger.log('  スキップ件数    : ' + skipCount + ' 件');
  Logger.log('');

  if (obtainedCount > 0) {
    Logger.log('  ✅ ' + obtainedCount + ' 件の追跡番号を取得しました。');
    Logger.log('     R列 / S列 / T列 / U列 を確認してください。');
  } else {
    Logger.log('  ℹ️ 新規取得した追跡番号はありませんでした。');
  }

  Logger.log('');
  Logger.log('  ★ eBayへの追跡番号提出は行いませんでした（別途実装予定）');
  Logger.log('━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━');
  Logger.log('');
}

// ============================================================
// ■ 3. dryRunCheckMcfTrackingNumbersAuto
//       確認用：書き込みなし・LINE通知なし
// ============================================================

function dryRunCheckMcfTrackingNumbersAuto() {
  const NOW_JST = new Date();

  Logger.log('');
  Logger.log('━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━');
  Logger.log('【dryRunCheckMcfTrackingNumbersAuto】追跡番号確認（DRY_RUN）');
  Logger.log('  ★ シートへの書き込みなし / LINE通知なし');
  Logger.log('  ★ GetFulfillmentOrder API は呼び出します（読み取りのみ）');
  Logger.log('  実行日時: ' + Utilities.formatDate(NOW_JST, 'Asia/Tokyo', 'yyyy/MM/dd HH:mm:ss'));
  Logger.log('━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━');
  Logger.log('');

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const eBayLog = ss.getSheetByName('eBay受注ログ');

  if (!eBayLog) {
    Logger.log('❌ 「eBay受注ログ」シートが見つかりません');
    return;
  }

  const allData = eBayLog.getDataRange().getValues();

  Logger.log('[eBay受注ログ スキャン]');
  Logger.log('  総行数（ヘッダー含む）: ' + allData.length);
  Logger.log('');

  let checkedCount = 0;
  let wouldObtain  = 0;
  let pendingCount = 0;
  let errorCount   = 0;
  let skipCount    = 0;

  const obtainList = [];
  const pendingList = [];
  const errorList = [];

  for (let i = 1; i < allData.length; i++) {
    const row       = allData[i];
    const orderId   = String(row[MCF_TRACKING_COL_.ORDER_ID] || '').trim();
    const rawMcfId  = String(row[MCF_TRACKING_COL_.MCF_ID] || '').trim();
    const status    = String(row[MCF_TRACKING_COL_.STATUS] || '').trim();
    const tracking  = String(row[MCF_TRACKING_COL_.TRACKING] || '').trim();
    const sku       = String(row[MCF_TRACKING_COL_.SKU] || '').trim();
    const shipByRaw = row[MCF_TRACKING_COL_.SHIP_BY];

    const shipByDate = shipByRaw instanceof Date
      ? Utilities.formatDate(shipByRaw, 'Asia/Tokyo', 'yyyy/MM/dd')
      : String(shipByRaw || '');

    if (!orderId) continue;

    const skipReasons = [];

    if (!rawMcfId) {
      skipReasons.push('P列にMCF IDなし');
    }

    if (status !== 'MCF作成済') {
      skipReasons.push('Q列ステータスが「' + status + '」');
    }

    if (tracking) {
      skipReasons.push('S列に既存Tracking（' + tracking + '）');
    }

    if (skipReasons.length > 0) {
      Logger.log(
        '  ⏭  行' +
        (i + 1) +
        ': ' +
        orderId +
        ' → スキップ（' +
        skipReasons.join(' / ') +
        '）'
      );
      skipCount++;
      continue;
    }

    const mcfId = normalizeMcfId_(rawMcfId);

    if (!mcfId) {
      skipCount++;
      Logger.log(
        '  ⏭  行' +
        (i + 1) +
        ': ' +
        orderId +
        ' → スキップ（P列が有効なMCF IDではありません: ' +
        rawMcfId +
        '）'
      );
      continue;
    }

    checkedCount++;

    Logger.log('');
    Logger.log('  🔍 行' + (i + 1) + ': ' + orderId + ' → API確認対象');
    Logger.log('     MCF ID（正規化後）: ' + mcfId);
    Logger.log('     SKU     : ' + sku);
    Logger.log('     発送期限: ' + shipByDate);

    const result = fetchMcfTracking_(mcfId);

    if (!result.success) {
      Logger.log('     ❌ APIエラー: ' + result.error);
      errorCount++;
      errorList.push(orderId + '（' + result.error.substring(0, 80) + '）');
      continue;
    }

    Logger.log('     fulfillmentOrderStatus: ' + result.fulfillmentOrderStatus);
    Logger.log('     shipments数           : ' + result.shipmentsCount);
    Logger.log('     packageStatus         : ' + (result.packageStatus || '（なし）'));

    if (result.hasTracking) {
      Logger.log('     ✅ 追跡番号あり');
      Logger.log('        Carrier : ' + result.carrierCode);
      Logger.log('        Tracking: ' + result.trackingNumber);
      Logger.log('        [DRY_RUN] 本番実行時 → R/S/T/U列に書き込み + LINE通知（TRACKING_OBTAINED）');

      wouldObtain++;
      obtainList.push(orderId + '（' + result.carrierCode + ': ' + result.trackingNumber + '）');

    } else {
      Logger.log('     ℹ️ 追跡番号未発行 → 書き込みなし・LINE通知なし');

      pendingCount++;
      pendingList.push(orderId + '（Status: ' + result.fulfillmentOrderStatus + ' / ' + result.packageStatus + '）');
    }
  }

  Logger.log('');
  Logger.log('━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━');
  Logger.log('【DRY_RUN サマリー】');
  Logger.log('  確認対象件数    : ' + checkedCount + ' 件');
  Logger.log('  取得見込み件数  : ' + wouldObtain + ' 件  ← 本番実行時に書き込み＋LINE通知される');
  Logger.log('  未発行件数      : ' + pendingCount + ' 件  ← Amazon処理待ち');
  Logger.log('  エラー件数      : ' + errorCount + ' 件');
  Logger.log('  スキップ件数    : ' + skipCount + ' 件  （対象条件不一致）');
  Logger.log('');

  if (obtainList.length > 0) {
    Logger.log('  【✅ 取得見込み一覧】');
    obtainList.forEach(function(s) {
      Logger.log('     - ' + s);
    });
    Logger.log('');
  }

  if (pendingList.length > 0) {
    Logger.log('  【ℹ️ 未発行一覧】');
    pendingList.forEach(function(s) {
      Logger.log('     - ' + s);
    });
    Logger.log('');
  }

  if (errorList.length > 0) {
    Logger.log('  【❌ エラー一覧】');
    errorList.forEach(function(s) {
      Logger.log('     - ' + s);
    });
    Logger.log('');
  }

  Logger.log('  ★ シートへの書き込みは行いませんでした（DRY_RUN）');
  Logger.log('  ★ LINE通知は行いませんでした（DRY_RUN）');
  Logger.log('');

  Logger.log('【次のステップ】');

  if (wouldObtain > 0) {
    Logger.log('  取得見込みが ' + wouldObtain + ' 件あります。');
    Logger.log('  内容を確認して、問題なければ checkMcfTrackingNumbersAuto() を実行してください。');
  } else if (pendingCount > 0) {
    Logger.log('  現在すべて追跡番号未発行です。しばらく後に再実行してください。');
  } else {
    Logger.log('  現在対象の注文はありません。');
  }

  Logger.log('━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━');
  Logger.log('');
}

// ============================================================
// ■ LWAアクセストークン取得（MCF専用）
//
//   Script Properties から LWA認証情報を読み取り
//   アクセストークンを返す。
//
//   キー名（優先順）:
//     SPAPI_CLIENT_ID / SPAPI_CLIENT_SECRET / SPAPI_REFRESH_TOKEN
// ============================================================
function getMcfLwaAccessToken_() {
  const props = PropertiesService.getScriptProperties();

  // キー名は SPAPI_CLIENT_ID を優先、なければ AMAZON_CLIENT_ID を試みる
  const clientId     = sampleCsvProps_().getProperty('SPAPI_CLIENT_ID')
                    || sampleCsvProps_().getProperty('AMAZON_CLIENT_ID');
  const clientSecret = sampleCsvProps_().getProperty('SPAPI_CLIENT_SECRET')
                    || sampleCsvProps_().getProperty('AMAZON_CLIENT_SECRET');
  const refreshToken = sampleCsvProps_().getProperty('SPAPI_REFRESH_TOKEN')
                    || sampleCsvProps_().getProperty('SP_REFRESH_TOKEN');

  if (!clientId || !clientSecret || !refreshToken) {
    throw new Error(
      'Script Properties に LWA認証情報が不足しています。' +
      'SPAPI_CLIENT_ID / SPAPI_CLIENT_SECRET / SPAPI_REFRESH_TOKEN を確認してください。'
    );
  }

  const tokenResponse = UrlFetchApp.fetch('https://api.amazon.com/auth/o2/token', {
    method     : 'POST',
    contentType: 'application/x-www-form-urlencoded',
    payload    : {
      grant_type   : 'refresh_token',
      refresh_token: refreshToken,
      client_id    : clientId,
      client_secret: clientSecret
    },
    muteHttpExceptions: true
  });

  const tokenCode = tokenResponse.getResponseCode();
  const tokenBody = tokenResponse.getContentText();

  if (tokenCode !== 200) {
    throw new Error(
      'LWAトークン取得失敗（HTTP ' + tokenCode + '）: ' + tokenBody
    );
  }

  const tokenJson = JSON.parse(tokenBody);
  if (!tokenJson.access_token) {
    throw new Error('LWAレスポンスに access_token が含まれていません: ' + tokenBody);
  }

  return tokenJson.access_token;
}

/**
 * MCF_ErrorValidation_Work.gs
 * 
 * Phase 1：dry-run専用コード（作業用GAS）
 * 
 * 目的：
 *   eBay受注のMCF作成前にエラーを検知・検証
 *   dry-runで全受注を診断（本実行なし）
 * 
 * 禁止事項（この段階では実行しない）：
 *   ✗ autoCreateMcfOrders() への変更
 *   ✗ MCF作成API呼び出し
 *   ✗ eBay受注ログの更新
 *   ✗ LINE通知の送信
 *   ✗ clasp push・本番反映
 * 
 * 実行可能：
 *   ✓ dryRunValidateMcfOrders() - 診断実行
 *   ✓ MCF_SKU_ERROR_LOG_DRY_RUN への記録
 *   ✓ Logger ログでの結果確認
 */

// ============================================================================
// ■ 1. SKU情報取得関数
// ============================================================================

/**
 * eBay SKU から SKU情報を取得（MCF_SKU_MAP優先）
 * 
 * 参照優先順：
 *   Step 1: MCF_SKU_MAP（eBay SKU列でマッチ）
 *   Step 2: FBA在庫_API（sellerSku列でマッチ）
 *   Step 3: FBA在庫_DL（eBay SKU / sku 列でマッチ）
 * 
 * @param {SpreadsheetApp.Spreadsheet} ss - スプレッドシート
 * @param {string} sku - eBay SKU
 * @returns {object|null} SKU情報 or null
 */
function getSkuInfoBySku(ss, sku) {
  if (!sku) return null;

  // ★Step 1: MCF_SKU_MAP から検索（最優先）
  const mcfMap = ss.getSheetByName('MCF_SKU_MAP');
  if (mcfMap) {
    const data = mcfMap.getDataRange().getValues();
    if (data.length < 2) return null;

    const headers = data[0];
    const skuCol = headers.indexOf('eBay SKU');
    const sellerSkuCol = headers.indexOf('Amazon sellerSku');
    const asinCol = headers.indexOf('ASIN');
    const nameCol = headers.indexOf('商品名');
    const activeCol = headers.indexOf('MCF利用可否');

    // 列が見つからない場合は null
    if (skuCol === -1 || sellerSkuCol === -1) return null;

    // SKU でマッチを探す
    for (let i = 1; i < data.length; i++) {
      const cellValue = String(data[i][skuCol] || '').trim();
      if (cellValue === sku) {
        return {
          source: 'MCF_SKU_MAP',
          eBaySkU: sku,
          sellerSku: String(data[i][sellerSkuCol] || '').trim(),
          asin: asinCol !== -1 ? String(data[i][asinCol] || '').trim() : '',
          productName: nameCol !== -1 ? String(data[i][nameCol] || '').trim() : '',
          active: activeCol !== -1 ? (data[i][activeCol] === '✅') : true,
          rowIndex: i + 1
        };
      }
    }
  }

  // ★Step 2: FBA在庫_API から検索（sellerSku列でマッチ）
  //   MCF_SKU_MAP に存在しない SKU でも、FBA在庫_API にデータがあれば返す。
  //   fulfillableQuantity / totalQuantity を含めて返す。
  const fbaApiSheet = ss.getSheetByName('FBA在庫_API');
  if (fbaApiSheet) {
    const apiData = fbaApiSheet.getDataRange().getValues();
    if (apiData.length >= 2) {
      const apiHeaders      = apiData[0];
      const apiSellerSkuCol = apiHeaders.indexOf('sellerSku');
      const apiAsinCol      = apiHeaders.indexOf('asin');
      const apiProductCol   = apiHeaders.indexOf('productName');
      const apiCondCol      = apiHeaders.indexOf('condition');
      const apiTotalQtyCol  = apiHeaders.indexOf('totalQuantity');
      const apiFulfillCol   = apiHeaders.indexOf('fulfillableQuantity');

      if (apiSellerSkuCol !== -1) {
        for (let i = 1; i < apiData.length; i++) {
          const cellValue = String(apiData[i][apiSellerSkuCol] || '').trim();
          if (cellValue === sku) {
            return {
              source             : 'FBA在庫_API',
              eBaySkU            : sku,
              sellerSku          : cellValue,
              asin               : apiAsinCol     !== -1 ? String(apiData[i][apiAsinCol]    || '').trim() : '',
              productName        : apiProductCol  !== -1 ? String(apiData[i][apiProductCol] || '').trim() : '',
              active             : true,
              rowIndex           : i + 1,
              duplicated         : false,
              fulfillableQuantity: apiFulfillCol  !== -1 ? Number(apiData[i][apiFulfillCol]  || 0) : 0,
              totalQuantity      : apiTotalQtyCol !== -1 ? Number(apiData[i][apiTotalQtyCol] || 0) : 0
            };
          }
        }
      }
    }
  }

  // ★Step 3: FBA在庫_DL から検索（互換性フォールバック）
  //   SKU列のヘッダーは 'eBay SKU' または 'sku' のどちらかを受け付ける。
  const fbaStockDl = ss.getSheetByName('FBA在庫_DL');
  if (fbaStockDl) {
    const dlData = fbaStockDl.getDataRange().getValues();
    if (dlData.length >= 2) {
      const dlHeaders = dlData[0];

      // 'eBay SKU' を優先し、なければ 'sku' を使う
      let dlSkuCol = dlHeaders.indexOf('eBay SKU');
      if (dlSkuCol === -1) dlSkuCol = dlHeaders.indexOf('sku');

      if (dlSkuCol !== -1) {
        const dlAsinCol    = dlHeaders.indexOf('asin');
        const dlProductCol = dlHeaders.indexOf('product-name');
        const dlCondCol    = dlHeaders.indexOf('condition');
        const dlFulfillCol = dlHeaders.indexOf('afn-fulfillable-quantity');
        const dlTotalCol   = dlHeaders.indexOf('afn-total-quantity');

        for (let i = 1; i < dlData.length; i++) {
          const cellValue = String(dlData[i][dlSkuCol] || '').trim();
          if (cellValue === sku) {
            return {
              source             : 'FBA在庫_DL',
              eBaySkU            : sku,
              sellerSku          : String(dlData[i][1] || '').trim(),
              asin               : dlAsinCol    !== -1 ? String(dlData[i][dlAsinCol]    || '').trim() : String(dlData[i][2] || '').trim(),
              productName        : dlProductCol !== -1 ? String(dlData[i][dlProductCol] || '').trim() : String(dlData[i][3] || '').trim(),
              active             : true,
              rowIndex           : i + 1,
              duplicated         : false,
              fulfillableQuantity: dlFulfillCol !== -1 ? Number(dlData[i][dlFulfillCol] || 0) : 0,
              totalQuantity      : dlTotalCol   !== -1 ? Number(dlData[i][dlTotalCol]   || 0) : 0
            };
          }
        }
      }
    }
  }

  // 見つからない
  return null;
}

// ============================================================================
// ■ 2. FBA在庫取得関数
// ============================================================================

/**
 * FBA在庫_API から FBA在庫情報を取得
 * 
 * @param {SpreadsheetApp.Spreadsheet} ss
 * @param {string} sellerSku - Amazon seller SKU
 * @returns {object|null} {fulfillableQuantity, lastFetchTime, ...} or null
 */
function getFbaStockBySku_(ss, sellerSku) {
  if (!sellerSku) return null;

  const fbaSheet = ss.getSheetByName('FBA在庫_API');
  if (!fbaSheet) return null;

  const data = fbaSheet.getDataRange().getValues();
  if (data.length < 2) {
    return null;
  }

  const headers = data[0];
  const timeCol = headers.indexOf('取得日時');
  const sellerSkuCol = headers.indexOf('sellerSku');
  const fulfillableCol = headers.indexOf('fulfillableQuantity');

  if (sellerSkuCol === -1 || fulfillableCol === -1) return null;

  for (let i = 1; i < data.length; i++) {
    const cellValue = String(data[i][sellerSkuCol] || '').trim();
    if (cellValue === sellerSku) {
      let lastFetchTime = null;

      if (timeCol !== -1 && data[i][timeCol]) {
        const timeValue = data[i][timeCol];
        if (timeValue instanceof Date) {
          lastFetchTime = timeValue;
        } else if (typeof timeValue === 'string') {
          lastFetchTime = new Date(timeValue);
        } else if (typeof timeValue === 'number') {
          lastFetchTime = new Date(timeValue);
        }
      }

      return {
        sellerSku: sellerSku,
        fulfillableQuantity: Number(data[i][fulfillableCol] || 0),
        lastFetchTime: lastFetchTime,
        rowIndex: i + 1
      };
    }
  }

  return null;
}

// ============================================================================
// ■ 3. エラー判定関数（メイン）
// ============================================================================

/**
 * MCF作成前のエラー総合判定
 */
function validateMcfSkuAndStock_(ss, orderRow, skuInfo) {
  const result = {
    isValid: true,
    errorType: null,
    errorInfo: null
  };

  const colDefs = {
    eBayOrder: 1,
    eBaySkU: 2,
    productName: 3,
    quantity: 4,
    shipByDate: 14
  };

  // Step 1: SKU が見つかったか
  if (!skuInfo) {
    result.isValid = false;
    result.errorType = 'SKU_NOT_FOUND';
    result.errorInfo = {
      errorType: 'SKU_NOT_FOUND',
      severity: 3,
      eBayOrderId: String(orderRow[colDefs.eBayOrder] || '').trim(),
      eBaySkU: String(orderRow[colDefs.eBaySkU] || '').trim(),
      amazonSellerSku: null,
      asin: null,
      productName: String(orderRow[colDefs.productName] || '').trim(),
      orderQuantity: orderRow[colDefs.quantity] || 0,
      fulfillableQuantity: null,
      shipByDate: orderRow[colDefs.shipByDate] || '',
      sourceSheet: 'eBay受注ログ'
    };
    return result;
  }

  // Step 2: sellerSku が空欄か
  if (!skuInfo.sellerSku) {
    result.isValid = false;
    result.errorType = 'SELLER_SKU_EMPTY';
    result.errorInfo = {
      errorType: 'SELLER_SKU_EMPTY',
      severity: 3,
      eBayOrderId: String(orderRow[colDefs.eBayOrder] || '').trim(),
      eBaySkU: skuInfo.eBaySkU,
      amazonSellerSku: null,
      asin: skuInfo.asin,
      productName: skuInfo.productName,
      orderQuantity: orderRow[colDefs.quantity] || 0,
      fulfillableQuantity: null,
      shipByDate: orderRow[colDefs.shipByDate] || '',
      sourceSheet: skuInfo.source
    };
    return result;
  }

  // Step 3: MCF利用可否が ✅ か
  if (!skuInfo.active) {
    result.isValid = false;
    result.errorType = 'SKU_INACTIVE';
    result.errorInfo = {
      errorType: 'SKU_INACTIVE',
      severity: 2,
      eBayOrderId: String(orderRow[colDefs.eBayOrder] || '').trim(),
      eBaySkU: skuInfo.eBaySkU,
      amazonSellerSku: skuInfo.sellerSku,
      asin: skuInfo.asin,
      productName: skuInfo.productName,
      orderQuantity: orderRow[colDefs.quantity] || 0,
      fulfillableQuantity: null,
      shipByDate: orderRow[colDefs.shipByDate] || '',
      sourceSheet: skuInfo.source
    };
    return result;
  }

  // Step 4: FBA在庫を確認
  const fbaStock = getFbaStockBySku_(ss, skuInfo.sellerSku);

  if (fbaStock === null) {
    result.isValid = false;
    result.errorType = 'FBA_API_NOT_FOUND';
    result.errorInfo = {
      errorType: 'FBA_API_NOT_FOUND',
      severity: 2,
      eBayOrderId: String(orderRow[colDefs.eBayOrder] || '').trim(),
      eBaySkU: skuInfo.eBaySkU,
      amazonSellerSku: skuInfo.sellerSku,
      asin: skuInfo.asin,
      productName: skuInfo.productName,
      orderQuantity: orderRow[colDefs.quantity] || 0,
      fulfillableQuantity: null,
      shipByDate: orderRow[colDefs.shipByDate] || '',
      sourceSheet: 'FBA在庫_API'
    };
    return result;
  }

  // Step 5: FBA在庫が0か
  if (fbaStock.fulfillableQuantity === 0) {
    result.isValid = false;
    result.errorType = 'FBA_STOCK_ZERO';
    result.errorInfo = {
      errorType: 'FBA_STOCK_ZERO',
      severity: 3,
      eBayOrderId: String(orderRow[colDefs.eBayOrder] || '').trim(),
      eBaySkU: skuInfo.eBaySkU,
      amazonSellerSku: skuInfo.sellerSku,
      asin: skuInfo.asin,
      productName: skuInfo.productName,
      orderQuantity: orderRow[colDefs.quantity] || 0,
      fulfillableQuantity: fbaStock.fulfillableQuantity,
      shipByDate: orderRow[colDefs.shipByDate] || '',
      sourceSheet: 'FBA在庫_API'
    };
    return result;
  }

  // Step 6: FBA在庫が注文数より少ないか
  if (fbaStock.fulfillableQuantity < (orderRow[colDefs.quantity] || 0)) {
    result.isValid = false;
    result.errorType = 'FBA_STOCK_SHORTAGE';
    result.errorInfo = {
      errorType: 'FBA_STOCK_SHORTAGE',
      severity: 3,
      eBayOrderId: String(orderRow[colDefs.eBayOrder] || '').trim(),
      eBaySkU: skuInfo.eBaySkU,
      amazonSellerSku: skuInfo.sellerSku,
      asin: skuInfo.asin,
      productName: skuInfo.productName,
      orderQuantity: orderRow[colDefs.quantity] || 0,
      fulfillableQuantity: fbaStock.fulfillableQuantity,
      shipByDate: orderRow[colDefs.shipByDate] || '',
      sourceSheet: 'FBA在庫_API'
    };
    return result;
  }

  // Step 7: FBA在庫_APIの鮮度をチェック（24時間以内か）
  if (fbaStock.lastFetchTime) {
    const now = new Date();
    const hours = (now - fbaStock.lastFetchTime) / (1000 * 60 * 60);
    if (hours > 24) {
      result.isValid = false;
      result.errorType = 'FBA_API_STALE';
      result.errorInfo = {
        errorType: 'FBA_API_STALE',
        severity: 1,
        eBayOrderId: String(orderRow[colDefs.eBayOrder] || '').trim(),
        eBaySkU: skuInfo.eBaySkU,
        amazonSellerSku: skuInfo.sellerSku,
        asin: skuInfo.asin,
        productName: skuInfo.productName,
        orderQuantity: orderRow[colDefs.quantity] || 0,
        fulfillableQuantity: fbaStock.fulfillableQuantity,
        shipByDate: orderRow[colDefs.shipByDate] || '',
        sourceSheet: 'FBA在庫_API',
        lastFetchTime: fbaStock.lastFetchTime,
        hoursOld: hours.toFixed(1)
      };
      return result;
    }
  }

  // すべてOK：MCF作成可能
  result.isValid = true;
  return result;
}

// ============================================================================
// ■ 4. dry-run用ログ記録関数
// ============================================================================

/**
 * dry-run専用：エラーを MCF_SKU_ERROR_LOG_DRY_RUN に記録
 */
function logMcfSkuError_DryRun_(ss, errorInfo) {
  let dryRunLog = ss.getSheetByName('MCF_SKU_ERROR_LOG_DRY_RUN');

  if (!dryRunLog) {
    dryRunLog = ss.insertSheet('MCF_SKU_ERROR_LOG_DRY_RUN');
    dryRunLog.appendRow([
      '実行日時', '実行モード', 'run_id', 'eBay注文番号', 'eBay SKU',
      'Amazon sellerSku', '商品名', '参照元シート', '判定結果', '在庫数',
      'エラー種別', 'エラー詳細', 'ソース行', '作成停止', 'LINE通知対象', '次にやること'
    ]);
  }

  const now = new Date();
  const row = [
    Utilities.formatDate(now, 'Asia/Tokyo', 'yyyy/MM/dd HH:mm:ss'),
    'DRY_RUN',
    errorInfo.runId || '',
    errorInfo.eBayOrderId || '',
    errorInfo.eBaySkU || '',
    errorInfo.amazonSellerSku || '',
    errorInfo.productName || '',
    errorInfo.sourceSheet || '',
    '',
    errorInfo.fulfillableQuantity !== undefined && errorInfo.fulfillableQuantity !== null
      ? errorInfo.fulfillableQuantity : '',
    errorInfo.errorType || '',
    (errorInfo.errorType || '') + (errorInfo.errorDetail ? ' ' + errorInfo.errorDetail : ''),
    errorInfo.rowIndex || '',
    'YES',
    ['SKU_NOT_FOUND','SELLER_SKU_EMPTY','FBA_STOCK_ZERO','FBA_STOCK_SHORTAGE','MCF_CREATE_FAILED','ORDER_SKU_EMPTY'].includes(errorInfo.errorType) ? 'YES' : 'NO',
    errorInfo.nextAction || ''
  ];

  dryRunLog.appendRow(row);
}

// ============================================================================
// ■ 5. dry-run診断関数（メイン）
// ============================================================================

/**
 * dryRunValidateMcfOrders
 *
 * dry-run：全eBay受注を検証（本実行なし）
 * 実行後、MCF_VALIDATION_RUN_LOG にサマリー1行を必ず記録する。
 */
function dryRunValidateMcfOrders() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const eBayLog = ss.getSheetByName('eBay受注ログ');

  if (!eBayLog) {
    Logger.log('❌ エラー：「eBay受注ログ」シートが見つかりません');
    return;
  }

  const data = eBayLog.getDataRange().getValues();

  // ── run_id を生成（実行ごとに一意）──────────────────────────
  const runId = 'RUN_' + Date.now() + '_' + Math.random().toString(36).substring(2, 11);
  const runStartTime = new Date();

  // ── ログヘッダー出力 ─────────────────────────────────────────
  Logger.log('');
  Logger.log('━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━');
  Logger.log('【dry-run：MCF作成前エラー検知 検証開始】');
  Logger.log('━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━');
  Logger.log('実行日時：' + runStartTime.toLocaleString('ja-JP'));
  Logger.log('run_id  ：' + runId);
  Logger.log('実行モード：DRY_RUN（本実行なし）');
  Logger.log('');

  // ── 統計変数の初期化 ─────────────────────────────────────────
  let validCount = 0;
  let errorCount = 0;
  let processedCount = 0;
  let skipCount = 0;            // スキップ件数（既処理・MCF作成済み）
  const errorSummary = {};
  const errorTypeCount = {};

  // ── eBay受注ログをループ ──────────────────────────────────────
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const eBayOrder = String(row[1] || '').trim();
    const eBaySkU   = String(row[2] || '').trim();

    if (!eBayOrder) continue;

    // 既に処理済みの受注をスキップ
    const mcfStatus = String(row[15] || '').trim();
    if (mcfStatus && mcfStatus !== '未作成') {
      skipCount++;
      continue;
    }

    processedCount++;

    const skuInfo    = getSkuInfoBySku(ss, eBaySkU);
    const validation = validateMcfSkuAndStock_(ss, row, skuInfo);

    if (validation.isValid) {
      validCount++;
      Logger.log(`✅ 行${i+1}：${eBayOrder} - MCF作成可能`);
    } else {
      errorCount++;
      const errType = validation.errorType;
      errorSummary[errType]   = (errorSummary[errType]   || 0) + 1;
      errorTypeCount[errType] = (errorTypeCount[errType] || 0) + 1;

      let symbol = '';
      if (['SKU_NOT_FOUND','SELLER_SKU_EMPTY','FBA_STOCK_ZERO','FBA_STOCK_SHORTAGE','MCF_CREATE_FAILED'].includes(errType)) {
        symbol = '📢';
      } else if (['SKU_INACTIVE','SKU_DUPLICATED','FBA_API_NOT_FOUND'].includes(errType)) {
        symbol = '⚠️ ';
      } else {
        symbol = '  ';
      }

      Logger.log(`❌ 行${i+1}：${eBayOrder} - ${symbol} ${errType}`);

      // run_id を errorInfo に付与してログ記録
      if (validation.errorInfo) {
        validation.errorInfo.runId = runId;
        validation.errorInfo.rowIndex = i + 1;
      }
      logMcfSkuError_DryRun_(ss, validation.errorInfo);
    }
  }

  // ── 集計結果 ─────────────────────────────────────────────────
  Logger.log('');
  Logger.log('【集計結果】');
  Logger.log('処理対象受注：' + processedCount + '件');
  Logger.log('MCF作成可能：' + validCount + '件');
  Logger.log('エラー：' + errorCount + '件');
  Logger.log('スキップ：' + skipCount + '件');
  Logger.log('');

  Logger.log('【エラー分布】');
  if (errorCount === 0) {
    Logger.log('  （エラーなし）');
  } else {
    for (const [errType, count] of Object.entries(errorSummary)) {
      Logger.log(`  ${errType}：${count}件`);
    }
  }
  Logger.log('');

  // ── 警告・注記 ────────────────────────────────────────────────
  Logger.log('【⚠️  注記】');
  if ((errorSummary['FBA_API_NOT_FOUND'] || 0) > 0) {
    const fbaNotFoundCount = errorSummary['FBA_API_NOT_FOUND'];
    Logger.log(`FBA在庫_APIがヘッダーのみのため FBA_API_NOT_FOUND が ${fbaNotFoundCount} 件出ています`);
    Logger.log('（既知の状態、本実装時に「集計通知」または「初期OFF」で対応）');
    Logger.log('');
  }

  // ── LINE通知対象サマリー ─────────────────────────────────────
  const lineNotifyTargets = ['SKU_NOT_FOUND','SELLER_SKU_EMPTY','FBA_STOCK_ZERO','FBA_STOCK_SHORTAGE','MCF_CREATE_FAILED'];
  let lineNotifyCount = 0;
  for (const target of lineNotifyTargets) {
    lineNotifyCount += (errorSummary[target] || 0);
  }
  Logger.log('【LINE通知対象エラー（集計）】');
  Logger.log(`  ${lineNotifyCount} 件のエラーが LINE通知対象です`);
  Logger.log('  （本実装時に重複防止ロジックで最終判定）');
  Logger.log('');

  // ── 次のステップ ─────────────────────────────────────────────
  Logger.log('【✅ 次のステップ】');
  Logger.log('1. 上記の検証結果を確認してください');
  Logger.log('2. false positive（誤検知）がないか確認');
  Logger.log('3. 「エラーの種類・内容が正しい」と判定できたら');
  Logger.log('4. Shin が「本実装OK」の指示を出す');
  Logger.log('5. その後、autoCreateMcfOrders() の修正に進む');
  Logger.log('');

  Logger.log('━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━');
  Logger.log('【dry-run：検証完了】');
  Logger.log('━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━');
  Logger.log('');

  // ── MCF_VALIDATION_RUN_LOG にサマリー1行を記録 ────────────────
  // ★ try-catch で保護。記録失敗時はログに出力する
  try {
    logValidationRunSummary_(ss, {
      runId         : runId,
      runStartTime  : runStartTime,
      processedCount: processedCount,
      validCount    : validCount,
      errorCount    : errorCount,
      skipCount     : skipCount,
      errorSummary  : errorSummary
    });
  } catch (e) {
    Logger.log('⚠️ [MCF_VALIDATION_RUN_LOG 記録失敗] ' + e.toString());
  }
}

// ============================================================================
// ■ 6. MCF_VALIDATION_RUN_LOG：dry-run 実行サマリー記録
// ============================================================================

/**
 * logValidationRunSummary_
 *
 * dryRunValidateMcfOrders() の実行ごとに MCF_VALIDATION_RUN_LOG シートへ
 * サマリー1行を記録する。
 *
 * ヘッダー（8列）：
 *   A: 実行日時  B: run_id  C: 処理対象件数  D: MCF作成可能件数
 *   E: エラー件数  F: スキップ件数  G: 判定結果  H: 次にやること
 */
function logValidationRunSummary_(ss, stats) {

  // ── シートを取得または新規作成 ──────────────────────────────
  let runLog = ss.getSheetByName('MCF_VALIDATION_RUN_LOG');

  if (!runLog) {
    runLog = ss.insertSheet('MCF_VALIDATION_RUN_LOG');
    runLog.setTabColor('#4FC3F7');  // 規則：自動生成シートのタブ色

    runLog.appendRow([
      '実行日時',         // A
      'run_id',          // B
      '処理対象件数',     // C
      'MCF作成可能件数',  // D
      'エラー件数',       // E
      'スキップ件数',     // F
      '判定結果',         // G
      '次にやること'      // H
    ]);

    const header = runLog.getRange(1, 1, 1, 8);
    header.setBackground('#1a73e8');
    header.setFontColor('#ffffff');
    header.setFontWeight('bold');
    runLog.setFrozenRows(1);
  }

  // ── 判定結果を自動生成 ───────────────────────────────────────
  const { errorSummary, errorCount } = stats;
  let judgeText = '';

  if (errorCount === 0) {
    judgeText = '✅ 全件MCF作成可能';
  } else {
    const errTypes = Object.keys(errorSummary);
    const critical = ['SKU_NOT_FOUND','SELLER_SKU_EMPTY','FBA_STOCK_ZERO','FBA_STOCK_SHORTAGE'];
    const hasCritical = errTypes.some(t => critical.includes(t));

    if (hasCritical) {
      judgeText = '❌ 要対応 ' + errTypes.length + '種のエラー';
    } else if (errTypes.every(t => t === 'ORDER_QUANTITY_INVALID')) {
      judgeText = '⚠️ 数量入力待ち';
    } else if (errTypes.every(t => t === 'FBA_API_NOT_FOUND')) {
      judgeText = '⚠️ FBA在庫_API未反映';
    } else if (errTypes.every(t => t === 'ORDER_SKU_EMPTY')) {
      judgeText = '⚠️ SKU空欄あり（手動対応）';
    } else {
      judgeText = '⚠️ 確認推奨 ' + errTypes.join(' / ');
    }
  }

  // ── 次にやることを自動生成 ──────────────────────────────────
  let nextAction = '';

  if (errorCount === 0) {
    nextAction = '設計AIにレビューを依頼し、本実装OK判断を Shin へ確認';
  } else {
    const actions = [];
    if (errorSummary['ORDER_QUANTITY_INVALID'])
      actions.push('MCF_QuantityPatch.gs Phase A→B で数量補完');
    if (errorSummary['ORDER_SKU_EMPTY'])
      actions.push('eBay受注ログのSKUを手動補完');
    if (errorSummary['FBA_API_NOT_FOUND'])
      actions.push('FBA_InventoryPatch.gs Phase B を実行して FBA在庫_API に反映');
    if (errorSummary['SKU_NOT_FOUND'] || errorSummary['SELLER_SKU_EMPTY'])
      actions.push('MCF_SKU_MAP にSKUを追加登録');
    if (errorSummary['FBA_STOCK_ZERO'] || errorSummary['FBA_STOCK_SHORTAGE'])
      actions.push('Amazon FBA補充または手動出荷を確認');
    if (actions.length === 0)
      actions.push('MCF_SKU_ERROR_LOG_DRY_RUN を確認');
    nextAction = actions.join(' ／ ');
  }

  // ── 1行記録 ────────────────────────────────────────────────
  const row = [
    Utilities.formatDate(stats.runStartTime, 'Asia/Tokyo', 'yyyy/MM/dd HH:mm:ss'), // A
    stats.runId,           // B
    stats.processedCount,  // C
    stats.validCount,      // D
    stats.errorCount,      // E
    stats.skipCount,       // F
    judgeText,             // G
    nextAction             // H
  ];

  runLog.appendRow(row);

  // 判定結果セルの文字色
  const lastRow = runLog.getLastRow();
  const judgeCell = runLog.getRange(lastRow, 7);
  if (errorCount === 0) {
    judgeCell.setFontColor('#137333');  // 緑
  } else if (Object.keys(errorSummary).some(t =>
    ['SKU_NOT_FOUND','SELLER_SKU_EMPTY','FBA_STOCK_ZERO','FBA_STOCK_SHORTAGE'].includes(t))) {
    judgeCell.setFontColor('#c5221f');  // 赤
  } else {
    judgeCell.setFontColor('#e37400');  // オレンジ
  }

  Logger.log('');
  Logger.log('【MCF_VALIDATION_RUN_LOG 記録完了】');
  Logger.log('  run_id       : ' + stats.runId);
  Logger.log('  処理対象件数 : ' + stats.processedCount);
  Logger.log('  MCF作成可能  : ' + stats.validCount);
  Logger.log('  エラー件数   : ' + stats.errorCount);
  Logger.log('  スキップ件数 : ' + stats.skipCount);
  Logger.log('  判定結果     : ' + judgeText);
  Logger.log('  次にやること : ' + nextAction);
  Logger.log('');
}

// ============================================================================
// ■ 7. ユーティリティ関数
// ============================================================================

/**
 * テスト関数：単一SKUのエラー判定を確認
 */
function testSingleSku() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const testSku = 'UT-VJKJ-CCTQ';

  Logger.log('');
  Logger.log('【テスト：単一SKU検証】');
  Logger.log('対象SKU：' + testSku);

  const skuInfo = getSkuInfoBySku(ss, testSku);
  if (!skuInfo) {
    Logger.log('❌ SKU が見つかりません');
    return;
  }

  Logger.log('✅ SKU情報：');
  Logger.log('  source: ' + skuInfo.source);
  Logger.log('  sellerSku: ' + skuInfo.sellerSku);
  Logger.log('  asin: ' + skuInfo.asin);
  Logger.log('  productName: ' + skuInfo.productName);
  Logger.log('  active: ' + skuInfo.active);

  const fbaStock = getFbaStockBySku_(ss, skuInfo.sellerSku);
  if (!fbaStock) {
    Logger.log('ℹ  FBA在庫：見つかりません（FBA在庫_APIがヘッダーのみ）');
  } else {
    Logger.log('✅ FBA在庫情報：');
    Logger.log('  fulfillableQuantity: ' + fbaStock.fulfillableQuantity);
    Logger.log('  lastFetchTime: ' + (fbaStock.lastFetchTime || 'N/A'));
  }
}

/**
 * 作業用GASの確認：関数リストを表示
 */
function checkInstalledFunctions() {
  Logger.log('');
  Logger.log('【インストール確認】');
  Logger.log('✅ getSkuInfoBySku() - SKU情報取得（MCF_SKU_MAP→FBA在庫_API→FBA在庫_DL）');
  Logger.log('✅ getFbaStockBySku_() - FBA在庫取得');
  Logger.log('✅ validateMcfSkuAndStock_() - エラー判定');
  Logger.log('✅ logMcfSkuError_DryRun_() - dry-runエラーログ記録');
  Logger.log('✅ dryRunValidateMcfOrders() - dry-run診断（メイン）');
  Logger.log('✅ logValidationRunSummary_() - MCF_VALIDATION_RUN_LOG サマリー記録');
  Logger.log('✅ testSingleSku() - テスト関数');
  Logger.log('✅ testGetSkuInfoBySkuForValidation() - SKU参照先確認');
  Logger.log('✅ checkInstalledFunctions() - この関数');
  Logger.log('');
  Logger.log('実行方法：');
  Logger.log('  1. 関数を選択：dryRunValidateMcfOrders');
  Logger.log('  2. ▶︎ 実行ボタンをクリック');
  Logger.log('  3. 下部「実行トランスクリプト」でログを確認');
  Logger.log('  4. スプレッドシートの MCF_VALIDATION_RUN_LOG シートでサマリーを確認');
  Logger.log('');
}

// ============================================================================
// ■ 8. 単体確認：getSkuInfoBySku() の参照先修正確認
// ============================================================================

/**
 * testGetSkuInfoBySkuForValidation
 *
 * getSkuInfoBySku(ss, 'UT-VJKJ-CCTQ') が正しく動作するかを確認する。
 * シートへの書き込みは行わない。
 */
function testGetSkuInfoBySkuForValidation() {

  const TEST_SKUS = ['UT-VJKJ-CCTQ', '2E-JR3L-ZA0C', 'NOTEXIST-SKU-TEST'];

  Logger.log('');
  Logger.log('='.repeat(70));
  Logger.log('  testGetSkuInfoBySkuForValidation()');
  Logger.log('  getSkuInfoBySku() の参照先修正確認');
  Logger.log('  対象: ' + TEST_SKUS.join(' / '));
  Logger.log('='.repeat(70));
  Logger.log('');

  const ss = SpreadsheetApp.getActiveSpreadsheet();

  TEST_SKUS.forEach(function(sku, idx) {
    Logger.log('──────────────────────────────────────────────────────────────────────');
    Logger.log('  [テスト ' + (idx + 1) + '/' + TEST_SKUS.length + '] SKU: ' + sku);

    const info = getSkuInfoBySku(ss, sku);

    if (info === null) {
      Logger.log('  戻り値: null（SKU未発見）');
    } else {
      Logger.log('  戻り値: ✅ オブジェクト取得成功');
      Logger.log('    source             : ' + (info.source              || '（なし）'));
      Logger.log('    sellerSku          : ' + (info.sellerSku           || '（なし）'));
      Logger.log('    asin               : ' + (info.asin                || '（なし）'));
      Logger.log('    productName        : ' + (info.productName         || '（なし）').substring(0, 60));
      Logger.log('    active             : ' + info.active);
      Logger.log('    fulfillableQuantity: ' + (info.fulfillableQuantity !== undefined ? info.fulfillableQuantity : '（なし）'));
      Logger.log('    totalQuantity      : ' + (info.totalQuantity       !== undefined ? info.totalQuantity       : '（なし）'));
    }

    Logger.log('');
  });

  Logger.log('='.repeat(70));
  Logger.log('  テスト完了');
  Logger.log('='.repeat(70));
  Logger.log('');
}

// ============================================================================
// end of file  (MCF_ErrorValidation_Work.gs  v3.5.0)
// ============================================================================

/**
 * FBA_InventoryCore.gs
 *
 * FBA在庫を SP-API から全件取得する共通ヘルパー。
 *
 * 【残す理由】
 * - 他の処理から _fetchAllFbaInventory_(token) が呼ばれる可能性があるため
 * - FBA在庫_API 更新・MCF_SKU_MAP 登録・在庫確認系の共通取得処理として利用可能
 *
 * 【前提】
 * - SP_API_ENDPOINT が別ファイルで定義済み
 * - MARKETPLACE_ID が別ファイルで定義済み
 * - _spApiGetRaw_(url, token) が別ファイルで定義済み
 *
 * 【注意】
 * - この関数単体ではシート更新しない
 * - MCF作成APIは呼ばない
 * - LINE通知しない
 * - 取得した在庫配列を返すだけ
 */

/**
 * 全FBA在庫を取得する
 *
 * @param {string} token SP-API access token
 * @return {Array<Object>} FBA在庫一覧
 */
function _fetchAllFbaInventory_(token) {
  const results = [];
  let nextToken = null;
  let pageCount = 0;

  do {
    const url = SP_API_ENDPOINT + '/fba/inventory/v1/summaries'
      + '?details=true'
      + '&granularityType=Marketplace'
      + '&granularityId=' + MARKETPLACE_ID
      + '&marketplaceIds=' + MARKETPLACE_ID
      + (nextToken ? '&nextToken=' + encodeURIComponent(nextToken) : '');

    const res = _spApiGetRaw_(url, token);

    if (res.code !== 200) {
      Logger.log('⚠️ FBA在庫取得エラー：ページ '
        + (pageCount + 1)
        + ' / HTTP '
        + res.code);
      Logger.log('レスポンス先頭500文字：'
        + String(res.body || '').substring(0, 500));
      break;
    }

    const json = JSON.parse(res.body);
    const summaries = (json.payload && json.payload.inventorySummaries) || [];

    Logger.log('FBA在庫取得 ページ '
      + (pageCount + 1)
      + '：'
      + summaries.length
      + '件');

    summaries.forEach(function(s) {
      const d = s.inventoryDetails || {};

      results.push({
        sellerSku               : s.sellerSku || '',
        fnSku                   : s.fnSku || '',
        asin                    : s.asin || '',
        productName             : s.productName || '',
        condition               : s.condition || '',
        totalQuantity           : s.totalQuantity !== undefined ? s.totalQuantity : 0,
        fulfillableQuantity     : d.fulfillableQuantity || 0,
        reservedQuantity        : d.reservedQuantity || {},
        inboundWorkingQuantity  : d.inboundWorkingQuantity || 0,
        inboundShippedQuantity  : d.inboundShippedQuantity || 0,
        inboundReceivingQuantity: d.inboundReceivingQuantity || 0,
        lastUpdatedTime         : s.lastUpdatedTime || ''
      });
    });

    pageCount++;
    nextToken = (json.pagination && json.pagination.nextToken) || null;

    // SP-API rate limit対策
    Utilities.sleep(700);

  } while (nextToken && pageCount < 50);

  Logger.log('FBA在庫取得完了：'
    + results.length
    + '件 / '
    + pageCount
    + 'ページ');

  return results;
}/**
 * 現在のトリガー一覧を確認する
 */
function listMcfAutomationTriggers() {
  const triggers = ScriptApp.getProjectTriggers();

  Logger.log('====================================');
  Logger.log('現在のトリガー一覧');
  Logger.log('====================================');

  if (triggers.length === 0) {
    Logger.log('トリガーはありません');
    return;
  }

  triggers.forEach(function(trigger, index) {
    Logger.log(
      (index + 1) +
      '. 関数: ' +
      trigger.getHandlerFunction() +
      ' / 種類: ' +
      trigger.getEventType()
    );
  });

  Logger.log('====================================');
}


/**
 * 旧MCF系トリガーだけ削除する
 * 在庫同期・メール監視・eBayRaw更新などは削除しない
 */
function deleteLegacyMcfTriggersOnly() {
  const legacyFunctions = [
    'checkMcfTrackingNumbers',
    'sendMcfReminder',
    'checkMcfTrackingForRow9'
  ];

  let deletedCount = 0;

  ScriptApp.getProjectTriggers().forEach(function(trigger) {
    const fn = trigger.getHandlerFunction();

    if (legacyFunctions.indexOf(fn) !== -1) {
      ScriptApp.deleteTrigger(trigger);
      Logger.log('削除: 旧MCFトリガー ' + fn);
      deletedCount++;
    }
  });

  Logger.log('====================================');
  Logger.log('旧MCFトリガー削除完了: ' + deletedCount + '件');
  Logger.log('====================================');
}


/**
 * 本番用トリガーを作り直す
 * dryRun系は登録しない
 */
function setupMcfAutomationTriggers() {
  const activeFunctions = [
    'checkEbayOrders',
    'autoCreateMcfOrders',
    'checkMcfTrackingNumbersAuto'
  ];

  const legacyFunctions = [
    'checkMcfTrackingNumbers',
    'sendMcfReminder',
    'checkMcfTrackingForRow9'
  ];

  const deleteTargets = activeFunctions.concat(legacyFunctions);

  ScriptApp.getProjectTriggers().forEach(function(trigger) {
    const fn = trigger.getHandlerFunction();

    if (deleteTargets.indexOf(fn) !== -1) {
      ScriptApp.deleteTrigger(trigger);
      Logger.log('削除: ' + fn);
    }
  });

  ScriptApp.newTrigger('checkEbayOrders')
    .timeBased()
    .everyHours(1)
    .create();
  Logger.log('作成: checkEbayOrders / 1時間ごと');

  ScriptApp.newTrigger('autoCreateMcfOrders')
    .timeBased()
    .everyHours(1)
    .create();
  Logger.log('作成: autoCreateMcfOrders / 1時間ごと');

  ScriptApp.newTrigger('checkMcfTrackingNumbersAuto')
    .timeBased()
    .everyHours(1)
    .create();
  Logger.log('作成: checkMcfTrackingNumbersAuto / 1時間ごと');

  Logger.log('====================================');
  Logger.log('✅ MCF/eBay 本番トリガー再設定完了');
  Logger.log('====================================');
}function checkForbiddenWords() {
  // チェック対象のテキスト（例として、関数の文字化など）
  // 実際にはチェックしたい文字列や、シートから取得したテキストを入れます
  let targetText = "'Shimano 23 Vanquish C2000SHG spinning reel',";
  
  // 残っていてほしくない文言のリスト
  const forbiddenWords = ["テスト", "修正中", "TODO", "ダミー"];
  
  let foundWords = [];
  
  forbiddenWords.forEach(word => {
    if (targetText.includes(word)) {
      foundWords.push(word);
    }
  });
  
  if (foundWords.length > 0) {
    console.error("❌ 以下の文言が残っています: " + foundWords.join(", "));
  } else {
    console.log("✅ すべてクリア！特定の文言は見つかりませんでした。");
  }
}

/**
const MCF_MONITOR_CONFIG = {
  spreadsheetId: '1exGBAEx99-2Qc9d0DLRiZbygvgIo4FY_NoBG3Nkan-A',
  orderSheetName: 'eBay受注ログ',
  logSheetName: 'MCF監視ログ',
  notifyEmailProperty: 'NOTIFY_EMAIL',
  defaultNotifyEmail: 'shinicchee@gmail.com',
  timezone: 'Asia/Tokyo',
};

const MCF_MONITOR_HEADERS = [
  'MCF監視判定',
  'MCF監視理由',
  'MCF経過時間',
  '発送期限残り時間',
  'MCF通知レベル',
  'MCF通知済み',
  'MCF最終通知日時',
];

const MCF_LEVEL_SCORE = {
  '対象外': 0,
  '正常': 0,
  '注意': 1,
  '警告': 2,
  '危険': 3,
};

function judgeMcfRow_(row, headerMap, now) {
  const orderDate = parseDate_(row[headerMap['受注日時(JST)'] - 1]);
  const ebayOrderId = String(row[headerMap['eBay注文番号'] - 1] || '').trim();
  const deadline = parseDate_(row[headerMap['発送期限'] - 1]);
  const mcfCreated = String(row[headerMap['MCF作成'] - 1] || '').trim();
  const status = String(row[headerMap['ステータス'] - 1] || '').trim();
  const tracking = String(row[headerMap['Tracking'] - 1] || '').trim();
  const trackingStatus = String(row[headerMap['Tracking確認ステータス'] - 1] || '').trim();

  if (!ebayOrderId) return result_('対象外', '注文番号なし', '', '');

  if (
    tracking ||
    trackingStatus === '追跡番号取得済' ||
    trackingStatus === 'チェック不要' ||
    status === '手動対応済み'
  ) {
    return result_('対象外', '追跡取得済みまたはチェック不要', '', formatDeadlineText_(deadline, now));
  }

  let level = '正常';
  const reasons = [];

  if (mcfCreated === '未作成' && status === '受注済') {
    const h = hoursBetween_(orderDate, now);
    if (h >= 24) {
      level = maxLevel_(level, '危険');
      reasons.push('MCF未作成 24h超');
    } else if (h >= 12) {
      level = maxLevel_(level, '警告');
      reasons.push('MCF未作成 12h超');
    } else if (h >= 6) {
      level = maxLevel_(level, '注意');
      reasons.push('MCF未作成 6h超');
    }
  }

  if (status === 'MCF作成済' && !tracking) {
    const baseDate = orderDate ? new Date(orderDate.getTime() + 12 * 60 * 60 * 1000) : null;
    const h = hoursBetween_(baseDate, now);
    if (h >= 48) {
      level = maxLevel_(level, '危険');
      reasons.push('MCF作成済み追跡なし 48h超');
    } else if (h >= 24) {
      level = maxLevel_(level, '警告');
      reasons.push('MCF作成済み追跡なし 24h超');
    } else if (h >= 12) {
      level = maxLevel_(level, '注意');
      reasons.push('MCF作成済み追跡なし 12h超');
    }
  }

  const deadlineHours = hoursBetween_(now, deadline);
  if (deadline) {
    if (deadlineHours < 0) {
      level = maxLevel_(level, '危険');
      reasons.push('eBay発送期限超過');
    } else if (deadlineHours <= 24) {
      level = maxLevel_(level, '警告');
      reasons.push('eBay発送期限24h以内');
    }
  }

  return result_(
    level,
    reasons.length ? reasons.join(' / ') : '異常なし',
    orderDate ? Math.floor(hoursBetween_(orderDate, now)) + 'h' : '',
    formatDeadlineText_(deadline, now)
  );
}

function sendMcfMonitorEmail_(notifyRows, headerMap) {
  const props = PropertiesService.getScriptProperties();
  const to = sampleCsvProps_().getProperty(MCF_MONITOR_CONFIG.notifyEmailProperty) || MCF_MONITOR_CONFIG.defaultNotifyEmail;

  const lines = [];
  lines.push('MCF監視で通知対象が発生しました。');
  lines.push('');
  lines.push('対象件数: ' + notifyRows.length);
  lines.push('');

  notifyRows.forEach((item) => {
    const row = item.row;
    lines.push('------------------------------');
    lines.push('判定: ' + item.result.level);
    lines.push('理由: ' + item.result.reason);
    lines.push('eBay注文番号: ' + valueByHeader_(row, headerMap, 'eBay注文番号'));
    lines.push('SKU: ' + valueByHeader_(row, headerMap, 'SKU'));
    lines.push('商品名: ' + valueByHeader_(row, headerMap, '商品名'));
    lines.push('MCF作成: ' + valueByHeader_(row, headerMap, 'MCF作成'));
    lines.push('ステータス: ' + valueByHeader_(row, headerMap, 'ステータス'));
    lines.push('Tracking確認ステータス: ' + valueByHeader_(row, headerMap, 'Tracking確認ステータス'));
    lines.push('経過時間: ' + item.result.elapsedText);
    lines.push('発送期限: ' + valueByHeader_(row, headerMap, '発送期限'));
    lines.push('発送期限残り: ' + item.result.deadlineText);
  });

  MailApp.sendEmail({
    to,
    subject: '【MCF監視】通知対象 ' + notifyRows.length + '件',
    body: lines.join('\n'),
  });
}

function ensureMcfMonitorColumns_(sheet) {
  const lastCol = sheet.getLastColumn();
  const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0].map(String);
  let col = lastCol + 1;

  MCF_MONITOR_HEADERS.forEach((h) => {
    if (headers.indexOf(h) === -1) {
      sheet.getRange(1, col).setValue(h);
      col++;
    }
  });
}

function getHeaderMap_(sheet) {
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const map = {};
  headers.forEach((h, i) => {
    map[String(h).trim()] = i + 1;
  });
  return map;
}

function result_(level, reason, elapsedText, deadlineText) {
  return { level, reason, elapsedText, deadlineText };
}

function maxLevel_(a, b) {
  return MCF_LEVEL_SCORE[b] > MCF_LEVEL_SCORE[a] ? b : a;
}

function hoursBetween_(from, to) {
  if (!from || !to) return 0;
  return (to.getTime() - from.getTime()) / 3600000;
}

function formatDeadlineText_(deadline, now) {
  if (!deadline) return '';
  const h = hoursBetween_(now, deadline);
  if (h < 0) return '超過' + Math.abs(Math.floor(h)) + 'h';
  return '残り' + Math.floor(h) + 'h';
}

function valueByHeader_(row, headerMap, header) {
  const col = headerMap[header];
  if (!col) return '';
  return row[col - 1] || '';
}

function parseDate_(value) {
  if (!value) return null;
  if (value instanceof Date) return value;

  const s = String(value).trim();

  const normalized = s
    .replace(/Jan/i, 'January')
    .replace(/Feb/i, 'February')
    .replace(/Mar/i, 'March')
    .replace(/Apr/i, 'April')
    .replace(/Jun/i, 'June')
    .replace(/Jul/i, 'July')
    .replace(/Aug/i, 'August')
    .replace(/Sep/i, 'September')
    .replace(/Oct/i, 'October')
    .replace(/Nov/i, 'November')
    .replace(/Dec/i, 'December');

  const d = new Date(normalized);
  return isNaN(d.getTime()) ? null : d;
}

function testMcfMailPermission() {
  MailApp.sendEmail({
    to: 'shinicchee@gmail.com',
    subject: 'MCF監視メール権限テスト',
    body: 'MailApp.sendEmail の権限確認テストです。'
  });
}

/**
 * MCF_AnomalyMonitor.gs  v2
 *
 * eBay受注ログ MCF異常監視システム
 *
 * 【メイン関数】
 *   runMcfAnomalyMonitor()
 *
 * 【処理概要】
 *   1. eBay受注ログを読み取る
 *   2. 監視列（V〜AB）がなければ右端に作成
 *   3. 各行を監視判定する
 *   4. 判定結果を監視列へ書き込む
 *   5. 注意・警告・危険があれば shinicchee@gmail.com にメール通知する
 *
 * 【追加列】（U列の右）
 *   V: MCF監視判定      （正常 / 注意 / 警告 / 危険 / 対象外）
 *   W: MCF監視理由
 *   X: MCF経過時間
 *   Y: 発送期限残り時間
 *   Z: MCF通知レベル
 *   AA: MCF通知済み
 *   AB: MCF最終通知日時
 *
 * 【判定ルール】（v2 変更後）
 *
 *   A. MCF未作成 かつ ステータス=受注済
 *      - 受注から 18h超 → 注意
 *      - 受注から 24h超 → 警告
 *      - 受注から 36h超 → 危険
 *
 *   B. MCF作成済 かつ Tracking空欄（発送期限ベース）
 *      - 発送期限まで 72h以内 → 注意
 *      - 発送期限まで 24h以内 → 警告
 *      - 発送期限まで 12h以内 → 危険
 *      - 発送期限超過          → 危険
 *      ※ MCF作成からの経過時間判定は廃止
 *
 * 【v2 変更点】
 *   - ANOMALY_THRESHOLD_：しきい値を全面改訂
 *   - getMcfAnomalyJudgment_：判定ロジックを変更
 *     - 判定A：6/12/24h → 18/24/36h
 *     - 判定B：MCF作成経過時間ベース → 発送期限残り時間ベース（72/24/12h）
 *     - 旧判定C/D（独立した発送期限チェック）を削除 → 判定Bに統合済み
 *   - その他（通知・ログ・列管理・メイン関数）は変更なし
 *
 * 【除外条件】
 *   - Trackingが入っている行
 *   - Tracking確認ステータス = 追跡番号取得済
 *   - Tracking確認ステータス = チェック不要
 *   - ステータス = 手動対応済み
 *   - ステータス = 確認通知済み かつ Trackingあり
 *
 * 【メール重複防止】
 *   MCF通知済み列に同じレベルがある場合は再通知しない。
 *   レベルが上がった（注意→警告、警告→危険など）場合のみ再通知。
 *
 * 【安全条件】
 *   ★ MCF作成APIは呼ばない
 *   ★ eBay追跡番号提出APIは呼ばない
 *   ★ LINE通知は使わない（MailApp.sendEmail を使用）
 *   ★ 既存列（A〜U）は変更しない
 *   ★ 既存トリガーは変更しない
 *   ★ 書き込みは監視列（V〜AB）と MCF監視ログ シートのみ
 */

// ============================================================
// ■ 定数
// ============================================================

const ANOMALY_NOTIFY_EMAIL_    = 'shinicchee@gmail.com';
const ANOMALY_LOG_SHEET_NAME_  = 'MCF監視ログ';
const ANOMALY_EBAY_LOG_SHEET_  = 'eBay受注ログ';

// eBay受注ログ 既存列インデックス（0始まり）
const ANOMALY_COL_ = {
  RECEIVED_AT   :  0,  // A: 受注日時(JST)
  ORDER_ID      :  1,  // B: eBay注文番号
  SKU           :  2,  // C: SKU
  PRODUCT_NAME  :  3,  // D: 商品名
  SHIP_BY       : 14,  // O: 発送期限
  MCF_ID        : 15,  // P: MCF作成
  STATUS        : 16,  // Q: ステータス
  CARRIER       : 17,  // R: Carrier
  TRACKING      : 18,  // S: Tracking
  TRACKING_AT   : 19,  // T: Tracking取得日時
  TRACKING_STS  : 20   // U: Tracking確認ステータス
};

// 追加する監視列のヘッダー名（順序を維持）
const ANOMALY_MONITOR_HEADERS_ = [
  'MCF監視判定',       // V
  'MCF監視理由',       // W
  'MCF経過時間',       // X
  '発送期限残り時間',  // Y
  'MCF通知レベル',     // Z
  'MCF通知済み',       // AA
  'MCF最終通知日時'    // AB
];

// 通知レベルの優先度（数値が大きいほど高優先）
const ANOMALY_LEVEL_PRIORITY_ = {
  '対象外': -1,
  '正常'  :  0,
  '注意'  :  1,
  '警告'  :  2,
  '危険'  :  3
};

// ★ v2 変更：判定ルールのしきい値（時間）
const ANOMALY_THRESHOLD_ = {
  // 判定A：MCF未作成 & ステータス=受注済（受注からの経過時間）
  MCF_NOT_CREATED_CAUTION : 18,  // 受注後 18h超 → 注意
  MCF_NOT_CREATED_WARNING : 24,  // 受注後 24h超 → 警告
  MCF_NOT_CREATED_DANGER  : 36,  // 受注後 36h超 → 危険

  // 判定B：MCF作成済 & Tracking空欄（発送期限までの残り時間）
  // ★ 方向に注意：残り時間が「以下」のときにトリガー
  TRACKING_CAUTION        : 72,  // 発送期限まで 72h以内 → 注意
  TRACKING_WARNING        : 24,  // 発送期限まで 24h以内 → 警告
  TRACKING_DANGER         : 12   // 発送期限まで 12h以内 → 危険 / 超過 → 危険
};

// ============================================================
// ■ 0. ensureAnomalyColumns_
//       監視列がなければ右端に追加し、ヘッダー名→列番号（1始まり）の
//       マップを返す
// ============================================================

/**
 * @param {SpreadsheetApp.Sheet} sheet - eBay受注ログシート
 * @returns {Object} colMap  例: { 'MCF監視判定': 22, 'MCF監視理由': 23, ... }
 */
function ensureAnomalyColumns_(sheet) {
  const lastCol = sheet.getLastColumn();
  const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];

  const colMap = {};

  // 既存ヘッダーから監視列の列番号を探す
  ANOMALY_MONITOR_HEADERS_.forEach(function(name) {
    const idx = headers.indexOf(name);
    if (idx >= 0) {
      colMap[name] = idx + 1; // 1始まり
    }
  });

  // 存在しない監視列を右端に追加
  let nextCol = lastCol + 1;
  ANOMALY_MONITOR_HEADERS_.forEach(function(name) {
    if (!colMap[name]) {
      sheet.getRange(1, nextCol).setValue(name);
      const hdrCell = sheet.getRange(1, nextCol);
      hdrCell.setBackground('#e8eaf6');
      hdrCell.setFontWeight('bold');
      colMap[name] = nextCol;
      nextCol++;
    }
  });

  // 列幅を設定
  if (colMap['MCF監視判定'])    sheet.setColumnWidth(colMap['MCF監視判定'],    100);
  if (colMap['MCF監視理由'])    sheet.setColumnWidth(colMap['MCF監視理由'],    300);
  if (colMap['MCF経過時間'])    sheet.setColumnWidth(colMap['MCF経過時間'],    90);
  if (colMap['発送期限残り時間']) sheet.setColumnWidth(colMap['発送期限残り時間'], 110);
  if (colMap['MCF通知レベル'])  sheet.setColumnWidth(colMap['MCF通知レベル'],  90);
  if (colMap['MCF通知済み'])    sheet.setColumnWidth(colMap['MCF通知済み'],    80);
  if (colMap['MCF最終通知日時']) sheet.setColumnWidth(colMap['MCF最終通知日時'], 140);

  return colMap;
}

// ============================================================
// ■ 1. getMcfAnomalyJudgment_  ★ v2 変更
//       1行分の受注データを監視判定して結果を返す内部関数
// ============================================================

/**
 * getMcfAnomalyJudgment_
 *
 * v2 変更点：
 *   - 判定A（MCF未作成）のしきい値を 6/12/24h → 18/24/36h に変更
 *   - 判定B（MCF作成済・Tracking空欄）を発送期限残り時間ベースに変更
 *     72h以内→注意 / 24h以内→警告 / 12h以内→危険 / 超過→危険
 *   - 旧判定C/D（独立発送期限チェック）を削除（判定Bに統合）
 *
 * @param {Array} row - allData[i]（0始まり配列）
 * @returns {{
 *   level            : string,
 *   reason           : string,
 *   mcfElapsedStr    : string,
 *   deadlineRemaining: string
 * }}
 */
function getMcfAnomalyJudgment_(row) {
  const receivedRaw  = row[ANOMALY_COL_.RECEIVED_AT];
  const mcfId        = String(row[ANOMALY_COL_.MCF_ID]        || '').trim();
  const status       = String(row[ANOMALY_COL_.STATUS]         || '').trim();
  const tracking     = String(row[ANOMALY_COL_.TRACKING]       || '').trim();
  const trackingSts  = String(row[ANOMALY_COL_.TRACKING_STS]   || '').trim();
  const shipByRaw    = row[ANOMALY_COL_.SHIP_BY];

  const now = new Date();

  // ── 除外条件チェック ──────────────────────────────────────
  if (tracking)
    return { level: '対象外', reason: 'Tracking取得済み',          mcfElapsedStr: '', deadlineRemaining: '' };
  if (trackingSts === '追跡番号取得済')
    return { level: '対象外', reason: '追跡番号取得済',              mcfElapsedStr: '', deadlineRemaining: '' };
  if (trackingSts === 'チェック不要')
    return { level: '対象外', reason: 'チェック不要',                mcfElapsedStr: '', deadlineRemaining: '' };
  if (status === '手動対応済み')
    return { level: '対象外', reason: '手動対応済み',                mcfElapsedStr: '', deadlineRemaining: '' };
  if (status === '確認通知済み' && tracking)
    return { level: '対象外', reason: '確認通知済み（Tracking有）',  mcfElapsedStr: '', deadlineRemaining: '' };

  // ── 受注日時のパース ──────────────────────────────────────
  let receivedAt = null;
  try {
    if (receivedRaw instanceof Date) {
      receivedAt = receivedRaw;
    } else if (receivedRaw) {
      receivedAt = new Date(String(receivedRaw));
      if (isNaN(receivedAt.getTime())) {
        receivedAt = Utilities.parseDate(
          String(receivedRaw).trim(), 'Asia/Tokyo', 'yyyy-MM-dd HH:mm'
        );
      }
    }
  } catch (e) {
    receivedAt = null;
  }

  const hoursFromOrder = (receivedAt && !isNaN(receivedAt.getTime()))
    ? (now - receivedAt) / (1000 * 60 * 60)
    : null;

  const mcfElapsedStr = hoursFromOrder !== null
    ? hoursFromOrder.toFixed(1) + 'h'
    : '不明';

  // ── 発送期限のパース ──────────────────────────────────────
  let shipByDate = null;
  try {
    if (shipByRaw instanceof Date) {
      shipByDate = shipByRaw;
    } else if (shipByRaw) {
      shipByDate = new Date(String(shipByRaw));
      if (isNaN(shipByDate.getTime())) shipByDate = null;
    }
  } catch (e) {
    shipByDate = null;
  }

  const hoursToDeadline = (shipByDate && !isNaN(shipByDate.getTime()))
    ? (shipByDate - now) / (1000 * 60 * 60)
    : null;

  const deadlineRemaining = hoursToDeadline !== null
    ? (hoursToDeadline >= 0
        ? hoursToDeadline.toFixed(1) + 'h'
        : '超過 ' + Math.abs(hoursToDeadline).toFixed(1) + 'h')
    : '不明';

  // ── 判定変数 ──────────────────────────────────────────────
  let level  = '正常';
  let reason = '';

  // ──────────────────────────────────────────────────────────
  // ★ 判定A（v2）：MCF未作成 かつ ステータス = 受注済
  //    しきい値変更：6/12/24h → 18/24/36h
  // ──────────────────────────────────────────────────────────
  if (!mcfId && status === '受注済') {
    if (hoursFromOrder !== null) {
      if (hoursFromOrder > ANOMALY_THRESHOLD_.MCF_NOT_CREATED_DANGER) {
        level  = '危険';
        reason = 'MCF未作成・受注から ' + hoursFromOrder.toFixed(1) + 'h経過（36h超）';
      } else if (hoursFromOrder > ANOMALY_THRESHOLD_.MCF_NOT_CREATED_WARNING) {
        level  = '警告';
        reason = 'MCF未作成・受注から ' + hoursFromOrder.toFixed(1) + 'h経過（24h超）';
      } else if (hoursFromOrder > ANOMALY_THRESHOLD_.MCF_NOT_CREATED_CAUTION) {
        level  = '注意';
        reason = 'MCF未作成・受注から ' + hoursFromOrder.toFixed(1) + 'h経過（18h超）';
      }
    }
  }

  // ──────────────────────────────────────────────────────────
  // ★ 判定B（v2）：MCF作成済 かつ Tracking空欄
  //    発送期限残り時間ベースに変更（MCF作成経過時間判定は廃止）
  //    発送期限超過  → 危険
  //    12h以内       → 危険
  //    24h以内       → 警告
  //    72h以内       → 注意
  //    72h超         → 正常（追跡番号発行待ちとして通常扱い）
  // ──────────────────────────────────────────────────────────
  if (mcfId && status === 'MCF作成済' && !tracking) {
    if (hoursToDeadline !== null) {
      let newLevel  = '正常';
      let newReason = '';

      if (hoursToDeadline < 0) {
        // 発送期限超過
        newLevel  = '危険';
        newReason = 'MCF作成済・Tracking未取得（発送期限超過 '
          + Math.abs(hoursToDeadline).toFixed(1) + 'h）';
      } else if (hoursToDeadline <= ANOMALY_THRESHOLD_.TRACKING_DANGER) {
        // 12h以内
        newLevel  = '危険';
        newReason = 'MCF作成済・Tracking未取得（発送期限まで '
          + hoursToDeadline.toFixed(1) + 'h）';
      } else if (hoursToDeadline <= ANOMALY_THRESHOLD_.TRACKING_WARNING) {
        // 24h以内
        newLevel  = '警告';
        newReason = 'MCF作成済・Tracking未取得（発送期限まで '
          + hoursToDeadline.toFixed(1) + 'h）';
      } else if (hoursToDeadline <= ANOMALY_THRESHOLD_.TRACKING_CAUTION) {
        // 72h以内
        newLevel  = '注意';
        newReason = 'MCF作成済・Tracking未取得（発送期限まで '
          + hoursToDeadline.toFixed(1) + 'h）';
      }
      // 72h超 → newLevel = '正常'（通常の追跡番号発行待ち）

      if (ANOMALY_LEVEL_PRIORITY_[newLevel] > ANOMALY_LEVEL_PRIORITY_[level]) {
        level  = newLevel;
        reason = newReason;
      }
    } else {
      // 発送期限が取得できない場合は正常扱い（判定不能）
      Logger.log('     ⚠️ 発送期限が取得できないため判定B をスキップ');
    }
  }

  return {
    level            : level,
    reason           : reason || '異常なし',
    mcfElapsedStr    : mcfElapsedStr,
    deadlineRemaining: deadlineRemaining
  };
}

// ============================================================
// ■ 2. shouldSendEmail_
//       メール通知の要否を判断（レベルが上がった場合のみtrue）
// ============================================================

function shouldSendEmail_(alreadyNotified, newLevel) {
  const currentPriority = ANOMALY_LEVEL_PRIORITY_[alreadyNotified] !== undefined
    ? ANOMALY_LEVEL_PRIORITY_[alreadyNotified]
    : -1;
  const newPriority = ANOMALY_LEVEL_PRIORITY_[newLevel] || 0;
  return newPriority > 0 && newPriority > currentPriority;
}

// ============================================================
// ■ 3. sendAnomalyEmail_
//       異常検知メールを送信する
// ============================================================

function sendAnomalyEmail_(anomalyRows) {
  if (!anomalyRows || anomalyRows.length === 0) return;

  const dangerRows  = anomalyRows.filter(function(r) { return r.level === '危険'; });
  const warningRows = anomalyRows.filter(function(r) { return r.level === '警告'; });
  const cautionRows = anomalyRows.filter(function(r) { return r.level === '注意'; });

  const subject = '【MCF監視】eBay受注異常 '
    + (dangerRows.length  > 0 ? '危険' + dangerRows.length  + '件 ' : '')
    + (warningRows.length > 0 ? '警告' + warningRows.length + '件 ' : '')
    + (cautionRows.length > 0 ? '注意' + cautionRows.length + '件'  : '');

  let body = 'MCF監視システムからの通知です。\n\n';
  body += '実行日時: '
    + Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy/MM/dd HH:mm:ss')
    + '\n\n';

  [
    { label: '🔴 危険', rows: dangerRows  },
    { label: '🟡 警告', rows: warningRows },
    { label: '🔵 注意', rows: cautionRows }
  ].forEach(function(section) {
    if (section.rows.length === 0) return;
    body += '■ ' + section.label + '（' + section.rows.length + '件）\n';
    body += '─────────────────────────────────────\n';
    section.rows.forEach(function(r) {
      body += '行' + r.rowNum + ': ' + r.orderId;
      if (r.sku) body += '  SKU: ' + r.sku;
      body += '\n';
      body += '  理由: ' + r.reason + '\n';
      body += '  MCF経過: ' + r.mcfElapsedStr
        + '  発送期限残り: ' + r.deadlineRemaining + '\n';
      body += '\n';
    });
  });

  body += '\n以上です。eBay受注ログをご確認ください。\n';

  MailApp.sendEmail({
    to     : ANOMALY_NOTIFY_EMAIL_,
    subject: subject,
    body   : body
  });
}

// ============================================================
// ■ 4. logAnomalyMonitorResult_
//       MCF監視ログシートに実行結果を記録する
// ============================================================

function logAnomalyMonitorResult_(ss, checkedCount, notifyCount, errorCount, anomalyCount) {
  let logSheet = ss.getSheetByName(ANOMALY_LOG_SHEET_NAME_);

  if (!logSheet) {
    logSheet = ss.insertSheet(ANOMALY_LOG_SHEET_NAME_);
    logSheet.setTabColor('#ff9800');

    logSheet.appendRow(['実行日時', '対象件数', '異常件数', 'メール通知件数', 'エラー件数']);
    const hdr = logSheet.getRange(1, 1, 1, 5);
    hdr.setBackground('#e65100');
    hdr.setFontColor('#ffffff');
    hdr.setFontWeight('bold');
    logSheet.setFrozenRows(1);
    logSheet.setColumnWidth(1, 160);
    logSheet.setColumnWidth(2, 80);
    logSheet.setColumnWidth(3, 80);
    logSheet.setColumnWidth(4, 110);
    logSheet.setColumnWidth(5, 80);
  }

  const now = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy/MM/dd HH:mm:ss');
  logSheet.appendRow([now, checkedCount, anomalyCount, notifyCount, errorCount]);
}

// ============================================================
// ■ 5. runMcfAnomalyMonitor（メイン関数）
//       eBay受注ログを全件監視判定し、結果を書き込む
// ============================================================

function runMcfAnomalyMonitor() {
  const NOW_JST = new Date();

  Logger.log('');
  Logger.log('━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━');
  Logger.log('【runMcfAnomalyMonitor v2】MCF異常監視');
  Logger.log('  実行日時: ' + Utilities.formatDate(NOW_JST, 'Asia/Tokyo', 'yyyy/MM/dd HH:mm:ss'));
  Logger.log('━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━');
  Logger.log('');

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const eBayLog = ss.getSheetByName(ANOMALY_EBAY_LOG_SHEET_);
  if (!eBayLog) {
    Logger.log('❌ 「' + ANOMALY_EBAY_LOG_SHEET_ + '」シートが見つかりません');
    return;
  }

  Logger.log('[監視列の確認・追加]');
  const colMap = ensureAnomalyColumns_(eBayLog);
  ANOMALY_MONITOR_HEADERS_.forEach(function(name) {
    Logger.log('  ' + name + ' → 列' + colMap[name]);
  });
  Logger.log('');

  const allData = eBayLog.getDataRange().getValues();
  Logger.log('[eBay受注ログ スキャン]');
  Logger.log('  総行数（ヘッダー含む）: ' + allData.length);
  Logger.log('');

  let checkedCount = 0;
  let anomalyCount = 0;
  let notifyCount  = 0;
  let errorCount   = 0;

  const anomalyRowsToNotify = [];

  for (let i = 1; i < allData.length; i++) {
    const row     = allData[i];
    const orderId = String(row[ANOMALY_COL_.ORDER_ID] || '').trim();
    if (!orderId) continue;

    checkedCount++;
    const sku = String(row[ANOMALY_COL_.SKU] || '').trim();

    try {
      const judgment = getMcfAnomalyJudgment_(row);
      const sheetRow = i + 1;

      // ── 監視列に書き込み ──────────────────────────────────
      eBayLog.getRange(sheetRow, colMap['MCF監視判定']    ).setValue(judgment.level);
      eBayLog.getRange(sheetRow, colMap['MCF監視理由']    ).setValue(judgment.reason);
      eBayLog.getRange(sheetRow, colMap['MCF経過時間']    ).setValue(judgment.mcfElapsedStr);
      eBayLog.getRange(sheetRow, colMap['発送期限残り時間']).setValue(judgment.deadlineRemaining);
      eBayLog.getRange(sheetRow, colMap['MCF通知レベル']  ).setValue(judgment.level);

      // ── MCF監視判定セルに色をつける ────────────────────────
      const levelCell = eBayLog.getRange(sheetRow, colMap['MCF監視判定']);
      switch (judgment.level) {
        case '危険' : levelCell.setBackground('#fce8e6').setFontColor('#c5221f'); break;
        case '警告' : levelCell.setBackground('#fef7e0').setFontColor('#e37400'); break;
        case '注意' : levelCell.setBackground('#e8f0fe').setFontColor('#1a73e8'); break;
        case '正常' : levelCell.setBackground('#e6f4ea').setFontColor('#137333'); break;
        case '対象外': levelCell.setBackground('#f1f3f4').setFontColor('#5f6368'); break;
      }

      // ── Logger 出力 ────────────────────────────────────────
      const symbol = { '危険':'🔴', '警告':'🟡', '注意':'🔵', '正常':'✅', '対象外':'⏭ ' };
      Logger.log('  ' + (symbol[judgment.level] || '  ') + ' 行' + sheetRow
        + ' [' + judgment.level + '] ' + orderId);
      if (judgment.level !== '正常' && judgment.level !== '対象外') {
        Logger.log('     ' + judgment.reason);
      }

      // ── 異常カウント ──────────────────────────────────────
      if (['注意', '警告', '危険'].includes(judgment.level)) {
        anomalyCount++;
      }

      // ── メール通知の要否判断 ──────────────────────────────
      if (['注意', '警告', '危険'].includes(judgment.level)) {
        const alreadyNotified = String(
          row[colMap['MCF通知済み'] - 1] || ''
        ).trim();

        if (shouldSendEmail_(alreadyNotified, judgment.level)) {
          const nowStr = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy/MM/dd HH:mm:ss');
          eBayLog.getRange(sheetRow, colMap['MCF通知済み']    ).setValue(judgment.level);
          eBayLog.getRange(sheetRow, colMap['MCF最終通知日時']).setValue(nowStr);

          anomalyRowsToNotify.push({
            rowNum           : sheetRow,
            orderId          : orderId,
            sku              : sku,
            level            : judgment.level,
            reason           : judgment.reason,
            mcfElapsedStr    : judgment.mcfElapsedStr,
            deadlineRemaining: judgment.deadlineRemaining
          });

          Logger.log('     → メール通知対象（前回: "' + alreadyNotified + '" → 今回: "' + judgment.level + '"）');
        } else if (alreadyNotified === judgment.level) {
          Logger.log('     → 通知スキップ（同レベル "' + judgment.level + '" 通知済み）');
        }
      }

    } catch (e) {
      Logger.log('  ❌ 行' + (i + 1) + ' 処理エラー: ' + e.toString());
      errorCount++;
    }
  }

  // ── メール送信 ─────────────────────────────────────────────
  Logger.log('');
  if (anomalyRowsToNotify.length > 0) {
    try {
      sendAnomalyEmail_(anomalyRowsToNotify);
      notifyCount = anomalyRowsToNotify.length;
      Logger.log('✅ メール通知送信完了 → ' + ANOMALY_NOTIFY_EMAIL_);
      Logger.log('   通知件数: ' + notifyCount + ' 件');
    } catch (e) {
      Logger.log('❌ メール送信エラー: ' + e.toString());
      errorCount++;
    }
  } else {
    Logger.log('ℹ️ メール通知対象なし（新規・レベルアップなし）');
  }

  // ── MCF監視ログに記録 ──────────────────────────────────────
  try {
    logAnomalyMonitorResult_(ss, checkedCount, notifyCount, errorCount, anomalyCount);
    Logger.log('✅ MCF監視ログに記録しました');
  } catch (e) {
    Logger.log('⚠️ MCF監視ログ記録エラー: ' + e.toString());
  }

  // ── Loggerサマリー ─────────────────────────────────────────
  Logger.log('');
  Logger.log('━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━');
  Logger.log('【サマリー】');
  Logger.log('  判定対象件数    : ' + checkedCount + ' 件');
  Logger.log('  異常件数（合計）: ' + anomalyCount + ' 件（注意+警告+危険）');
  Logger.log('  メール通知件数  : ' + notifyCount  + ' 件');
  Logger.log('  エラー件数      : ' + errorCount   + ' 件');
  Logger.log('');
  Logger.log('  ★ 既存列（A〜U）は変更しませんでした');
  Logger.log('  ★ MCF作成API・eBay提出APIは呼びませんでした');
  Logger.log('  ★ 既存トリガーは変更しませんでした');
  Logger.log('━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━');
  Logger.log('');
}

// ============================================================
// ■ 判定ロジック補足（v2）
// ============================================================
//
// 【判定A：MCF未作成 & 受注済（v2）】
//   受注日時からの経過時間で判定：
//   18h超 → 注意 / 24h超 → 警告 / 36h超 → 危険
//
// 【判定B：MCF作成済 & Tracking空欄（v2）】
//   発送期限までの残り時間で判定（MCF作成経過時間判定は廃止）：
//   72h超        → 正常（追跡番号発行待ち、通常状態）
//   72h以内      → 注意
//   24h以内      → 警告
//   12h以内      → 危険
//   期限超過     → 危険
//
//   例：09-14755-76592 の期待動作
//     発送期限まで72h超  → 正常
//     発送期限まで72h以内 → 注意
//     発送期限まで24h以内 → 警告
//     発送期限まで12h以内または超過 → 危険
//
//   例：26-14732-00119（MCF未作成・受注済）の期待動作
//     受注後18h未満 → 正常
//     受注後18h超  → 注意
//     受注後24h超  → 警告
//     受注後36h超  → 危険
//
// 【メール重複防止】
//   AA列（MCF通知済み）: 最後にメール送信したレベルを記録。
//   注意→注意 = 再送しない / 注意→警告 = 再送する / 警告→危険 = 再送する
//
// 【発送期限フォーマット】
//   O列は 'Jun 10, 2026' 形式の文字列または Date型に対応。
//   new Date(String(shipByRaw)) でパースする。
//
// 【トリガー設定（このファイルでは設定しない）】
//   手動実行で動作確認後、以下のトリガーを設定することを推奨：
//   - runMcfAnomalyMonitor を 2時間おきに実行
//   - または毎朝 9:00 / 夕方 18:00 など

function appendMcfMonitorLogV2_(ss, targetCount, notifyCount, message) {
  const sheet = ss.getSheetByName(MCF_MONITOR_CONFIG.logSheetName) || ss.insertSheet(MCF_MONITOR_CONFIG.logSheetName);

  if (sheet.getLastRow() === 0) {
    sheet.appendRow(['実行日時', '対象件数', '通知件数', 'メッセージ']);
  }

  sheet.appendRow([
    new Date(),
    targetCount,
    notifyCount,
    message || 'MCF監視完了'
  ]);
}

/**
function dryRunRefreshFbaStockListFromSpApi() {
  Logger.log('=== dryRunRefreshFbaStockListFromSpApi 開始 ===');
  Logger.log('FBA Stock List は更新しません。レポート取得と安全チェックのみ実行します。');

  var ss = SpreadsheetApp.getActiveSpreadsheet();

  // ── レポート取得 ──
  Logger.log('--- SP-API レポート取得 ---');
  var report = fsl_fetchReport_();
  if (!report.ok) {
    Logger.log('❌ レポート取得失敗: ' + report.error);
    fsl_writeExecLog_(ss, 'DRY_RUN', 'FETCH_FAILED', report.error, 0);
    return;
  }
  Logger.log('✅ レポート取得成功: ' + report.rows.length + '行（ヘッダー除く）');

  // ── 安全チェック（全7項目）──
  Logger.log('--- 安全チェック ---');
  var sheet    = ss.getSheetByName(FSL.SHEET_STOCK);
  var prevCount = fsl_getCurrentDataRowCount_(sheet);
  Logger.log('  現在の FBA Stock List データ行数: ' + prevCount);

  var check = fsl_runSafetyChecks_(report, prevCount);
  check.items.forEach(function(item) { Logger.log('  ' + item); });

  if (!check.ok) {
    Logger.log('❌ 安全チェック失敗: ' + check.failReason);
    fsl_writeExecLog_(ss, 'DRY_RUN', 'SAFETY_FAILED', check.failReason, report.rows.length);
    return;
  }
  Logger.log('✅ 安全チェック全通過 (' + check.formattedRows.length + '行)');

  // ── DRY_RUN 結果サマリー ──
  Logger.log('--- DRY_RUN サマリー ---');
  Logger.log('  取得行数: '           + report.rows.length);
  Logger.log('  整形後行数: '          + check.formattedRows.length);
  Logger.log('  前回件数: '            + prevCount);
  Logger.log('  件数変化: '            + prevCount + ' → ' + check.formattedRows.length);
  Logger.log('  先頭3行（SKU確認）:');
  check.formattedRows.slice(0, 3).forEach(function(r, i) {
    Logger.log('    [' + i + '] sku=' + r[0] + '  afn-fulfillable=' + r[10]);
  });
  Logger.log('');
  Logger.log('FBA Stock List は更新していません。');
  Logger.log('本番更新: refreshFbaStockListFromSpApi() を実行してください。');

  fsl_writeExecLog_(ss, 'DRY_RUN', 'OK',
    '安全チェック全通過 取得:' + report.rows.length + '行 整形:' + check.formattedRows.length + '行（更新なし）',
    check.formattedRows.length);

  Logger.log('=== dryRunRefreshFbaStockListFromSpApi 終了 ===');
}

// ═══════════════════════════════════════════════════════════
// 公開関数 2: refreshFbaStockListFromSpApi()
//   安全チェック通過時のみ FBA Stock List を更新する。
// ═══════════════════════════════════════════════════════════

/**
 * SP-API からレポートを取得し、安全チェック通過時のみ
 * FBA Stock List シートを更新する。
 *
 * 更新方法（clearContents 禁止）:
 *   1. 新データを A1:V（ヘッダー行 + データ行 全行）へ setValues
 *   2. setValues 成功後に旧データの余り行だけ clearContent
 *   3. 異常時は既存 FBA Stock List を残す
 */
function refreshFbaStockListFromSpApi() {
  Logger.log('=== refreshFbaStockListFromSpApi 開始 ===');

  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(FSL.SHEET_STOCK);

  if (!sheet) {
    Logger.log('❌ シート「' + FSL.SHEET_STOCK + '」が見つかりません。');
    fsl_writeExecLog_(ss, 'UPDATE', 'ERROR', 'シートが見つかりません: ' + FSL.SHEET_STOCK, 0);
    return;
  }

  // ── レポート取得 ──
  Logger.log('--- SP-API レポート取得 ---');
  var report = fsl_fetchReport_();
  if (!report.ok) {
    Logger.log('❌ レポート取得失敗（既存 FBA Stock List を保持します）: ' + report.error);
    fsl_writeExecLog_(ss, 'UPDATE', 'FETCH_FAILED',
      '既存データ保持。取得失敗: ' + report.error, 0);
    return;
  }
  Logger.log('✅ レポート取得成功: ' + report.rows.length + '行');

  // ── 安全チェック ──
  Logger.log('--- 安全チェック ---');
  var prevCount = fsl_getCurrentDataRowCount_(sheet);
  Logger.log('  現在の FBA Stock List データ行数: ' + prevCount);

  var check = fsl_runSafetyChecks_(report, prevCount);
  check.items.forEach(function(item) { Logger.log('  ' + item); });

  if (!check.ok) {
    Logger.log('❌ 安全チェック失敗（既存 FBA Stock List を保持します）: ' + check.failReason);
    fsl_writeExecLog_(ss, 'UPDATE', 'SAFETY_FAILED',
      '既存データ保持。安全チェック失敗: ' + check.failReason, report.rows.length);
    return;
  }
  Logger.log('✅ 安全チェック全通過: ' + check.formattedRows.length + '行');

  // ── FBA Stock List 更新 ──
  Logger.log('--- FBA Stock List 更新 ---');
  var newRowCount    = check.formattedRows.length;  // データ行数（ヘッダー除く）
  var newTotalRows   = newRowCount + 1;              // ヘッダー行込み
  var oldTotalRows   = prevCount + 1;               // 旧ヘッダー行込み

  // 書き込む2次元配列（ヘッダー行 + データ行）
  var writeData = [FSL.HEADERS].concat(check.formattedRows);

  // ── Step1: 新データを A1:V（新行数分）へ setValues ──
  // setValues 単独で try-catch する。失敗時は既存データを保持して終了。
  try {
    sheet.getRange(1, 1, newTotalRows, FSL.HEADERS.length)
         .setValues(writeData);
    Logger.log('✅ setValues 完了: ' + newTotalRows + '行 × ' + FSL.HEADERS.length + '列');
  } catch (e) {
    Logger.log('❌ setValues 失敗（既存データが保護されています）: ' + e.message);
    fsl_writeExecLog_(ss, 'UPDATE', 'WRITE_ERROR',
      'setValues失敗（既存データ保持）: ' + e.message, 0);
    return;
  }

  // ── Step2: 旧データより新データが少ない場合のみ、余り行をクリア ──
  // clearContent は setValues 成功後のみ実行する。
  // clearContent 単独で try-catch する。
  // 失敗しても新データの書き込み自体は完了しているため、
  // PARTIAL_SUCCESS / WARN として記録する（古いSKUが残存する可能性あり）。
  if (oldTotalRows > newTotalRows) {
    var excessRows = oldTotalRows - newTotalRows;
    var clearStart = newTotalRows + 1;
    try {
      sheet.getRange(clearStart, 1, excessRows, FSL.HEADERS.length)
           .clearContent();
      Logger.log('✅ 余り行クリア: ' + clearStart + '行目〜' + excessRows + '行分');
    } catch (e) {
      // clearContent 失敗: 新データは書き込み済みだが旧余り行が残存
      Logger.log('⚠️  余り行クリア失敗: ' + clearStart + '行目〜' + excessRows + '行分: ' + e.message);
      Logger.log('⚠️  FBA Stock List更新は成功。ただし旧余り行のクリアに失敗。手動確認してください。');
      fsl_writeExecLog_(ss, 'UPDATE', 'WARN',
        'FBA Stock List更新は成功。ただし旧余り行のクリアに失敗。手動確認してください。'
        + ' (clearStart=' + clearStart + ' rows=' + excessRows + ')', newRowCount);
      Logger.log('=== refreshFbaStockListFromSpApi 終了（部分成功）===');
      return;
    }
  } else {
    Logger.log('  余り行なし（旧行数:' + oldTotalRows + ' ≦ 新行数:' + newTotalRows + '）');
  }

  Logger.log('✅ FBA Stock List 更新完了: ' + prevCount + '件 → ' + newRowCount + '件');
  fsl_writeExecLog_(ss, 'UPDATE', 'SUCCESS',
    prevCount + '件 → ' + newRowCount + '件 更新完了', newRowCount);

  Logger.log('=== refreshFbaStockListFromSpApi 終了 ===');
}

// ═══════════════════════════════════════════════════════════
// 公開関数 3: auditFbaStockListRefreshTrigger()
//   トリガー監査（変更なし）
// ═══════════════════════════════════════════════════════════

/**
 * FBA Stock List 更新トリガーの状態を監査してログ出力する。
 * トリガーの作成・削除は行わない。
 */
function auditFbaStockListRefreshTrigger() {
  Logger.log('=== auditFbaStockListRefreshTrigger 開始 ===');

  var TARGET_FUNCS = [
    'refreshFbaStockListFromSpApi',
    'dryRunRefreshFbaStockListFromSpApi',
  ];
  var PROTECTED_FUNCS = [
    'runInventorySyncAllAuto',
    'updateEbayRawFromTradingApi',
    'updateStockSyncFromEbayRaw',
    'checkEbayOrders',
    'runDailySampleListing',
    'endDailySampleListing',
  ];

  var triggers = ScriptApp.getProjectTriggers();
  Logger.log('総トリガー数: ' + triggers.length);
  Logger.log('');

  var counts = { target: 0, protected: 0, other: 0 };
  triggers.forEach(function(t, i) {
    var fn        = t.getHandlerFunction();
    var isTarget  = TARGET_FUNCS.indexOf(fn)   !== -1;
    var isProtect = PROTECTED_FUNCS.indexOf(fn) !== -1;
    var label = isTarget  ? '[FBA更新] '
              : isProtect ? '[保護]    '
              : '[その他]  ';
    if (isTarget)  { counts.target++;    }
    else if (isProtect) { counts.protected++; }
    else           { counts.other++;     }

    Logger.log('  [' + (i+1) + '] ' + label + fn
      + '  EventType:' + t.getEventType()
      + '  ID:' + t.getUniqueId());
  });

  Logger.log('');
  Logger.log('=== 集計 ===');
  Logger.log('  [FBA更新] refreshFbaStockListFromSpApi系: ' + counts.target    + '件（正常: 1件）');
  Logger.log('  [保護] 在庫同期・MCF等: '                  + counts.protected  + '件');
  Logger.log('  [その他]: '                                 + counts.other      + '件');

  // FBA Stock List の現在の状態も出力
  Logger.log('');
  Logger.log('=== FBA Stock List 現在の状態 ===');
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(FSL.SHEET_STOCK);
  if (sheet) {
    var dataCount = fsl_getCurrentDataRowCount_(sheet);
    Logger.log('  データ行数: ' + dataCount + '件');
    Logger.log('  最終更新行: ' + sheet.getLastRow() + '行');
  } else {
    Logger.log('  ❌ シート「' + FSL.SHEET_STOCK + '」が見つかりません。');
  }

  Logger.log('=== auditFbaStockListRefreshTrigger 終了 ===');
}

// ═══════════════════════════════════════════════════════════
// 公開関数 4: setupFbaStockListRefreshTrigger()
//   トリガー設定（return ガード付き。最初は実行不可）
// ═══════════════════════════════════════════════════════════

/**
 * refreshFbaStockListFromSpApi の日次トリガーを設定する。
 *
 * 【このフェーズでは実行しない】
 * dryRunRefreshFbaStockListFromSpApi() で動作確認後に実行する。
 *
 * 設定するトリガー:
 *   refreshFbaStockListFromSpApi  毎日 atHour(3)  JST 3時台（深夜）
 *
 * 保護対象（削除しない）:
 *   runInventorySyncAllAuto および既存の全トリガー
 */
function setupFbaStockListRefreshTrigger() {
  Logger.log('=== setupFbaStockListRefreshTrigger ===');
  Logger.log('【注意】このフェーズでは実行しません。');
  Logger.log('dryRunRefreshFbaStockListFromSpApi() で確認後に実行してください。');
  Logger.log('');
  Logger.log('設定予定: refreshFbaStockListFromSpApi  毎日 atHour(3)  JST 3時台');
  return;  // ← DRY_RUN 確認完了後にこの行をコメントアウトする

  /* ─── 確認完了後に有効になるコード ─── */
  var TARGET_FUNC  = 'refreshFbaStockListFromSpApi';
  var PROTECTED    = ['runInventorySyncAllAuto'];

  // 対象関数の既存トリガーのみ削除（保護対象は触れない）
  ScriptApp.getProjectTriggers().forEach(function(t) {
    if (PROTECTED.indexOf(t.getHandlerFunction()) !== -1) { return; }
    if (t.getHandlerFunction() === TARGET_FUNC) {
      ScriptApp.deleteTrigger(t);
      Logger.log('🗑️  既存トリガー削除: ' + t.getUniqueId());
    }
  });

  ScriptApp.newTrigger(TARGET_FUNC)
    .timeBased().everyDays(1).atHour(3).create();
  Logger.log('✅ トリガー作成: ' + TARGET_FUNC + '  atHour(3) JST 3時台');

  Logger.log('現在のトリガー一覧:');
  ScriptApp.getProjectTriggers().forEach(function(t) {
    Logger.log('  ' + t.getHandlerFunction() + '  ID:' + t.getUniqueId());
  });
}

// ═══════════════════════════════════════════════════════════
// 内部: SP-API レポート取得フロー
// ═══════════════════════════════════════════════════════════

/**
 * GET_FBA_MYI_UNSUPPRESSED_INVENTORY_DATA レポートを取得して
 * 行データの配列として返す。
 *
 * フロー:
 *   1. createReport でレポート作成リクエスト
 *   2. reportId を取得
 *   3. ポーリングで DONE になるまで待機
 *   4. reportDocumentId を取得
 *   5. getReportDocument でダウンロード URL を取得
 *   6. TSV テキストをダウンロードしてパース
 *
 * @returns {{ok, headers, rows, error}}
 *   headers: string[]  TSV の全ヘッダー
 *   rows   : Object[]  ヘッダー名キーの各行オブジェクト
 */
function fsl_fetchReport_() {
  // Access Token 取得
  var token = fsl_getSpApiToken_();
  if (!token) {
    return { ok: false, error: 'SP-API Access Token取得失敗' };
  }
  Logger.log('  SP-API Token: 取得成功（値はログに出力しません）');

  // Step1: レポート作成
  Logger.log('  Step1: createReport');
  var createBody = JSON.stringify({
    reportType   : FSL.REPORT_TYPE,
    marketplaceIds: [FSL.MARKETPLACE_ID],
  });
  var createResp = UrlFetchApp.fetch(
    FSL.SP_API_ENDPOINT + '/reports/2021-06-30/reports', {
      method            : 'post',
      headers           : fsl_buildHeaders_(token),
      payload           : createBody,
      muteHttpExceptions: true,
    }
  );
  var createJson = fsl_parseJson_(createResp.getContentText());
  if (createResp.getResponseCode() !== 202 || !createJson.reportId) {
    return {
      ok   : false,
      error: 'createReport 失敗 HTTP=' + createResp.getResponseCode()
           + ' body=' + createResp.getContentText().substring(0, 200),
    };
  }
  var reportId = createJson.reportId;
  Logger.log('  Step1 完了: reportId=' + reportId);

  // Step2: ポーリングで DONE を待つ
  Logger.log('  Step2: ポーリング開始（最大' + FSL.POLL_MAX_RETRY + '回）');
  var reportDocumentId = '';
  for (var i = 0; i < FSL.POLL_MAX_RETRY; i++) {
    Utilities.sleep(FSL.POLL_INTERVAL_MS);
    var pollResp = UrlFetchApp.fetch(
      FSL.SP_API_ENDPOINT + '/reports/2021-06-30/reports/' + reportId, {
        method            : 'get',
        headers           : fsl_buildHeaders_(token),
        muteHttpExceptions: true,
      }
    );
    var pollJson = fsl_parseJson_(pollResp.getContentText());
    var status   = pollJson.processingStatus || '';
    Logger.log('  ポーリング[' + (i+1) + '/' + FSL.POLL_MAX_RETRY + ']: status=' + status);

    if (status === 'DONE') {
      reportDocumentId = pollJson.reportDocumentId || '';
      break;
    }
    if (status === 'FATAL' || status === 'CANCELLED') {
      return { ok: false, error: 'レポート処理失敗: status=' + status };
    }
  }
  if (!reportDocumentId) {
    return { ok: false, error: 'ポーリングタイムアウト: DONE にならず終了' };
  }
  Logger.log('  Step2 完了: reportDocumentId=' + reportDocumentId);

  // Step3: ドキュメントURL取得
  Logger.log('  Step3: getReportDocument');
  var docResp = UrlFetchApp.fetch(
    FSL.SP_API_ENDPOINT + '/reports/2021-06-30/documents/' + reportDocumentId, {
      method            : 'get',
      headers           : fsl_buildHeaders_(token),
      muteHttpExceptions: true,
    }
  );
  var docJson = fsl_parseJson_(docResp.getContentText());
  if (!docJson.url) {
    return {
      ok   : false,
      error: 'getReportDocument URL取得失敗: ' + docResp.getContentText().substring(0, 200),
    };
  }
  Logger.log('  Step3 完了: document URL取得成功');

  // Step4: TSV ダウンロード
  Logger.log('  Step4: TSVダウンロード');
  var tsvResp = UrlFetchApp.fetch(docJson.url, { muteHttpExceptions: true });
  if (tsvResp.getResponseCode() !== 200) {
    return {
      ok   : false,
      error: 'TSVダウンロード失敗 HTTP=' + tsvResp.getResponseCode(),
    };
  }
  var tsvText = tsvResp.getContentText('UTF-8');
  Logger.log('  Step4 完了: TSV ' + tsvText.length + ' chars');

  // Step5: TSV パース
  return fsl_parseTsv_(tsvText);
}

// ═══════════════════════════════════════════════════════════
// 内部: 安全チェック（全7項目）
// ═══════════════════════════════════════════════════════════

/**
 * 取得レポートデータに対して全7項目の安全チェックを実行する。
 *
 * @param {{headers, rows}} report - fsl_fetchReport_() の戻り値
 * @param {number} prevCount - 現在の FBA Stock List データ行数
 * @returns {{ok, items, failReason, formattedRows}}
 */
function fsl_runSafetyChecks_(report, prevCount) {
  var items      = [];
  var failReason = '';

  // ① レポート取得成功（呼び出し元で確認済み。ここでは再確認）
  if (!report || !report.ok) {
    return { ok: false, items: ['❌ ①レポート取得失敗'], failReason: '①レポート取得失敗', formattedRows: [] };
  }
  items.push('✅ ①レポート取得成功');

  // ② データ行数が 0 件ではない
  if (!report.rows || report.rows.length === 0) {
    items.push('❌ ②データ行数 0 件');
    return { ok: false, items: items, failReason: '②データ行数0件', formattedRows: [] };
  }
  items.push('✅ ②データ行数: ' + report.rows.length + '件');

  // ③ 必須ヘッダー 11 列がすべて存在する
  var missingHeaders = FSL.REQUIRED_HEADERS.filter(function(h) {
    return report.headers.indexOf(h) === -1;
  });
  if (missingHeaders.length > 0) {
    var msg = '③必須ヘッダー不足: ' + missingHeaders.join(', ');
    items.push('❌ ' + msg);
    return { ok: false, items: items, failReason: msg, formattedRows: [] };
  }
  items.push('✅ ③必須ヘッダー ' + FSL.REQUIRED_HEADERS.length + '列確認');

  // ④ sku 列が空ばかりではない
  var nonEmptySkuCount = report.rows.filter(function(r) {
    return String(r['sku'] || '').trim() !== '';
  }).length;
  if (nonEmptySkuCount === 0) {
    items.push('❌ ④sku列がすべて空欄');
    return { ok: false, items: items, failReason: '④sku列が空ばかり', formattedRows: [] };
  }
  items.push('✅ ④sku非空行数: ' + nonEmptySkuCount + '件');

  // ⑤ afn-fulfillable-quantity が数値として読める行が存在する
  var numericQtyCount = report.rows.filter(function(r) {
    var v = r['afn-fulfillable-quantity'];
    return v !== undefined && v !== '' && !isNaN(parseInt(String(v), 10));
  }).length;
  if (numericQtyCount === 0) {
    items.push('❌ ⑤afn-fulfillable-quantityが数値として読める行が0件');
    return { ok: false, items: items, failReason: '⑤afn-fulfillable-quantity数値行0件', formattedRows: [] };
  }
  items.push('✅ ⑤afn-fulfillable-quantity数値行: ' + numericQtyCount + '件');

  // ⑥ 前回件数から 80% 以上減っていない
  if (prevCount > 0) {
    var retentionRate = report.rows.length / prevCount;
    if (retentionRate < FSL.MIN_RETENTION_RATE) {
      var msg6 = '⑥件数が前回比' + Math.round(retentionRate * 100) + '%に激減（前回:'
        + prevCount + '件 → 今回:' + report.rows.length + '件）。更新停止。';
      items.push('❌ ' + msg6);
      return { ok: false, items: items, failReason: msg6, formattedRows: [] };
    }
    items.push('✅ ⑥件数維持率: ' + Math.round(retentionRate * 100) + '%（'
      + prevCount + '→' + report.rows.length + '件）');
  } else {
    items.push('✅ ⑥初回取得のため件数比較スキップ（前回0件）');
  }

  // ⑦ 22列形式への整形
  var formattedRows = [];
  var formatErrors  = 0;
  report.rows.forEach(function(r) {
    var row = FSL.HEADERS.map(function(h) {
      var v = r[h];
      return (v !== undefined && v !== null) ? String(v) : '';
    });
    // 列数が正しいか確認
    if (row.length !== 22) {
      formatErrors++;
    } else {
      formattedRows.push(row);
    }
  });
  if (formatErrors > 0) {
    var msg7 = '⑦整形失敗行数: ' + formatErrors;
    items.push('❌ ' + msg7);
    return { ok: false, items: items, failReason: msg7, formattedRows: [] };
  }
  if (formattedRows.length === 0) {
    items.push('❌ ⑦整形後データが0件');
    return { ok: false, items: items, failReason: '⑦整形後0件', formattedRows: [] };
  }
  items.push('✅ ⑦22列整形成功: ' + formattedRows.length + '行');

  return { ok: true, items: items, failReason: '', formattedRows: formattedRows };
}

// ═══════════════════════════════════════════════════════════
// 内部: TSV パーサー
// ═══════════════════════════════════════════════════════════

/**
 * TSV テキストを行オブジェクトの配列に変換する。
 * 1行目をヘッダーとして使用する。
 *
 * @param {string} tsvText
 * @returns {{ok, headers, rows, error}}
 */
function fsl_parseTsv_(tsvText) {
  try {
    var lines = tsvText.replace(/\r\n/g, '\n').replace(/\r/g, '\n').split('\n');
    // 末尾の空行を除去
    while (lines.length > 0 && lines[lines.length - 1].trim() === '') {
      lines.pop();
    }
    if (lines.length < 1) {
      return { ok: false, error: 'TSVが空です', headers: [], rows: [] };
    }

    var headers = lines[0].split('\t').map(function(h) { return h.trim(); });
    var rows    = [];

    for (var i = 1; i < lines.length; i++) {
      var cells = lines[i].split('\t');
      var rowObj = {};
      headers.forEach(function(h, idx) {
        rowObj[h] = idx < cells.length ? cells[idx].trim() : '';
      });
      rows.push(rowObj);
    }

    Logger.log('  TSVパース: ヘッダー' + headers.length + '列 データ' + rows.length + '行');
    return { ok: true, headers: headers, rows: rows, error: null };

  } catch (e) {
    return { ok: false, error: 'TSVパース例外: ' + e.message, headers: [], rows: [] };
  }
}

// ═══════════════════════════════════════════════════════════
// 内部: FBA Stock List の現在のデータ行数を取得
// ═══════════════════════════════════════════════════════════

/**
 * FBA Stock List シートの現在のデータ行数（ヘッダー除く）を返す。
 * シートが存在しない場合は 0 を返す。
 */
function fsl_getCurrentDataRowCount_(sheet) {
  if (!sheet) { return 0; }
  var lastRow = sheet.getLastRow();
  if (lastRow <= 1) { return 0; }  // ヘッダーのみ
  return lastRow - 1;  // ヘッダー1行を除いたデータ行数
}

// ═══════════════════════════════════════════════════════════
// 内部: SP-API 認証
// ═══════════════════════════════════════════════════════════

/**
 * SP-API Access Token を取得する。
 * 既存の認証関数（getSpApiAccessToken_ / getSpApiAccessToken）を優先使用する。
 * 見つからない場合は Script Properties から内蔵フローを実行する。
 * ★ Token 値はログに出力しない。
 */
function fsl_getSpApiToken_() {
  // 既存認証関数を優先使用
  var candidates = ['getSpApiAccessToken_', 'getSpApiAccessToken', 'getSpApiToken_'];
  for (var i = 0; i < candidates.length; i++) {
    var fnName = candidates[i];
    try {
      // GASの eval でグローバル関数を検索
      var fn = eval('typeof ' + fnName);
      if (fn === 'function') {
        Logger.log('  既存認証関数を使用: ' + fnName + '()');
        var token = eval(fnName + '()');
        if (token) { return token; }
        Logger.log('  ⚠️ ' + fnName + '() が null を返しました。内蔵フローへフォールバック。');
      }
    } catch (e) { /* 存在しない場合は次へ */ }
  }

  // 内蔵フロー
  Logger.log('  内蔵LWAフローを使用します。');
  return fsl_fetchSpApiTokenInternal_();
}

/**
 * LWA で SP-API Access Token を取得する（内蔵フロー）。
 * Script Properties の値は名前のみログに出力し、値は出力しない。
 */
function fsl_fetchSpApiTokenInternal_() {
  var props = PropertiesService.getScriptProperties();

  var clientId     = sampleCsvProps_().getProperty(FSL.PROP_CLIENT_ID)     || sampleCsvProps_().getProperty(FSL.PROP_CLIENT_ID_ALT);
  var clientSecret = sampleCsvProps_().getProperty(FSL.PROP_CLIENT_SECRET) || sampleCsvProps_().getProperty(FSL.PROP_CLIENT_SECRET_ALT);
  var refreshToken = sampleCsvProps_().getProperty(FSL.PROP_REFRESH_TOKEN) || sampleCsvProps_().getProperty(FSL.PROP_REFRESH_TOKEN_ALT);

  // 設定状況のみログ出力（値は出力しない）
  Logger.log('  Script Properties 確認:');
  Logger.log('    ' + FSL.PROP_CLIENT_ID     + ': ' + (clientId     ? '設定あり' : '❌ 未設定'));
  Logger.log('    ' + FSL.PROP_CLIENT_SECRET + ': ' + (clientSecret ? '設定あり' : '❌ 未設定'));
  Logger.log('    ' + FSL.PROP_REFRESH_TOKEN + ': ' + (refreshToken ? '設定あり' : '❌ 未設定'));

  if (!clientId || !clientSecret || !refreshToken) {
    Logger.log('  ❌ 認証情報が不足しています。');
    return null;
  }

  try {
    var resp = UrlFetchApp.fetch('https://api.amazon.com/auth/o2/token', {
      method  : 'post',
      headers : { 'Content-Type': 'application/x-www-form-urlencoded' },
      payload : {
        grant_type   : 'refresh_token',
        refresh_token: refreshToken,
        client_id    : clientId,
        client_secret: clientSecret,
      },
      muteHttpExceptions: true,
    });
    var json = fsl_parseJson_(resp.getContentText());
    if (!json.access_token) {
      Logger.log('  ❌ Token取得失敗: ' + resp.getResponseCode());
      return null;
    }
    return json.access_token;
  } catch (e) {
    Logger.log('  ❌ Token取得例外: ' + e.message);
    return null;
  }
}

/**
 * SP-API リクエスト用ヘッダーを生成する。
 */
function fsl_buildHeaders_(token) {
  return {
    'x-amz-access-token': token,
    'Content-Type'      : 'application/json',
    'Accept'            : 'application/json',
  };
}

// ═══════════════════════════════════════════════════════════
// 内部: 実行ログ
// ═══════════════════════════════════════════════════════════

/**
 * 実行ログシートに記録する。
 * 列構成: [0]日時 [1]関数名 [2]モード [3]ステータス [4]詳細 [5]件数
 */
function fsl_writeExecLog_(ss, mode, status, detail, rowCount) {
  var sheet = ss.getSheetByName(FSL.SHEET_EXEC_LOG);
  if (!sheet) {
    Logger.log('⚠️  実行ログシート「' + FSL.SHEET_EXEC_LOG + '」が見つかりません。');
    return;
  }
  sheet.appendRow([
    fsl_nowJst_(),
    'FbaStockListRefresh',
    mode,
    status,
    String(detail).substring(0, 300),
    rowCount,
  ]);
}

// ═══════════════════════════════════════════════════════════
// 内部: ユーティリティ（fsl_ 接頭辞）
// ═══════════════════════════════════════════════════════════

function fsl_parseJson_(text) {
  try { return JSON.parse(text); }
  catch (e) { return {}; }
}

function fsl_nowJst_() {
  return Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy/MM/dd HH:mm:ss');
}

/**
 * FbaStockListInventoryUpdate.gs
 * ─────────────────────────────────────────────────────────────────
 * 目的   : FBA Stock List に既に存在するSKUについて、
 *          SP-API Inventory API から在庫数を取得し、
 *          afn-fulfillable-quantity 等の在庫列のみ更新する。
 *
 * ─── 流用する既存関数（SpApiInventory.gs に存在）───
 *   getSpApiAccessToken_()          SP-API認証トークン取得
 *   _fetchInventoryForSkus_()       一括取得（sellerSkus複数指定）
 *   _fetchInventoryForSingleSku_()  単体補完取得
 *   chunkArray_()                   配列をチャンク分割
 *   _normSku_()                     SKU正規化（マッチング精度向上）
 *
 * ─── 対象シート（この2シートのみ使用）───
 *   FBA Stock List  ← 在庫列のみ更新
 *   実行ログ         ← 結果記録
 *
 * ─── FBA Stock List の列とSP-APIレスポンスの対応 ───
 *   【本番更新する列】                SP-APIレスポンスフィールド
 *   ─────────────────────────    ──────────────────────────────────
 *   afn-fulfillable-quantity   ← inventoryDetails.fulfillableQuantity  ★必須
 *   afn-reserved-quantity      ← reservedQuantity.totalReservedQuantity
 *   afn-total-quantity         ← totalQuantity
 *
 *   【更新しない列（既存値を維持）】
 *   afn-warehouse-quantity     → 不更新。fulfillableQuantity と同値とは限らないため
 *   afn-unsellable-quantity    → 不更新。SP-API未提供のため 0 で上書きしない
 *
 * ─── 仕様 ───
 *   ・ヘッダー行から列インデックスを動的に特定（列番号固定禁止）
 *   ・product-name / your-price / asin / fnsku は上書きしない
 *   ・FBA Stock List を clearContents しない
 *   ・既存行数を変えない
 *   ・取得できたSKUのみ更新（未取得SKUは既存値を維持）
 *   ・DRY_RUNでは更新予定件数・未取得件数をログ出力のみ
 *   ・本番関数のみ setValues する
 *
 * ─── 安全条件 ───
 *   ・LockService で二重実行防止
 *   ・対象SKU数0なら停止
 *   ・必須ヘッダー（sku / afn-fulfillable-quantity）がなければ停止
 *   ・SP-API取得件数0なら本番更新しない
 *   ・取得成功率が対象SKU数100件以上で5%未満（5件未満）の場合は停止
 *
 * ─── 禁止事項 ───
 *   ✗ updateFbaInventoryFromSpApi の変更
 *   ✗ runInventorySyncAllAuto の変更
 *   ✗ 既存トリガーの削除
 *   ✗ clearContents
 *   ✗ 行の追加・削除
 *   ✗ sku / product-name / your-price / asin / fnsku 列の上書き
 * ─────────────────────────────────────────────────────────────────
 */

// ═══════════════════════════════════════════════════════════
// 定数
// ═══════════════════════════════════════════════════════════

var FSLI = {
  // ── 対象シート ──
  SHEET_STOCK   : 'FBA Stock List',
  SHEET_EXEC_LOG: '実行ログ',

  // ── 必須ヘッダー ──
  HDR_SKU          : 'sku',
  HDR_FULFILLABLE  : 'afn-fulfillable-quantity',   // ★最重要・必須

  // ── 本番更新対象ヘッダー（3列のみ）──
  UPDATE_HEADERS: [
    'afn-fulfillable-quantity',   // ← SP-API: fulfillableQuantity  ★必須
    'afn-reserved-quantity',      // ← SP-API: reservedQuantity.totalReservedQuantity
    'afn-total-quantity',         // ← SP-API: totalQuantity
  ],

  // ── 更新しない列（シートに存在しても既存値を維持）──
  // DRY_RUNログで「更新対象外・既存値維持」と表示する
  SKIP_HEADERS: [
    'afn-warehouse-quantity',     // fulfillableQuantityと同値とは限らないため不更新
    'afn-unsellable-quantity',    // SP-API未提供のため0で上書きしない
  ],

  // ── 上書き禁止ヘッダー ──
  READONLY_HEADERS: ['sku', 'fnsku', 'asin', 'product-name', 'condition', 'your-price'],

  // ── SP-API 一括取得のチャンクサイズ ──
  SKU_CHUNK_SIZE: 50,

  // ── チャンク間スリープ (ms) ──
  SLEEP_BETWEEN_MS: 500,

  // ── 安全条件: 対象SKU数がこの値以上のとき、取得成功率チェックを行う ──
  MIN_SKU_FOR_RATE_CHECK: 100,

  // ── 安全条件: 取得成功率の最低閾値 (5%) ──
  MIN_SUCCESS_RATE: 0.05,
};

/**
 * updateAllSkusInventoryFromSpApi()  修正版
 * ===============================================================
 * 在庫同期管理シートの全SKUのAmazon在庫をSP-APIから取得してD/H/I列を更新する。
 * eBay出品の有無に関係なく全SKUを対象とする点のみ updateFbaInventoryFromSpApi() と異なる。
 *
 * 【既存関数をそのまま利用】
 *   _fetchInventoryForSkus_()      … 一括取得（summaries配列を返す）
 *   _fetchInventoryForSingleSku_() … 単体補完
 *   getSpApiAccessToken_()         … LWAトークン取得
 *   _normSku_()                    … SKU正規化
 *   chunkArray_()                  … 配列分割
 *   定数: SHEET_STOCK_SYNC / SKU_CHUNK_SIZE / SLEEP_BETWEEN_MS /
 *         SLEEP_SINGLE_MS / SKU_EXCLUDE / TEST_SKU_TOKENS
 *
 * 【修正ポイント（旧コードからの変更点）】
 *   1. _fetchInventoryForSkus_() の戻り値は summaries配列。
 *      旧コードはマップを期待していたため在庫が取れなかった。→ 正しく処理する
 *   2. SKU照合に _normSku_() を使用（大文字小文字・スペースのズレを吸収）
 *   3. SP-API未発見時は 0 でなく '' を書き込む（取得失敗と在庫0を区別）
 * ===============================================================
 */
function updateAllSkusInventoryFromSpApi() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_STOCK_SYNC); // 既存定数を使用

  if (!sheet) {
    Logger.log('❌ [在庫同期管理] シートが見つかりません。処理を中断します。');
    return;
  }

  var lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    Logger.log('ℹ️ [在庫同期管理] データ行（2行目以降）が存在しません。');
    return;
  }

  var numRows = lastRow - 1;

  // ── STEP1：A列からSKUを読み込み、rowSkus と uniqueSkus を構築 ──
  var skuValues = sheet.getRange(2, 1, numRows, 1).getValues();

  var rowSkus = [];       // { row0, norm, excluded } 全行分
  var uniqueSkus = {};    // norm → raw（SP-API問い合わせ用）
  var excludedCount = 0;

  for (var i = 0; i < skuValues.length; i++) {
    var raw  = skuValues[i][0];
    var norm = _normSku_(raw); // 既存の正規化関数を使用
    if (!norm) continue;
    if (SKU_EXCLUDE.indexOf(norm.toUpperCase()) !== -1) continue; // 既存定数を使用

    if (_isTestSku_(norm)) { // 既存のテストSKU判定関数を使用
      rowSkus.push({ row0: i, norm: norm, excluded: true });
      excludedCount++;
      continue;
    }

    rowSkus.push({ row0: i, norm: norm, excluded: false });
    uniqueSkus[norm] = raw;
  }

  var skus = Object.keys(uniqueSkus);
  Logger.log('🚀 全件在庫更新開始 対象SKU=' + skus.length + '件 監視除外=' + excludedCount + '件');

  if (skus.length === 0 && excludedCount === 0) {
    Logger.log('ℹ️ 有効なSKUがありません。処理を終了します。');
    return;
  }

  // ── STEP2：SP-API 一括取得 ──
  var fbaMap    = {}; // norm → fulfillableQuantity
  var fbaSource = {}; // norm → 'batch' | 'single'
  var batchFoundCount = 0;
  var missingSkus = [];
  var singleSuccess = 0, singleFail = 0;

  if (skus.length > 0) {
    var token = getSpApiAccessToken_(); // 既存のトークン取得関数

    var chunks = chunkArray_(skus, SKU_CHUNK_SIZE); // 既存の配列分割関数
    for (var c = 0; c < chunks.length; c++) {
      var chunk = chunks[c];

      // ★修正ポイント①：_fetchInventoryForSkus_ は summaries配列を返す
      //   旧コード: var fetchedChunkMap = _fetchInventoryForSkus_(chunk, token);
      //             inventoryMap[skuKey] = fetchedChunkMap[skuKey]; ← マップ扱いで取れない
      var summaries = _fetchInventoryForSkus_(chunk, token);

      Logger.log('[CHUNK ' + (c + 1) + '/' + chunks.length + '] '
        + '送信=' + chunk.length + '件 / レスポンス=' + summaries.length + '件');

      // summaries配列を正しく処理してfbaMapへ格納
      for (var s = 0; s < summaries.length; s++) {
        var sum  = summaries[s];
        // ★修正ポイント②：_normSku_() で正規化してから照合
        var nsku = _normSku_(sum.sellerSku);
        if (!nsku) continue;
        var q = (sum.inventoryDetails || {}).fulfillableQuantity;
        if (typeof q === 'number') {
          fbaMap[nsku]    = q;
          fbaSource[nsku] = 'batch';
        }
      }

      if (c < chunks.length - 1) Utilities.sleep(SLEEP_BETWEEN_MS);
    }
    batchFoundCount = Object.keys(fbaMap).length;

    // ── STEP3：一括未取得SKUを抽出 ──
    for (var m = 0; m < skus.length; m++) {
      if (!fbaMap.hasOwnProperty(skus[m])) missingSkus.push(skus[m]);
    }
    Logger.log('一括取得: ' + batchFoundCount + '件 / 一括未発見: ' + missingSkus.length + '件');

    // ── STEP4：単体補完 ──
    for (var k = 0; k < missingSkus.length; k++) {
      var msku = missingSkus[k];
      var got  = _fetchInventoryForSingleSku_(msku, token); // 既存の単体補完関数
      if (got.found && typeof got.qty === 'number') {
        fbaMap[msku]    = got.qty;
        fbaSource[msku] = 'single';
        singleSuccess++;
        Logger.log('  [単体補完OK] ' + msku + ' qty=' + got.qty);
      } else {
        singleFail++;
        Logger.log('  [単体補完NG] ' + msku + (got.note ? ' (' + got.note + ')' : ''));
      }
      if (k < missingSkus.length - 1) Utilities.sleep(SLEEP_SINGLE_MS);
    }
  } else {
    Logger.log('実在SKU=0件のためSP-API呼び出しをスキップ（監視除外のみ）');
  }

  // ── STEP5：シート反映 ──
  var now   = new Date();
  var dCol  = sheet.getRange(2, 4, numRows, 1).getValues(); // D列: Amazon在庫
  var hiCol = sheet.getRange(2, 8, numRows, 2).getValues(); // H列:日時 I列:確認結果
  var excludedWritten = 0;

  for (var r = 0; r < rowSkus.length; r++) {
    var idx  = rowSkus[r].row0;
    var norm = rowSkus[r].norm;

    // 監視除外SKU：D列は更新せずI列のみ
    if (rowSkus[r].excluded) {
      hiCol[idx][1] = '監視除外';
      excludedWritten++;
      continue;
    }

    if (fbaMap.hasOwnProperty(norm)) {
      dCol[idx][0]  = fbaMap[norm];
      hiCol[idx][0] = now;
      hiCol[idx][1] = (fbaSource[norm] === 'single')
        ? 'SP-API取得済（単体補完）'
        : 'SP-API取得済(全件一括)';
    } else {
      // ★修正ポイント③：未発見時は '' を入れる（0にすると在庫0と区別不可）
      dCol[idx][0]  = '';
      hiCol[idx][0] = now;
      hiCol[idx][1] = 'SP-API未発見';
    }
  }

  // 一括書き込み
  sheet.getRange(2, 4, numRows, 1).setValues(dCol);
  sheet.getRange(2, 8, numRows, 2).setValues(hiCol);

  // ── STEP6：完了ログ ──
  Logger.log('===== 全件在庫更新 完了 =====');
  Logger.log('一括取得        : ' + batchFoundCount + '件');
  Logger.log('一括未発見      : ' + missingSkus.length + '件');
  Logger.log('単体補完成功    : ' + singleSuccess + '件');
  Logger.log('単体補完失敗    : ' + singleFail + '件');
  Logger.log('監視除外        : ' + excludedWritten + '件');
  Logger.log('最終取得済(合計): ' + (batchFoundCount + singleSuccess) + ' / ' + skus.length + '件');
  Logger.log('最終未発見      : ' + (missingSkus.length - singleSuccess) + '件');
}

/**

/**
 * --- Unified eBay listing module ---
 * Old eBay listing blocks were removed; this module is appended once.
 */
/**
 * eBayAutoListing_Unified.gs
 *
 * eBay listing code unified into one production path.
 *
 * Public entry points:
 * - enableEbayAutoListingEvery15Minutes()
 * - disableEbayAutoListing()
 * - auditEbayAutoListingStatus()
 * - runEbayScheduledListingPipeline()
 * - publishSampleCsvListingNow_Integrated()
 * - endSampleCsvListingBySavedItemId_Integrated()
 * - auditSampleCsvListingState_Integrated()
 *
 * Removed from this unified module:
 * - previewEbayListingQueue / verifyApprovedEbayListing / publishApprovedEbayListing
 * - runEbayManualListing
 * - runEbayDailyListingFromCsv
 * - Amazon-to-eBay candidate listing functions
 */

var EAL_FBA = {
  SHEET_NAME: 'FBA Stock List',
  HEADER_SKU: 'sku',
  HEADER_STOCK: 'afn-fulfillable-quantity',
  HEADER_SEARCH_ROWS: 10
};

var EAL = {
  SHEET_QUEUE: 'eBay出品キュー',
  SHEET_RAW: 'eBayRaw',
  SHEET_LIST: 'eBay出品リスト',
  SHEET_LOG: 'eBay_LISTING_LOG',
  SHEET_EXEC_LOG: '実行ログ',
  DATA_START_ROW: 5,

  COL: {
    TASK_ID: 1,
    HUMAN_APPROVAL: 2,
    STATUS: 3,
    SKU: 4,
    ASIN: 5,
    PRODUCT_NAME: 6,
    EBAY_TITLE: 7,
    CATEGORY_ID: 8,
    CATEGORY_NAME: 9,
    CONDITION_ID: 10,
    QUANTITY: 11,
    PRICE_USD: 12,
    PICTURE_URL: 13,
    SHIPPING_ID: 14,
    RETURN_ID: 15,
    PAYMENT_ID: 16,
    CITY: 17,
    STATE: 18,
    COUNTRY: 19,
    DESCRIPTION: 20,
    PREVIEW_RESULT: 21,
    VERIFY_STATUS: 22,
    VERIFY_ACK: 23,
    VERIFY_MESSAGE: 24,
    ITEM_ID: 25,
    LISTING_URL: 26,
    ERROR_CODE: 27,
    ERROR_DETAIL: 28,
    CREATED_DATE: 29,
    UPDATED_DATE: 30,
    NOTES: 31,
    PUBLISH_DATETIME: 32,
    END_TYPE: 33,
    END_REASON: 34,
    END_STATUS: 35,
    END_DATETIME: 36,
    END_ACK: 37,
    END_MESSAGE: 38,
    LISTING_MODE: 39,
    SCHEDULED_DATETIME: 40,
    MANUAL_STATUS: 41,
    EXEC_SOURCE: 42,
    EXEC_RESULT: 43
  },

  STATUS: {
    AUTO_WAIT: '自動出品待ち',
    VERIFY_SUCCESS: 'Verify成功',
    VERIFY_FAILED: 'Verify失敗',
    PUBLISHED: '出品済み',
    ERROR: 'エラー',
    DRY_RUN: 'DRY_RUN済み'
  },

  LISTING_MODE: {
    AUTO: '自動',
    STOP: '停止'
  },

  EXEC_SOURCE: {
    SCHEDULED: '自動トリガー'
  },

  EXEC_RESULT: {
    SUCCESS: 'SUCCESS',
    VERIFY_FAILED: 'VERIFY_FAILED',
    ADD_FAILED: 'ADD_FAILED',
    ERROR: 'ERROR',
    DRY_RUN: 'DRY_RUN'
  },

  TRADING_ENDPOINT: 'https://api.ebay.com/ws/api.dll',
  COMPATIBILITY_LEVEL: '1225',
  SITE_ID: '0',
  PROP_ENABLED: 'EBAY_AUTO_LISTING_ENABLED'
};

var EAT = {
  FUNC_SCHEDULED: 'runEbayScheduledListingPipeline',
  PROP_LISTING_ENABLED: 'EBAY_AUTO_LISTING_ENABLED',
  DELETE_TARGETS: [
    'runAutoListing',
    'checkEmailReplies',
    'runEbayAutoListingPipeline',
    'runEbayScheduledListingPipeline'
  ],
  PROTECTED_FUNCS: [
    'runEbayAutoEndPipeline',
    'runInventorySyncAllAuto',
    'updateEbayRawFromTradingApi',
    'updateStockSyncFromEbayRaw',
    'checkEbayOrders'
  ]
};

var SAMPLE_LISTING = {
  FOLDER_ID: '1q3nMR-tUGXz9mEG3qyEln1uwi9gipWjb',
  CSV_FILE_NAME: 'sample_listing.csv',
  IMAGE_FILE_NAME: 'sample_image.jpg',
  SKU_PREFIX: 'SAMPLE-CSV-',
  PROP_ITEM_ID: 'SAMPLE_CSV_ITEM_ID',
  PROP_LISTED_AT: 'SAMPLE_CSV_LISTED_AT',
  PROP_TASK_ID: 'SAMPLE_CSV_TASK_ID',
  PROP_QUEUE_ROW: 'SAMPLE_CSV_QUEUE_ROW_INDEX',
  PROP_END_TRIGGER_ID: 'SAMPLE_CSV_END_TRIGGER_ID',
  END_HANDLER: 'endSampleCsvListingBySavedItemId_Integrated',
  END_AFTER_MS: 60 * 60 * 1000,
  DEFAULT_TITLE: 'Canon PF-03 Printhead for imagePROGRAF Printers',
  DEFAULT_CATEGORY_ID: 47078,
  DEFAULT_CONDITION_ID: 3000,
  DEFAULT_PRICE_USD: 100,
  DEFAULT_SHIPPING_ID: 285608251015,
  DEFAULT_RETURN_ID: 142913939015,
  DEFAULT_PAYMENT_ID: 142913940015,
  DEFAULT_CITY: 'Kashiwa',
  DEFAULT_COUNTRY: 'JP'
};

function enableEbayAutoListingEvery15Minutes() {
  var props = PropertiesService.getScriptProperties();
  props.setProperty(EAT.PROP_LISTING_ENABLED, 'true');
  eat_deleteTargetTriggers_(EAT.DELETE_TARGETS);

  ScriptApp.newTrigger(EAT.FUNC_SCHEDULED)
    .timeBased()
    .everyMinutes(15)
    .create();

  Logger.log('eBay auto listing enabled.');
  Logger.log(EAT.PROP_LISTING_ENABLED + '=true');
  Logger.log('Trigger: ' + EAT.FUNC_SCHEDULED + ' every 15 minutes');
}

function disableEbayAutoListing() {
  PropertiesService.getScriptProperties().setProperty(EAT.PROP_LISTING_ENABLED, 'false');
  var deleted = eat_deleteTargetTriggers_(EAT.DELETE_TARGETS);
  Logger.log('eBay auto listing disabled.');
  Logger.log(EAT.PROP_LISTING_ENABLED + '=false');
  Logger.log('Deleted listing triggers: ' + deleted);
}

function auditEbayAutoListingStatus() {
  var props = PropertiesService.getScriptProperties();
  var enabled = String(props.getProperty(EAT.PROP_LISTING_ENABLED) || '').toLowerCase() === 'true';
  var triggers = ScriptApp.getProjectTriggers().filter(function(trigger) {
    return trigger.getHandlerFunction() === EAT.FUNC_SCHEDULED;
  });

  Logger.log('=== eBay auto listing status ===');
  Logger.log(EAT.PROP_LISTING_ENABLED + ': ' + (enabled ? 'true / production' : 'false / dry-run'));
  Logger.log('Listing trigger count: ' + triggers.length);
  triggers.forEach(function(trigger, index) {
    Logger.log('[' + (index + 1) + '] ' + trigger.getEventType() + ' / ' + trigger.getUniqueId());
  });

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(EAL.SHEET_QUEUE);
  if (!sheet) {
    Logger.log('Missing sheet: ' + EAL.SHEET_QUEUE);
    return;
  }

  var now = new Date();
  var targets = eal_findScheduledTargets_(sheet, now);
  Logger.log('Ready target rows: ' + targets.length);
  targets.slice(0, 10).forEach(function(target) {
    Logger.log('Row ' + target.rowNum + ' SKU=' + String(target.data[EAL.COL.SKU - 1] || ''));
  });
}

function runEbayScheduledListingPipeline() {
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(10000)) {
    Logger.log('Skipped because another listing run is active.');
    return;
  }

  try {
    var isDryRun = eal_isDryRun_();
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(EAL.SHEET_QUEUE);
    if (!sheet) {
      Logger.log('Missing sheet: ' + EAL.SHEET_QUEUE);
      return;
    }

    Logger.log('=== runEbayScheduledListingPipeline start ===');
    Logger.log('Mode: ' + (isDryRun ? 'DRY_RUN' : 'PRODUCTION'));

    var targets = eal_findScheduledTargets_(sheet, new Date());
    if (targets.length === 0) {
      Logger.log('No listing target.');
      eal_writeExecLog_(ss, EAL.EXEC_SOURCE.SCHEDULED, 'SKIP', 'No target rows', isDryRun);
      return;
    }

    eal_executeListing_(targets[0], EAL.EXEC_SOURCE.SCHEDULED, sheet, ss, isDryRun);
  } finally {
    lock.releaseLock();
  }

  Logger.log('=== runEbayScheduledListingPipeline end ===');
}

function publishSampleCsvListingNow_Integrated() {
  var props = PropertiesService.getScriptProperties();
  var logs = [];

  if (String(props.getProperty(EAT.PROP_LISTING_ENABLED) || '').toLowerCase() !== 'true') {
    logs.push('EBAY_AUTO_LISTING_ENABLED is not true. Sample listing stopped.');
    Logger.log(logs.join('\n'));
    return { status: 'disabled', logs: logs };
  }

  var existingItemId = props.getProperty(SAMPLE_LISTING.PROP_ITEM_ID);
  if (existingItemId) {
    logs.push('Sample listing already exists. ItemID=' + existingItemId);
    Logger.log(logs.join('\n'));
    return { status: 'already_listed', itemId: existingItemId, logs: logs };
  }

  var csvFile = sample_findFileInFolder_(SAMPLE_LISTING.FOLDER_ID, SAMPLE_LISTING.CSV_FILE_NAME);
  if (!csvFile) throw new Error('sample_listing.csv not found');

  var imageFile = sample_findFileInFolder_(SAMPLE_LISTING.FOLDER_ID, SAMPLE_LISTING.IMAGE_FILE_NAME);
  if (!imageFile) throw new Error('sample_image.jpg not found');

  var csvData = sample_readFirstCsvData_(csvFile);
  var rowIndex = sample_importCsvToQueue_(csvData, imageFile);
  props.setProperty(SAMPLE_LISTING.PROP_QUEUE_ROW, String(rowIndex));
  logs.push('Sample row added to queue. Row=' + rowIndex);

  runEbayScheduledListingPipeline();

  var itemId = sample_getItemIdFromQueueRow_(rowIndex);
  if (!itemId) {
    logs.push('Sample listing did not publish. Check queue row ' + rowIndex);
    Logger.log(logs.join('\n'));
    return { status: 'not_published', queueRowIndex: rowIndex, logs: logs };
  }

  var taskId = String(SpreadsheetApp.getActiveSpreadsheet()
    .getSheetByName(EAL.SHEET_QUEUE)
    .getRange(rowIndex, EAL.COL.TASK_ID)
    .getValue() || '');
  var listedAt = eal_nowJst_();
  props.setProperty(SAMPLE_LISTING.PROP_ITEM_ID, itemId);
  props.setProperty(SAMPLE_LISTING.PROP_TASK_ID, taskId);
  props.setProperty(SAMPLE_LISTING.PROP_LISTED_AT, listedAt);

  var triggerId = sample_scheduleEndAfterOneHour_();
  logs.push('Sample listing published. ItemID=' + itemId);
  logs.push('End trigger created. TriggerID=' + triggerId);
  Logger.log(logs.join('\n'));

  return {
    status: 'published',
    itemId: itemId,
    queueRowIndex: rowIndex,
    listedAt: listedAt,
    endTriggerId: triggerId,
    logs: logs
  };
}

function endSampleCsvListingBySavedItemId_Integrated() {
  var props = PropertiesService.getScriptProperties();
  var itemId = props.getProperty(SAMPLE_LISTING.PROP_ITEM_ID) || sample_recoverLatestSampleItemId_();
  if (!itemId) {
    Logger.log('No sample ItemID to end.');
    return { status: 'skip', message: 'No sample ItemID' };
  }

  var isDryRun = eal_isDryRun_();
  if (isDryRun) {
    Logger.log('DRY_RUN: sample EndFixedPriceItem skipped. ItemID=' + itemId);
    return { status: 'dry_run', itemId: itemId };
  }

  var accessToken = eal_getAccessToken_();
  if (!accessToken) throw new Error('Access token fetch failed');

  var xml = sample_buildEndFixedPriceItemXml_(itemId);
  var response = eal_callTradingApi_('EndFixedPriceItem', xml, accessToken);
  var ack = eal_extractXml_(response, 'Ack');
  var message = eal_extractXml_(response, 'ShortMessage') || '';
  var errorCode = eal_extractXml_(response, 'ErrorCode') || '';

  if (ack === 'Success' || ack === 'Warning') {
    sample_markEndedInQueue_(itemId, ack, message);
    sample_clearProperties_();
    sample_deleteEndTrigger_();
    Logger.log('Sample listing ended. ItemID=' + itemId);
    return { status: 'ended', itemId: itemId, ack: ack };
  }

  Logger.log('Sample listing end failed. ItemID=' + itemId + ' Error=' + errorCode + ' ' + message);
  return { status: 'failed', itemId: itemId, ack: ack, errorCode: errorCode, message: message };
}

function auditSampleCsvListingState_Integrated() {
  var props = PropertiesService.getScriptProperties();
  var triggerExists = ScriptApp.getProjectTriggers().some(function(trigger) {
    return trigger.getHandlerFunction() === SAMPLE_LISTING.END_HANDLER;
  });
  var state = {
    csvExists: !!sample_findFileInFolder_(SAMPLE_LISTING.FOLDER_ID, SAMPLE_LISTING.CSV_FILE_NAME),
    imageExists: !!sample_findFileInFolder_(SAMPLE_LISTING.FOLDER_ID, SAMPLE_LISTING.IMAGE_FILE_NAME),
    itemId: props.getProperty(SAMPLE_LISTING.PROP_ITEM_ID) || '',
    listedAt: props.getProperty(SAMPLE_LISTING.PROP_LISTED_AT) || '',
    taskId: props.getProperty(SAMPLE_LISTING.PROP_TASK_ID) || '',
    queueRowIndex: props.getProperty(SAMPLE_LISTING.PROP_QUEUE_ROW) || '',
    endTriggerExists: triggerExists
  };
  Logger.log(JSON.stringify(state));
  return state;
}

function eal_findScheduledTargets_(sheet, now) {
  return eal_getDataRows_(sheet).filter(function(row) {
    var data = row.data;
    var status = String(data[EAL.COL.STATUS - 1] || '').trim();
    var mode = String(data[EAL.COL.LISTING_MODE - 1] || '').trim();
    var scheduledRaw = data[EAL.COL.SCHEDULED_DATETIME - 1];
    var itemId = String(data[EAL.COL.ITEM_ID - 1] || '').trim();

    if (status !== EAL.STATUS.AUTO_WAIT) return false;
    if (mode !== EAL.LISTING_MODE.AUTO) return false;
    if (itemId !== '') return false;
    if (!scheduledRaw || String(scheduledRaw).trim() === '') return false;

    var scheduledAt = new Date(scheduledRaw);
    return !isNaN(scheduledAt.getTime()) && scheduledAt <= now;
  });
}

function eal_executeListing_(target, execSource, sheet, ss, isDryRun) {
  var taskId = String(target.data[EAL.COL.TASK_ID - 1] || '');
  var sku = String(target.data[EAL.COL.SKU - 1] || '').trim();
  var context = '[' + execSource + '] ' + taskId + ' SKU=' + sku;

  Logger.log('Target row: ' + target.rowNum + ' / ' + context);

  var checks = eal_validateRow_(target.data, ss);
  var errors = checks.filter(function(item) { return item.indexOf('ERROR:') === 0; });
  checks.forEach(function(item) { Logger.log(item); });

  if (errors.length > 0) {
    var validationMessage = errors.join(' / ').substring(0, 300);
    eal_recordResult_(sheet, target.rowNum, {
      status: EAL.STATUS.ERROR,
      errDetail: validationMessage,
      source: execSource,
      result: EAL.EXEC_RESULT.ERROR
    });
    eal_writeExecLog_(ss, execSource, 'ERROR', context + ' validation failed: ' + errors[0], isDryRun);
    return;
  }

  if (isDryRun) {
    eal_recordResult_(sheet, target.rowNum, {
      status: EAL.STATUS.DRY_RUN,
      source: execSource,
      result: EAL.EXEC_RESULT.DRY_RUN
    });
    eal_writeExecLog_(ss, execSource, 'DRY_RUN', context + ' checks passed', true);
    Logger.log('DRY_RUN completed. No eBay API call was made.');
    return;
  }

  var accessToken = eal_getAccessToken_();
  if (!accessToken) {
    eal_recordResult_(sheet, target.rowNum, {
      status: EAL.STATUS.ERROR,
      errDetail: 'Access token fetch failed',
      source: execSource,
      result: EAL.EXEC_RESULT.ERROR
    });
    eal_writeExecLog_(ss, execSource, 'ERROR', context + ' token fetch failed', isDryRun);
    return;
  }

  var price = String(target.data[EAL.COL.PRICE_USD - 1]);
  var pictureUrl = String(target.data[EAL.COL.PICTURE_URL - 1]);

  var verifyXml = eal_buildXml_(target.data, price, pictureUrl, false);
  var verifyResp = eal_callTradingApi_('VerifyAddFixedPriceItem', verifyXml, accessToken);
  var verifyAck = eal_extractXml_(verifyResp, 'Ack');
  var verifyMsg = eal_extractXml_(verifyResp, 'ShortMessage') || '';
  var verifyErr = eal_extractXml_(verifyResp, 'ErrorCode') || '';

  sheet.getRange(target.rowNum, EAL.COL.VERIFY_ACK).setValue(verifyAck);
  sheet.getRange(target.rowNum, EAL.COL.VERIFY_MESSAGE).setValue(verifyMsg.substring(0, 300));

  if (verifyAck !== 'Success' && verifyAck !== 'Warning') {
    eal_recordResult_(sheet, target.rowNum, {
      status: EAL.STATUS.VERIFY_FAILED,
      verifyStatus: EAL.STATUS.VERIFY_FAILED,
      errCode: verifyErr,
      errDetail: verifyMsg.substring(0, 300),
      source: execSource,
      result: EAL.EXEC_RESULT.VERIFY_FAILED
    });
    eal_writeListingLog_(ss, taskId, sku, 'VERIFY_FAILED', '', verifyMsg);
    eal_writeExecLog_(ss, execSource, 'VERIFY_FAILED', context + ' ' + verifyMsg, isDryRun);
    Logger.log('Verify failed: ' + verifyMsg);
    return;
  }

  sheet.getRange(target.rowNum, EAL.COL.VERIFY_STATUS).setValue(EAL.STATUS.VERIFY_SUCCESS);

  var addXml = eal_buildXml_(target.data, price, pictureUrl, true);
  var addResp = eal_callTradingApi_('AddFixedPriceItem', addXml, accessToken);
  var addAck = eal_extractXml_(addResp, 'Ack');
  var itemId = eal_extractXml_(addResp, 'ItemID') || '';
  var addMsg = eal_extractXml_(addResp, 'ShortMessage') || '';
  var addErr = eal_extractXml_(addResp, 'ErrorCode') || '';

  if (addAck === 'Success' || addAck === 'Warning') {
    var listingUrl = itemId ? 'https://www.ebay.com/itm/' + itemId : '';
    eal_recordResult_(sheet, target.rowNum, {
      status: EAL.STATUS.PUBLISHED,
      verifyStatus: EAL.STATUS.VERIFY_SUCCESS,
      itemId: itemId,
      listingUrl: listingUrl,
      publishDatetime: eal_nowJst_(),
      source: execSource,
      result: EAL.EXEC_RESULT.SUCCESS
    });
    eal_writeListingLog_(ss, taskId, sku, 'SUCCESS', itemId, addMsg);
    eal_writeExecLog_(ss, execSource, 'SUCCESS', context + ' ItemID=' + itemId, isDryRun);
    Logger.log('Listing published. ItemID=' + itemId);
    return;
  }

  eal_recordResult_(sheet, target.rowNum, {
    status: EAL.STATUS.ERROR,
    errCode: addErr,
    errDetail: addMsg.substring(0, 300),
    source: execSource,
    result: EAL.EXEC_RESULT.ADD_FAILED
  });
  eal_writeListingLog_(ss, taskId, sku, 'ADD_FAILED', '', addMsg);
  eal_writeExecLog_(ss, execSource, 'ADD_FAILED', context + ' ' + addMsg, isDryRun);
  Logger.log('Listing failed: ' + addMsg);
}

function eal_validateRow_(data, ss) {
  var results = [];
  var required = [
    { col: EAL.COL.SKU, label: 'SKU' },
    { col: EAL.COL.EBAY_TITLE, label: 'eBay title' },
    { col: EAL.COL.CATEGORY_ID, label: 'CategoryID' },
    { col: EAL.COL.CONDITION_ID, label: 'ConditionID' },
    { col: EAL.COL.QUANTITY, label: 'Quantity' },
    { col: EAL.COL.PRICE_USD, label: 'Price USD' },
    { col: EAL.COL.PICTURE_URL, label: 'Picture URL' },
    { col: EAL.COL.SHIPPING_ID, label: 'ShippingProfileID' },
    { col: EAL.COL.RETURN_ID, label: 'ReturnProfileID' },
    { col: EAL.COL.PAYMENT_ID, label: 'PaymentProfileID' },
    { col: EAL.COL.CITY, label: 'City' },
    { col: EAL.COL.COUNTRY, label: 'Country' },
    { col: EAL.COL.DESCRIPTION, label: 'Description' }
  ];

  required.forEach(function(item) {
    var value = String(data[item.col - 1] || '').trim();
    if (!value) results.push('ERROR: Missing ' + item.label);
  });

  var sku = String(data[EAL.COL.SKU - 1] || '').trim();
  var itemId = String(data[EAL.COL.ITEM_ID - 1] || '').trim();
  var title = String(data[EAL.COL.EBAY_TITLE - 1] || '').trim();
  var price = Number(data[EAL.COL.PRICE_USD - 1] || 0);
  var quantity = Number(data[EAL.COL.QUANTITY - 1] || 0);
  var pictureUrl = String(data[EAL.COL.PICTURE_URL - 1] || '').trim();

  if (itemId) results.push('ERROR: ItemID already exists: ' + itemId);
  if (title.length > 80) results.push('ERROR: eBay title is over 80 chars: ' + title.length);
  if (!(price > 0)) results.push('ERROR: Price must be greater than 0');
  if (!(quantity > 0)) results.push('ERROR: Quantity must be greater than 0');
  if (pictureUrl && !/^https?:\/\//i.test(pictureUrl)) results.push('ERROR: Picture URL must start with http(s)');

  if (sku) {
    if (sku.indexOf(SAMPLE_LISTING.SKU_PREFIX) !== 0) {
      var fba = eal_getFbaFulfillable_(sku, ss);
      if (fba.error) results.push('ERROR: FBA stock check failed: ' + fba.error);
      else if (!(Number(fba.qty) > 0)) results.push('ERROR: FBA stock is zero or blank: ' + fba.qty);
    }

    var dup = eal_checkSkuDup_(sku, ss);
    if (dup.found) results.push('ERROR: Duplicate SKU found in ' + dup.source);
  }

  if (results.length === 0) results.push('OK: validation passed');
  return results;
}

function eal_getFbaFulfillable_(sku, ss) {
  var sheet = ss.getSheetByName(EAL_FBA.SHEET_NAME);
  if (!sheet) return { qty: null, error: 'Missing sheet: ' + EAL_FBA.SHEET_NAME };

  var lastRow = sheet.getLastRow();
  var lastCol = sheet.getLastColumn();
  if (lastRow < 1 || lastCol < 1) return { qty: null, error: 'FBA sheet is empty' };

  var searchRows = Math.min(EAL_FBA.HEADER_SEARCH_ROWS, lastRow);
  var headerValues = sheet.getRange(1, 1, searchRows, lastCol).getValues();
  var headerRow = -1;
  var skuCol = -1;
  var stockCol = -1;

  for (var r = 0; r < headerValues.length; r++) {
    for (var c = 0; c < headerValues[r].length; c++) {
      var header = String(headerValues[r][c] || '').trim().toLowerCase();
      if (header === EAL_FBA.HEADER_SKU) skuCol = c;
      if (header === EAL_FBA.HEADER_STOCK) stockCol = c;
    }
    if (skuCol >= 0 && stockCol >= 0) {
      headerRow = r;
      break;
    }
  }

  if (headerRow < 0) {
    return { qty: null, error: 'Required FBA headers not found' };
  }

  var dataStartRow = headerRow + 2;
  if (dataStartRow > lastRow) return { qty: null, error: 'FBA sheet has no data rows' };

  var values = sheet.getRange(dataStartRow, 1, lastRow - dataStartRow + 1, lastCol).getValues();
  var wanted = String(sku || '').trim();
  for (var i = 0; i < values.length; i++) {
    if (String(values[i][skuCol] || '').trim() === wanted) {
      return { qty: Number(values[i][stockCol] || 0), error: null };
    }
  }

  return { qty: null, error: 'SKU not found in FBA Stock List: ' + wanted };
}

function eal_checkSkuDup_(sku, ss) {
  var rawSheet = ss.getSheetByName(EAL.SHEET_RAW);
  if (rawSheet) {
    var rawValues = rawSheet.getDataRange().getValues();
    for (var i = 1; i < rawValues.length; i++) {
      if (String(rawValues[i][1] || '').trim() === sku) {
        return { found: true, source: EAL.SHEET_RAW + ' row ' + (i + 1) };
      }
    }
  }

  var listSheet = ss.getSheetByName(EAL.SHEET_LIST);
  if (listSheet) {
    var listValues = listSheet.getDataRange().getValues();
    for (var j = 1; j < listValues.length; j++) {
      if (String(listValues[j][3] || '').trim() === sku) {
        return { found: true, source: EAL.SHEET_LIST + ' row ' + (j + 1) };
      }
    }
  }

  return { found: false, source: '' };
}

function eal_buildXml_(data, priceStr, pictureUrl, isAdd) {
  var tag = isAdd ? 'AddFixedPriceItem' : 'VerifyAddFixedPriceItem';
  var title = eal_esc_(String(data[EAL.COL.EBAY_TITLE - 1] || ''));
  var sku = eal_esc_(String(data[EAL.COL.SKU - 1] || ''));
  var categoryId = eal_esc_(String(data[EAL.COL.CATEGORY_ID - 1] || ''));
  var conditionId = eal_esc_(String(data[EAL.COL.CONDITION_ID - 1] || ''));
  var quantity = eal_esc_(String(data[EAL.COL.QUANTITY - 1] || '1'));
  var shippingId = eal_esc_(String(data[EAL.COL.SHIPPING_ID - 1] || ''));
  var returnId = eal_esc_(String(data[EAL.COL.RETURN_ID - 1] || ''));
  var paymentId = eal_esc_(String(data[EAL.COL.PAYMENT_ID - 1] || ''));
  var city = eal_esc_(String(data[EAL.COL.CITY - 1] || ''));
  var country = eal_esc_(String(data[EAL.COL.COUNTRY - 1] || ''));
  var description = String(data[EAL.COL.DESCRIPTION - 1] || '');

  return '<?xml version="1.0" encoding="utf-8"?>\n'
    + '<' + tag + 'Request xmlns="urn:ebay:apis:eBLBaseComponents">\n'
    + '  <Item>\n'
    + '    <Title>' + title + '</Title>\n'
    + '    <SKU>' + sku + '</SKU>\n'
    + '    <PrimaryCategory><CategoryID>' + categoryId + '</CategoryID></PrimaryCategory>\n'
    + '    <StartPrice currencyID="USD">' + eal_esc_(priceStr) + '</StartPrice>\n'
    + '    <Quantity>' + quantity + '</Quantity>\n'
    + '    <ListingType>FixedPriceItem</ListingType>\n'
    + '    <ListingDuration>GTC</ListingDuration>\n'
    + '    <OutOfStockControl>true</OutOfStockControl>\n'
    + '    <ConditionID>' + conditionId + '</ConditionID>\n'
    + '    <PictureDetails><PictureURL>' + eal_esc_(pictureUrl) + '</PictureURL></PictureDetails>\n'
    + '    <Location>' + city + '</Location>\n'
    + '    <Currency>USD</Currency>\n'
    + '    <Country>' + country + '</Country>\n'
    + '    <SellerProfiles>\n'
    + '      <SellerShippingProfile><ShippingProfileID>' + shippingId + '</ShippingProfileID></SellerShippingProfile>\n'
    + '      <SellerReturnProfile><ReturnProfileID>' + returnId + '</ReturnProfileID></SellerReturnProfile>\n'
    + '      <SellerPaymentProfile><PaymentProfileID>' + paymentId + '</PaymentProfileID></SellerPaymentProfile>\n'
    + '    </SellerProfiles>\n'
    + '    <ItemSpecifics>\n'
    + '      <NameValueList><Name>Brand</Name><Value>Shimano</Value></NameValueList>\n'
    + '      <NameValueList><Name>Type</Name><Value>Spinning</Value></NameValueList>\n'
    + '    </ItemSpecifics>\n'
    + '    <Description><![CDATA[\n' + description + '\n    ]]></Description>\n'
    + '  </Item>\n'
    + '</' + tag + 'Request>';
}

function eal_callTradingApi_(callName, xml, accessToken) {
  var props = PropertiesService.getScriptProperties();
  var appId = props.getProperty('EBAY_CLIENT_ID') || '';
  var certId = props.getProperty('EBAY_CLIENT_SECRET') || '';
  var devId = props.getProperty('EBAY_DEV_ID') || '';

  try {
    var response = UrlFetchApp.fetch(EAL.TRADING_ENDPOINT, {
      method: 'post',
      headers: {
        'X-EBAY-API-SITEID': EAL.SITE_ID,
        'X-EBAY-API-COMPATIBILITY-LEVEL': EAL.COMPATIBILITY_LEVEL,
        'X-EBAY-API-CALL-NAME': callName,
        'X-EBAY-API-APP-ID': appId,
        'X-EBAY-API-DEV-ID': devId,
        'X-EBAY-API-CERT-ID': certId,
        'X-EBAY-API-IAF-TOKEN': accessToken,
        'Content-Type': 'text/xml'
      },
      payload: xml,
      muteHttpExceptions: true
    });
    return response.getContentText();
  } catch (error) {
    Logger.log('Trading API exception: ' + callName + ' / ' + error.message);
    return '<Ack>Failure</Ack><ShortMessage>' + eal_esc_(error.message) + '</ShortMessage>';
  }
}

function eal_getAccessToken_() {
  try {
    var props = PropertiesService.getScriptProperties();
    var clientId = props.getProperty('EBAY_CLIENT_ID');
    var clientSecret = props.getProperty('EBAY_CLIENT_SECRET');
    var refreshToken = props.getProperty('EBAY_REFRESH_TOKEN');

    if (!clientId || !clientSecret || !refreshToken) {
      Logger.log('Missing Script Properties: EBAY_CLIENT_ID / EBAY_CLIENT_SECRET / EBAY_REFRESH_TOKEN');
      return null;
    }

    var response = UrlFetchApp.fetch('https://api.ebay.com/identity/v1/oauth2/token', {
      method: 'post',
      headers: {
        Authorization: 'Basic ' + Utilities.base64Encode(clientId + ':' + clientSecret),
        'Content-Type': 'application/x-www-form-urlencoded'
      },
      payload: {
        grant_type: 'refresh_token',
        refresh_token: refreshToken
      },
      muteHttpExceptions: true
    });

    var json = JSON.parse(response.getContentText());
    return json.access_token || null;
  } catch (error) {
    Logger.log('Access token exception: ' + error.message);
    return null;
  }
}

function eal_recordResult_(sheet, rowNum, opts) {
  if (opts.status) sheet.getRange(rowNum, EAL.COL.STATUS).setValue(opts.status);
  if (opts.verifyStatus) sheet.getRange(rowNum, EAL.COL.VERIFY_STATUS).setValue(opts.verifyStatus);
  if (opts.errCode) sheet.getRange(rowNum, EAL.COL.ERROR_CODE).setValue(opts.errCode);
  if (opts.errDetail) sheet.getRange(rowNum, EAL.COL.ERROR_DETAIL).setValue(opts.errDetail);
  if (opts.itemId) sheet.getRange(rowNum, EAL.COL.ITEM_ID).setValue(opts.itemId);
  if (opts.listingUrl) sheet.getRange(rowNum, EAL.COL.LISTING_URL).setValue(opts.listingUrl);
  if (opts.publishDatetime) sheet.getRange(rowNum, EAL.COL.PUBLISH_DATETIME).setValue(opts.publishDatetime);
  if (opts.source) sheet.getRange(rowNum, EAL.COL.EXEC_SOURCE).setValue(opts.source);
  if (opts.result) sheet.getRange(rowNum, EAL.COL.EXEC_RESULT).setValue(opts.result);
  sheet.getRange(rowNum, EAL.COL.UPDATED_DATE).setValue(eal_nowJst_());
}

function eal_writeListingLog_(ss, taskId, sku, status, itemId, message) {
  var sheet = ss.getSheetByName(EAL.SHEET_LOG);
  if (!sheet) return;
  sheet.appendRow([
    eal_nowJst_(),
    sku,
    taskId,
    '',
    '',
    status,
    status === 'SUCCESS' ? 'TRUE' : 'FALSE',
    status,
    itemId,
    '',
    String(message || '').substring(0, 200)
  ]);
}

function eal_writeExecLog_(ss, funcName, status, detail, isDryRun) {
  var sheet = ss.getSheetByName(EAL.SHEET_EXEC_LOG);
  if (!sheet) return;
  sheet.appendRow([eal_nowJst_(), funcName, isDryRun ? 'DRY_RUN' : 'PRODUCTION', status, detail]);
}

function eal_getDataRows_(sheet) {
  var lastRow = sheet.getLastRow();
  if (lastRow < EAL.DATA_START_ROW) return [];

  var lastCol = Math.max(sheet.getLastColumn(), EAL.COL.EXEC_RESULT);
  var numRows = lastRow - EAL.DATA_START_ROW + 1;
  var values = sheet.getRange(EAL.DATA_START_ROW, 1, numRows, lastCol).getValues();
  var rows = [];

  values.forEach(function(row, index) {
    if (String(row[EAL.COL.TASK_ID - 1] || '').trim() === '') return;
    rows.push({ rowNum: index + EAL.DATA_START_ROW, data: row });
  });

  return rows;
}

function eal_isDryRun_() {
  var value = PropertiesService.getScriptProperties().getProperty(EAL.PROP_ENABLED);
  return String(value || '').trim().toLowerCase() !== 'true';
}

function eal_esc_(value) {
  return String(value)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&apos;');
}

function eal_extractXml_(xml, tag) {
  var match = String(xml || '').match(new RegExp('<' + tag + '>([^<]*)</' + tag + '>'));
  return match ? match[1].trim() : '';
}

function eal_nowJst_() {
  return Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy/MM/dd HH:mm:ss');
}

function eat_deleteTargetTriggers_(funcNames) {
  var triggers = ScriptApp.getProjectTriggers();
  var deleted = 0;

  triggers.forEach(function(trigger) {
    var handler = trigger.getHandlerFunction();
    if (EAT.PROTECTED_FUNCS.indexOf(handler) !== -1) return;
    if (funcNames.indexOf(handler) === -1) return;

    ScriptApp.deleteTrigger(trigger);
    deleted++;
  });

  return deleted;
}

function sample_findFileInFolder_(folderId, fileName) {
  try {
    var folder = DriveApp.getFolderById(folderId);
    var files = folder.getFilesByName(fileName);
    return files.hasNext() ? files.next() : null;
  } catch (error) {
    Logger.log('Drive file lookup failed: ' + fileName + ' / ' + error.message);
    return null;
  }
}

function sample_readFirstCsvData_(csvFile) {
  var content = csvFile.getBlob().getDataAsString('UTF-8');
  var rows = Utilities.parseCsv(content);
  if (!rows || rows.length < 2) throw new Error('sample CSV has no data row');

  var headers = rows[0];
  var values = rows[1];
  var data = {};
  headers.forEach(function(header, index) {
    data[String(header || '').trim()] = values[index] ? String(values[index]).trim() : '';
  });

  return {
    title: data.Title || data.title || '',
    sku: data['Custom label (SKU)'] || data.sku || '',
    category: data['Category ID'] || data.category || '',
    condition: data['Condition ID'] || data.condition || '',
    price: data['Start price'] || data.price || '',
    quantity: data.Quantity || data.quantity || '',
    imageUrl: data['Item photo URL'] || data.image_url || '',
    description: data.Description || data.description || ''
  };
}

function sample_importCsvToQueue_(csvData, imageFile) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(EAL.SHEET_QUEUE);
  if (!sheet) throw new Error('Missing sheet: ' + EAL.SHEET_QUEUE);

  var now = new Date();
  var stamp = Utilities.formatDate(now, 'Asia/Tokyo', 'yyyyMMdd-HHmmss');
  var sku = SAMPLE_LISTING.SKU_PREFIX + stamp;
  var row = sheet.getLastRow() + 1;
  var imageUrl = csvData.imageUrl || ('https://drive.google.com/uc?id=' + imageFile.getId());

  sheet.getRange(row, EAL.COL.TASK_ID).setValue('SAMPLE-' + Date.now());
  sheet.getRange(row, EAL.COL.STATUS).setValue(EAL.STATUS.AUTO_WAIT);
  sheet.getRange(row, EAL.COL.SKU).setValue(sku);
  sheet.getRange(row, EAL.COL.EBAY_TITLE).setValue(csvData.title || SAMPLE_LISTING.DEFAULT_TITLE);
  sheet.getRange(row, EAL.COL.CATEGORY_ID).setValue(Number(csvData.category) || SAMPLE_LISTING.DEFAULT_CATEGORY_ID);
  sheet.getRange(row, EAL.COL.CONDITION_ID).setValue(Number(csvData.condition) || SAMPLE_LISTING.DEFAULT_CONDITION_ID);
  sheet.getRange(row, EAL.COL.QUANTITY).setValue(Number(csvData.quantity) || 1);
  sheet.getRange(row, EAL.COL.PRICE_USD).setValue(Number(csvData.price) || SAMPLE_LISTING.DEFAULT_PRICE_USD);
  sheet.getRange(row, EAL.COL.PICTURE_URL).setValue(imageUrl);
  sheet.getRange(row, EAL.COL.SHIPPING_ID).setValue(SAMPLE_LISTING.DEFAULT_SHIPPING_ID);
  sheet.getRange(row, EAL.COL.RETURN_ID).setValue(SAMPLE_LISTING.DEFAULT_RETURN_ID);
  sheet.getRange(row, EAL.COL.PAYMENT_ID).setValue(SAMPLE_LISTING.DEFAULT_PAYMENT_ID);
  sheet.getRange(row, EAL.COL.CITY).setValue(SAMPLE_LISTING.DEFAULT_CITY);
  sheet.getRange(row, EAL.COL.COUNTRY).setValue(SAMPLE_LISTING.DEFAULT_COUNTRY);
  sheet.getRange(row, EAL.COL.DESCRIPTION).setValue(csvData.description || csvData.title || SAMPLE_LISTING.DEFAULT_TITLE);
  sheet.getRange(row, EAL.COL.CREATED_DATE).setValue(eal_nowJst_());
  sheet.getRange(row, EAL.COL.UPDATED_DATE).setValue(eal_nowJst_());
  sheet.getRange(row, EAL.COL.LISTING_MODE).setValue(EAL.LISTING_MODE.AUTO);
  sheet.getRange(row, EAL.COL.SCHEDULED_DATETIME).setValue(now);

  return row;
}

function sample_getItemIdFromQueueRow_(rowIndex) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(EAL.SHEET_QUEUE);
  if (!sheet) return '';
  return String(sheet.getRange(rowIndex, EAL.COL.ITEM_ID).getValue() || '').trim();
}

function sample_scheduleEndAfterOneHour_() {
  sample_deleteEndTrigger_();
  var trigger = ScriptApp.newTrigger(SAMPLE_LISTING.END_HANDLER)
    .timeBased()
    .after(SAMPLE_LISTING.END_AFTER_MS)
    .create();
  PropertiesService.getScriptProperties()
    .setProperty(SAMPLE_LISTING.PROP_END_TRIGGER_ID, trigger.getUniqueId());
  return trigger.getUniqueId();
}

function sample_deleteEndTrigger_() {
  ScriptApp.getProjectTriggers().forEach(function(trigger) {
    if (trigger.getHandlerFunction() === SAMPLE_LISTING.END_HANDLER) {
      ScriptApp.deleteTrigger(trigger);
    }
  });
}

function sample_recoverLatestSampleItemId_() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(EAL.SHEET_QUEUE);
  if (!sheet) return '';

  var rows = eal_getDataRows_(sheet);
  for (var i = rows.length - 1; i >= 0; i--) {
    var data = rows[i].data;
    var sku = String(data[EAL.COL.SKU - 1] || '').trim();
    var itemId = String(data[EAL.COL.ITEM_ID - 1] || '').trim();
    var status = String(data[EAL.COL.STATUS - 1] || '').trim();
    if (sku.indexOf(SAMPLE_LISTING.SKU_PREFIX) === 0 && itemId && status === EAL.STATUS.PUBLISHED) {
      PropertiesService.getScriptProperties().setProperty(SAMPLE_LISTING.PROP_ITEM_ID, itemId);
      return itemId;
    }
  }
  return '';
}

function sample_buildEndFixedPriceItemXml_(itemId) {
  return '<?xml version="1.0" encoding="utf-8"?>\n'
    + '<EndFixedPriceItemRequest xmlns="urn:ebay:apis:eBLBaseComponents">\n'
    + '  <EndingReason>NotAvailable</EndingReason>\n'
    + '  <ItemID>' + eal_esc_(itemId) + '</ItemID>\n'
    + '</EndFixedPriceItemRequest>';
}

function sample_markEndedInQueue_(itemId, ack, message) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(EAL.SHEET_QUEUE);
  if (!sheet) return;

  var rows = eal_getDataRows_(sheet);
  rows.forEach(function(row) {
    if (String(row.data[EAL.COL.ITEM_ID - 1] || '').trim() === String(itemId)) {
      sheet.getRange(row.rowNum, EAL.COL.END_TYPE).setValue('サンプル自動終了');
      sheet.getRange(row.rowNum, EAL.COL.END_REASON).setValue('1時間後自動終了');
      sheet.getRange(row.rowNum, EAL.COL.END_STATUS).setValue('終了済み');
      sheet.getRange(row.rowNum, EAL.COL.END_DATETIME).setValue(eal_nowJst_());
      sheet.getRange(row.rowNum, EAL.COL.END_ACK).setValue(ack);
      sheet.getRange(row.rowNum, EAL.COL.END_MESSAGE).setValue(String(message || '').substring(0, 300));
      sheet.getRange(row.rowNum, EAL.COL.UPDATED_DATE).setValue(eal_nowJst_());
    }
  });
}

function sample_clearProperties_() {
  var props = PropertiesService.getScriptProperties();
  props.deleteProperty(SAMPLE_LISTING.PROP_ITEM_ID);
  props.deleteProperty(SAMPLE_LISTING.PROP_LISTED_AT);
  props.deleteProperty(SAMPLE_LISTING.PROP_TASK_ID);
  props.deleteProperty(SAMPLE_LISTING.PROP_QUEUE_ROW);
  props.deleteProperty(SAMPLE_LISTING.PROP_END_TRIGGER_ID);
}

/**
 * EbayAutoEndPipeline.gs  v2
 * ─────────────────────────────────────────────────────────────────
 * 修正点（差し戻し対応）:
 *   問題3: 実行ログ文字列はコードに含まない
 *   問題4: 処理対象名を eae_ 命名規則に統一（1箇所修正）
 *   問題5: DATA_START_ROW = 5（データ行が5行目から始まる）
 *
 * 終了処理の2条件（優先順位あり：条件A → 条件B → 処理終了）:
 *
 * 条件A: 完全終了（出品者がシートに「完全終了」をセットした時）
 *   → AG列（終了区分）= 「完全終了」をセットした場合
 *   → Y列（ItemID）あり
 *   → AI列（終了状態）が「終了済み」でない
 *
 * 条件B: sample自動終了（出品後を自動で終了させるための実装）
 *   → タイトルまたはSKUが 単一で「sample」（大文字小文字問わず）
 *     ‥例示: sample product / sample-001 等は不一致（=除外）
 *   → Y列（ItemID）あり
 *   → AF列（出品日時）が記録されている
 *   → 出品日時から1時間以上経過している
 *   → AI列（終了状態）が「終了済み」でない
 *
 * 在庫警告:
 *   FBA在庫0の処理は EndFixedPriceItem を呼ばない（在庫管理機能が別途ある）
 *
 * 注記      : LINE通知なし / メール通知なし
 *             EBAY_AUTO_END_ENABLED=true の時のみAPI実行
 * 認証方式  : X-EBAY-API-IAF-TOKEN（Trading API）
 *             → X-EBAY-C-IAF-TOKEN / X-EBAY-IAF-TOKEN は不使用（非推奨）
 * ─────────────────────────────────────────────────────────────────
 */

// ════════════════════════════════════════════════════════════════
// 定数（EbayAutoListingPipeline.gs の EAL.COL を引き継ぎ）
// ════════════════════════════════════════════════════════════════

var EAE = {
  // ── シート名 ──
  SHEET_QUEUE   : 'eBay出品Aキュー',
  SHEET_LOG     : 'eBay_LISTING_LOG',
  SHEET_EXEC_LOG: '実行ログ',

  // ── データ開始行（1-3行目=ヘッダー・説明、4=ダミー、5=データ行）──
  DATA_START_ROW: 5,

  // ── 列番号（EAL.COL を引き継ぎ）──
  COL: {
    TASK_ID         :  1,   // A
    HUMAN_APPROVAL  :  2,   // B
    STATUS          :  3,   // C
    SKU             :  4,   // D
    ASIN            :  5,   // E
    PRODUCT_NAME    :  6,   // F
    EBAY_TITLE      :  7,   // G
    CATEGORY_ID     :  8,   // H
    CATEGORY_NAME   :  9,   // I
    CONDITION_ID    : 10,   // J
    QUANTITY        : 11,   // K
    PRICE_USD       : 12,   // L
    PICTURE_URL     : 13,   // M
    SHIPPING_ID     : 14,   // N
    RETURN_ID       : 15,   // O
    PAYMENT_ID      : 16,   // P
    CITY            : 17,   // Q
    STATE           : 18,   // R
    COUNTRY         : 19,   // S
    DESCRIPTION     : 20,   // T
    PREVIEW_RESULT  : 21,   // U
    VERIFY_STATUS   : 22,   // V
    VERIFY_ACK      : 23,   // W
    VERIFY_MESSAGE  : 24,   // X
    ITEM_ID         : 25,   // Y
    LISTING_URL     : 26,   // Z
    ERROR_CODE      : 27,   // AA
    ERROR_DETAIL    : 28,   // AB
    CREATED_DATE    : 29,   // AC
    UPDATED_DATE    : 30,   // AD
    NOTES           : 31,   // AE
    PUBLISH_DATETIME: 32,   // AF ☆出品日時（sample1時間判定の起点）
    END_TYPE        : 33,   // AG 終了区分（なし／完全終了）
    END_REASON      : 34,   // AH 終了理由
    END_STATUS      : 35,   // AI 終了状態
    END_DATETIME    : 36,   // AJ 終了日時
    END_ACK         : 37,   // AK 終了Ack
    END_MESSAGE     : 38,   // AL 終了メッセージ
  },

  // ── ステータス値・終了区分値 ──
  STATUS    : { ENDED: '終了済み/', ERROR: 'エラー', DRY_RUN: 'DRY_RUN済み' },
  END_TYPE  : { NONE: 'なし', FULL_END: '完全終了' },
  END_STATUS: { ENDED: '終了済み' },

  // ── Trading API ──
  TRADING_ENDPOINT   : 'https://api.ebay.com/ws/api.dll',
  COMPATIBILITY_LEVEL: '1225',
  SITE_ID            : '0',

  // ── Script Properties キー ──
  PROP_ENABLED: 'EBAY_AUTO_END_ENABLED',

  // ── sample 1時間経過閾値（ミリ秒）──
  SAMPLE_MIN_ELAPSED_MS: 60 * 60 * 1000,
};

// ════════════════════════════════════════════════════════════════
// メイン: runEbayAutoEndPipeline()
// ════════════════════════════════════════════════════════════════

/**
 * 条件A（完全終了）または条件B（sample自動終了）を求めて行を
 * 最大1件ずつ EndFixedPriceItem を呼び出す。
 * LockService で2重実行防止。
 */
function runEbayAutoEndPipeline() {
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(10000)) {
    Logger.log('⚠️  別の実行が処理中のためスキップ（LockService）。');
    return;
  }
  try {
    eae_runInternal_();
  } finally {
    lock.releaseLock();
  }
}

function eae_runInternal_() {
  var isDryRun = eae_isDryRun_();
  Logger.log('=== runEbayAutoEndPipeline 開始 ===');
  Logger.log('モード: ' + (isDryRun ? 'DRY_RUN（API実行なし）' : '本番'));

  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(EAE.SHEET_QUEUE);
  if (!sheet) {
    Logger.log('✗ シート「' + EAE.SHEET_QUEUE + '」が見つかりません。');
    return;
  }

  var rows = eae_getDataRows_(sheet);

  // 優先順位: 条件A → 条件B
  var targetA = null;
  var targetB = null;
  for (var i = 0; i < rows.length; i++) {
    if (!targetA && eae_isConditionA_(rows[i].data)) { targetA = rows[i]; }
    if (!targetB && eae_isConditionB_(rows[i].data)) { targetB = rows[i]; }
    if (targetA) { break; }  // 条件Aが見つかれば優先
  }

  var target = targetA || targetB;
  var reason = targetA ? '条件A（完全終了）' : (targetB ? '条件B（sample自動終了）' : null);

  if (!target) {
    Logger.log('ℹ️   終了処理の対象がありません。');
    eae_writeExecLog_(ss, 'AUTO_END', 'SKIP', '対象なし', isDryRun);
    return;
  }

  var taskId = String(target.data[EAE.COL.TASK_ID - 1] || '');
  var itemId = String(target.data[EAE.COL.ITEM_ID  - 1] || '').trim();
  Logger.log('対象タスク: ' + taskId + '  シート行: ' + target.rowNum + '  終了理由: ' + reason);

  // DRY_RUN 時はここで終了
  if (isDryRun) {
    Logger.log('✅ DRY_RUN: 終了対象を確認しました。本番時は EndFixedPriceItem を実行します。');
    Logger.log('   ItemID: ' + itemId + '  理由: ' + reason);
    sheet.getRange(target.rowNum, EAE.COL.END_REASON).setValue(reason);
    sheet.getRange(target.rowNum, EAE.COL.UPDATED_DATE).setValue(eae_nowJst_());
    eae_writeExecLog_(ss, 'AUTO_END', 'DRY_RUN',
      taskId + ' ItemID:' + itemId + ' 理由:' + reason, isDryRun);
    return;
  }

  // Access Token 取得
  var accessToken = eae_getAccessToken_();
  if (!accessToken) {
    Logger.log('✗ Access Token取得失敗。');
    eae_writeExecLog_(ss, 'AUTO_END', 'ERROR', taskId + ' Token取得失敗', isDryRun);
    return;
  }

  // EndFixedPriceItem 実行
  Logger.log('--- EndFixedPriceItem 実行 ItemID: ' + itemId + ' ---');
  var xml      = eae_buildEndXml_(itemId);
  var respText = eae_callTradingApi_('EndFixedPriceItem', xml, accessToken);
  var ack      = eae_extractXml_(respText, 'Ack');
  var errCode  = eae_extractXml_(respText, 'ErrorCode')    || '';
  var errMsg   = eae_extractXml_(respText, 'ShortMessage') || '';
  Logger.log('EndFixedPriceItem Ack: ' + ack);

  var now = eae_nowJst_();
  if (ack === 'Success' || ack === 'Warning') {
    sheet.getRange(target.rowNum, EAE.COL.STATUS).setValue(EAE.STATUS.ENDED);
    sheet.getRange(target.rowNum, EAE.COL.END_REASON).setValue(reason);
    sheet.getRange(target.rowNum, EAE.COL.END_STATUS).setValue(EAE.END_STATUS.ENDED);
    sheet.getRange(target.rowNum, EAE.COL.END_DATETIME).setValue(now);
    sheet.getRange(target.rowNum, EAE.COL.END_ACK).setValue(ack);
    sheet.getRange(target.rowNum, EAE.COL.END_MESSAGE).setValue(errMsg.substring(0, 300));
    sheet.getRange(target.rowNum, EAE.COL.UPDATED_DATE).setValue(now);
    Logger.log('✅ 終了成功 ItemID: ' + itemId);
    eae_writeListingLog_(ss, taskId, String(target.data[EAE.COL.SKU - 1]), 'END_SUCCESS', itemId, reason);
    eae_writeExecLog_(ss, 'AUTO_END', 'SUCCESS',
      taskId + ' ItemID:' + itemId + ' ' + reason, isDryRun);
  } else {
    sheet.getRange(target.rowNum, EAE.COL.STATUS).setValue(EAE.STATUS.ERROR);
    sheet.getRange(target.rowNum, EAE.COL.END_REASON).setValue(reason);
    sheet.getRange(target.rowNum, EAE.COL.END_ACK).setValue(ack);
    sheet.getRange(target.rowNum, EAE.COL.END_MESSAGE).setValue(errMsg.substring(0, 300));
    sheet.getRange(target.rowNum, EAE.COL.ERROR_CODE).setValue(errCode);
    sheet.getRange(target.rowNum, EAE.COL.ERROR_DETAIL).setValue(errMsg.substring(0, 300));
    sheet.getRange(target.rowNum, EAE.COL.UPDATED_DATE).setValue(now);
    Logger.log('✗ 終了失敗: ' + errMsg);
    eae_writeExecLog_(ss, 'AUTO_END', 'FAILED', taskId + ' ' + errMsg, isDryRun);
  }

  Logger.log('=== runEbayAutoEndPipeline 終了 ===');
}

// ════════════════════════════════════════════════════════════════
// 終了条件判定
// ════════════════════════════════════════════════════════════════

/**
 * 条件A: 完全終了（出品者がシートに「完全終了」をセットした場合）
 */
function eae_isConditionA_(data) {
  var endType   = String(data[EAE.COL.END_TYPE  - 1] || '').trim();
  var itemId    = String(data[EAE.COL.ITEM_ID   - 1] || '').trim();
  var endStatus = String(data[EAE.COL.END_STATUS- 1] || '').trim();
  return endType === EAE.END_TYPE.FULL_END
      && itemId  !== ''
      && endStatus !== EAE.END_STATUS.ENDED;
}

/**
 * 条件B: sample自動終了
 * タイトルまたはSKUが単一で「sample」（大文字小文字問わず）
 * ItemIDあり / AF列（出品日時）あり / 出品から1時間以上経過 / 終了済みでない
 *
 * 【例示一致条件】
 *   title.toLowerCase() === 'sample' の時のみ true
 *   'sample product'.toLowerCase() === 'sample' → false（不一致）
 *   'test sample'.toLowerCase() === 'sample'    → false（不一致）
 */
function eae_isConditionB_(data) {
  var title     = String(data[EAE.COL.EBAY_TITLE- 1] || '').trim();
  var sku       = String(data[EAE.COL.SKU       - 1] || '').trim();
  var itemId    = String(data[EAE.COL.ITEM_ID   - 1] || '').trim();
  var pubRaw    = data[EAE.COL.PUBLISH_DATETIME - 1];
  var endStatus = String(data[EAE.COL.END_STATUS- 1] || '').trim();

  var titleMatch = (title.toLowerCase() === 'sample');
  var skuMatch   = (sku.toLowerCase()   === 'sample');
  if (!titleMatch && !skuMatch) { return false; }

  if (itemId === '') { return false; }

  if (!pubRaw || String(pubRaw).trim() === '') { return false; }

  var publishedAt = new Date(pubRaw);
  if (isNaN(publishedAt.getTime())) { return false; }
  var elapsed = Date.now() - publishedAt.getTime();
  if (elapsed < EAE.SAMPLE_MIN_ELAPSED_MS) {
    Logger.log('ℹ️  sample出品から1時間未満のためスキップ: ItemID=' + itemId
             + ' 経過=' + Math.floor(elapsed / 60000) + '分');
    return false;
  }

  if (endStatus === EAE.END_STATUS.ENDED) { return false; }

  return true;
}

// ════════════════════════════════════════════════════════════════
// 補助: EndFixedPriceItem XML 生成
// ════════════════════════════════════════════════════════════════

function eae_buildEndXml_(itemId) {
  return '<?xml version="1.0" encoding="utf-8"?>\n'
    + '<EndFixedPriceItemRequest xmlns="urn:ebay:apis:eBLBaseComponents">\n'
    + '  <!-- 認証: X-EBAY-API-IAF-TOKEN ヘッダーで Access Token を渡す -->\n'
    + '  <!-- RequesterCredentials は使用しない（IAFトークン方式）-->\n'
    + '  <ItemID>' + eae_esc_(itemId) + '</ItemID>\n'
    + '  <EndingReason>NotAvailable</EndingReason>\n'
    + '</EndFixedPriceItemRequest>';
}

// ════════════════════════════════════════════════════════════════
// 補助: Trading API 呼び出し
// ════════════════════════════════════════════════════════════════

function eae_callTradingApi_(callName, xml, accessToken) {
  var props  = PropertiesService.getScriptProperties();
  var appId  = props.getProperty('EBAY_CLIENT_ID')     || '';
  var certId = props.getProperty('EBAY_CLIENT_SECRET') || '';
  var devId  = props.getProperty('EBAY_DEV_ID')        || '';
  try {
    var resp = UrlFetchApp.fetch(EAE.TRADING_ENDPOINT, {
      method : 'post',
      headers: {
        'X-EBAY-API-SITEID'             : EAE.SITE_ID,
        'X-EBAY-API-COMPATIBILITY-LEVEL': EAE.COMPATIBILITY_LEVEL,
        'X-EBAY-API-CALL-NAME'          : callName,
        'X-EBAY-API-APP-ID'             : appId,
        'X-EBAY-API-DEV-ID'             : devId,
        'X-EBAY-API-CERT-ID'            : certId,
        // Trading API: X-EBAY-API-IAF-TOKEN を使用
        // → X-EBAY-C-IAF-TOKEN / X-EBAY-IAF-TOKEN は不使用（非推奨）
        'X-EBAY-API-IAF-TOKEN'          : accessToken,
        'Content-Type'                  : 'text/xml',
      },
      payload             : xml,
      muteHttpExceptions  : true,
    });
    return resp.getContentText();
  } catch (e) {
    Logger.log('✗ eae_callTradingApi_ 失敗(' + callName + '): ' + e.message);
    return '<Ack>Failure</Ack><ShortMessage>' + e.message + '</ShortMessage>';
  }
}

// ════════════════════════════════════════════════════════════════
// 補助: ログ書き込み（eae_ 命名規則で一本化）
// ════════════════════════════════════════════════════════════════

function eae_writeListingLog_(ss, taskId, sku, status, itemId, message) {
  var logSheet = ss.getSheetByName(EAE.SHEET_LOG);
  if (!logSheet) { return; }
  logSheet.appendRow([
    eae_nowJst_(), sku, taskId, '', '', status,
    (status === 'END_SUCCESS') ? 'TRUE' : 'FALSE',
    status, itemId, '', message.substring(0, 200)
  ]);
}

function eae_writeExecLog_(ss, funcName, status, detail, isDryRun) {
  var execSheet = ss.getSheetByName(EAE.SHEET_EXEC_LOG);
  if (!execSheet) { return; }
  execSheet.appendRow([eae_nowJst_(), funcName, isDryRun ? 'DRY_RUN' : '本番', status, detail]);
}

/** データ行取得（DATA_START_ROW=5行目から） */
function eae_getDataRows_(sheet) {
  var lastRow = sheet.getLastRow();
  if (lastRow < EAE.DATA_START_ROW) { return []; }
  var lastCol = Math.max(sheet.getLastColumn(), EAE.COL.END_MESSAGE);
  var numRows = lastRow - EAE.DATA_START_ROW + 1;
  var values  = sheet.getRange(EAE.DATA_START_ROW, 1, numRows, lastCol).getValues();
  var result  = [];
  values.forEach(function(row, idx) {
    if (String(row[EAE.COL.TASK_ID - 1] || '').trim() === '') { return; }
    result.push({ rowNum: idx + EAE.DATA_START_ROW, data: row });
  });
  return result;
}

function eae_isDryRun_() {
  var val = PropertiesService.getScriptProperties().getProperty(EAE.PROP_ENABLED);
  return (String(val || '').trim().toLowerCase() !== 'true');
}

function eae_getAccessToken_() {
  try {
    var props     = PropertiesService.getScriptProperties();
    var clientId  = props.getProperty('EBAY_CLIENT_ID');
    var clientSec = props.getProperty('EBAY_CLIENT_SECRET');
    var refreshTk = props.getProperty('EBAY_REFRESH_TOKEN');
    if (!clientId || !clientSec || !refreshTk) {
      Logger.log('✗ Script Properties未設定: EBAY_CLIENT_ID / EBAY_CLIENT_SECRET / EBAY_REFRESH_TOKEN');
      return null;
    }
    var resp = UrlFetchApp.fetch('https://api.ebay.com/identity/v1/oauth2/token', {
      method  : 'post',
      headers : {
        'Authorization' : 'Basic ' + Utilities.base64Encode(clientId + ':' + clientSec),
        'Content-Type'  : 'application/x-www-form-urlencoded',
      },
      payload : { grant_type: 'refresh_token', refresh_token: refreshTk },
      muteHttpExceptions: true,
    });
    var json = JSON.parse(resp.getContentText());
    return json.access_token || null;
  } catch (e) {
    Logger.log('✗ eae_getAccessToken_ 失敗: ' + e.message);
    return null;
  }
}

function eae_esc_(str) {
  return String(str)
    .replace(/&/g,'&amp;').replace(/</g,'&lt;')
    .replace(/>/g,'&gt;').replace(/"/g,'&quot;').replace(/'/g,'&apos;');
}

function eae_extractXml_(xml, tag) {
  var m = String(xml || '').match(new RegExp('<' + tag + '>([^<]*)</' + tag + '>'));
  return m ? m[1].trim() : '';
}

function eae_nowJst_() {
  return Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy/MM/dd HH:mm:ss');
}

// Replace the same-named functions in the main GAS project with these.

function checkEbayOrders() {
  const threads = GmailApp.search(
    '(from:ebay.com OR from:ebayshopicj@gmail.com) ' +
    '(subject:"You made the sale" OR subject:"Fwd: You made the sale") ' +
    'is:unread',
    0,
    20
  );

  if (threads.length === 0) {
    Logger.log('No new eBay orders');
    return;
  }

  threads.forEach(function(thread) {
    thread.getMessages().forEach(function(msg) {
      if (!msg.isUnread()) return;

      const order = parseEbayEmail(msg);
      if (!order || !order.orderId) {
        Logger.log('Invalid order email. Marking as read: ' + msg.getId());
        msg.markRead();
        return;
      }

      if (isOrderAlreadyLogged(order.orderId)) {
        Logger.log('Duplicate order: ' + order.orderId);
        msg.markRead();
        return;
      }

      const detail = getOrderDetail(order.orderId);

      // Keep the email unread so the next hourly run retries the API call.
      if (!detail || !detail.sku) {
        Logger.log(
          'Order detail/SKU unavailable. Keep unread for retry: ' +
          order.orderId
        );
        return;
      }

      order.phone = detail.phone || order.phone;
      order.phone2 = detail.phone2 || order.phone2;
      order.email = detail.email || order.email;
      order.sku = detail.sku;
      order.itemId = detail.itemId || order.itemId;
      order.buyerName = detail.buyerName || order.buyerName;

      if (!logOrder(order)) {
        Logger.log('Sheet write failed. Keep unread for retry: ' + order.orderId);
        return;
      }

      notifyOrderReceived(order);
      scheduleMcfReminder(order);
      msg.markRead();

      Logger.log('Order flow completed: ' + order.orderId);
    });
  });
}

function getOrderDetail(orderId) {
  const maxAttempts = 3;

  for (let attempt = 1; attempt <= maxAttempts; attempt++) {
    try {
      const props = PropertiesService.getScriptProperties();
      const creds = Utilities.base64Encode(
        props.getProperty('EBAY_CLIENT_ID') + ':' +
        props.getProperty('EBAY_CLIENT_SECRET')
      );

      const tokenRes = UrlFetchApp.fetch(
        'https://api.ebay.com/identity/v1/oauth2/token',
        {
          method: 'post',
          headers: {
            'Authorization': 'Basic ' + creds,
            'Content-Type': 'application/x-www-form-urlencoded'
          },
          payload:
            'grant_type=refresh_token&refresh_token=' +
            encodeURIComponent(props.getProperty('EBAY_REFRESH_TOKEN')),
          muteHttpExceptions: true
        }
      );

      if (tokenRes.getResponseCode() !== 200) {
        throw new Error(
          'eBay token HTTP ' + tokenRes.getResponseCode() + ': ' +
          tokenRes.getContentText().substring(0, 500)
        );
      }

      const accessToken = JSON.parse(tokenRes.getContentText()).access_token;
      if (!accessToken) throw new Error('eBay access token missing');

      const res = UrlFetchApp.fetch(
        'https://api.ebay.com/sell/fulfillment/v1/order/' +
        encodeURIComponent(orderId),
        {
          method: 'get',
          headers: {
            'Authorization': 'Bearer ' + accessToken,
            'Content-Type': 'application/json'
          },
          muteHttpExceptions: true
        }
      );

      if (res.getResponseCode() !== 200) {
        throw new Error(
          'eBay Fulfillment HTTP ' + res.getResponseCode() + ': ' +
          res.getContentText().substring(0, 500)
        );
      }

      const o = JSON.parse(res.getContentText());
      const result = {};
      const shipTo =
        o.fulfillmentStartInstructions &&
        o.fulfillmentStartInstructions[0] &&
        o.fulfillmentStartInstructions[0].shippingStep &&
        o.fulfillmentStartInstructions[0].shippingStep.shipTo;

      if (shipTo && shipTo.primaryPhone) {
        result.phone = shipTo.primaryPhone.phoneNumber;
      }

      if (o.buyer && o.buyer.buyerRegistrationAddress) {
        const buyer = o.buyer.buyerRegistrationAddress;
        result.email = buyer.email || '';
        result.buyerName = buyer.fullName || '';

        if (!result.phone && buyer.primaryPhone) {
          result.phone = buyer.primaryPhone.phoneNumber;
        }

        result.phone2 =
          buyer.secondaryPhone && buyer.secondaryPhone.phoneNumber || '';
      }

      if (o.lineItems && o.lineItems[0]) {
        const line = o.lineItems[0];
        const orderSku = line.sku || '';
        const itemId =
          line.legacyItemId || extractEbayItemId_(line.itemId) || '';

        result.itemId = itemId;
        result.sku = itemId
          ? getCurrentListingSkuByItemId_(itemId, accessToken) || orderSku
          : orderSku;
      }

      if (!result.sku) {
        throw new Error('SKU missing in eBay order detail');
      }

      Logger.log(
        'eBay detail success: order=' + orderId +
        ' sku=' + result.sku +
        ' attempt=' + attempt
      );

      return result;

    } catch (e) {
      Logger.log(
        'eBay detail attempt ' + attempt + '/' + maxAttempts +
        ' failed for ' + orderId + ': ' + e.message
      );

      if (attempt < maxAttempts) {
        Utilities.sleep(2000 * attempt);
      }
    }
  }

  return null;
}

function logOrder(order) {
  try {
    const ss = SpreadsheetApp.openById(SS_ID);
    let sheet = ss.getSheetByName(ORDER_SHEET);

    if (!sheet) {
      sheet = ss.insertSheet(ORDER_SHEET);
      const headers = [
        'Received at (JST)', 'eBay order ID', 'SKU', 'Item', 'Quantity',
        'Amount', 'Buyer', 'Street', 'City', 'State', 'ZIP', 'Phone',
        'Phone 2', 'Email', 'Ship by', 'MCF order', 'Status', 'Carrier',
        'Tracking', 'Tracking fetched at', 'Tracking check status'
      ];

      sheet.appendRow(headers);
      sheet.getRange(1, 1, 1, headers.length)
        .setFontWeight('bold')
        .setBackground('#1F4E79')
        .setFontColor('#FFFFFF');
      sheet.setFrozenRows(1);
    }

    const jst = new Date(order.receivedAt.getTime() + 9 * 60 * 60 * 1000);

    sheet.appendRow([
      Utilities.formatDate(jst, 'Asia/Tokyo', 'yyyy-MM-dd HH:mm'),
      order.orderId,
      order.sku,
      order.itemTitle,
      order.quantity,
      order.amount,
      order.buyerName,
      order.street,
      order.city,
      order.state,
      order.zip,
      order.phone || '',
      order.phone2 || '',
      order.email || '',
      order.shipBy,
      '\u672a\u4f5c\u6210',
      '\u53d7\u6ce8\u6e08',
      '', '', '', ''
    ]);

    Logger.log('Order logged: ' + order.orderId);
    return true;

  } catch (e) {
    Logger.log('Order log error: ' + e.message);
    return false;
  }
}

function getMcfSpApiAccessToken_() {
  const props = PropertiesService.getScriptProperties();

  try {
    const response = UrlFetchApp.fetch(
      'https://api.amazon.com/auth/o2/token',
      {
        method: 'post',
        contentType: 'application/x-www-form-urlencoded',
        payload: {
          grant_type: 'refresh_token',
          refresh_token: props.getProperty('SP_REFRESH_TOKEN'),
          client_id: props.getProperty('AMAZON_CLIENT_ID'),
          client_secret: props.getProperty('AMAZON_CLIENT_SECRET')
        },
        muteHttpExceptions: true
      }
    );

    const status = response.getResponseCode();
    const text = response.getContentText();

    if (status < 200 || status >= 300) {
      Logger.log('MCF token HTTP ' + status + ': ' + text.substring(0, 500));
      return null;
    }

    const json = JSON.parse(text);

    if (!json.access_token) {
      Logger.log('MCF access token missing');
      return null;
    }

    Logger.log('MCF SP-API token success');
    return json.access_token;

  } catch (e) {
    Logger.log('MCF token exception: ' + e.message);
    return null;
  }
}

function sampleCsvProps_() {
  return PropertiesService.getScriptProperties();
}

/* ===================================================================
 * v7 追加ブロック（安全版）  ファイルの「一番下」に貼り付けるだけ
 * -------------------------------------------------------------------
 * ・既存の AUTO_ 定数（AUTO_SHEET_STOCK_SYNC 等）を再利用します。
 *   ここでは再定義しません（重複宣言エラー回避）。
 * ・SKU正規化は専用の _autoNormSku_ を使用（_normSku_ の有無に依存しない）。
 * ・在庫0変更には一切関与しません（D列の更新のみ）。
 * =================================================================== */

var FBA_PASTE_SS_ID        = '1zdEo_spDLplWVc2ddNUL5udeN8zeQcvZb9xIKqHw0dk';
var FBA_PASTE_SHEET_NAME   = 'FBA在庫_貼付_';
var FBA_PASTE_SHEET_GID    = 776037102;
var FBA_PASTE_HEADER_SKU   = 'sellersku';
var FBA_PASTE_HEADER_QTY   = 'fulfillablequantity';
var FBA_PASTE_HEADER_SCAN_ROWS = 10;
var FBA_PASTE_MARK_UNMATCHED   = true;


function updateStockSyncDFromFbaPaste_() {
  var srcSs = SpreadsheetApp.openById(FBA_PASTE_SS_ID);
  var srcSheet = srcSs.getSheetByName(FBA_PASTE_SHEET_NAME);
  if (!srcSheet) {
    var sheets = srcSs.getSheets();
    for (var s = 0; s < sheets.length; s++) {
      if (sheets[s].getSheetId() === FBA_PASTE_SHEET_GID) { srcSheet = sheets[s]; break; }
    }
  }
  if (!srcSheet) {
    throw new Error('FBA paste sheet not found (D column kept)');
  }

  var srcValues = srcSheet.getDataRange().getValues();
  if (!srcValues || srcValues.length < 2) {
    throw new Error('FBA paste sheet has no data rows (D column kept)');
  }

  var headerRowIdx = -1, skuCol = -1, qtyCol = -1;
  var scanMax = Math.min(FBA_PASTE_HEADER_SCAN_ROWS, srcValues.length);
  for (var r = 0; r < scanMax; r++) {
    var hdr = srcValues[r].map(function (x) {
      return String(x == null ? '' : x).trim().toLowerCase();
    });
    var si = hdr.indexOf(FBA_PASTE_HEADER_SKU);
    var qi = hdr.indexOf(FBA_PASTE_HEADER_QTY);
    if (si !== -1 && qi !== -1) { headerRowIdx = r; skuCol = si; qtyCol = qi; break; }
  }
  if (headerRowIdx === -1) {
    throw new Error('FBA paste header mismatch (D column kept)');
  }

  var qtyMap = {};
  for (var i = headerRowIdx + 1; i < srcValues.length; i++) {
    var nsku = _autoNormSku_(srcValues[i][skuCol]);
    if (!nsku) continue;
    var q = Number(srcValues[i][qtyCol]);
    qtyMap[nsku] = isNaN(q) ? '' : q;
  }

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(AUTO_SHEET_STOCK_SYNC);
  if (!sheet) throw new Error('Sheet not found: ' + AUTO_SHEET_STOCK_SYNC);

  var lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    Logger.log('Stock sync sheet has no data rows');
    return { matched: 0, unmatched: 0, processed: 0 };
  }

  var numRows = lastRow - 1;
  var skuColVals = sheet.getRange(2, AUTO_COL_SKU, numRows, 1).getValues();
  var dCol  = sheet.getRange(2, AUTO_COL_AMZ_QTY, numRows, 1).getValues();
  var hiCol = sheet.getRange(2, AUTO_COL_LASTCHECK, numRows, 2).getValues();

  var now = _autoNowStr_();
  var matched = 0, unmatched = 0, processed = 0;

  for (var k = 0; k < numRows; k++) {
    var sku = String(skuColVals[k][0] == null ? '' : skuColVals[k][0]).trim();
    if (!sku) continue;
    if (AUTO_SKU_EXCLUDE.indexOf(sku.toUpperCase()) !== -1) continue;
    if (_autoIsTestSku_(sku)) continue;

    processed++;
    var nsku = _autoNormSku_(sku);

    if (qtyMap.hasOwnProperty(nsku)) {
      dCol[k][0]  = qtyMap[nsku];
      hiCol[k][0] = now;
      hiCol[k][1] = 'FBA貼付取得済';
      matched++;
    } else {
      if (FBA_PASTE_MARK_UNMATCHED) {
        hiCol[k][0] = now;
        hiCol[k][1] = 'FBA貼付未発見';
      }
      unmatched++;
    }
  }

  sheet.getRange(2, AUTO_COL_AMZ_QTY, numRows, 1).setValues(dCol);
  sheet.getRange(2, AUTO_COL_LASTCHECK, numRows, 2).setValues(hiCol);

  Logger.log('FBA paste to D column: matched=' + matched + ' unmatched=' + unmatched + ' processed=' + processed);
  return { matched: matched, unmatched: unmatched, processed: processed };
}


function dryRunFbaPasteLookup_RUN() {
  var srcSs = SpreadsheetApp.openById(FBA_PASTE_SS_ID);
  var srcSheet = srcSs.getSheetByName(FBA_PASTE_SHEET_NAME);
  if (!srcSheet) {
    var sheets = srcSs.getSheets();
    for (var s = 0; s < sheets.length; s++) {
      if (sheets[s].getSheetId() === FBA_PASTE_SHEET_GID) { srcSheet = sheets[s]; break; }
    }
  }
  if (!srcSheet) { Logger.log('NG: source sheet not found'); return; }

  var srcValues = srcSheet.getDataRange().getValues();
  var headerRowIdx = -1, skuCol = -1, qtyCol = -1;
  var scanMax = Math.min(FBA_PASTE_HEADER_SCAN_ROWS, srcValues.length);
  for (var r = 0; r < scanMax; r++) {
    var hdr = srcValues[r].map(function (x) { return String(x == null ? '' : x).trim().toLowerCase(); });
    var si = hdr.indexOf(FBA_PASTE_HEADER_SKU);
    var qi = hdr.indexOf(FBA_PASTE_HEADER_QTY);
    if (si !== -1 && qi !== -1) { headerRowIdx = r; skuCol = si; qtyCol = qi; break; }
  }
  if (headerRowIdx === -1) { Logger.log('NG: header mismatch'); return; }

  var qtyMap = {};
  for (var i = headerRowIdx + 1; i < srcValues.length; i++) {
    var nsku = _autoNormSku_(srcValues[i][skuCol]);
    if (!nsku) continue;
    qtyMap[nsku] = srcValues[i][qtyCol];
  }

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(AUTO_SHEET_STOCK_SYNC);
  var lastRow = sheet.getLastRow();
  var numRows = lastRow - 1;
  var skuColVals = sheet.getRange(2, AUTO_COL_SKU, numRows, 1).getValues();

  var matched = 0, unmatched = 0, processed = 0, samples = [];
  for (var k = 0; k < numRows; k++) {
    var sku = String(skuColVals[k][0] == null ? '' : skuColVals[k][0]).trim();
    if (!sku) continue;
    if (AUTO_SKU_EXCLUDE.indexOf(sku.toUpperCase()) !== -1) continue;
    if (_autoIsTestSku_(sku)) continue;
    processed++;
    var nsku = _autoNormSku_(sku);
    if (qtyMap.hasOwnProperty(nsku)) {
      matched++;
      if (samples.length < 10) samples.push(sku + ' -> ' + qtyMap[nsku]);
    } else {
      unmatched++;
    }
  }

  Logger.log('=== DRY-RUN (no write) ===');
  Logger.log('source headerRow=' + (headerRowIdx + 1) + ' skuCol=' + (skuCol + 1) + ' qtyCol=' + (qtyCol + 1));
  Logger.log('processed=' + processed + ' matched=' + matched + ' unmatched=' + unmatched);
  Logger.log('--- matched samples (max 10) ---\n' + (samples.length ? samples.join('\n') : '(none)'));
}


function _autoNormSku_(sku) {
  if (typeof normalizeSku_ === 'function') {
    try { return normalizeSku_(sku); } catch (e) {}
  }
  return String(sku == null ? '' : sku).trim();
}

function applyFbaPasteToDColumn_RUN() {
  var r = updateStockSyncDFromFbaPaste_();
  Logger.log('result: ' + JSON.stringify(r));
}

function _autoNotifyOnce_(key, signatureText, message) {
  var props = PropertiesService.getScriptProperties();
  var k = 'LAST_NOTIFY_' + key;
  var sig = Utilities.base64Encode(
    Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, signatureText || '')
  );
  if (props.getProperty(k) === sig) { Logger.log('同一内容→通知スキップ:' + key); return false; }
  sendLine(message);
  props.setProperty(k, sig);
  return true;
}
function _autoClearNotify_(key) {
  PropertiesService.getScriptProperties().deleteProperty('LAST_NOTIFY_' + key);
}

/* ===================================================================
 * Q3 Stage1：在庫承認 候補検出（eBay書き込みなし・安全）
 * -------------------------------------------------------------------
 * ・在庫同期管理から候補を検出し、「eBay在庫承認」シートに書き出すだけ。
 * ・eBayには一切触りません（検出のみ）。
 * ・実行は detectInventoryApprovalCandidates_RUN を選んで実行。
 *
 * 候補条件（あとで調整可）：
 *   0→1（出品再開）：eBay在庫=0 かつ Amazon在庫 >= 基準数（基準が空なら Amazon>=1）
 *   1→0（在庫切れ）：eBay在庫>=1 かつ 基準数>=1 かつ Amazon在庫 < 基準数
 *
 * 使う列（在庫同期管理）：A=SKU, D=Amazon在庫, E=eBay設定在庫, F=通知基準数
 * ItemID は eBayRaw から SKU で引く（ヘッダー自動判別）。
 * =================================================================== */

var APPROVAL_SHEET_NAME   = 'eBay在庫承認';
var APPROVAL_SRC_STOCK    = '在庫同期管理';
var APPROVAL_SRC_EBAYRAW  = 'eBayRaw';

// 在庫同期管理の列（1始まり）
var APV_COL_SKU       = 1; // A
var APV_COL_AMZ       = 4; // D Amazon在庫
var APV_COL_EBAY      = 5; // E eBay設定在庫
var APV_COL_THRESHOLD = 6; // F 通知基準数

// 除外
var APV_SKU_EXCLUDE = ['SKU', 'WAREHOUSES', 'DELETED', 'ITEMID', 'TITLE', 'AVAILABLEQUANTITY'];
var APV_TEST_TOKENS = ['SAMPLE', 'TEST', 'DUMMY'];

// eBayRaw のヘッダー候補（小文字照合）
var APV_RAW_HEADER_SKU    = ['sku', 'customlabel', 'custom label'];
var APV_RAW_HEADER_ITEMID = ['itemid', 'item id', 'item number'];


/**
 * 候補を検出して「eBay在庫承認」シートに書き出す（eBay書き込みなし）
 */
function detectInventoryApprovalCandidates_RUN() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  // ItemIDマップ（eBayRaw から SKU->ItemID）
  var itemIdMap = _apvBuildItemIdMap_(ss);

  // 在庫同期管理 読み込み
  var stock = ss.getSheetByName(APPROVAL_SRC_STOCK);
  if (!stock) { Logger.log('NG: シートなし ' + APPROVAL_SRC_STOCK); return; }
  var lastRow = stock.getLastRow();
  if (lastRow < 2) { Logger.log('在庫同期管理にデータなし'); return; }

  var numRows = lastRow - 1;
  var data = stock.getRange(2, 1, numRows, 6).getValues();

  var now = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy/MM/dd HH:mm');
  var candidates = [];
  var to1 = 0, to0 = 0, noItemId = 0;

  for (var i = 0; i < data.length; i++) {
    var row = data[i];
    var sku = String(row[APV_COL_SKU - 1] == null ? '' : row[APV_COL_SKU - 1]).trim();
    if (!sku) continue;
    if (APV_SKU_EXCLUDE.indexOf(sku.toUpperCase()) !== -1) continue;
    if (_apvIsTestSku_(sku)) continue;

    var amz = _apvNum_(row[APV_COL_AMZ - 1]);
    var ebay = _apvNum_(row[APV_COL_EBAY - 1]);
    var th = _apvNum_(row[APV_COL_THRESHOLD - 1]);
    if (isNaN(amz) || isNaN(ebay)) continue;

    var direction = '', target = null;

    // 0->1: eBay=0 かつ Amazon在庫が基準以上
    if (ebay === 0) {
      var ok01 = (!isNaN(th) && th >= 1) ? (amz >= th) : (amz >= 1);
      if (ok01) { direction = '0→1'; target = 1; }
    }
    // 1->0: eBay>=1 かつ 基準>=1 かつ Amazon<基準
    if (!direction && ebay >= 1 && !isNaN(th) && th >= 1 && amz < th) {
      direction = '1→0'; target = 0;
    }

    if (!direction) continue;

    var nsku = _apvNorm_(sku);
    var itemId = itemIdMap.hasOwnProperty(nsku) ? itemIdMap[nsku] : '';
    if (!itemId) noItemId++;
    if (direction === '0→1') to1++; else to0++;

    candidates.push([
      now, sku, itemId, direction, amz, ebay, target,
      '承認待ち', '', '', (itemId ? '' : 'ItemID未取得')
    ]);
  }

  // 承認シートへ書き出し（ヘッダー＋本体を作り直す）
  var sheet = ss.getSheetByName(APPROVAL_SHEET_NAME);
  if (!sheet) sheet = ss.insertSheet(APPROVAL_SHEET_NAME);
  sheet.clearContents();

  var header = ['検出日時','SKU','ItemID','方向','Amazon在庫','eBay現在','予定値','状態','実行日時','結果','エラー'];
  sheet.getRange(1, 1, 1, header.length).setValues([header]);

  if (candidates.length > 0) {
    sheet.getRange(2, 1, candidates.length, header.length).setValues(candidates);
  }

  Logger.log('=== 在庫承認 候補検出（eBay書き込みなし）===');
  Logger.log('0→1（出品再開）: ' + to1 + '件');
  Logger.log('1→0（在庫切れ）: ' + to0 + '件');
  Logger.log('ItemID未取得: ' + noItemId + '件');
  Logger.log('「' + APPROVAL_SHEET_NAME + '」シートに ' + candidates.length + '件 書き出しました。');
}


/** eBayRaw から SKU->ItemID のマップを作る（ヘッダー自動判別） */
function _apvBuildItemIdMap_(ss) {
  var map = {};
  var raw = ss.getSheetByName(APPROVAL_SRC_EBAYRAW);
  if (!raw) { Logger.log('注意: eBayRawなし。ItemIDは空。'); return map; }

  var values = raw.getDataRange().getValues();
  if (!values || values.length < 2) return map;

  var hdr = values[0].map(function (x) { return String(x == null ? '' : x).trim().toLowerCase(); });
  var skuCol = _apvFindCol_(hdr, APV_RAW_HEADER_SKU);
  var idCol = _apvFindCol_(hdr, APV_RAW_HEADER_ITEMID);

  if (skuCol === -1 || idCol === -1) {
    Logger.log('注意: eBayRawのSKU/ItemID列が不明 (sku=' + skuCol + ', id=' + idCol + ')。ItemIDは空。');
    return map;
  }

  for (var i = 1; i < values.length; i++) {
    var nsku = _apvNorm_(values[i][skuCol]);
    if (!nsku) continue;
    var id = String(values[i][idCol] == null ? '' : values[i][idCol]).trim();
    if (id && !map.hasOwnProperty(nsku)) map[nsku] = id;
  }
  return map;
}

function _apvFindCol_(hdrLower, candidates) {
  for (var c = 0; c < candidates.length; c++) {
    var idx = hdrLower.indexOf(candidates[c]);
    if (idx !== -1) return idx;
  }
  return -1;
}

function _apvNorm_(sku) {
  if (typeof normalizeSku_ === 'function') {
    try { return normalizeSku_(sku); } catch (e) {}
  }
  return String(sku == null ? '' : sku).trim();
}

function _apvIsTestSku_(sku) {
  var u = String(sku == null ? '' : sku).toUpperCase();
  for (var i = 0; i < APV_TEST_TOKENS.length; i++) {
    if (u.indexOf(APV_TEST_TOKENS[i]) !== -1) return true;
  }
  return false;
}

function _apvNum_(v) {
  if (v === '' || v === null || v === undefined) return NaN;
  var n = Number(v);
  return isNaN(n) ? NaN : n;
}

/* ===================================================================
 * Q3 Stage2：在庫承認 実行部品（eBay書き込み・安全装置付き）
 * -------------------------------------------------------------------
 * 「eBay在庫承認」シートの 状態='承認済' の行だけを実行する。
 *
 * ★安全装置（必ず確認）★
 *   INVENTORY_EXEC_DRYRUN  = true  → プレビューのみ。eBayに一切触れない。
 *   INVENTORY_EXEC_ENABLED = false → 本番実行しない（DRYRUN=falseでも安全側で停止）
 *   INVENTORY_EXEC_LIMIT   = 5     → 1回の実行上限
 *
 *   ＝ 実際にeBayを変えるのは「DRYRUN=false かつ ENABLED=true」の時だけ。
 *
 * 実行：executeApprovedInventory_RUN を選んで実行
 * 結果は「eBay在庫承認」シートの H/I/J/K に書き戻す。
 * =================================================================== */

var INVENTORY_EXEC_DRYRUN  = true;   // ★まずは true（プレビュー）
var INVENTORY_EXEC_ENABLED = false;  // ★まずは false（本番停止）
var INVENTORY_EXEC_LIMIT   = 5;      // 1回の上限

// 承認シートの列（Stage1と同じ）
var APV_C_DATE   = 1;  // A 検出日時
var APV_C_SKU    = 2;  // B SKU
var APV_C_ITEMID = 3;  // C ItemID
var APV_C_DIR    = 4;  // D 方向
var APV_C_AMZ    = 5;  // E Amazon在庫
var APV_C_EBAY   = 6;  // F eBay現在
var APV_C_TARGET = 7;  // G 予定値
var APV_C_STATUS = 8;  // H 状態
var APV_C_EXECAT = 9;  // I 実行日時
var APV_C_RESULT = 10; // J 結果
var APV_C_ERROR  = 11; // K エラー


/**
 * 状態='承認済' の行を実行（DRY-RUN/本番は上のフラグで切替）
 */
function executeApprovedInventory_RUN() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(APPROVAL_SHEET_NAME); // Stage1の定数を流用
  if (!sheet) { Logger.log('NG: シートなし ' + APPROVAL_SHEET_NAME); return; }

  var lastRow = sheet.getLastRow();
  if (lastRow < 2) { Logger.log('承認シートにデータなし'); return; }

  var numRows = lastRow - 1;
  var data = sheet.getRange(2, 1, numRows, 11).getValues();

  // 対象＝状態が「承認済」かつ ItemIDあり
  var targets = [];
  for (var i = 0; i < data.length; i++) {
    var status = String(data[i][APV_C_STATUS - 1] == null ? '' : data[i][APV_C_STATUS - 1]).trim();
    var itemId = String(data[i][APV_C_ITEMID - 1] == null ? '' : data[i][APV_C_ITEMID - 1]).trim();
    if (status === '承認済' && itemId) {
      targets.push({ rowIndex: i + 2, data: data[i] });
    }
  }

  if (targets.length === 0) {
    Logger.log('実行対象（承認済）がありません。');
    Logger.log('※テスト時は、承認シートのH列を手動で「承認済」にしてから実行してください。');
    return;
  }

  // 件数上限
  var willProcess = targets.slice(0, INVENTORY_EXEC_LIMIT);

  Logger.log('=== 在庫承認 実行 ===');
  Logger.log('DRY-RUN=' + INVENTORY_EXEC_DRYRUN + ' / ENABLED=' + INVENTORY_EXEC_ENABLED
    + ' / 対象=' + targets.length + '件（上限' + INVENTORY_EXEC_LIMIT + 'で' + willProcess.length + '件処理）');

  // 本番実行する条件
  var doRealUpdate = (INVENTORY_EXEC_DRYRUN === false && INVENTORY_EXEC_ENABLED === true);

  var token = '';
  if (doRealUpdate) {
    token = _apvGetEbayToken_();
    if (!token) {
      Logger.log('NG: eBayトークン取得失敗。実行中止（シートは変更しません）。');
      return;
    }
  }

  var now = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy/MM/dd HH:mm');
  var okCount = 0, ngCount = 0;

  for (var t = 0; t < willProcess.length; t++) {
    var r = willProcess[t];
    var sku = String(r.data[APV_C_SKU - 1]).trim();
    var itemId = String(r.data[APV_C_ITEMID - 1]).trim();
    var dir = String(r.data[APV_C_DIR - 1]).trim();
    var target = _apvNum_(r.data[APV_C_TARGET - 1]);
    if (isNaN(target)) target = (dir === '0→1') ? 1 : 0;

    if (!doRealUpdate) {
      // DRY-RUN：ログのみ。シートは変えない。
      Logger.log('[DRY-RUN] ' + sku + ' / ' + dir + ' / ItemID=' + itemId
        + ' → eBay数量を ' + target + ' にする（予定）');
      continue;
    }

    // 本番：eBay更新
    var res = reviseEbayQuantity_(itemId, target, token);
    if (res && res.success) {
      sheet.getRange(r.rowIndex, APV_C_STATUS).setValue('実行済');
      sheet.getRange(r.rowIndex, APV_C_EXECAT).setValue(now);
      sheet.getRange(r.rowIndex, APV_C_RESULT).setValue('成功');
      sheet.getRange(r.rowIndex, APV_C_ERROR).setValue('');
      okCount++;
      Logger.log('OK ' + sku + ' / ' + dir + ' → ' + target);
    } else {
      sheet.getRange(r.rowIndex, APV_C_STATUS).setValue('失敗');
      sheet.getRange(r.rowIndex, APV_C_EXECAT).setValue(now);
      sheet.getRange(r.rowIndex, APV_C_RESULT).setValue('失敗');
      sheet.getRange(r.rowIndex, APV_C_ERROR).setValue(res ? String(res.message || res.ack || 'error').substring(0, 200) : 'no response');
      ngCount++;
      Logger.log('NG ' + sku + ' / ' + dir + ' → ' + (res ? (res.message || res.ack) : 'no response'));
    }
    Utilities.sleep(500); // eBay負荷軽減
  }

  if (doRealUpdate) {
    Logger.log('実行完了：成功' + okCount + ' / 失敗' + ngCount);
  } else {
    Logger.log('DRY-RUN完了：上記は「予定」です。eBayは変更していません。');
    Logger.log('本番実行するには INVENTORY_EXEC_DRYRUN=false かつ INVENTORY_EXEC_ENABLED=true に変更。');
  }
}


/**
 * eBay数量を指定値に変更（ReviseFixedPriceItem）。0でも1でも可。
 * 既存 reviseEbayAvailableToZero_ の数量を引数化したもの。
 */
function reviseEbayQuantity_(itemId, qty, accessToken) {
  var xml =
    '<?xml version="1.0" encoding="UTF-8"?>' +
    '<ReviseFixedPriceItemRequest xmlns="urn:ebay:apis:eBLBaseComponents">' +
      '<Item>' +
        '<ItemID>' + itemId + '</ItemID>' +
        '<Quantity>' + qty + '</Quantity>' +
      '</Item>' +
    '</ReviseFixedPriceItemRequest>';

  try {
    var response = UrlFetchApp.fetch('https://api.ebay.com/ws/api.dll', {
      method: 'post',
      headers: {
        'X-EBAY-API-IAF-TOKEN': accessToken,
        'X-EBAY-API-CALL-NAME': 'ReviseFixedPriceItem',
        'X-EBAY-API-SITEID': '0',
        'X-EBAY-API-COMPATIBILITY-LEVEL': '1199',
        'Content-Type': 'text/xml'
      },
      payload: xml,
      muteHttpExceptions: true
    });

    var httpCode = response.getResponseCode();
    var body = response.getContentText();
    var ack = (body.match(/<Ack>([^<]+)<\/Ack>/) || [])[1] || '';
    var shortMessage = (body.match(/<ShortMessage>([\s\S]*?)<\/ShortMessage>/) || [])[1] || '';

    Logger.log('ReviseFixedPriceItem [' + httpCode + '] ItemID=' + itemId + ' qty=' + qty + ' Ack=' + ack
      + (shortMessage ? (' / ' + shortMessage) : ''));

    return {
      success: (httpCode === 200 && (ack === 'Success' || ack === 'Warning')),
      httpCode: httpCode,
      ack: ack,
      message: shortMessage
    };
  } catch (err) {
    return { success: false, httpCode: 0, ack: '', message: String(err) };
  }
}


/** eBayアクセストークン取得（既存関数を優先利用） */
function _apvGetEbayToken_() {
  try {
    if (typeof getEbayAccessTokenForFbaZeroAuto_ === 'function') {
      return getEbayAccessTokenForFbaZeroAuto_();
    }
    if (typeof getEbayAccessTokenFromRefreshToken === 'function') {
      return getEbayAccessTokenFromRefreshToken();
    }
  } catch (e) {
    Logger.log('トークン取得エラー: ' + e);
  }
  return '';
}

/* ===================================================================
 * Q3 Stage3：1日1回 LINE通知（候補件数）  ※eBay書き込みなし
 * -------------------------------------------------------------------
 * 1日1回、候補を検出して「eBay在庫承認」シートを更新し、
 * 承認待ちが1件以上あれば LINE を1通だけ送る。
 *
 * 実行：dailyInventoryApprovalNotify（手動テスト可）
 * トリガー設定：setupDailyInventoryApprovalTrigger を1回実行（毎日8時頃）
 *
 * 既存を流用：detectInventoryApprovalCandidates_RUN（Stage1）/ sendLine
 * =================================================================== */

function dailyInventoryApprovalNotify() {
  // 1) 候補を検出して「eBay在庫承認」シートを更新（eBayには触れない）
  detectInventoryApprovalCandidates_RUN();

  // 2) シートから「承認待ち」を方向別に集計
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(APPROVAL_SHEET_NAME);
  if (!sheet) { Logger.log('NG: シートなし ' + APPROVAL_SHEET_NAME); return; }

  var lastRow = sheet.getLastRow();
  if (lastRow < 2) { Logger.log('候補なし → LINE送信せず'); return; }

  var data = sheet.getRange(2, 1, lastRow - 1, 8).getValues();
  var to1 = 0, to0 = 0;
  for (var i = 0; i < data.length; i++) {
    var status = String(data[i][7] == null ? '' : data[i][7]).trim(); // H 状態
    if (status !== '承認待ち') continue;
    var dir = String(data[i][3] == null ? '' : data[i][3]).trim();     // D 方向
    if (dir === '0→1') to1++;
    else if (dir === '1→0') to0++;
  }

  // 3) 承認待ちが無ければ送らない（LINE節約）
  if (to1 === 0 && to0 === 0) {
    Logger.log('承認待ち候補なし → LINE送信せず');
    return;
  }

  // 4) LINEを1通だけ送る
  var msg =
    '📦 eBay在庫 承認待ち\n' +
    '0→1（出品再開）: ' + to1 + '件\n' +
    '1→0（在庫切れ）: ' + to0 + '件\n\n' +
    '「進めてください」で実行\n' +
    '「停めてください」で破棄\n\n' +
    '（詳細は「' + APPROVAL_SHEET_NAME + '」シート）';

  sendLine(msg);
  Logger.log('LINE通知送信: 0→1=' + to1 + ' / 1→0=' + to0);
}


/** 毎日8時頃に dailyInventoryApprovalNotify を動かすトリガーを作る（1回だけ実行）*/
function setupDailyInventoryApprovalTrigger() {
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === 'dailyInventoryApprovalNotify') {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }
  ScriptApp.newTrigger('dailyInventoryApprovalNotify')
    .timeBased()
    .everyDays(1)
    .atHour(8)
    .nearMinute(0)
    .inTimezone('Asia/Tokyo')
    .create();

  Logger.log('日次トリガー設定完了: dailyInventoryApprovalNotify 毎日8時頃');
}

//テスト前
//テスト後