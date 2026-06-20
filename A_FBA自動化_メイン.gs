/**
 * Ship&Co LINE Bot - v12
 * 送料確認 ＋ 賞味期限OCR ＋ JP向け荷姿登録 ＋ 返信文簡略化
 *
 * v12 修正点:
 *   [JP向け]
 *   - 国コードが JP の場合は Ship&Co 送料計算をしない
 *   - JP の場合は FBA送料一覧シートへ送料記録しない
 *   - JP の場合は 発送G に管理番号グループ・荷姿のみ登録する
 *   - JP食品あり:
 *       初回返信は「賞味期限画像を送ってください」
 *       賞味期限OK後、通常成功時は「登録完了」
 *   - JP食品なし:
 *       初回返信は「登録完了」
 *
 *   [USなど海外向け]
 *   - 送料候補メッセージの最後に
 *       「本明に、転送してください」
 *     を追加
 *   - 送料番号選択後、通常成功時は「登録完了」
 *   - 賞味期限OK後、通常成功時は「登録完了」
 *
 * スクリプトプロパティ（必須）:
 *   LINE_TOKEN, SHIPANDCO_TOKEN, CLAUDE_API_KEY, SHIN_USER_ID
 */

var SORYO_SHEET_ID = '1xJ8PiYwv_T_FVLEXBUi1Fs8MYqAav7CjknzqieLhbkg';  // FBA送料一覧
var HASSO_SHEET_ID = '1l6Vs-5el4-N3xe0msWoqOTWAg6J_RAARtvVJF73VdsE'; // そー草加ファイル
var HASSO_SHEET_NAME = '発送G';
var EXPIRY_UPLOAD_ENABLED = false;  // false=賞味期限OCR停止 / true=再開

// ============================================================
// Webhook受信
// ============================================================

function doGet(e) {
  return ContentService.createTextOutput('OK');
}

function doPost(e) {
  var output = ContentService.createTextOutput('OK');

  try {
    var body = JSON.parse(e.postData.contents);
    var events = body.events;
    if (!events || events.length === 0) return output;

    var event = events[0];

    var senderUserId = (event.source && event.source.userId) ? event.source.userId : '';
    console.log('senderUserId: ' + senderUserId);

    if (event.type !== 'message') return output;

    // ===== 画像メッセージ → 賞味期限OCR =====
    if (event.message.type === 'image') {
      if (!EXPIRY_UPLOAD_ENABLED) {
        replyToLine(event.replyToken, '賞味期限の読み取りは現在停止中です。');
        return output;
      }
      console.log('→ 賞味期限画像を受信');
      handleExpiryImage(event.message.id, event.replyToken);
      return output;
    }

    // ===== テキストメッセージ =====
    if (event.message.type !== 'text') return output;

    var replyToken = event.replyToken;
    var userMessage = event.message.text.trim();

    console.log('受信メッセージ（正規化前）: ' + userMessage);

    // タブと連続スペースは詰めるが、改行は保持する
    userMessage = userMessage
      .replace(/\t+/g, ' ')
      .replace(/[ \u3000]+/g, ' ')
      .replace(/×\s+/g, '×')
      .replace(/[ ]*\n[ ]*/g, '\n')
      .trim();

    console.log('受信メッセージ（正規化後）: ' + userMessage);

    // 賞味期限の確認OK（書き込み実行）
    if (
      userMessage === 'OK' ||
      userMessage === 'ok' ||
      userMessage === 'Ok' ||
      userMessage === 'オーケー' ||
      userMessage === 'おーけー'
    ) {
      confirmExpiryWrite(replyToken);
      return output;
    }

    // 賞味期限のキャンセル
    if (userMessage === '賞味期限キャンセル' || userMessage === 'NG' || userMessage === 'ng') {
      PropertiesService.getScriptProperties().deleteProperty('expiry_context');
      replyToLine(replyToken, '賞味期限の登録をキャンセルしました。もう一度シートを撮影して送ってください。');
      return output;
    }

    if (userMessage === 'ping') {
      replyToLine(replyToken, 'pong v12 動作中');
      return output;
    }

    if (userMessage === 'myid') {
      PropertiesService.getScriptProperties().setProperty('SHIN_USER_ID', senderUserId);
      replyToLine(replyToken, 'あなたのID:\n' + senderUserId + '\n\nSHIN_USER_IDに登録しました');
      return output;
    }

    if (userMessage.startsWith('削除 ')) {
      deleteFromSoryoSheet(userMessage.replace('削除 ', '').trim(), replyToken);
      return output;
    }

    if (userMessage.startsWith('確認 ')) {
      checkSoryoSheet(userMessage.replace('確認 ', '').trim(), replyToken);
      return output;
    }

    if (/^\d+$/.test(userMessage)) {
      handleSoryoReply(userMessage, replyToken);
      return output;
    }

    if (userMessage === 'help' || userMessage === 'ヘルプ') {
      replyToLine(replyToken, getHelpMessage());
      return output;
    }

    // ========================================================
    // 通常の荷姿情報解析
    // ========================================================

    console.log('→ Claude解析開始');

// ===== 在庫承認（進めて / 停めて）=====
    if (userMessage === '進めてください' || userMessage === '進めて') {
      handleInventoryApproval_('進める', replyToken);
      return output;
    }
    if (userMessage === '停めてください' || userMessage === '停めて' ||
        userMessage === '止めてください' || userMessage === '止めて') {
      handleInventoryApproval_('停める', replyToken);
      return output;
    }

    var dims = parseDimensionsWithClaude(userMessage);

    if (!dims) {
      replyToLine(
        replyToken,
        '❌ サイズ情報を認識できませんでした。\n\n例:\n「14525 60×40×40 14350g US」\n「14111 14112 14113 60×40×40 14350g JP」\n\n「help」でヘルプ表示'
      );
      return output;
    }

    // この箱の管理番号リストを作る
    var allKanriNumbers = dims.order_id;
    var kanriNums = [];

    if (dims.order_id) {
      var firstLine = userMessage.split('\n')[0];
      var tokens = firstLine.split(/\s+/);

      for (var i = 0; i < tokens.length; i++) {
        if (/^\d{5}$/.test(tokens[i])) {
          kanriNums.push(tokens[i]);
        } else {
          break;
        }
      }

      allKanriNumbers = kanriNums.length > 0 ? kanriNums.join(' ') : dims.order_id;
    }

    var dimForSheet = {
      length: dims.length,
      width: dims.width,
      height: dims.height,
      weight: dims.weight / 1000
    };

    var countryCode = String(dims.country || '').trim().toUpperCase();
    console.log('判定用 countryCode: [' + countryCode + ']');

    // ========================================================
    // JP向け：送料計算なし・FBA送料一覧記録なし・発送Gだけ登録
    // ========================================================

    if (countryCode === 'JP') {
      var jpProps = PropertiesService.getScriptProperties();

      // JPでは送料選択を使わないため、前回の送料選択コンテキストを消す
      jpProps.deleteProperty('soryo_context');

      // 発送Gへ管理番号グループ・荷姿を書き込み
      var jpHassoWarning = writeToHassoG(allKanriNumbers, dimForSheet);

      if (dims.has_food) {
        // 賞味期限OCR時の照合用に、この箱の正しい管理番号リストを保存
        jpProps.setProperty(
          'expected_kanri',
          JSON.stringify({
            list: kanriNums,
            country: 'JP',
            timestamp: new Date().toISOString()
          })
        );

        // 通常成功時は、倉庫担当へ余計な説明を出さない
        if (jpHassoWarning) {
          replyToLine(replyToken, jpHassoWarning + '\n\n賞味期限画像を送ってください');
        } else {
          replyToLine(replyToken, '賞味期限画像を送ってください');
        }

      } else {
        // 食品なしでは賞味期限照合リストを残さない
        jpProps.deleteProperty('expected_kanri');

        // 通常成功時は「登録完了」のみ
        if (jpHassoWarning) {
          replyToLine(replyToken, jpHassoWarning);
        } else {
          replyToLine(replyToken, '登録完了');
        }
      }

      return output;
    }

    // ========================================================
    // USなど海外向け：従来通り送料計算・送料選択待ち
    // ========================================================

    var ratesResult = getShipAndCoRates(dims);
    var ratesMessage = formatRatesMessage(ratesResult, dims);

    if (dims.order_id) {
    saveSoryoContext(dims.order_id, ratesMessage, dimForSheet, allKanriNumbers, dims.has_food);
    }

    // 食品ありなら、送料メッセージに賞味期限の案内を追記
    var replyMsg = ratesMessage;

    if (dims.has_food) {
      var foodList = kanriNums.length > 0 ? kanriNums.join(' ') : allKanriNumbers;

      replyMsg += '\n─────────────────\n🍱 食品が含まれます\n対象管理番号: ' + foodList
        + '\n→ 賞味期限記入シートに上記を記入し、撮影して送ってください。';

      PropertiesService.getScriptProperties().setProperty(
        'expected_kanri',
        JSON.stringify({
          list: kanriNums,
          country: countryCode,
          timestamp: new Date().toISOString()
        })
      );

    } else {
      PropertiesService.getScriptProperties().deleteProperty('expected_kanri');
    }

    // 送料候補がある場合だけ、倉庫担当への指示を最後に追加
    if (hasRateOptions(ratesMessage)) {
      replyMsg += '\n─────────────────\n本明に、転送してください';
    }

    replyToLine(replyToken, replyMsg);

  } catch (err) {
    console.log('doPost error: ' + err.toString());
  }

  return output;
}

// ============================================================
// 賞味期限：画像受信 → Claude読み取り → 確認待ち
// ============================================================

function handleExpiryImage(messageId, replyToken) {
  try {
    var token = PropertiesService.getScriptProperties().getProperty('LINE_TOKEN');

    // LINEコンテンツAPIで画像を取得
    var res = UrlFetchApp.fetch('https://api-data.line.me/v2/bot/message/' + messageId + '/content', {
      method: 'GET',
      headers: {
        'Authorization': 'Bearer ' + token
      },
      muteHttpExceptions: true
    });

    if (res.getResponseCode() !== 200) {
      replyToLine(replyToken, '❌ 画像の取得に失敗しました（' + res.getResponseCode() + '）。もう一度送ってください。');
      return;
    }

    var blob = res.getBlob();
    var base64 = Utilities.base64Encode(blob.getBytes());
    var mime = blob.getContentType() === 'image/png' ? 'image/png' : 'image/jpeg';

    // Claudeで読み取り
    var items = analyzeExpirySheet(base64, mime);

    if (!items || items.length === 0) {
      replyToLine(replyToken, '❌ 賞味期限を読み取れませんでした。\n明るい場所で、紙全体が入るように真上から撮り直してください。');
      return;
    }

    // 直前の荷姿確認で送られた、この箱の正しい管理番号リスト
    var expectedList = null;
    var expCtx = JSON.parse(PropertiesService.getScriptProperties().getProperty('expected_kanri') || 'null');

    if (expCtx && expCtx.list && expCtx.list.length > 0) {
      expectedList = expCtx.list;
    }

    // 日付妥当性チェック＆確認メッセージ生成
    var lines = [
      '📋 賞味期限の読み取り結果',
      '─────────────────'
    ];

    var warnings = [];

    for (var i = 0; i < items.length; i++) {
      var it = items[i];
      var dateStr = it.display; // 表示は yyyy/mm/dd
      var mark = '';

      if (!it.kanriValid) {
        mark += ' ⚠️管理番号が5桁でない';
        warnings.push('管理番号「' + it.kanri + '」が5桁ではありません（読み間違いの可能性）');
      }

      // この箱の管理番号リストに無い番号は誤読の可能性
      it.inExpected = true; // 照合リストが無いときは判定しない＝trueのまま

      if (expectedList && it.kanriValid) {
        it.inExpected = expectedList.indexOf(it.kanri) >= 0;

        if (!it.inExpected) {
          mark += ' ⚠️この箱の番号ではない';
          warnings.push('管理番号「' + it.kanri + '」は直前の荷姿確認の箱に含まれていません（誤読の可能性）');
        }
      }

      if (!it.valid) {
        mark += ' ⚠️不正な日付';
        warnings.push(it.kanri + ' の日付（' + (it.raw || '空') + '）を確認してください');
      }

      lines.push(it.kanri + ' → ' + (dateStr || '（空欄）') + mark);
    }

    lines.push('─────────────────');

    if (warnings.length > 0) {
      lines.push('⚠️ 要確認 ' + warnings.length + '件:');

      for (var w = 0; w < warnings.length; w++) {
        lines.push('・' + warnings[w]);
      }

      lines.push('─────────────────');
    }

lines.push('');
lines.push('');
lines.push('よろしいですか？');
lines.push('');
lines.push('「OK」か「NG」　と返信');

    // コンテキスト保存（OKで書き込む）
    PropertiesService.getScriptProperties().setProperty(
      'expiry_context',
      JSON.stringify({
        items: items,
        country: expCtx && expCtx.country ? expCtx.country : '',
        timestamp: new Date().toISOString()
      })
    );

    replyToLine(replyToken, lines.join('\n'));

  } catch (err) {
    console.log('handleExpiryImage error: ' + err.toString());
    replyToLine(replyToken, '❌ 賞味期限の処理でエラーが発生しました: ' + err.toString());
  }
}

function analyzeExpirySheet(base64Image, mime) {
  var apiKey = PropertiesService.getScriptProperties().getProperty('CLAUDE_API_KEY');

  var prompt = 'これは賞味期限記入シートの写真です。各行の「管理番号」と「賞味期限（年/月/日）」を読み取り、'
    + 'JSON配列のみで返してください。説明文は不要。\n\n'
    + '表の構造: 左から「No.」「管理番号(5マス)」「賞味期限の年(4マス)/月(2マス)/日(2マス)」。'
    + '管理番号エリアと賞味期限エリアの間には太い縦の区切り線があります。\n\n'
    + '各行の形式: {"kanri":"管理番号5桁","year":"年4桁または空","month":"月または空","day":"日または空"}\n'
    + 'ルール（厳守）:\n'
    + '- 管理番号は必ず5桁の数字です。6桁以上になった場合は読み間違いなので、5マス分だけを管理番号として読んでください。\n'
    + '- 管理番号エリア(5マス)と賞味期限の年エリア(4マス)を絶対に混同しないこと。太い縦線の左が管理番号、右が賞味期限の年です。\n'
    + '- 管理番号が空欄の行（5マスすべて空）は出力しない\n'
    + '- 年は必ず西暦4桁。月日は1〜2桁。数字が読めないマスは空文字\n'
    + '- 出力はJSON配列のみ。例: [{"kanri":"14594","year":"2027","month":"1","day":"13"}]';

  var res = UrlFetchApp.fetch('https://api.anthropic.com/v1/messages', {
    method: 'POST',
    headers: {
      'Content-Type': 'application/json',
      'x-api-key': apiKey,
      'anthropic-version': '2023-06-01'
    },
    payload: JSON.stringify({
      model: 'claude-haiku-4-5-20251001',
      max_tokens: 1000,
      messages: [
        {
          role: 'user',
          content: [
            {
              type: 'image',
              source: {
                type: 'base64',
                media_type: mime,
                data: base64Image
              }
            },
            {
              type: 'text',
              text: prompt
            }
          ]
        }
      ]
    }),
    muteHttpExceptions: true
  });

  var data = JSON.parse(res.getContentText());

  if (data.error || !data.content || !data.content[0]) {
    console.log('Claude(expiry) error: ' + res.getContentText());
    return null;
  }

  var raw = data.content[0].text.replace(/```json|```/g, '').trim();
  var start = raw.indexOf('[');
  var end = raw.lastIndexOf(']');

  if (start === -1 || end === -1) return null;

  var arr;

  try {
    arr = JSON.parse(raw.substring(start, end + 1));
  } catch (e) {
    return null;
  }

  // 整形＋妥当性チェック
  var out = [];

  for (var i = 0; i < arr.length; i++) {
    var r = arr[i];

    if (!r.kanri) continue;

    var kanri = String(r.kanri).trim();
    var kanriValid = /^\d{5}$/.test(kanri);

    var y = parseInt(r.year, 10);
    var m = parseInt(r.month, 10);
    var d = parseInt(r.day, 10);

    var valid = false;
    var formatted = '';   // 書き込み用 MM/DD/YYYY
    var display = '';     // 表示用 yyyy/mm/dd
    var rawStr = (r.year || '?') + '/' + (r.month || '?') + '/' + (r.day || '?');

    if (y && m && d) {
      valid = isValidDate(y, m, d);

      if (valid) {
        formatted = pad2(m) + '/' + pad2(d) + '/' + y; // H列: MM/DD/YYYY
        display = y + '/' + pad2(m) + '/' + pad2(d);   // 表示: yyyy/mm/dd
      }
    }

    out.push({
      kanri: kanri,
      kanriValid: kanriValid,
      year: y,
      month: m,
      day: d,
      formatted: formatted,
      display: display,
      valid: valid,
      raw: rawStr
    });
  }

  return out;
}

function isValidDate(y, m, d) {
  if (m < 1 || m > 12) return false;
  if (d < 1 || d > 31) return false;

  var dim = [
    31,
    isLeap(y) ? 29 : 28,
    31,
    30,
    31,
    30,
    31,
    31,
    30,
    31,
    30,
    31
  ];

  return d <= dim[m - 1];
}

function isLeap(y) {
  return (y % 4 === 0 && y % 100 !== 0) || (y % 400 === 0);
}

function pad2(n) {
  return (n < 10 ? '0' : '') + n;
}

// ============================================================
// 賞味期限：確認OK → 発送GのH列に書き込み
// ============================================================

function confirmExpiryWrite(replyToken) {
  var props = PropertiesService.getScriptProperties();
  var ctx = JSON.parse(props.getProperty('expiry_context') || 'null');

  if (!ctx || !ctx.items) {
    replyToLine(replyToken, '⚠️ 登録対象が見つかりません。先に賞味期限シートの写真を送ってください。');
    return;
  }

  var result = writeExpiryToHassoG(ctx.items);

  props.deleteProperty('expiry_context');

  // JP/US問わず、通常成功時は「登録完了」だけ返す
  if (isSuccessfulExpiryWriteResult(result)) {
    replyToLine(replyToken, '登録完了');
  } else {
    // 警告・エラーがある場合は詳細を返す
    replyToLine(replyToken, result);
  }
}

function isSuccessfulExpiryWriteResult(result) {
  var msg = String(result || '');

  if (msg.indexOf('✅') !== 0) return false;
  if (msg.indexOf('⚠️') >= 0) return false;
  if (msg.indexOf('⏭') >= 0) return false;
  if (msg.indexOf('登録: なし') >= 0) return false;
  if (msg.indexOf('エラー') >= 0) return false;

  return true;
}

function writeExpiryToHassoG(items) {
  try {
    var ss = SpreadsheetApp.openById(HASSO_SHEET_ID);
    var sheet = ss.getSheetByName(HASSO_SHEET_NAME);

    if (!sheet) return '⚠️ 発送Gシートが見つかりません';

    var lastRow = sheet.getLastRow();
    var aCol = sheet.getRange(1, 1, lastRow, 1).getValues();

    var rowOf = {};

    for (var i = 0; i < aCol.length; i++) {
      var key = String(aCol[i][0]).trim();

      if (key !== '' && !(key in rowOf)) {
        rowOf[key] = i + 1;
      }
    }

    var written = [];
    var notFound = [];
    var skipped = [];
    var badKanri = [];
    var notInBox = [];

    for (var j = 0; j < items.length; j++) {
      var it = items[j];

      // 5桁でない番号は書き込まない
      if (!it.kanriValid) {
        badKanri.push(it.kanri);
        continue;
      }

      // この箱の番号でない＝誤読の可能性
      if (it.inExpected === false) {
        notInBox.push(it.kanri);
        continue;
      }

      if (!it.valid || !it.formatted) {
        skipped.push(it.kanri);
        continue;
      }

      var row = rowOf[it.kanri];

      if (!row) {
        notFound.push(it.kanri);
        continue;
      }

      sheet.getRange(row, 8).setValue(it.formatted); // H=8列目（MM/DD/YYYY）
      written.push(it.kanri + '→' + it.display);     // 表示は yyyy/mm/dd
    }

    var msg = '✅ 賞味期限を発送Gに登録しました\n─────────────────\n';

    msg += '登録: ' + (written.length > 0 ? written.join('\n') : 'なし');

    if (badKanri.length > 0) {
      msg += '\n─────────────────\n⚠️ 管理番号が5桁でないため書き込まず: ' + badKanri.join(', ')
        + '\n（撮り直すか手入力してください）';
    }

    if (notInBox.length > 0) {
      msg += '\n─────────────────\n⚠️ この箱の管理番号でないため書き込まず: ' + notInBox.join(', ')
        + '\n（誤読の可能性。撮り直すか確認してください）';
    }

    if (skipped.length > 0) {
      msg += '\n⏭ 日付不正でスキップ: ' + skipped.join(', ');
    }

    if (notFound.length > 0) {
      msg += '\n⚠️ 発送Gに見つからない管理番号: ' + notFound.join(', ');
    }

    return msg;

  } catch (err) {
    console.log('writeExpiryToHassoG error: ' + err.toString());
    return '⚠️ 賞味期限書き込みエラー: ' + err.toString();
  }
}

// ============================================================
// 送料確認まわり
// ============================================================

function parseDimensionsWithClaude(text) {
  var apiKey = PropertiesService.getScriptProperties().getProperty('CLAUDE_API_KEY');

  var prompt = 'LINEメッセージから梱包情報を抽出してJSONのみ返してください。説明不要。\n\n'
    + 'メッセージ: "' + text + '"\n\n'
    + 'このメッセージは次の構造（複数行）であることが多い:\n'
    + '  1行目: 管理番号（5桁の数字。スペース区切りで複数並ぶことがある）\n'
    + '  2行目: 「食品 あり」または「食品 なし」\n'
    + '  3行目: サイズ3つ（例「39 54 31」スペース区切り、または「39×54×31」）\n'
    + '  4行目: 重量（例「10kg」「10000g」）\n'
    + '  5行目: 国（例「US」「JP」）\n'
    + 'ただし1行にまとまっている場合もある（例「14525 60×40×40 14350g US」）。どちらの形式も読み取ること。\n\n'
    + '返答形式（このJSONのみ、余分なテキストなし）:\n'
    + '{"length_cm":数値,"width_cm":数値,"height_cm":数値,"weight_g":数値,"country_code":"2文字ISO","zip":"郵便番号または空文字","order_id":"最初の管理番号または空文字","has_food":true または false}\n\n'
    + '変換ルール:\n'
    + '- 管理番号（5桁の数字）はサイズではない。サイズは「食品 あり/なし」の後にある数字3つ、または×で区切られた数字3つ\n'
    + '- 日本/国内/JP→JP, アメリカ/米国/USA→US, イギリス/UK→GB, オーストラリア→AU, カナダ→CA, ドイツ→DE, フランス→FR, オランダ→NL, シンガポール→SG, 台湾→TW, 韓国→KR, 香港→HK\n'
    + '- kgはgに変換（1.5kg→1500、10kg→10000）\n'
    + '- サイズ順不明なら最大値=length, 中間=width, 最小=height\n'
    + '- 国が不明ならUS\n'
    + '- サイズか重量が本当に見つからないときだけnull（文字列）を返す\n'
    + '- order_idは最初の管理番号（5桁）\n'
    + '- 管理番号がなければ空文字\n'
    + '- 「食品 あり」「食品あり」等があれば has_food=true、「食品 なし」「食品なし」等なら has_food=false、記載なしはfalse';

  try {
    var response = UrlFetchApp.fetch('https://api.anthropic.com/v1/messages', {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
        'x-api-key': apiKey,
        'anthropic-version': '2023-06-01'
      },
      payload: JSON.stringify({
        model: 'claude-haiku-4-5-20251001',
        max_tokens: 200,
        messages: [
          {
            role: 'user',
            content: prompt
          }
        ]
      }),
      muteHttpExceptions: true
    });

    var result = JSON.parse(response.getContentText());

    if (!result.content || !result.content[0]) return null;

    var rawText = result.content[0].text.trim();
    var start = rawText.indexOf('{');
    var end = rawText.lastIndexOf('}');

    if (start === -1 || end === -1) return null;

    var parsed = JSON.parse(rawText.substring(start, end + 1));

    if (!parsed.length_cm || !parsed.width_cm || !parsed.height_cm || !parsed.weight_g) {
      return null;
    }

    return {
      length: parsed.length_cm,
      width: parsed.width_cm,
      height: parsed.height_cm,
      weight: parsed.weight_g,
      country: String(parsed.country_code || 'US').trim().toUpperCase(),
      zip: parsed.zip || '',
      order_id: parsed.order_id || '',
      has_food: parsed.has_food === true
    };

  } catch (err) {
    console.log('Claude error: ' + err.toString());
    return null;
  }
}

function getDefaultAddress(country, zip) {
  var addresses = {
    'US': {
      country: 'US',
      full_name: 'Test Buyer',
      phone: '2125551234',
      address1: '1600 Amphitheatre Pkwy',
      city: 'Mountain View',
      province: 'CA',
      zip: zip || '94043'
    },
    'GB': {
      country: 'GB',
      full_name: 'Test Buyer',
      phone: '02071234567',
      address1: '10 Downing Street',
      city: 'London',
      province: 'ENG',
      zip: zip || 'SW1A2AA'
    },
    'AU': {
      country: 'AU',
      full_name: 'Test Buyer',
      phone: '0298765432',
      address1: '1 Martin Place',
      city: 'Sydney',
      province: 'NSW',
      zip: zip || '2000'
    },
    'CA': {
      country: 'CA',
      full_name: 'Test Buyer',
      phone: '4165551234',
      address1: '100 King Street West',
      city: 'Toronto',
      province: 'ON',
      zip: zip || 'M5X1A9'
    },
    'DE': {
      country: 'DE',
      full_name: 'Test Buyer',
      phone: '03012345678',
      address1: 'Unter den Linden 1',
      city: 'Berlin',
      zip: zip || '10117'
    },
    'FR': {
      country: 'FR',
      full_name: 'Test Buyer',
      phone: '0142345678',
      address1: '1 Rue de Rivoli',
      city: 'Paris',
      zip: zip || '75001'
    },
    'NL': {
      country: 'NL',
      full_name: 'Test Buyer',
      phone: '0201234567',
      address1: 'Dam 1',
      city: 'Amsterdam',
      zip: zip || '1012JS'
    },
    'SG': {
      country: 'SG',
      full_name: 'Test Buyer',
      phone: '65123456',
      address1: '1 Raffles Place',
      city: 'Singapore',
      zip: zip || '048616'
    },
    'TW': {
      country: 'TW',
      full_name: 'Test Buyer',
      phone: '0212345678',
      address1: '1 Zhongzheng Road',
      city: 'Taipei',
      zip: zip || '100'
    },
    'KR': {
      country: 'KR',
      full_name: 'Test Buyer',
      phone: '0212345678',
      address1: '1 Sejong-daero',
      city: 'Seoul',
      zip: zip || '04524'
    },
    'HK': {
      country: 'HK',
      full_name: 'Test Buyer',
      phone: '21234567',
      address1: '1 Connaught Place',
      city: 'Hong Kong',
      zip: zip || '999077'
    }
  };

  var addr = addresses[country];

  if (!addr) {
    addr = {
      country: country,
      full_name: 'Test Buyer',
      phone: '0000000000',
      address1: '1 Main Street',
      city: 'City',
      zip: zip || '00000'
    };
  }

  return addr;
}

function getShipAndCoRates(dims) {
  var token = PropertiesService.getScriptProperties().getProperty('SHIPANDCO_TOKEN');

  var payload = {
    setup: {
      currency: 'USD'
    },
    from_address: {
      country: 'JP',
      full_name: 'Sky Crew Japan',
      company: 'Sky Crew Japan',
      phone: '0489000000',
      email: 'shinicchee@gmail.com',
      address1: '1-1 Soka',
      city: 'Soka',
      province: 'Saitama',
      zip: '340-0001'
    },
    to_address: getDefaultAddress(dims.country, dims.zip),
    parcels: [
      {
        width: dims.width,
        height: dims.height,
        depth: dims.length,
        weight: dims.weight,
        amount: 1
      }
    ],
    products: [
      {
        name: 'Item',
        price: 10,
        quantity: 1,
        country: 'JP',
        currency: 'USD'
      }
    ]
  };

  try {
    var response = UrlFetchApp.fetch('https://api.shipandco.com/v1/rates', {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
        'x-access-token': token
      },
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    });

    var sc = response.getResponseCode();

    if (sc !== 200) {
      return {
        error: 'HTTP ' + sc
      };
    }

    return {
      success: true,
      data: JSON.parse(response.getContentText())
    };

  } catch (err) {
    return {
      error: err.toString()
    };
  }
}

function formatRatesMessage(ratesResult, dims) {
  var volWeight = Math.round(dims.length * dims.width * dims.height / 5000 * 1000);
  var chargeWeight = Math.max(dims.weight, volWeight);
  var isVol = volWeight > dims.weight;

  var lines = [];

  lines.push('📦 送料確認結果' + (dims.order_id ? ' ／ ' + dims.order_id : ''));
  lines.push('─────────────────');
  lines.push(dims.length + '×' + dims.width + '×' + dims.height + ' cm');
  lines.push('実重量: ' + dims.weight + 'g　容積: ' + volWeight + 'g' + (isVol ? ' ⚠️' : ''));
  lines.push('課金対象: ' + chargeWeight + 'g / 発送先: ' + dims.country);
  lines.push('─────────────────');

  if (ratesResult.error) {
    lines.push('❌ エラー: ' + ratesResult.error);
    return lines.join('\n');
  }

  var rateList = ratesResult.data.filter(function (r) {
    return !r.errors;
  });

  if (rateList.length === 0) {
    lines.push('利用可能なサービスがありませんでした');
    return lines.join('\n');
  }

  rateList.sort(function (a, b) {
    return a.price - b.price;
  });

  lines.push('【送料一覧】');

  rateList.slice(0, 10).forEach(function (r, i) {
    var name = r.service
      .replace('japanpost_', '郵便/')
      .replace('fedex_international_', 'FedEx/')
      .replace('dhl_express_', 'DHL/')
      .replace('ups_worldwide_', 'UPS/')
      .replace('ups_', 'UPS/');

    var surcharge = r.surcharges && r.surcharges.length > 0
      ? ' (+¥' + r.surcharges.reduce(function (s, x) {
        return s + x.price;
      }, 0) + ')'
      : '';

    lines.push((i + 1) + '. ' + name);
    lines.push('   ¥' + Math.round(r.price).toLocaleString() + surcharge);
  });

  return lines.join('\n');
}

function hasRateOptions(message) {
  return /\n1\.\s+/.test(String(message || ''));
}

function saveSoryoContext(kanriNumber, optionsText, dimensions, allKanriNumbers, hasFood) {
  var props = PropertiesService.getScriptProperties();
  var options = [];
  var lines = optionsText.split('\n');

  for (var i = 0; i < lines.length; i++) {
    var cm = lines[i].match(/^(\d+)\.\s+(.+)/);

    if (cm) {
      var price = '';

      if (i + 1 < lines.length) {
        var pm = lines[i + 1].match(/(¥[\d,]+)/);

        if (pm) {
          price = pm[1];
        }
      }

      options.push({
        num: cm[1],
        carrier: cm[2].trim(),
        price: price
      });
    }
  }

  props.setProperty(
    'soryo_context',
    JSON.stringify({
      kanriNumber: kanriNumber,
      allKanriNumbers: allKanriNumbers || kanriNumber,
      options: options,
      dimensions: dimensions || {},
      hasFood: hasFood === true,
      timestamp: new Date().toISOString()
    })
  );
}

function handleSoryoReply(replyText, replyToken) {
  var props = PropertiesService.getScriptProperties();
  var context = JSON.parse(props.getProperty('soryo_context') || 'null');

  if (!context) {
    replyToLine(replyToken, '⚠️ 送料選択の対象が見つかりません。\n先に送料確認を実行してください。');
    return;
  }

  var selected = context.options.find(function (o) {
    return o.num === replyText;
  });

  if (!selected) {
    replyToLine(replyToken, '⚠️ 番号が正しくありません。\n1〜' + context.options.length + ' の番号で返信してください。');
    return;
  }

  var parts = selected.carrier.split('/');
  var carrier = parts[0];
  var service = parts[1] || '';
  var price = selected.price.replace('¥', '').replace(/,/g, '');

  recordToSoryoSheet(
    context.kanriNumber,
    replyText,
    carrier,
    service,
    price,
    context.dimensions,
    context.allKanriNumbers
  );

  var hassoWarning = writeToHassoG(context.allKanriNumbers, context.dimensions);

  props.deleteProperty('soryo_context');

// US食品ありの場合は、送料選択後に賞味期限画像を依頼する
if (hassoWarning) {
  replyToLine(replyToken, hassoWarning);
} else if (context.hasFood === true) {
  replyToLine(replyToken, '賞味期限画像を送ってください');
} else {
  replyToLine(replyToken, '登録完了');
}
}

function recordToSoryoSheet(kanriNumber, selectedNum, carrier, service, price, dimensions, allKanriNumbers) {
  var ss = SpreadsheetApp.openById(SORYO_SHEET_ID);
  var sheet = ss.getSheetByName('FBA送料一覧');
  var data = sheet.getDataRange().getValues();

  for (var i = data.length - 1; i >= 1; i--) {
    if (String(data[i][1]) === String(kanriNumber)) {
      sheet.deleteRow(i + 1);
    }
  }

  var dim = dimensions || {};

  sheet.appendRow([
    new Date(),
    kanriNumber,
    dim.length || '',
    dim.width || '',
    dim.height || '',
    dim.weight || '',
    selectedNum,
    carrier,
    service,
    '¥' + Number(price).toLocaleString(),
    allKanriNumbers || kanriNumber,
    '選択済',
    ''
  ]);
}

function writeToHassoG(allKanriNumbers, dimensions) {
  try {
    var dim = dimensions || {};
    var nums = String(allKanriNumbers || '').split(/\s+/).filter(function (n) {
      return n !== '';
    });

    if (nums.length === 0) {
      return '⚠️ 発送G: 管理番号が空のため書き込みスキップ';
    }

    var groupStr = nums.join(' ');
    var minNum = nums.reduce(function (min, cur) {
      return Number(cur) < Number(min) ? cur : min;
    }, nums[0]);

    var ss = SpreadsheetApp.openById(HASSO_SHEET_ID);
    var sheet = ss.getSheetByName(HASSO_SHEET_NAME);

    if (!sheet) {
      return '⚠️ 発送G: シートが見つかりません';
    }

    var lastRow = sheet.getLastRow();

    if (lastRow < 1) {
      return '⚠️ 発送G: データがありません';
    }

    var aCol = sheet.getRange(1, 1, lastRow, 1).getValues();
    var rowOf = {};

    for (var i = 0; i < aCol.length; i++) {
      var key = String(aCol[i][0]).trim();

      if (key !== '' && !(key in rowOf)) {
        rowOf[key] = i + 1;
      }
    }

    var notFound = [];

    for (var j = 0; j < nums.length; j++) {
      var num = nums[j];
      var row = rowOf[num];

      if (!row) {
        notFound.push(num);
        continue;
      }

      // C列：同じ箱の管理番号グループ
      sheet.getRange(row, 3).setValue(groupStr);

      // 荷姿は最小管理番号の行だけに入れる
      if (num === minNum) {
        sheet.getRange(row, 4).setValue(dim.weight || '');
        sheet.getRange(row, 5).setValue(dim.length || '');
        sheet.getRange(row, 6).setValue(dim.width || '');
        sheet.getRange(row, 7).setValue(dim.height || '');
      }
    }

    if (notFound.length > 0) {
      return '⚠️ 発送Gに見つからない管理番号: ' + notFound.join(', ') + '\n（見つかった分は書き込み済み）';
    }

    return '';

  } catch (err) {
    return '⚠️ 発送G書き込みエラー: ' + err.toString();
  }
}

function deleteFromSoryoSheet(kanriNumber, replyToken) {
  var ss = SpreadsheetApp.openById(SORYO_SHEET_ID);
  var sheet = ss.getSheetByName('FBA送料一覧');
  var data = sheet.getDataRange().getValues();

  var deleted = 0;

  for (var i = data.length - 1; i >= 1; i--) {
    if (String(data[i][1]) === String(kanriNumber)) {
      sheet.deleteRow(i + 1);
      deleted++;
    }
  }

  if (deleted > 0) {
    replyToLine(
      replyToken,
      '🗑️ 削除完了\n─────────────────\n管理番号: ' + kanriNumber + '\n' + deleted + '件のデータを削除しました'
    );
  } else {
    replyToLine(replyToken, '⚠️ 該当データなし\n管理番号: ' + kanriNumber);
  }
}

function checkSoryoSheet(kanriNumber, replyToken) {
  var ss = SpreadsheetApp.openById(SORYO_SHEET_ID);
  var sheet = ss.getSheetByName('FBA送料一覧');
  var data = sheet.getDataRange().getValues();

  var found = null;

  for (var i = 1; i < data.length; i++) {
    if (String(data[i][1]) === String(kanriNumber)) {
      found = data[i];
      break;
    }
  }

  if (found) {
    replyToLine(
      replyToken,
      '📋 登録内容確認\n'
      + '─────────────────\n'
      + '管理番号: ' + found[1] + '\n'
      + '全管理番号: ' + found[10] + '\n'
      + '荷姿: ' + found[2] + '×' + found[3] + '×' + found[4] + 'cm / ' + found[5] + 'kg\n'
      + '採用業者: ' + found[7] + '\n'
      + 'サービス: ' + found[8] + '\n'
      + '送料: ' + found[9] + '\n'
      + 'ステータス: ' + found[11]
    );
  } else {
    replyToLine(replyToken, '⚠️ 該当データなし\n管理番号: ' + kanriNumber);
  }
}

function replyToLine(replyToken, text) {
  var token = PropertiesService.getScriptProperties().getProperty('LINE_TOKEN');

  UrlFetchApp.fetch('https://api.line.me/v2/bot/message/reply', {
    method: 'POST',
    headers: {
      'Content-Type': 'application/json',
      'Authorization': 'Bearer ' + token
    },
    payload: JSON.stringify({
      replyToken: replyToken,
      messages: [
        {
          type: 'text',
          text: text
        }
      ]
    }),
    muteHttpExceptions: true
  });
}

function pushToLine(userId, text) {
  var token = PropertiesService.getScriptProperties().getProperty('LINE_TOKEN');

  UrlFetchApp.fetch('https://api.line.me/v2/bot/message/push', {
    method: 'POST',
    headers: {
      'Content-Type': 'application/json',
      'Authorization': 'Bearer ' + token
    },
    payload: JSON.stringify({
      to: userId,
      messages: [
        {
          type: 'text',
          text: text
        }
      ]
    }),
    muteHttpExceptions: true
  });
}

function getHelpMessage() {
  return [
    '📦 送料チェッカー',
    '─────────────────',
    '【USなど海外向け：送料確認】',
    '「14525 60×40×40 14350g US」',
    '「14111 14112 14113 60×40×40 14350g US」',
    '数字のみ → 送料番号を選択',
    '',
    '【JP向け：荷姿登録のみ】',
    '「14525 60×40×40 14350g JP」',
    '「14111 14112 14113 食品なし 60×40×40 10kg JP」',
    'JPは送料計算なし・FBA送料一覧記録なし',
    '',
    '【賞味期限】',
    '食品ありの場合、賞味期限画像を送信',
    '→ 読み取り結果を確認し「OK」で登録',
    '',
    '【その他】',
    '「削除 14525」/「確認 14525」',
    '─────────────────'
  ].join('\n');
}

// ============================================================
// テスト用
// ============================================================

function testExpiryValidate() {
  // 妥当性チェックの動作確認
  console.log('2026/9/31 valid? ' + isValidDate(2026, 9, 31)); // false
  console.log('2026/9/30 valid? ' + isValidDate(2026, 9, 30)); // true
  console.log('2027/2/29 valid? ' + isValidDate(2027, 2, 29)); // false
  console.log('2028/2/29 valid? ' + isValidDate(2028, 2, 29)); // true
  console.log('format 2027/1/13 → ' + pad2(1) + '/' + pad2(13) + '/' + 2027);
}

/**
 * Ship&Co LINE Bot - v12
 * 送料確認 ＋ 賞味期限OCR ＋ JP向け荷姿登録 ＋ 返信文簡略化
 *
 * v12 修正点:
 *   [JP向け]
 *   - 国コードが JP の場合は Ship&Co 送料計算をしない
 *   - JP の場合は FBA送料一覧シートへ送料記録しない
 *   - JP の場合は 発送G に管理番号グループ・荷姿のみ登録する
 *   - JP食品あり:
 *       初回返信は「賞味期限画像を送ってください」
 *       賞味期限OK後、通常成功時は「登録完了」
 *   - JP食品なし:
 *       初回返信は「登録完了」
 *
 *   [USなど海外向け]
 *   - 送料候補メッセージの最後に
 *       「本明に、転送してください」
 *     を追加
 *   - 送料番号選択後、通常成功時は「登録完了」
 *   - 賞味期限OK後、通常成功時は「登録完了」
 *
 * スクリプトプロパティ（必須）:
 *   LINE_TOKEN, SHIPANDCO_TOKEN, CLAUDE_API_KEY, SHIN_USER_ID
 */

var SORYO_SHEET_ID = '1xJ8PiYwv_T_FVLEXBUi1Fs8MYqAav7CjknzqieLhbkg';  // FBA送料一覧
var HASSO_SHEET_ID = '1l6Vs-5el4-N3xe0msWoqOTWAg6J_RAARtvVJF73VdsE'; // そー草加ファイル
var HASSO_SHEET_NAME = '発送G';

// ============================================================
// Webhook受信
// ============================================================

function doGet(e) {
  return ContentService.createTextOutput('OK');
}

function doPost(e) {
  var output = ContentService.createTextOutput('OK');

  try {
    var body = JSON.parse(e.postData.contents);
    var events = body.events;
    if (!events || events.length === 0) return output;

    var event = events[0];

    var senderUserId = (event.source && event.source.userId) ? event.source.userId : '';
    console.log('senderUserId: ' + senderUserId);

    if (event.type !== 'message') return output;

    // ===== 画像メッセージ → 賞味期限OCR =====
    if (event.message.type === 'image') {
      console.log('→ 賞味期限画像を受信');
      handleExpiryImage(event.message.id, event.replyToken);
      return output;
    }

    // ===== テキストメッセージ =====
    if (event.message.type !== 'text') return output;

    var replyToken = event.replyToken;
    var userMessage = event.message.text.trim();

    console.log('受信メッセージ（正規化前）: ' + userMessage);

    // タブと連続スペースは詰めるが、改行は保持する
    userMessage = userMessage
      .replace(/\t+/g, ' ')
      .replace(/[ \u3000]+/g, ' ')
      .replace(/×\s+/g, '×')
      .replace(/[ ]*\n[ ]*/g, '\n')
      .trim();

    console.log('受信メッセージ（正規化後）: ' + userMessage);

    // 賞味期限の確認OK（書き込み実行）
    if (
      userMessage === 'OK' ||
      userMessage === 'ok' ||
      userMessage === 'Ok' ||
      userMessage === 'オーケー' ||
      userMessage === 'おーけー'
    ) {
      confirmExpiryWrite(replyToken);
      return output;
    }

    // 賞味期限のキャンセル
    if (userMessage === '賞味期限キャンセル' || userMessage === 'NG' || userMessage === 'ng') {
      PropertiesService.getScriptProperties().deleteProperty('expiry_context');
      replyToLine(replyToken, '賞味期限の登録をキャンセルしました。もう一度シートを撮影して送ってください。');
      return output;
    }

    if (userMessage === 'ping') {
      replyToLine(replyToken, 'pong v12 動作中');
      return output;
    }

    if (userMessage === 'myid') {
      PropertiesService.getScriptProperties().setProperty('SHIN_USER_ID', senderUserId);
      replyToLine(replyToken, 'あなたのID:\n' + senderUserId + '\n\nSHIN_USER_IDに登録しました');
      return output;
    }

    if (userMessage.startsWith('削除 ')) {
      deleteFromSoryoSheet(userMessage.replace('削除 ', '').trim(), replyToken);
      return output;
    }

    if (userMessage.startsWith('確認 ')) {
      checkSoryoSheet(userMessage.replace('確認 ', '').trim(), replyToken);
      return output;
    }

    if (/^\d+$/.test(userMessage)) {
      handleSoryoReply(userMessage, replyToken);
      return output;
    }

    if (userMessage === 'help' || userMessage === 'ヘルプ') {
      replyToLine(replyToken, getHelpMessage());
      return output;
    }

    // ========================================================
    // 通常の荷姿情報解析
    // ========================================================

    console.log('→ Claude解析開始');

// ===== 在庫承認（進めて / 停めて）=====
    if (userMessage === '進めてください' || userMessage === '進めて') {
      handleInventoryApproval_('進める', replyToken);
      return output;
    }
    if (userMessage === '停めてください' || userMessage === '停めて' ||
        userMessage === '止めてください' || userMessage === '止めて') {
      handleInventoryApproval_('停める', replyToken);
      return output;
    }

    var dims = parseDimensionsWithClaude(userMessage);

    if (!dims) {
      replyToLine(
        replyToken,
        '❌ サイズ情報を認識できませんでした。\n\n例:\n「14525 60×40×40 14350g US」\n「14111 14112 14113 60×40×40 14350g JP」\n\n「help」でヘルプ表示'
      );
      return output;
    }

    // この箱の管理番号リストを作る
    var allKanriNumbers = dims.order_id;
    var kanriNums = [];

    if (dims.order_id) {
      var firstLine = userMessage.split('\n')[0];
      var tokens = firstLine.split(/\s+/);

      for (var i = 0; i < tokens.length; i++) {
        if (/^\d{5}$/.test(tokens[i])) {
          kanriNums.push(tokens[i]);
        } else {
          break;
        }
      }

      allKanriNumbers = kanriNums.length > 0 ? kanriNums.join(' ') : dims.order_id;
    }

    var dimForSheet = {
      length: dims.length,
      width: dims.width,
      height: dims.height,
      weight: dims.weight / 1000
    };

    var countryCode = String(dims.country || '').trim().toUpperCase();
    console.log('判定用 countryCode: [' + countryCode + ']');

    // ========================================================
    // JP向け：送料計算なし・FBA送料一覧記録なし・発送Gだけ登録
    // ========================================================

    if (countryCode === 'JP') {
      var jpProps = PropertiesService.getScriptProperties();

      // JPでは送料選択を使わないため、前回の送料選択コンテキストを消す
      jpProps.deleteProperty('soryo_context');

      // 発送Gへ管理番号グループ・荷姿を書き込み
      var jpHassoWarning = writeToHassoG(allKanriNumbers, dimForSheet);

      if (dims.has_food) {
        // 賞味期限OCR時の照合用に、この箱の正しい管理番号リストを保存
        jpProps.setProperty(
          'expected_kanri',
          JSON.stringify({
            list: kanriNums,
            country: 'JP',
            timestamp: new Date().toISOString()
          })
        );

        // 通常成功時は、倉庫担当へ余計な説明を出さない
        if (jpHassoWarning) {
          replyToLine(replyToken, jpHassoWarning + '\n\n賞味期限画像を送ってください');
        } else {
          replyToLine(replyToken, '賞味期限画像を送ってください');
        }

      } else {
        // 食品なしでは賞味期限照合リストを残さない
        jpProps.deleteProperty('expected_kanri');

        // 通常成功時は「登録完了」のみ
        if (jpHassoWarning) {
          replyToLine(replyToken, jpHassoWarning);
        } else {
          replyToLine(replyToken, '登録完了');
        }
      }

      return output;
    }

    // ========================================================
    // USなど海外向け：従来通り送料計算・送料選択待ち
    // ========================================================

    var ratesResult = getShipAndCoRates(dims);
    var ratesMessage = formatRatesMessage(ratesResult, dims);

    if (dims.order_id) {
    saveSoryoContext(dims.order_id, ratesMessage, dimForSheet, allKanriNumbers, dims.has_food);
    }

    // 食品ありなら、送料メッセージに賞味期限の案内を追記
    var replyMsg = ratesMessage;

    if (dims.has_food) {
      var foodList = kanriNums.length > 0 ? kanriNums.join(' ') : allKanriNumbers;

      replyMsg += '\n─────────────────\n🍱 食品が含まれます\n対象管理番号: ' + foodList
        + '\n→ 賞味期限記入シートに上記を記入し、撮影して送ってください。';

      PropertiesService.getScriptProperties().setProperty(
        'expected_kanri',
        JSON.stringify({
          list: kanriNums,
          country: countryCode,
          timestamp: new Date().toISOString()
        })
      );

    } else {
      PropertiesService.getScriptProperties().deleteProperty('expected_kanri');
    }

    // 送料候補がある場合だけ、倉庫担当への指示を最後に追加
    if (hasRateOptions(ratesMessage)) {
      replyMsg += '\n─────────────────\n本明に、転送してください';
    }

    replyToLine(replyToken, replyMsg);

  } catch (err) {
    console.log('doPost error: ' + err.toString());
  }

  return output;
}

// ============================================================
// 賞味期限：画像受信 → Claude読み取り → 確認待ち
// ============================================================

function handleExpiryImage(messageId, replyToken) {
  try {
    var token = PropertiesService.getScriptProperties().getProperty('LINE_TOKEN');

    // LINEコンテンツAPIで画像を取得
    var res = UrlFetchApp.fetch('https://api-data.line.me/v2/bot/message/' + messageId + '/content', {
      method: 'GET',
      headers: {
        'Authorization': 'Bearer ' + token
      },
      muteHttpExceptions: true
    });

    if (res.getResponseCode() !== 200) {
      replyToLine(replyToken, '❌ 画像の取得に失敗しました（' + res.getResponseCode() + '）。もう一度送ってください。');
      return;
    }

    var blob = res.getBlob();
    var base64 = Utilities.base64Encode(blob.getBytes());
    var mime = blob.getContentType() === 'image/png' ? 'image/png' : 'image/jpeg';

    // Claudeで読み取り
    var items = analyzeExpirySheet(base64, mime);

    if (!items || items.length === 0) {
      replyToLine(replyToken, '❌ 賞味期限を読み取れませんでした。\n明るい場所で、紙全体が入るように真上から撮り直してください。');
      return;
    }

    // 直前の荷姿確認で送られた、この箱の正しい管理番号リスト
    var expectedList = null;
    var expCtx = JSON.parse(PropertiesService.getScriptProperties().getProperty('expected_kanri') || 'null');

    if (expCtx && expCtx.list && expCtx.list.length > 0) {
      expectedList = expCtx.list;
    }

    // 日付妥当性チェック＆確認メッセージ生成
    var lines = [
      '📋 賞味期限の読み取り結果',
      '─────────────────'
    ];

    var warnings = [];

    for (var i = 0; i < items.length; i++) {
      var it = items[i];
      var dateStr = it.display; // 表示は yyyy/mm/dd
      var mark = '';

      if (!it.kanriValid) {
        mark += ' ⚠️管理番号が5桁でない';
        warnings.push('管理番号「' + it.kanri + '」が5桁ではありません（読み間違いの可能性）');
      }

      // この箱の管理番号リストに無い番号は誤読の可能性
      it.inExpected = true; // 照合リストが無いときは判定しない＝trueのまま

      if (expectedList && it.kanriValid) {
        it.inExpected = expectedList.indexOf(it.kanri) >= 0;

        if (!it.inExpected) {
          mark += ' ⚠️この箱の番号ではない';
          warnings.push('管理番号「' + it.kanri + '」は直前の荷姿確認の箱に含まれていません（誤読の可能性）');
        }
      }

      if (!it.valid) {
        mark += ' ⚠️不正な日付';
        warnings.push(it.kanri + ' の日付（' + (it.raw || '空') + '）を確認してください');
      }

      lines.push(it.kanri + ' → ' + (dateStr || '（空欄）') + mark);
    }

    lines.push('─────────────────');

    if (warnings.length > 0) {
      lines.push('⚠️ 要確認 ' + warnings.length + '件:');

      for (var w = 0; w < warnings.length; w++) {
        lines.push('・' + warnings[w]);
      }

      lines.push('─────────────────');
    }

lines.push('');
lines.push('');
lines.push('よろしいですか？');
lines.push('');
lines.push('「OK」か「NG」　と返信');

    // コンテキスト保存（OKで書き込む）
    PropertiesService.getScriptProperties().setProperty(
      'expiry_context',
      JSON.stringify({
        items: items,
        country: expCtx && expCtx.country ? expCtx.country : '',
        timestamp: new Date().toISOString()
      })
    );

    replyToLine(replyToken, lines.join('\n'));

  } catch (err) {
    console.log('handleExpiryImage error: ' + err.toString());
    replyToLine(replyToken, '❌ 賞味期限の処理でエラーが発生しました: ' + err.toString());
  }
}

function analyzeExpirySheet(base64Image, mime) {
  var apiKey = PropertiesService.getScriptProperties().getProperty('CLAUDE_API_KEY');

  var prompt = 'これは賞味期限記入シートの写真です。各行の「管理番号」と「賞味期限（年/月/日）」を読み取り、'
    + 'JSON配列のみで返してください。説明文は不要。\n\n'
    + '表の構造: 左から「No.」「管理番号(5マス)」「賞味期限の年(4マス)/月(2マス)/日(2マス)」。'
    + '管理番号エリアと賞味期限エリアの間には太い縦の区切り線があります。\n\n'
    + '各行の形式: {"kanri":"管理番号5桁","year":"年4桁または空","month":"月または空","day":"日または空"}\n'
    + 'ルール（厳守）:\n'
    + '- 管理番号は必ず5桁の数字です。6桁以上になった場合は読み間違いなので、5マス分だけを管理番号として読んでください。\n'
    + '- 管理番号エリア(5マス)と賞味期限の年エリア(4マス)を絶対に混同しないこと。太い縦線の左が管理番号、右が賞味期限の年です。\n'
    + '- 管理番号が空欄の行（5マスすべて空）は出力しない\n'
    + '- 年は必ず西暦4桁。月日は1〜2桁。数字が読めないマスは空文字\n'
    + '- 出力はJSON配列のみ。例: [{"kanri":"14594","year":"2027","month":"1","day":"13"}]';

  var res = UrlFetchApp.fetch('https://api.anthropic.com/v1/messages', {
    method: 'POST',
    headers: {
      'Content-Type': 'application/json',
      'x-api-key': apiKey,
      'anthropic-version': '2023-06-01'
    },
    payload: JSON.stringify({
      model: 'claude-haiku-4-5-20251001',
      max_tokens: 1000,
      messages: [
        {
          role: 'user',
          content: [
            {
              type: 'image',
              source: {
                type: 'base64',
                media_type: mime,
                data: base64Image
              }
            },
            {
              type: 'text',
              text: prompt
            }
          ]
        }
      ]
    }),
    muteHttpExceptions: true
  });

  var data = JSON.parse(res.getContentText());

  if (data.error || !data.content || !data.content[0]) {
    console.log('Claude(expiry) error: ' + res.getContentText());
    return null;
  }

  var raw = data.content[0].text.replace(/```json|```/g, '').trim();
  var start = raw.indexOf('[');
  var end = raw.lastIndexOf(']');

  if (start === -1 || end === -1) return null;

  var arr;

  try {
    arr = JSON.parse(raw.substring(start, end + 1));
  } catch (e) {
    return null;
  }

  // 整形＋妥当性チェック
  var out = [];

  for (var i = 0; i < arr.length; i++) {
    var r = arr[i];

    if (!r.kanri) continue;

    var kanri = String(r.kanri).trim();
    var kanriValid = /^\d{5}$/.test(kanri);

    var y = parseInt(r.year, 10);
    var m = parseInt(r.month, 10);
    var d = parseInt(r.day, 10);

    var valid = false;
    var formatted = '';   // 書き込み用 MM/DD/YYYY
    var display = '';     // 表示用 yyyy/mm/dd
    var rawStr = (r.year || '?') + '/' + (r.month || '?') + '/' + (r.day || '?');

    if (y && m && d) {
      valid = isValidDate(y, m, d);

      if (valid) {
        formatted = pad2(m) + '/' + pad2(d) + '/' + y; // H列: MM/DD/YYYY
        display = y + '/' + pad2(m) + '/' + pad2(d);   // 表示: yyyy/mm/dd
      }
    }

    out.push({
      kanri: kanri,
      kanriValid: kanriValid,
      year: y,
      month: m,
      day: d,
      formatted: formatted,
      display: display,
      valid: valid,
      raw: rawStr
    });
  }

  return out;
}

function isValidDate(y, m, d) {
  if (m < 1 || m > 12) return false;
  if (d < 1 || d > 31) return false;

  var dim = [
    31,
    isLeap(y) ? 29 : 28,
    31,
    30,
    31,
    30,
    31,
    31,
    30,
    31,
    30,
    31
  ];

  return d <= dim[m - 1];
}

function isLeap(y) {
  return (y % 4 === 0 && y % 100 !== 0) || (y % 400 === 0);
}

function pad2(n) {
  return (n < 10 ? '0' : '') + n;
}

// ============================================================
// 賞味期限：確認OK → 発送GのH列に書き込み
// ============================================================

function confirmExpiryWrite(replyToken) {
  var props = PropertiesService.getScriptProperties();
  var ctx = JSON.parse(props.getProperty('expiry_context') || 'null');

  if (!ctx || !ctx.items) {
    replyToLine(replyToken, '⚠️ 登録対象が見つかりません。先に賞味期限シートの写真を送ってください。');
    return;
  }

  var result = writeExpiryToHassoG(ctx.items);

  props.deleteProperty('expiry_context');

  // JP/US問わず、通常成功時は「登録完了」だけ返す
  if (isSuccessfulExpiryWriteResult(result)) {
    replyToLine(replyToken, '登録完了');
  } else {
    // 警告・エラーがある場合は詳細を返す
    replyToLine(replyToken, result);
  }
}

function isSuccessfulExpiryWriteResult(result) {
  var msg = String(result || '');

  if (msg.indexOf('✅') !== 0) return false;
  if (msg.indexOf('⚠️') >= 0) return false;
  if (msg.indexOf('⏭') >= 0) return false;
  if (msg.indexOf('登録: なし') >= 0) return false;
  if (msg.indexOf('エラー') >= 0) return false;

  return true;
}

function writeExpiryToHassoG(items) {
  try {
    var ss = SpreadsheetApp.openById(HASSO_SHEET_ID);
    var sheet = ss.getSheetByName(HASSO_SHEET_NAME);

    if (!sheet) return '⚠️ 発送Gシートが見つかりません';

    var lastRow = sheet.getLastRow();
    var aCol = sheet.getRange(1, 1, lastRow, 1).getValues();

    var rowOf = {};

    for (var i = 0; i < aCol.length; i++) {
      var key = String(aCol[i][0]).trim();

      if (key !== '' && !(key in rowOf)) {
        rowOf[key] = i + 1;
      }
    }

    var written = [];
    var notFound = [];
    var skipped = [];
    var badKanri = [];
    var notInBox = [];

    for (var j = 0; j < items.length; j++) {
      var it = items[j];

      // 5桁でない番号は書き込まない
      if (!it.kanriValid) {
        badKanri.push(it.kanri);
        continue;
      }

      // この箱の番号でない＝誤読の可能性
      if (it.inExpected === false) {
        notInBox.push(it.kanri);
        continue;
      }

      if (!it.valid || !it.formatted) {
        skipped.push(it.kanri);
        continue;
      }

      var row = rowOf[it.kanri];

      if (!row) {
        notFound.push(it.kanri);
        continue;
      }

      sheet.getRange(row, 8).setValue(it.formatted); // H=8列目（MM/DD/YYYY）
      written.push(it.kanri + '→' + it.display);     // 表示は yyyy/mm/dd
    }

    var msg = '✅ 賞味期限を発送Gに登録しました\n─────────────────\n';

    msg += '登録: ' + (written.length > 0 ? written.join('\n') : 'なし');

    if (badKanri.length > 0) {
      msg += '\n─────────────────\n⚠️ 管理番号が5桁でないため書き込まず: ' + badKanri.join(', ')
        + '\n（撮り直すか手入力してください）';
    }

    if (notInBox.length > 0) {
      msg += '\n─────────────────\n⚠️ この箱の管理番号でないため書き込まず: ' + notInBox.join(', ')
        + '\n（誤読の可能性。撮り直すか確認してください）';
    }

    if (skipped.length > 0) {
      msg += '\n⏭ 日付不正でスキップ: ' + skipped.join(', ');
    }

    if (notFound.length > 0) {
      msg += '\n⚠️ 発送Gに見つからない管理番号: ' + notFound.join(', ');
    }

    return msg;

  } catch (err) {
    console.log('writeExpiryToHassoG error: ' + err.toString());
    return '⚠️ 賞味期限書き込みエラー: ' + err.toString();
  }
}

// ============================================================
// 送料確認まわり
// ============================================================

function parseDimensionsWithClaude(text) {
  var apiKey = PropertiesService.getScriptProperties().getProperty('CLAUDE_API_KEY');

  var prompt = 'LINEメッセージから梱包情報を抽出してJSONのみ返してください。説明不要。\n\n'
    + 'メッセージ: "' + text + '"\n\n'
    + 'このメッセージは次の構造（複数行）であることが多い:\n'
    + '  1行目: 管理番号（5桁の数字。スペース区切りで複数並ぶことがある）\n'
    + '  2行目: 「食品 あり」または「食品 なし」\n'
    + '  3行目: サイズ3つ（例「39 54 31」スペース区切り、または「39×54×31」）\n'
    + '  4行目: 重量（例「10kg」「10000g」）\n'
    + '  5行目: 国（例「US」「JP」）\n'
    + 'ただし1行にまとまっている場合もある（例「14525 60×40×40 14350g US」）。どちらの形式も読み取ること。\n\n'
    + '返答形式（このJSONのみ、余分なテキストなし）:\n'
    + '{"length_cm":数値,"width_cm":数値,"height_cm":数値,"weight_g":数値,"country_code":"2文字ISO","zip":"郵便番号または空文字","order_id":"最初の管理番号または空文字","has_food":true または false}\n\n'
    + '変換ルール:\n'
    + '- 管理番号（5桁の数字）はサイズではない。サイズは「食品 あり/なし」の後にある数字3つ、または×で区切られた数字3つ\n'
    + '- 日本/国内/JP→JP, アメリカ/米国/USA→US, イギリス/UK→GB, オーストラリア→AU, カナダ→CA, ドイツ→DE, フランス→FR, オランダ→NL, シンガポール→SG, 台湾→TW, 韓国→KR, 香港→HK\n'
    + '- kgはgに変換（1.5kg→1500、10kg→10000）\n'
    + '- サイズ順不明なら最大値=length, 中間=width, 最小=height\n'
    + '- 国が不明ならUS\n'
    + '- サイズか重量が本当に見つからないときだけnull（文字列）を返す\n'
    + '- order_idは最初の管理番号（5桁）\n'
    + '- 管理番号がなければ空文字\n'
    + '- 「食品 あり」「食品あり」等があれば has_food=true、「食品 なし」「食品なし」等なら has_food=false、記載なしはfalse';

  try {
    var response = UrlFetchApp.fetch('https://api.anthropic.com/v1/messages', {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
        'x-api-key': apiKey,
        'anthropic-version': '2023-06-01'
      },
      payload: JSON.stringify({
        model: 'claude-haiku-4-5-20251001',
        max_tokens: 200,
        messages: [
          {
            role: 'user',
            content: prompt
          }
        ]
      }),
      muteHttpExceptions: true
    });

    var result = JSON.parse(response.getContentText());

    if (!result.content || !result.content[0]) return null;

    var rawText = result.content[0].text.trim();
    var start = rawText.indexOf('{');
    var end = rawText.lastIndexOf('}');

    if (start === -1 || end === -1) return null;

    var parsed = JSON.parse(rawText.substring(start, end + 1));

    if (!parsed.length_cm || !parsed.width_cm || !parsed.height_cm || !parsed.weight_g) {
      return null;
    }

    return {
      length: parsed.length_cm,
      width: parsed.width_cm,
      height: parsed.height_cm,
      weight: parsed.weight_g,
      country: String(parsed.country_code || 'US').trim().toUpperCase(),
      zip: parsed.zip || '',
      order_id: parsed.order_id || '',
      has_food: parsed.has_food === true
    };

  } catch (err) {
    console.log('Claude error: ' + err.toString());
    return null;
  }
}

function getDefaultAddress(country, zip) {
  var addresses = {
    'US': {
      country: 'US',
      full_name: 'Test Buyer',
      phone: '2125551234',
      address1: '1600 Amphitheatre Pkwy',
      city: 'Mountain View',
      province: 'CA',
      zip: zip || '94043'
    },
    'GB': {
      country: 'GB',
      full_name: 'Test Buyer',
      phone: '02071234567',
      address1: '10 Downing Street',
      city: 'London',
      province: 'ENG',
      zip: zip || 'SW1A2AA'
    },
    'AU': {
      country: 'AU',
      full_name: 'Test Buyer',
      phone: '0298765432',
      address1: '1 Martin Place',
      city: 'Sydney',
      province: 'NSW',
      zip: zip || '2000'
    },
    'CA': {
      country: 'CA',
      full_name: 'Test Buyer',
      phone: '4165551234',
      address1: '100 King Street West',
      city: 'Toronto',
      province: 'ON',
      zip: zip || 'M5X1A9'
    },
    'DE': {
      country: 'DE',
      full_name: 'Test Buyer',
      phone: '03012345678',
      address1: 'Unter den Linden 1',
      city: 'Berlin',
      zip: zip || '10117'
    },
    'FR': {
      country: 'FR',
      full_name: 'Test Buyer',
      phone: '0142345678',
      address1: '1 Rue de Rivoli',
      city: 'Paris',
      zip: zip || '75001'
    },
    'NL': {
      country: 'NL',
      full_name: 'Test Buyer',
      phone: '0201234567',
      address1: 'Dam 1',
      city: 'Amsterdam',
      zip: zip || '1012JS'
    },
    'SG': {
      country: 'SG',
      full_name: 'Test Buyer',
      phone: '65123456',
      address1: '1 Raffles Place',
      city: 'Singapore',
      zip: zip || '048616'
    },
    'TW': {
      country: 'TW',
      full_name: 'Test Buyer',
      phone: '0212345678',
      address1: '1 Zhongzheng Road',
      city: 'Taipei',
      zip: zip || '100'
    },
    'KR': {
      country: 'KR',
      full_name: 'Test Buyer',
      phone: '0212345678',
      address1: '1 Sejong-daero',
      city: 'Seoul',
      zip: zip || '04524'
    },
    'HK': {
      country: 'HK',
      full_name: 'Test Buyer',
      phone: '21234567',
      address1: '1 Connaught Place',
      city: 'Hong Kong',
      zip: zip || '999077'
    }
  };

  var addr = addresses[country];

  if (!addr) {
    addr = {
      country: country,
      full_name: 'Test Buyer',
      phone: '0000000000',
      address1: '1 Main Street',
      city: 'City',
      zip: zip || '00000'
    };
  }

  return addr;
}

function getShipAndCoRates(dims) {
  var token = PropertiesService.getScriptProperties().getProperty('SHIPANDCO_TOKEN');

  var payload = {
    setup: {
      currency: 'USD'
    },
    from_address: {
      country: 'JP',
      full_name: 'Sky Crew Japan',
      company: 'Sky Crew Japan',
      phone: '0489000000',
      email: 'shinicchee@gmail.com',
      address1: '1-1 Soka',
      city: 'Soka',
      province: 'Saitama',
      zip: '340-0001'
    },
    to_address: getDefaultAddress(dims.country, dims.zip),
    parcels: [
      {
        width: dims.width,
        height: dims.height,
        depth: dims.length,
        weight: dims.weight,
        amount: 1
      }
    ],
    products: [
      {
        name: 'Item',
        price: 10,
        quantity: 1,
        country: 'JP',
        currency: 'USD'
      }
    ]
  };

  try {
    var response = UrlFetchApp.fetch('https://api.shipandco.com/v1/rates', {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
        'x-access-token': token
      },
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    });

    var sc = response.getResponseCode();

    if (sc !== 200) {
      return {
        error: 'HTTP ' + sc
      };
    }

    return {
      success: true,
      data: JSON.parse(response.getContentText())
    };

  } catch (err) {
    return {
      error: err.toString()
    };
  }
}

function formatRatesMessage(ratesResult, dims) {
  var volWeight = Math.round(dims.length * dims.width * dims.height / 5000 * 1000);
  var chargeWeight = Math.max(dims.weight, volWeight);
  var isVol = volWeight > dims.weight;

  var lines = [];

  lines.push('📦 送料確認結果' + (dims.order_id ? ' ／ ' + dims.order_id : ''));
  lines.push('─────────────────');
  lines.push(dims.length + '×' + dims.width + '×' + dims.height + ' cm');
  lines.push('実重量: ' + dims.weight + 'g　容積: ' + volWeight + 'g' + (isVol ? ' ⚠️' : ''));
  lines.push('課金対象: ' + chargeWeight + 'g / 発送先: ' + dims.country);
  lines.push('─────────────────');

  if (ratesResult.error) {
    lines.push('❌ エラー: ' + ratesResult.error);
    return lines.join('\n');
  }

  var rateList = ratesResult.data.filter(function (r) {
    return !r.errors;
  });

  if (rateList.length === 0) {
    lines.push('利用可能なサービスがありませんでした');
    return lines.join('\n');
  }

  rateList.sort(function (a, b) {
    return a.price - b.price;
  });

  lines.push('【送料一覧】');

  rateList.slice(0, 10).forEach(function (r, i) {
    var name = r.service
      .replace('japanpost_', '郵便/')
      .replace('fedex_international_', 'FedEx/')
      .replace('dhl_express_', 'DHL/')
      .replace('ups_worldwide_', 'UPS/')
      .replace('ups_', 'UPS/');

    var surcharge = r.surcharges && r.surcharges.length > 0
      ? ' (+¥' + r.surcharges.reduce(function (s, x) {
        return s + x.price;
      }, 0) + ')'
      : '';

    lines.push((i + 1) + '. ' + name);
    lines.push('   ¥' + Math.round(r.price).toLocaleString() + surcharge);
  });

  return lines.join('\n');
}

function hasRateOptions(message) {
  return /\n1\.\s+/.test(String(message || ''));
}

function saveSoryoContext(kanriNumber, optionsText, dimensions, allKanriNumbers, hasFood) {
  var props = PropertiesService.getScriptProperties();
  var options = [];
  var lines = optionsText.split('\n');

  for (var i = 0; i < lines.length; i++) {
    var cm = lines[i].match(/^(\d+)\.\s+(.+)/);

    if (cm) {
      var price = '';

      if (i + 1 < lines.length) {
        var pm = lines[i + 1].match(/(¥[\d,]+)/);

        if (pm) {
          price = pm[1];
        }
      }

      options.push({
        num: cm[1],
        carrier: cm[2].trim(),
        price: price
      });
    }
  }

  props.setProperty(
    'soryo_context',
    JSON.stringify({
      kanriNumber: kanriNumber,
      allKanriNumbers: allKanriNumbers || kanriNumber,
      options: options,
      dimensions: dimensions || {},
      hasFood: hasFood === true,
      timestamp: new Date().toISOString()
    })
  );
}

function handleSoryoReply(replyText, replyToken) {
  var props = PropertiesService.getScriptProperties();
  var context = JSON.parse(props.getProperty('soryo_context') || 'null');

  if (!context) {
    replyToLine(replyToken, '⚠️ 送料選択の対象が見つかりません。\n先に送料確認を実行してください。');
    return;
  }

  var selected = context.options.find(function (o) {
    return o.num === replyText;
  });

  if (!selected) {
    replyToLine(replyToken, '⚠️ 番号が正しくありません。\n1〜' + context.options.length + ' の番号で返信してください。');
    return;
  }

  var parts = selected.carrier.split('/');
  var carrier = parts[0];
  var service = parts[1] || '';
  var price = selected.price.replace('¥', '').replace(/,/g, '');

  recordToSoryoSheet(
    context.kanriNumber,
    replyText,
    carrier,
    service,
    price,
    context.dimensions,
    context.allKanriNumbers
  );

  var hassoWarning = writeToHassoG(context.allKanriNumbers, context.dimensions);

  props.deleteProperty('soryo_context');

// US食品ありの場合は、送料選択後に賞味期限画像を依頼する
if (hassoWarning) {
  replyToLine(replyToken, hassoWarning);
} else if (context.hasFood === true) {
  replyToLine(replyToken, '賞味期限画像を送ってください');
} else {
  replyToLine(replyToken, '登録完了');
}
}

function recordToSoryoSheet(kanriNumber, selectedNum, carrier, service, price, dimensions, allKanriNumbers) {
  var ss = SpreadsheetApp.openById(SORYO_SHEET_ID);
  var sheet = ss.getSheetByName('FBA送料一覧');
  var data = sheet.getDataRange().getValues();

  for (var i = data.length - 1; i >= 1; i--) {
    if (String(data[i][1]) === String(kanriNumber)) {
      sheet.deleteRow(i + 1);
    }
  }

  var dim = dimensions || {};

  sheet.appendRow([
    new Date(),
    kanriNumber,
    dim.length || '',
    dim.width || '',
    dim.height || '',
    dim.weight || '',
    selectedNum,
    carrier,
    service,
    '¥' + Number(price).toLocaleString(),
    allKanriNumbers || kanriNumber,
    '選択済',
    ''
  ]);
}

function writeToHassoG(allKanriNumbers, dimensions) {
  try {
    var dim = dimensions || {};
    var nums = String(allKanriNumbers || '').split(/\s+/).filter(function (n) {
      return n !== '';
    });

    if (nums.length === 0) {
      return '⚠️ 発送G: 管理番号が空のため書き込みスキップ';
    }

    var groupStr = nums.join(' ');
    var minNum = nums.reduce(function (min, cur) {
      return Number(cur) < Number(min) ? cur : min;
    }, nums[0]);

    var ss = SpreadsheetApp.openById(HASSO_SHEET_ID);
    var sheet = ss.getSheetByName(HASSO_SHEET_NAME);

    if (!sheet) {
      return '⚠️ 発送G: シートが見つかりません';
    }

    var lastRow = sheet.getLastRow();

    if (lastRow < 1) {
      return '⚠️ 発送G: データがありません';
    }

    var aCol = sheet.getRange(1, 1, lastRow, 1).getValues();
    var rowOf = {};

    for (var i = 0; i < aCol.length; i++) {
      var key = String(aCol[i][0]).trim();

      if (key !== '' && !(key in rowOf)) {
        rowOf[key] = i + 1;
      }
    }

    var notFound = [];

    for (var j = 0; j < nums.length; j++) {
      var num = nums[j];
      var row = rowOf[num];

      if (!row) {
        notFound.push(num);
        continue;
      }

      // C列：同じ箱の管理番号グループ
      sheet.getRange(row, 3).setValue(groupStr);

      // 荷姿は最小管理番号の行だけに入れる
      if (num === minNum) {
        sheet.getRange(row, 4).setValue(dim.weight || '');
        sheet.getRange(row, 5).setValue(dim.length || '');
        sheet.getRange(row, 6).setValue(dim.width || '');
        sheet.getRange(row, 7).setValue(dim.height || '');
      }
    }

    if (notFound.length > 0) {
      return '⚠️ 発送Gに見つからない管理番号: ' + notFound.join(', ') + '\n（見つかった分は書き込み済み）';
    }

    return '';

  } catch (err) {
    return '⚠️ 発送G書き込みエラー: ' + err.toString();
  }
}

function deleteFromSoryoSheet(kanriNumber, replyToken) {
  var ss = SpreadsheetApp.openById(SORYO_SHEET_ID);
  var sheet = ss.getSheetByName('FBA送料一覧');
  var data = sheet.getDataRange().getValues();

  var deleted = 0;

  for (var i = data.length - 1; i >= 1; i--) {
    if (String(data[i][1]) === String(kanriNumber)) {
      sheet.deleteRow(i + 1);
      deleted++;
    }
  }

  if (deleted > 0) {
    replyToLine(
      replyToken,
      '🗑️ 削除完了\n─────────────────\n管理番号: ' + kanriNumber + '\n' + deleted + '件のデータを削除しました'
    );
  } else {
    replyToLine(replyToken, '⚠️ 該当データなし\n管理番号: ' + kanriNumber);
  }
}

function checkSoryoSheet(kanriNumber, replyToken) {
  var ss = SpreadsheetApp.openById(SORYO_SHEET_ID);
  var sheet = ss.getSheetByName('FBA送料一覧');
  var data = sheet.getDataRange().getValues();

  var found = null;

  for (var i = 1; i < data.length; i++) {
    if (String(data[i][1]) === String(kanriNumber)) {
      found = data[i];
      break;
    }
  }

  if (found) {
    replyToLine(
      replyToken,
      '📋 登録内容確認\n'
      + '─────────────────\n'
      + '管理番号: ' + found[1] + '\n'
      + '全管理番号: ' + found[10] + '\n'
      + '荷姿: ' + found[2] + '×' + found[3] + '×' + found[4] + 'cm / ' + found[5] + 'kg\n'
      + '採用業者: ' + found[7] + '\n'
      + 'サービス: ' + found[8] + '\n'
      + '送料: ' + found[9] + '\n'
      + 'ステータス: ' + found[11]
    );
  } else {
    replyToLine(replyToken, '⚠️ 該当データなし\n管理番号: ' + kanriNumber);
  }
}

function replyToLine(replyToken, text) {
  var token = PropertiesService.getScriptProperties().getProperty('LINE_TOKEN');

  UrlFetchApp.fetch('https://api.line.me/v2/bot/message/reply', {
    method: 'POST',
    headers: {
      'Content-Type': 'application/json',
      'Authorization': 'Bearer ' + token
    },
    payload: JSON.stringify({
      replyToken: replyToken,
      messages: [
        {
          type: 'text',
          text: text
        }
      ]
    }),
    muteHttpExceptions: true
  });
}

function pushToLine(userId, text) {
  var token = PropertiesService.getScriptProperties().getProperty('LINE_TOKEN');

  UrlFetchApp.fetch('https://api.line.me/v2/bot/message/push', {
    method: 'POST',
    headers: {
      'Content-Type': 'application/json',
      'Authorization': 'Bearer ' + token
    },
    payload: JSON.stringify({
      to: userId,
      messages: [
        {
          type: 'text',
          text: text
        }
      ]
    }),
    muteHttpExceptions: true
  });
}

function getHelpMessage() {
  return [
    '📦 送料チェッカー',
    '─────────────────',
    '【USなど海外向け：送料確認】',
    '「14525 60×40×40 14350g US」',
    '「14111 14112 14113 60×40×40 14350g US」',
    '数字のみ → 送料番号を選択',
    '',
    '【JP向け：荷姿登録のみ】',
    '「14525 60×40×40 14350g JP」',
    '「14111 14112 14113 食品なし 60×40×40 10kg JP」',
    'JPは送料計算なし・FBA送料一覧記録なし',
    '',
    '【賞味期限】',
    '食品ありの場合、賞味期限画像を送信',
    '→ 読み取り結果を確認し「OK」で登録',
    '',
    '【その他】',
    '「削除 14525」/「確認 14525」',
    '─────────────────'
  ].join('\n');
}

// ============================================================
// テスト用
// ============================================================

function testExpiryValidate() {
  // 妥当性チェックの動作確認
  console.log('2026/9/31 valid? ' + isValidDate(2026, 9, 31)); // false
  console.log('2026/9/30 valid? ' + isValidDate(2026, 9, 30)); // true
  console.log('2027/2/29 valid? ' + isValidDate(2027, 2, 29)); // false
  console.log('2028/2/29 valid? ' + isValidDate(2028, 2, 29)); // true
  console.log('format 2027/1/13 → ' + pad2(1) + '/' + pad2(13) + '/' + 2027);
}

/**
 * Amazon evaluation request automation for Amazon.co.jp (JAPAN).  ★修正版（丸ごと上書き用）
 * 設置先: A_FBA自動化_メイン (1xJ8PiYwv_T_FVLEXBUi1Fs8MYqAav7CjknzqieLhbkg)
 *
 * 変更点（ご指示8点）:
 *   1 runReviewRequestDryRun(): not_eligible を無条件スキップせず、注文後 RECHECK_WINDOW_DAYS(=40)
 *     日以内の not_eligible は再チェック対象に含める。
 *   2 persistNgDecisionsToExclusions_(): 廃止（関数削除）。NGを評価依頼_除外へ appendRow せず、
 *     注文側 exclude_flag/exclude_reason も書かない。
 *   3 buildReviewRequestCandidates() 冒頭の persistNgDecisionsToExclusions_() 呼び出し削除。
 *   4 sendApprovedReviewRequests() 冒頭の persistNgDecisionsToExclusions_() 呼び出し削除。
 *   5 send_decision=NG は評価依頼_結果に残すだけ（OKのみ送信＝送信対象外）。恒久除外には入れない。
 *   6 buildReviewRequestCandidates() は既存 send_decision を引き継ぐ（NGが空欄に戻らない）。
 *   7 cleanupFalseExclusions_NGdecision(): 誤除外2件(250-6490546-4785421 / 249-5741099-2453428)
 *     を復旧。DRY_RUN=true で確認 → false で本復旧。
 *   8 日本版設定維持: marketplace_id=A1VC38T7YXB528 / endpoint=FE / region=us-west-2。
 *
 * ★注意: 末尾の scheduledXXX ラッパーは「コアを呼ぶだけ」に再構成したものです。実際のラッパーに
 *   独自処理（ロック/バッチ/通知等）がある場合は、その関数だけ既存版を残してください。
 *
 * Required Script Properties:
 * - SPAPI_CLIENT_ID
 * - SPAPI_CLIENT_SECRET
 * - SPAPI_REFRESH_TOKEN
 * - LINE_CHANNEL_ACCESS_TOKEN (optional)
 * - LINE_USER_ID (optional)
 */
const REVIEW_JP = {
  spreadsheetUrl: 'https://docs.google.com/spreadsheets/d/1xJ8PiYwv_T_FVLEXBUi1Fs8MYqAav7CjknzqieLhbkg/edit',
  sheets: {
    orders: '評価依頼_注文',
    exclusions: '評価依頼_除外',
    results: '評価依頼_結果',
    logs: '評価依頼_ログ',
    flow: '評価依頼_フロー',
    settings: '評価依頼_設定',
  },
  defaults: {
    marketplace_id: 'A1VC38T7YXB528',
    spapi_endpoint: 'https://sellingpartnerapi-fe.amazon.com',
    aws_region: 'us-west-2',
    order_fetch_days: '60',
    solicitation_delay_ms: '1500',
    order_items_batch_size: '30',
    dry_run_batch_size: '25',
    refund_email_query: 'newer_than:180d (refund OR refunded OR return OR returned OR reimbursement OR reimbursed OR 返金 OR 返品 OR 払い戻し)',
    refund_email_batch_size: '50',
    dry_run: 'FALSE',
    send_enabled: 'TRUE',
    exclude_active_value: 'TRUE',
    source: 'SP-API Solicitations JP',
  },
};

// ★1: 注文日から何日以内の not_eligible を再チェック対象にするか（評価依頼の窓。配達ラグ込み）。35 にしたい場合はここを 35 に。
const RECHECK_WINDOW_DAYS = 40;

const ORDER_HEADERS = [
  'amazon_order_id',
  'order_date',
  'sku',
  'asin',
  'item_name',
  'order_status',
  'exclude_flag',
  'exclude_reason',
  'solicitation_status',
  'last_checked_at',
  'last_result',
  'request_sent',
  'sent_at',
  'send_result',
  'notes',
  'refund_flag',
  'refund_amount',
  'refund_checked_at',
  'refund_reason',
];
const RESULT_HEADERS = [
  'send_decision',
  'amazon_order_id',
  'seller_order_url',
  'order_date',
  'asin',
  'sku',
  'item_name',
  'solicitation_status',
  'exclude_check',
  'refund_flag',
  'request_sent',
  'sent_at',
  'send_result',
  'reason',
  'last_checked_at',
];
const EXCLUSION_HEADERS = [
  'type',
  'value',
  'reason',
  'active',
  'updated_at',
  'notes',
];
const LOG_HEADERS = [
  'timestamp',
  'level',
  'function_name',
  'amazon_order_id',
  'message',
  'details',
];
function onOpen() {
  const ui = SpreadsheetApp.getUi();

  ui.createMenu('評価依頼_JP')
    .addItem('認証設定チェック', 'checkReviewRequestProperties')
    .addSeparator()
    .addItem('注文一覧取得', 'fetchRecentFbaShippedOrders')
    .addItem('ASIN・商品名取得', 'fetchOrderItemsForReviewOrders')
    .addItem('返金メール取込', 'importRefundEmailsToExclusions')
    .addItem('ドライラン実行', 'runReviewRequestDryRun')
    .addItem('送信候補一覧作成', 'buildReviewRequestCandidates')
    .addItem('送信対象プレビュー(OK)', 'previewApprovedSendable')
    .addItem('候補トレース診断', 'debugBuildCandidatesTrace')
    .addSeparator()
    .addItem('■■■OK分だけ評価依頼送信', 'sendApprovedReviewRequests')
    .addToUi();

  ui.createMenu('ＦＤＡ書類作成')
    .addItem('メール情報取得', 'fetchEmailData')
    .addToUi();
}
function checkReviewRequestProperties() {
  const required = [
    'SPAPI_CLIENT_ID',
    'SPAPI_CLIENT_SECRET',
    'SPAPI_REFRESH_TOKEN',
  ];
  const props = PropertiesService.getScriptProperties();
  const missing = required.filter((key) => !props.getProperty(key));
  ensureBaseSheets_();
  upsertDefaultSettings_();
  const settings = getSettings_();
  const wrongSettings = [];
  Object.keys(REVIEW_JP.defaults).forEach((key) => {
    if (normalizeSettingValue_(settings[key]) !== normalizeSettingValue_(REVIEW_JP.defaults[key])) {
      wrongSettings.push(`${key}: ${displaySettingValue_(settings[key])} -> ${REVIEW_JP.defaults[key]}`);
    }
  });
  if (missing.length || wrongSettings.length) {
    log_('WARN', 'checkReviewRequestProperties', '', 'Configuration check found issues', {
      missing,
      wrongSettings,
    });
    SpreadsheetApp.getUi().alert(
      `確認が必要です。\n\nMissing Script Properties:\n${missing.join('\n') || 'なし'}\n\nSettings mismatch:\n${wrongSettings.join('\n') || 'なし'}`
    );
    return;
  }
  log_('INFO', 'checkReviewRequestProperties', '', 'Configuration check passed', {});
  SpreadsheetApp.getUi().alert('JP版の認証設定と基本設定はOKです。');
}
function testSpApiOrdersConnection() {
  ensureBaseSheets_();
  upsertDefaultSettings_();
  try {
    const token = getLwaAccessToken_();
    log_('INFO', 'testSpApiOrdersConnection', '', 'LWA access token was generated', {
      tokenLength: token ? token.length : 0,
    });
    const settings = getSettings_();
    const response = spApiRequest_('get', '/orders/v0/orders', {
      MarketplaceIds: settings.marketplace_id,
      CreatedAfter: new Date(Date.now() - 7 * 24 * 60 * 60 * 1000).toISOString(),
      OrderStatuses: 'Shipped',
      FulfillmentChannels: 'AFN',
    });
    log_('INFO', 'testSpApiOrdersConnection', '', 'Orders API connection succeeded', response);
    SpreadsheetApp.getUi().alert('Orders API接続OKです。次に fetchRecentFbaShippedOrders を実行してください。');
  } catch (error) {
    log_('ERROR', 'testSpApiOrdersConnection', '', error.message, {});
    SpreadsheetApp.getUi().alert(`Orders API接続テストでエラーです。\n\n${error.message}`);
  }
}
function forceApplyJpReviewRequestSettings() {
  ensureBaseSheets_();
  writeJpSettings_();
  log_('INFO', 'forceApplyJpReviewRequestSettings', '', 'JP settings were applied explicitly', {});
  SpreadsheetApp.getUi().alert('JP版の評価依頼設定を書き込みました。もう一度 checkReviewRequestProperties を実行してください。');
}
function fetchRecentFbaShippedOrders() {
  ensureBaseSheets_();
  upsertDefaultSettings_();
  const settings = getSettings_();
  const days = Number(settings.order_fetch_days || 60);
  const createdAfter = new Date(Date.now() - days * 24 * 60 * 60 * 1000).toISOString();
  const path = '/orders/v0/orders';
  const query = {
    MarketplaceIds: settings.marketplace_id,
    FulfillmentChannels: 'AFN',
    OrderStatuses: 'Shipped',
    CreatedAfter: createdAfter,
  };
  let added = 0;
  let nextToken = '';
  const sheet = getSheet_(REVIEW_JP.sheets.orders);
  const existing = readTable_(sheet);
  const existingIds = new Set(existing.rows.map((row) => row.amazon_order_id).filter(Boolean));
  do {
    const pageQuery = nextToken ? { NextToken: nextToken } : query;
    const response = spApiRequest_('get', path, pageQuery);
    const orders = response.Orders || [];
    const rows = [];
    orders.forEach((order) => {
      const orderId = order.AmazonOrderId;
      if (!orderId || existingIds.has(orderId)) return;
      existingIds.add(orderId);
      rows.push([
        orderId,
        order.PurchaseDate || '',
        '',
        '',
        '',
        order.OrderStatus || '',
        '',
        '',
        '',
        '',
        '',
        false,
        '',
        '',
        '',
        false,
        '',
        '',
        '',
      ]);
    });
    if (rows.length) {
      sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, ORDER_HEADERS.length).setValues(rows);
      added += rows.length;
    }
    nextToken = response.NextToken || '';
    Utilities.sleep(500);
  } while (nextToken);
  log_('INFO', 'fetchRecentFbaShippedOrders', '', `Fetched recent FBA shipped orders. Added: ${added}`, {});
}
function fetchOrderItemsForReviewOrders() {
  ensureBaseSheets_();
  upsertDefaultSettings_();
  const settings = getSettings_();
  const batchSize = Number(settings.order_items_batch_size || 30);
  const sheet = getSheet_(REVIEW_JP.sheets.orders);
  const table = readTable_(sheet);
  const targets = [];
  table.rows.forEach((row, index) => {
    if (targets.length >= batchSize) return;
    if (row.amazon_order_id && (!row.sku || !row.asin || !row.item_name)) {
      targets.push({ row, sheetRow: index + 2 });
    }
  });
  targets.forEach((target) => {
    try {
      const path = `/orders/v0/orders/${encodeURIComponent(target.row.amazon_order_id)}/orderItems`;
      const response = spApiRequest_('get', path, {});
      const item = (response.OrderItems || [])[0] || {};
      sheet.getRange(target.sheetRow, 3, 1, 3).setValues([[
        item.SellerSKU || target.row.sku || '',
        item.ASIN || target.row.asin || '',
        item.Title || target.row.item_name || '',
      ]]);
      log_('INFO', 'fetchOrderItemsForReviewOrders', target.row.amazon_order_id, 'Fetched order item', {});
      Utilities.sleep(500);
    } catch (error) {
      log_('ERROR', 'fetchOrderItemsForReviewOrders', target.row.amazon_order_id, error.message, {});
    }
  });
}
function importRefundEmailsToExclusions() {
  ensureBaseSheets_();
  upsertDefaultSettings_();
  const settings = getSettings_();
  const query = settings.refund_email_query || REVIEW_JP.defaults.refund_email_query;
  const limit = Number(settings.refund_email_batch_size || 50);
  const threads = GmailApp.search(query, 0, limit);
  const orderIds = new Set();
  const orderIdPattern = /\b\d{3}-\d{7}-\d{7}\b/g;
  threads.forEach((thread) => {
    thread.getMessages().forEach((message) => {
      const text = `${message.getSubject()}\n${message.getPlainBody()}`;
      const matches = text.match(orderIdPattern) || [];
      matches.forEach((orderId) => orderIds.add(orderId));
    });
  });
  const ordersSheet = getSheet_(REVIEW_JP.sheets.orders);
  const ordersTable = readTable_(ordersSheet);
  const exclusionsSheet = getSheet_(REVIEW_JP.sheets.exclusions);
  const exclusionsTable = readTable_(exclusionsSheet);
  const existingExclusions = new Set(
    exclusionsTable.rows
      .filter((row) => String(row.active).toUpperCase() !== 'FALSE')
      .map((row) => `${row.type}:${row.value}`)
  );
  const exclusionRows = [];
  ordersTable.rows.forEach((row, index) => {
    if (!orderIds.has(row.amazon_order_id)) return;
    ordersSheet.getRange(index + 2, 16, 1, 4).setValues([[true, '', new Date(), 'refund email matched']]);
    const key = `order_id:${row.amazon_order_id}`;
    if (!existingExclusions.has(key)) {
      existingExclusions.add(key);
      exclusionRows.push(['order_id', row.amazon_order_id, 'refund email matched', true, new Date(), '']);
    }
  });
  if (exclusionRows.length) {
    exclusionsSheet
      .getRange(exclusionsSheet.getLastRow() + 1, 1, exclusionRows.length, EXCLUSION_HEADERS.length)
      .setValues(exclusionRows);
  }
  log_('INFO', 'importRefundEmailsToExclusions', '', `Imported refund email exclusions: ${orderIds.size}`, {});
}
function debugRefundEmailForOrder() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.prompt('返金メール確認', '確認するAmazon注文番号を入力してください。', ui.ButtonSet.OK_CANCEL);
  if (response.getSelectedButton() !== ui.Button.OK) return;
  const orderId = response.getResponseText().trim();
  if (!orderId) return;
  const compactOrderId = orderId.replace(/-/g, '');
  const query = `newer_than:180d (${orderId} OR ${compactOrderId})`;
  const threads = GmailApp.search(query, 0, 20);
  const summaries = [];
  threads.forEach((thread) => {
    thread.getMessages().forEach((message) => {
      const subject = message.getSubject();
      const body = message.getPlainBody();
      const searchableText = `${subject}\n${body}`.toLowerCase();
      const refundLikeKeywords = ['refund', 'refunded', 'return', 'returned', 'reimbursement', 'reimbursed', '返金', '返品', '払い戻し'];
      const refundLike = refundLikeKeywords.some((keyword) => searchableText.indexOf(keyword.toLowerCase()) !== -1);
      summaries.push({
        date: message.getDate(),
        subject,
        refundLike,
      });
    });
  });
  log_('INFO', 'debugRefundEmailForOrder', orderId, `Gmail matches: ${summaries.length}`, summaries);
  ui.alert(
    `検索条件:\n${query}\n\n一致メール: ${summaries.length}件\n\n` +
    summaries.slice(0, 10).map((item) => `${item.date}\n${item.refundLike ? '[refund-like]' : '[not refund-like]'} ${item.subject}`).join('\n\n')
  );
}
function importRefundEmailForSingleOrder() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.prompt('単一注文の返金取込', '返金確認するAmazon注文番号を入力してください。', ui.ButtonSet.OK_CANCEL);
  if (response.getSelectedButton() !== ui.Button.OK) return;
  const orderId = response.getResponseText().trim();
  if (!orderId) return;
  const compactOrderId = orderId.replace(/-/g, '');
  const query = `newer_than:180d (${orderId} OR ${compactOrderId})`;
  const threads = GmailApp.search(query, 0, 20);
  const refundLikeKeywords = ['refund', 'refunded', 'return', 'returned', 'reimbursement', 'reimbursed', '返金', '返品', '払い戻し'];
  let matched = false;
  const subjects = [];
  threads.forEach((thread) => {
    thread.getMessages().forEach((message) => {
      const searchableText = `${message.getSubject()}\n${message.getPlainBody()}`.toLowerCase();
      const refundLike = refundLikeKeywords.some((keyword) => searchableText.indexOf(keyword.toLowerCase()) !== -1);
      if (!refundLike) return;
      matched = true;
      subjects.push(message.getSubject());
    });
  });
  if (!matched) {
    log_('WARN', 'importRefundEmailForSingleOrder', orderId, 'No refund-like email found for order', { query });
    ui.alert(`返金らしいメールは見つかりませんでした。\n\n検索条件:\n${query}`);
    return;
  }
  markOrderRefunded_(orderId, `refund email matched: ${subjects[0] || ''}`);
  log_('INFO', 'importRefundEmailForSingleOrder', orderId, 'Refund imported for single order', { query, subjects });
  ui.alert(`返金検知を反映しました。\n\n注文番号:\n${orderId}`);
}
function forceExcludeSingleOrder() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.prompt('単一注文の強制除外', '除外するAmazon注文番号を入力してください。', ui.ButtonSet.OK_CANCEL);
  if (response.getSelectedButton() !== ui.Button.OK) return;
  const orderId = response.getResponseText().trim();
  if (!orderId) return;
  markOrderRefunded_(orderId, 'manual exclusion / refund suspected');
  log_('INFO', 'forceExcludeSingleOrder', orderId, 'Order was manually excluded', {});
  ui.alert(`除外を反映しました。\n\n注文番号:\n${orderId}\n\n次に buildReviewRequestCandidates を実行してください。`);
}
function markOrderRefunded_(orderId, reason) {
  const ordersSheet = getSheet_(REVIEW_JP.sheets.orders);
  const ordersTable = readTable_(ordersSheet);
  const headerIndexes = getHeaderIndexes_(ordersSheet);
  let updated = false;
  ordersTable.rows.forEach((row, index) => {
    if (String(row.amazon_order_id) !== String(orderId)) return;
    setCellByHeader_(ordersSheet, headerIndexes, index + 2, 'refund_flag', true);
    setCellByHeader_(ordersSheet, headerIndexes, index + 2, 'refund_checked_at', new Date());
    setCellByHeader_(ordersSheet, headerIndexes, index + 2, 'refund_reason', reason);
    setCellByHeader_(ordersSheet, headerIndexes, index + 2, 'exclude_flag', true);
    setCellByHeader_(ordersSheet, headerIndexes, index + 2, 'exclude_reason', reason);
    updated = true;
  });
  const exclusionsSheet = getSheet_(REVIEW_JP.sheets.exclusions);
  const exclusionsTable = readTable_(exclusionsSheet);
  const exists = exclusionsTable.rows.some((row) => {
    return ['order_id', 'amazon_order_id', 'order', '注文番号'].includes(String(row.type).toLowerCase().trim()) &&
      String(row.value).trim() === String(orderId).trim() &&
      String(row.active).toUpperCase() !== 'FALSE';
  });
  if (!exists) {
    exclusionsSheet.appendRow(['order_id', orderId, reason, true, new Date(), '']);
  }
  if (!updated) {
    log_('WARN', 'markOrderRefunded_', orderId, 'Refund exclusion added but order was not found in orders sheet', {});
  }
}
function runReviewRequestDryRun() {
  ensureBaseSheets_();
  upsertDefaultSettings_();
  const settings = getSettings_();
  const batchSize = Number(settings.dry_run_batch_size || 25);
  const delay = Number(settings.solicitation_delay_ms || 1500);
  const sheet = getSheet_(REVIEW_JP.sheets.orders);
  const table = readTable_(sheet);
  const exclusions = getActiveExclusions_();
  const targets = [];
  table.rows.forEach((row, index) => {
    if (targets.length >= batchSize) return;
    if (!row.amazon_order_id) return;
    if (isTrue_(row.request_sent)) return;
    if (isTrue_(row.refund_flag)) return;
    if (isExcluded_(row, exclusions)) return;

    // ★1 修正: not_eligible を無条件スキップしない。窓内(注文後 RECHECK_WINDOW_DAYS 日以内)の not_eligible は再チェックする。
    const status = String(row.solicitation_status || '').toLowerCase();
    if (status === 'eligible' || status === 'excluded') return; // 候補/除外済みは再確認不要
    if (status === 'not_eligible') {
      const ageDays = orderAgeDays_(row.order_date);
      if (ageDays === null || ageDays > RECHECK_WINDOW_DAYS) return; // 窓外のみ確定スキップ
    }
    // '' / 'error' / 窓内 not_eligible はここを通過して再チェック対象になる

    targets.push({ row, sheetRow: index + 2 });
  });
  targets.forEach((target) => {
    try {
      const response = getSolicitationActions_(target.row.amazon_order_id);
      const actions = getSolicitationActionNames_(response);
      const eligible = actions.indexOf('productReviewAndSellerFeedback') !== -1;
      sheet.getRange(target.sheetRow, 9, 1, 3).setValues([[
        eligible ? 'eligible' : 'not_eligible',
        new Date(),
        JSON.stringify(response),
      ]]);
      log_('INFO', 'runReviewRequestDryRun', target.row.amazon_order_id, eligible ? 'eligible' : 'not_eligible', {});
      Utilities.sleep(delay);
    } catch (error) {
      sheet.getRange(target.sheetRow, 9, 1, 3).setValues([['error', new Date(), error.message]]);
      log_('ERROR', 'runReviewRequestDryRun', target.row.amazon_order_id, error.message, {});
      Utilities.sleep(delay);
    }
  });
}
function buildReviewRequestCandidates() {
  ensureBaseSheets_();
  upsertDefaultSettings_();
  // persistNgDecisionsToExclusions_() は呼ばない（NGは結果に残すだけ。恒久除外しない）。

  const ordersTable = readTable_(getSheet_(REVIEW_JP.sheets.orders));
  const exclusions = getActiveExclusions_();
  const resultSheet = getSheet_(REVIEW_JP.sheets.results);

  // 既存の send_decision(OK/NG) を引き継ぐ（NGが次回作成で空欄に戻らないようにする）。
  const priorDecisions = new Map();
  readTable_(resultSheet).rows.forEach((row) => {
    const id = String(row.amazon_order_id || '').trim();
    const decision = String(row.send_decision || '').trim();
    if (id && decision) priorDecisions.set(id, decision);
  });

  clearResultRows_(resultSheet);

  // ★v8: 診断カウンタ
  let cEligible = 0, dNotEligible = 0, dSent = 0, dRefund = 0, dExclude = 0, dReviewWindowError = 0;
  const dropEligible = [];
  const rows = [];
  ordersTable.rows.forEach((row) => {
    const orderId = String(row.amazon_order_id || '');
    const oid = orderId.trim();
    if (!oid) return;

    const status = String(row.solicitation_status || '').trim().toLowerCase();
    const rs = isTrue_(row.request_sent);
    const rf = isTrue_(row.refund_flag);
    const ef = isTrue_(row.exclude_flag);
    const exc = isExcluded_(row, exclusions);
    if (status === 'eligible') cEligible++;

    Logger.log(oid + ' | eligible=' + (status === 'eligible')
      + ' request_sent=' + rs + ' refund_flag=' + rf
      + ' exclude_flag=' + ef + ' isExcluded_=' + exc);

    let drop = '';
    if (status !== 'eligible') { drop = 'not_eligible(' + (status || 'empty') + ')'; dNotEligible++; }
    else if (rs) { drop = 'request_sent=TRUE'; dSent++; }
    else if (rf) { drop = 'refund_flag=TRUE'; dRefund++; }
    else if (exc) {
      const sub = [];
      if (ef) sub.push('exclude_flag');
      if (exclusions.orderIds.has(oid)) sub.push('除外order_id');
      if (exclusions.asins.has(String(row.asin || '').trim())) sub.push('除外asin');
      drop = 'isExcluded[' + sub.join('/') + ']';
      dExclude++;
    }
    else if (hasReviewWindowError_(row)) {
      drop = 'review_window_error';
      dReviewWindowError++;
    }
    if (drop) {
      if (status === 'eligible') dropEligible.push(oid + '(asin=' + String(row.asin || '').trim() + ')→' + drop);
      return;
    }

    rows.push([
      priorDecisions.get(oid) || '',
      orderId,
      buildSellerOrderUrl_(orderId),
      row.order_date || '',
      row.asin || '',
      row.sku || '',
      row.item_name || '',
      row.solicitation_status || '',
      'OK',
      row.refund_flag || false,
      row.request_sent || false,
      row.sent_at || '',
      row.send_result || '',
      'eligible / no exclusion',
      row.last_checked_at || '',
    ]);
  });
  if (rows.length) {
    resultSheet.getRange(2, 1, rows.length, RESULT_HEADERS.length).setValues(rows);
  }

  log_('INFO', 'buildReviewRequestCandidates', '',
    `候補書込: ${rows.length}件 / eligible:${cEligible} / not_eligible除外:${dNotEligible}`
    + ` / request_sent除外:${dSent} / refund除外:${dRefund} / exclude除外:${dExclude}`
    + ` / review_window_error除外:${dReviewWindowError}`, {});
  if (dropEligible.length) {
    log_('INFO', 'buildReviewRequestCandidates', '',
      `eligibleなのに除外(${dropEligible.length}件): ` + dropEligible.slice(0, 80).join(' ; '), {});
  }

  const sendable = getApprovedSendableCandidates_();
  if (sendable.length > 0) {
    sendLineMessageIfConfigured_(`Amazon評価依頼 JP: 送信対象(OK) ${sendable.length}件\n${REVIEW_JP.spreadsheetUrl}`);
    log_('INFO', 'buildReviewRequestCandidates', '', `OK送信対象 ${sendable.length}件 → 通知`, {});
  } else {
    log_('INFO', 'buildReviewRequestCandidates', '', 'OK送信対象0件 → 通知スキップ', {});
  }
}
function sendApprovedReviewRequests() {
  ensureBaseSheets_();
  upsertDefaultSettings_();
  const settings = getSettings_();
  if (String(settings.dry_run).toUpperCase() !== 'FALSE') {
    throw new Error('dry_run is not FALSE. Sending stopped.');
  }
  if (String(settings.send_enabled).toUpperCase() !== 'TRUE') {
    throw new Error('send_enabled is not TRUE. Sending stopped.');
  }
  const resultSheet = getSheet_(REVIEW_JP.sheets.results);
  const resultTable = readTable_(resultSheet);
  const orderSheet = getSheet_(REVIEW_JP.sheets.orders);
  const orderTable = readTable_(orderSheet);
  const orderIndex = new Map();
  const exclusions = getActiveExclusions_();
  orderTable.rows.forEach((row, index) => {
    if (row.amazon_order_id) orderIndex.set(row.amazon_order_id, { row, sheetRow: index + 2 });
  });
  resultTable.rows.forEach((candidate) => {
    if (String(candidate.send_decision || '').trim().toUpperCase() !== 'OK') return;
    const orderRecord = orderIndex.get(candidate.amazon_order_id);
    if (!orderRecord) {
      log_('WARN', 'sendApprovedReviewRequests', candidate.amazon_order_id, 'Order not found in orders sheet', {});
      return;
    }
    const row = orderRecord.row;
    if (String(row.solicitation_status || '').toLowerCase() !== 'eligible') return;
    if (isTrue_(row.request_sent)) return;
    if (isTrue_(row.refund_flag)) return;
    if (isExcluded_(row, exclusions)) return;
    if (hasReviewWindowError_(row)) return; // ★v8 送信直前ガード: 5-30日範囲外エラー注文は送信しない
    try {
      const response = sendProductReviewAndSellerFeedbackSolicitation_(row.amazon_order_id);
      orderSheet.getRange(orderRecord.sheetRow, 12, 1, 3).setValues([[true, new Date(), JSON.stringify(response)]]);
      log_('INFO', 'sendApprovedReviewRequests', row.amazon_order_id, 'Review request sent', response);
      Utilities.sleep(Number(settings.solicitation_delay_ms || 1500));
    } catch (error) {
      orderSheet.getRange(orderRecord.sheetRow, 14).setValue(error.message);
      log_('ERROR', 'sendApprovedReviewRequests', row.amazon_order_id, error.message, {});
    }
  });
}

/* ★2 修正: persistNgDecisionsToExclusions_() は削除しました（自動除外ループの発生源）。
   send_decision=NG は buildReviewRequestCandidates() の引き継ぎで保持され、送信(OKのみ)対象から外れます。
   恒久除外(評価依頼_除外)や注文の exclude_flag には一切書き込みません。 */

/**
 * ★7 復旧用（DRY_RUN付き・一度きり）:
 *   reason "send_decision NG / eligible / no exclusion" の誤除外行を active=FALSE にし、
 *   対応注文の exclude_flag / exclude_reason をクリアする。
 *   検証済み対象: 250-6490546-4785421 / 249-5741099-2453428（計2件）。
 *   ★この修正版を反映した後に、まず DRY_RUN=true で件数確認 → false にして本実行。
 */
function cleanupFalseExclusions_NGdecision() {
  const DRY_RUN = false; // 確認できたら false にして本実行

  const FALSE_REASON = 'send_decision NG / eligible / no exclusion';
  const book = SpreadsheetApp.openByUrl(REVIEW_JP.spreadsheetUrl);
  const idxOf = (headerRow) => {
    const m = {};
    headerRow.forEach((h, i) => {
      const k = String(h == null ? '' : h).trim();
      if (k && !(k in m)) m[k] = i;
    });
    return m;
  };

  const exSheet = book.getSheetByName(REVIEW_JP.sheets.exclusions);
  const exVals = exSheet.getDataRange().getValues();
  const ex = idxOf(exVals[0]);
  const recovered = {};
  const deactivateRows = [];
  for (let r = 1; r < exVals.length; r++) {
    const row = exVals[r];
    if (String(row[ex['reason']]).trim() !== FALSE_REASON) continue;
    if (!isTrue_(row[ex['active']])) continue;
    if (String(row[ex['type']]).trim() !== 'order_id') continue;
    const oid = String(row[ex['value']]).trim();
    if (oid) recovered[oid] = true;
    deactivateRows.push(r + 1);
  }

  const ordSheet = book.getSheetByName(REVIEW_JP.sheets.orders);
  const ordVals = ordSheet.getDataRange().getValues();
  const od = idxOf(ordVals[0]);
  const clearRows = [];
  for (let i = 1; i < ordVals.length; i++) {
    const oid = String(ordVals[i][od['amazon_order_id']]).trim();
    if (recovered[oid] && isTrue_(ordVals[i][od['exclude_flag']])) clearRows.push(i + 1);
  }

  Logger.log('[cleanup JP] DRY_RUN=' + DRY_RUN +
             ' / 除外解除(active->FALSE)=' + deactivateRows.length +
             ' / 注文exclude_flagクリア=' + clearRows.length);
  Logger.log('[cleanup JP] 復旧 order_id: ' + Object.keys(recovered).join(', '));
  if (DRY_RUN) return;

  deactivateRows.forEach((rowNum) => exSheet.getRange(rowNum, ex['active'] + 1).setValue(false));
  clearRows.forEach((rowNum) => {
    ordSheet.getRange(rowNum, od['exclude_flag'] + 1).setValue('');
    if (od['exclude_reason'] != null) ordSheet.getRange(rowNum, od['exclude_reason'] + 1).setValue('');
  });
  Logger.log('[cleanup JP] 本実行 完了');
}

/** ★1 で使用: order_date から現在までの経過日数。パース不能なら null */
function orderAgeDays_(orderDate) {
  if (!orderDate) return null;
  const d = (orderDate instanceof Date) ? orderDate : new Date(String(orderDate));
  if (isNaN(d.getTime())) return null;
  return Math.floor((Date.now() - d.getTime()) / 86400000);
}

function buildSellerOrderUrl_(amazonOrderId) {
  return amazonOrderId ? `https://sellercentral.amazon.co.jp/orders-v3/order/${amazonOrderId}` : '';
}

/* ============================================================================================
 * ▼▼ scheduledXXX ラッパー（★再構成：コアを呼ぶだけ）
 *    実際のラッパーに独自処理がある場合は、この節を既存のものに差し替えてください。
 * ============================================================================================ */
function scheduledFetchRecentOrders()    { fetchRecentFbaShippedOrders(); }
function scheduledFetchOrderItemsBatch() { fetchOrderItemsForReviewOrders(); }
function scheduledImportRefundEmails()   { importRefundEmailsToExclusions(); }
function scheduledRunDryRunBatch()       { runReviewRequestDryRun(); }
function scheduledBuildCandidates()      { buildReviewRequestCandidates(); }

function setupReviewRequestTriggers() {
  deleteReviewRequestTriggers_();
  ScriptApp.newTrigger('scheduledFetchRecentOrders').timeBased().everyDays(1).atHour(0).nearMinute(5).create();
  ScriptApp.newTrigger('scheduledFetchOrderItemsBatch').timeBased().everyDays(1).atHour(1).nearMinute(5).create();
  ScriptApp.newTrigger('scheduledFetchOrderItemsBatch').timeBased().everyDays(1).atHour(2).nearMinute(5).create();
  ScriptApp.newTrigger('scheduledFetchOrderItemsBatch').timeBased().everyDays(1).atHour(3).nearMinute(5).create();
  ScriptApp.newTrigger('scheduledImportRefundEmails').timeBased().everyDays(1).atHour(4).nearMinute(5).create();
  ScriptApp.newTrigger('scheduledRunDryRunBatch').timeBased().everyDays(1).atHour(5).nearMinute(5).create();
  ScriptApp.newTrigger('scheduledRunDryRunBatch').timeBased().everyDays(1).atHour(6).nearMinute(5).create();
  ScriptApp.newTrigger('scheduledBuildCandidates').timeBased().everyDays(1).atHour(6).nearMinute(45).create();
  log_('INFO', 'setupReviewRequestTriggers', '', 'Created automatic triggers (scheduledXXX) excluding send', {});
}
function deleteReviewRequestTriggers_() {
  const managed = new Set([
    'scheduledFetchRecentOrders',
    'scheduledFetchOrderItemsBatch',
    'scheduledImportRefundEmails',
    'scheduledRunDryRunBatch',
    'scheduledBuildCandidates',
  ]);
  ScriptApp.getProjectTriggers().forEach((trigger) => {
    if (managed.has(trigger.getHandlerFunction())) {
      ScriptApp.deleteTrigger(trigger);
    }
  });
}
/* ============================================================================================
 * ▲▲ scheduledXXX ラッパーここまで
 * ============================================================================================ */

function ensureBaseSheets_() {
  const ss = SpreadsheetApp.openByUrl(REVIEW_JP.spreadsheetUrl);
  ensureSheetWithHeaders_(ss, REVIEW_JP.sheets.orders, ORDER_HEADERS);
  ensureSheetWithHeaders_(ss, REVIEW_JP.sheets.exclusions, EXCLUSION_HEADERS);
  ensureSheetWithHeaders_(ss, REVIEW_JP.sheets.results, RESULT_HEADERS);
  ensureSheetWithHeaders_(ss, REVIEW_JP.sheets.logs, LOG_HEADERS);
  ensureSheetWithHeaders_(ss, REVIEW_JP.sheets.settings, ['key', 'value', 'notes', 'updated_at']);
  if (!ss.getSheetByName(REVIEW_JP.sheets.flow)) ss.insertSheet(REVIEW_JP.sheets.flow);
}
function ensureSheetWithHeaders_(ss, name, headers) {
  const sheet = ss.getSheetByName(name) || ss.insertSheet(name);
  const current = sheet.getRange(1, 1, 1, headers.length).getValues()[0];
  const hasHeaders = headers.every((header, index) => current[index] === header);
  if (!hasHeaders) {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.setFrozenRows(1);
  }
  return sheet;
}
function upsertDefaultSettings_() {
  writeJpSettings_();
}
function writeJpSettings_() {
  const sheet = getSheet_(REVIEW_JP.sheets.settings);
  const rows = [['key', 'value', 'notes', 'updated_at']];
  Object.keys(REVIEW_JP.defaults).forEach((key) => {
    rows.push([key, REVIEW_JP.defaults[key], jpSettingNote_(key), new Date()]);
  });
  sheet.getRange(1, 1, sheet.getMaxRows(), 4).clearContent();
  sheet.getRange(1, 1, rows.length, 4).setValues(rows);
  sheet.setFrozenRows(1);
}
function jpSettingNote_(key) {
  const notes = {
    marketplace_id: 'Amazon.co.jp',
    spapi_endpoint: 'Far East endpoint',
    aws_region: 'SP-API FE endpoint signing region',
    order_fetch_days: '注文取得期間（日）',
    dry_run: 'FALSE のとき送信可能',
    send_enabled: 'TRUE のとき送信可能',
  };
  return notes[key] || 'JP版設定';
}
function getSettings_() {
  const sheet = getSheet_(REVIEW_JP.sheets.settings);
  const table = readTable_(sheet);
  const settings = {};
  table.rows.forEach((row) => {
    if (row.key && row.value !== '') settings[row.key] = row.value;
  });
  return Object.assign({}, REVIEW_JP.defaults, settings);
}
function normalizeSettingValue_(value) {
  if (value === true) return 'TRUE';
  if (value === false) return 'FALSE';
  if (value === null || value === undefined) return '';
  return String(value).trim();
}
function displaySettingValue_(value) {
  const normalized = normalizeSettingValue_(value);
  return normalized === '' ? '(blank)' : normalized;
}
function getSheet_(name) {
  const ss = SpreadsheetApp.openByUrl(REVIEW_JP.spreadsheetUrl);
  const sheet = ss.getSheetByName(name);
  if (!sheet) throw new Error(`Missing sheet: ${name}`);
  return sheet;
}
function readTable_(sheet) {
  const values = sheet.getDataRange().getValues();
  if (!values.length) return { headers: [], rows: [] };
  const headers = values[0].map((header) => String(header || '').trim());
  const rows = values.slice(1).filter((row) => row.some((value) => value !== '')).map((row) => {
    const obj = {};
    headers.forEach((header, index) => {
      obj[header] = row[index];
    });
    return obj;
  });
  return { headers, rows };
}
function getHeaderIndexes_(sheet) {
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const indexes = {};
  headers.forEach((header, index) => {
    indexes[String(header || '').trim()] = index + 1;
  });
  return indexes;
}
function setCellByHeader_(sheet, headerIndexes, rowNumber, header, value) {
  const column = headerIndexes[header];
  if (!column) return;
  sheet.getRange(rowNumber, column).setValue(value);
}
function clearResultRows_(sheet) {
  const lastRow = sheet.getLastRow();
  if (lastRow > 1) {
    sheet.getRange(2, 1, lastRow - 1, sheet.getMaxColumns()).clearContent();
  }
}
function getActiveExclusions_() {
  const sheet = getSheet_(REVIEW_JP.sheets.exclusions);
  const table = readTable_(sheet);
  const orderIds = new Set();
  const asins = new Set();
  table.rows.forEach((row) => {
    // ★v8修正: active=TRUE の行だけ有効除外。boolean false / "FALSE" / 空欄 はすべて無効（安全側）。
    //   旧 String(row.active || '') は boolean false を '' に潰し active=FALSE 行を有効化していたため廃止。
    if (!isTrue_(row.active)) return;
    const type = String(row.type || '').toLowerCase().trim();
    const value = String(row.value || '').trim();
    if (!value) return; // value 空の行は無視（空文字混入による誤除外を防止）
    if (['order_id', 'amazon_order_id', 'order', '注文番号'].includes(type)) orderIds.add(value);
    if (type === 'asin') asins.add(value);
  });
  return { orderIds, asins };
}
/**
 * ★v8: 過去に「outside the 5-30 day range」(配達後5-30日の範囲外) エラーになった注文か判定。
 *   send_result に該当文字列を含むかで判定。除外シートには自動登録しない（候補一覧から非表示にするのみ）。
 */
function hasReviewWindowError_(row) {
  const text = String(row.send_result || '').toLowerCase();
  return text.indexOf('outside the 5-30 day range') !== -1;
}

/** ★v8: 実送信対象(OK)の一元判定（通知・プレビューで同一基準）。 */
function getApprovedSendableCandidates_() {
  const resultRows = readTable_(getSheet_(REVIEW_JP.sheets.results)).rows;
  const orderRows = readTable_(getSheet_(REVIEW_JP.sheets.orders)).rows;
  const exclusions = getActiveExclusions_();
  const orderIndex = new Map();
  orderRows.forEach((row) => {
    const id = String(row.amazon_order_id || '').trim();
    if (id) orderIndex.set(id, row);
  });
  const sendable = [];
  resultRows.forEach((c) => {
    if (String(c.send_decision || '').trim().toUpperCase() !== 'OK') return;
    const order = orderIndex.get(String(c.amazon_order_id || '').trim());
    if (!order) return;
    if (String(order.solicitation_status || '').trim().toLowerCase() !== 'eligible') return;
    if (isTrue_(order.request_sent)) return;
    if (isTrue_(order.refund_flag)) return;
    if (isExcluded_(order, exclusions)) return;
    if (hasReviewWindowError_(order)) return;
    sendable.push(c);
  });
  return sendable;
}

/** ★v8: ログ確認用（副作用なし）。実送信前の確認に使う。結果は実行ログへ出力。 */
function previewApprovedSendable() {
  const list = getApprovedSendableCandidates_();
  list.forEach((c) => Logger.log('OK  ' + c.amazon_order_id + '  ' + (c.asin || '') + '  ' + (c.item_name || '')));
  Logger.log('=== 実送信対象(OK): ' + list.length + '件 ===');
}

/** ★v8: 候補抽出トレース（読み取り専用）。判定/集計/除外理由を 評価依頼_ログ とアラートへ出力。 */
function debugBuildCandidatesTrace() {
  const ordersTable = readTable_(getSheet_(REVIEW_JP.sheets.orders));
  const exclusions = getActiveExclusions_();
  let eligible = 0, dNotEligible = 0, dSent = 0, dRefund = 0, dExclude = 0, dWindow = 0, candidate = 0;
  const eligibleDrops = [];
  ordersTable.rows.forEach((row) => {
    const oid = String(row.amazon_order_id || '').trim();
    if (!oid) return;
    const status = String(row.solicitation_status || '').trim().toLowerCase();
    const rs = isTrue_(row.request_sent);
    const rf = isTrue_(row.refund_flag);
    const exc = isExcluded_(row, exclusions);
    if (status === 'eligible') eligible++;
    let reason = '';
    if (status !== 'eligible') { reason = 'not_eligible(' + (status || 'empty') + ')'; dNotEligible++; }
    else if (rs) { reason = 'request_sent=TRUE'; dSent++; }
    else if (rf) { reason = 'refund_flag=TRUE'; dRefund++; }
    else if (exc) {
      const sub = [];
      if (isTrue_(row.exclude_flag)) sub.push('exclude_flag');
      if (exclusions.orderIds.has(oid)) sub.push('除外order_id');
      if (exclusions.asins.has(String(row.asin || '').trim())) sub.push('除外asin');
      reason = 'isExcluded[' + sub.join('/') + ']'; dExclude++;
    }
    else if (hasReviewWindowError_(row)) { reason = 'review_window_error'; dWindow++; }
    else { candidate++; }
    if (reason && status === 'eligible') eligibleDrops.push(oid + '(asin=' + String(row.asin || '').trim() + ')→' + reason);
  });
  const summary = '注文行数:' + ordersTable.rows.length + ' / eligible:' + eligible
    + ' / not_eligible除外:' + dNotEligible + ' / request_sent除外:' + dSent
    + ' / refund除外:' + dRefund + ' / exclude除外:' + dExclude
    + ' / review_window_error除外:' + dWindow + ' / 最終候補:' + candidate;
  log_('INFO', 'debugBuildCandidatesTrace', '', summary, {});
  eligibleDrops.slice(0, 50).forEach((s) => log_('INFO', 'debugBuildCandidatesTrace', '', s, {}));
  SpreadsheetApp.getUi().alert(summary + '\n\n■ eligibleなのに除外(' + eligibleDrops.length + '件):\n'
    + (eligibleDrops.slice(0, 30).join('\n') || '（なし）'));
}
function isExcluded_(row, exclusions) {
  return isTrue_(row.exclude_flag) ||
    exclusions.orderIds.has(String(row.amazon_order_id || '').trim()) ||
    exclusions.asins.has(String(row.asin || '').trim());
}
function isTrue_(value) {
  // ★v8修正: boolean / "TRUE"/"FALSE" / 空欄・null 混在でも正しく判定（前後空白に強い）。
  if (value === true) return true;
  if (value === false || value === null || value === undefined) return false;
  return String(value).trim().toUpperCase() === 'TRUE';
}
function getSolicitationActions_(amazonOrderId) {
  const settings = getSettings_();
  return spApiRequest_(
    'get',
    `/solicitations/v1/orders/${encodeURIComponent(amazonOrderId)}`,
    { marketplaceIds: settings.marketplace_id }
  );
}
function getSolicitationActionNames_(response) {
  const links = response && response._links ? response._links : {};
  const embedded = response && response._embedded ? response._embedded : {};
  const actions = embedded.actions || links.actions || [];
  return actions.map((action) => action.name || action.rel || '').filter(Boolean);
}
function sendProductReviewAndSellerFeedbackSolicitation_(amazonOrderId) {
  const settings = getSettings_();
  return spApiRequest_(
    'post',
    `/solicitations/v1/orders/${encodeURIComponent(amazonOrderId)}/solicitations/productReviewAndSellerFeedback`,
    { marketplaceIds: settings.marketplace_id },
    ''
  );
}
function spApiRequest_(method, path, query, payload) {
  const settings = getSettings_();
  const endpoint = String(settings.spapi_endpoint).replace(/\/$/, '');
  const accessToken = getLwaAccessToken_();
  const queryString = canonicalQueryString_(query || {});
  const url = `${endpoint}${path}${queryString ? `?${queryString}` : ''}`;
  const body = payload === undefined ? '' : payload;
  const headers = {
    Accept: 'application/json',
    'x-amz-access-token': accessToken,
  };
  const response = UrlFetchApp.fetch(url, {
    method,
    muteHttpExceptions: true,
    contentType: 'application/json',
    payload: method.toLowerCase() === 'get' ? undefined : body,
    headers,
  });
  const text = response.getContentText();
  const code = response.getResponseCode();
  if (code < 200 || code >= 300) {
    throw new Error(`SP-API HTTP ${code}: ${text}`);
  }
  const json = text ? JSON.parse(text) : {};
  return json && json.payload ? json.payload : json;
}
function getLwaAccessToken_() {
  const props = PropertiesService.getScriptProperties();
  const response = UrlFetchApp.fetch('https://api.amazon.com/auth/o2/token', {
    method: 'post',
    contentType: 'application/x-www-form-urlencoded',
    payload: {
      grant_type: 'refresh_token',
      refresh_token: props.getProperty('SPAPI_REFRESH_TOKEN'),
      client_id: props.getProperty('SPAPI_CLIENT_ID'),
      client_secret: props.getProperty('SPAPI_CLIENT_SECRET'),
    },
    muteHttpExceptions: true,
  });
  const text = response.getContentText();
  if (response.getResponseCode() !== 200) {
    throw new Error(`LWA token error: ${text}`);
  }
  return JSON.parse(text).access_token;
}
function signAwsRequest_(method, endpoint, path, query, payload, accessToken) {
  const settings = getSettings_();
  const props = PropertiesService.getScriptProperties();
  const host = endpoint.replace(/^https?:\/\//, '');
  const now = new Date();
  const amzDate = Utilities.formatDate(now, 'GMT', "yyyyMMdd'T'HHmmss'Z'");
  const dateStamp = Utilities.formatDate(now, 'GMT', 'yyyyMMdd');
  const service = 'execute-api';
  const region = settings.aws_region;
  const sessionToken = props.getProperty('AWS_SESSION_TOKEN');
  const payloadHash = sha256Hex_(payload || '');
  const canonicalHeaderMap = {
    host,
    'x-amz-access-token': accessToken,
    'x-amz-date': amzDate,
  };
  if (sessionToken) canonicalHeaderMap['x-amz-security-token'] = sessionToken;
  const signedHeaders = Object.keys(canonicalHeaderMap).sort().join(';');
  const canonicalHeaders = Object.keys(canonicalHeaderMap)
    .sort()
    .map((key) => `${key}:${canonicalHeaderMap[key]}\n`)
    .join('');
  const canonicalRequest = [
    method.toUpperCase(),
    encodePath_(path),
    canonicalQueryString_(query || {}),
    canonicalHeaders,
    signedHeaders,
    payloadHash,
  ].join('\n');
  const credentialScope = `${dateStamp}/${region}/${service}/aws4_request`;
  const stringToSign = [
    'AWS4-HMAC-SHA256',
    amzDate,
    credentialScope,
    sha256Hex_(canonicalRequest),
  ].join('\n');
  const signingKey = getSignatureKey_(props.getProperty('AWS_SECRET_ACCESS_KEY'), dateStamp, region, service);
  const signature = hmacHex_(signingKey, stringToSign);
  const authorization =
    `AWS4-HMAC-SHA256 Credential=${props.getProperty('AWS_ACCESS_KEY_ID')}/${credentialScope}, ` +
    `SignedHeaders=${signedHeaders}, Signature=${signature}`;
  const headers = {
    Authorization: authorization,
    Accept: 'application/json',
    'x-amz-access-token': accessToken,
    'x-amz-date': amzDate,
  };
  if (sessionToken) headers['x-amz-security-token'] = sessionToken;
  return headers;
}
function canonicalQueryString_(query) {
  return Object.keys(query)
    .filter((key) => query[key] !== undefined && query[key] !== '')
    .sort()
    .map((key) => `${encodeRfc3986_(key)}=${encodeRfc3986_(String(query[key]))}`)
    .join('&');
}
function encodePath_(path) {
  return path.split('/').map((part) => encodeRfc3986_(decodeURIComponent(part))).join('/');
}
function encodeRfc3986_(value) {
  return encodeURIComponent(value).replace(/[!'()*]/g, (char) => `%${char.charCodeAt(0).toString(16).toUpperCase()}`);
}
function sha256Hex_(value) {
  return bytesToHex_(Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, value, Utilities.Charset.UTF_8));
}
function hmacBytes_(key, value) {
  return Utilities.computeHmacSha256Signature(toBytes_(value), toBytes_(key));
}
function hmacHex_(key, value) {
  return bytesToHex_(hmacBytes_(key, value));
}
function getSignatureKey_(secretKey, dateStamp, regionName, serviceName) {
  const kDate = hmacBytes_(`AWS4${secretKey}`, dateStamp);
  const kRegion = hmacBytes_(kDate, regionName);
  const kService = hmacBytes_(kRegion, serviceName);
  return hmacBytes_(kService, 'aws4_request');
}
function toBytes_(value) {
  if (Array.isArray(value)) return value;
  return Utilities.newBlob(String(value)).getBytes();
}
function bytesToHex_(bytes) {
  return bytes.map((byte) => {
    const value = byte < 0 ? byte + 256 : byte;
    return (`0${value.toString(16)}`).slice(-2);
  }).join('');
}
function sendEmailNotification_(message) {
  const props = PropertiesService.getScriptProperties();
  const token = props.getProperty('LINE_CHANNEL_ACCESS_TOKEN');
  const userId = props.getProperty('LINE_USER_ID');
  if (!token || !userId) {
    log_('WARN', 'sendLineMessageIfConfigured_', '', 'LINE properties are not configured', {});
    return;
  }
  sendLineMessage_(message, token, userId);
}
function sendLineMessage_(message, token, userId) {
  const response = UrlFetchApp.fetch('https://api.line.me/v2/bot/message/push', {
    method: 'post',
    contentType: 'application/json',
    headers: { Authorization: `Bearer ${token}` },
    payload: JSON.stringify({
      to: userId,
      messages: [{ type: 'text', text: message }],
    }),
    muteHttpExceptions: true,
  });
  if (response.getResponseCode() < 200 || response.getResponseCode() >= 300) {
    throw new Error(`LINE push error: ${response.getResponseCode()} ${response.getContentText()}`);
  }
}
function log_(level, functionName, amazonOrderId, message, details) {
  const sheet = getSheet_(REVIEW_JP.sheets.logs);
  sheet.appendRow([
    new Date(),
    level,
    functionName,
    amazonOrderId || '',
    message || '',
    details ? JSON.stringify(details) : '',
  ]);
}

/* ★v8相当: onOpen は冒頭の単一定義に統合済み（旧FDA用onOpenはここにあったが統合のため削除）。 */

function fetchEmailData() {
  var ui = SpreadsheetApp.getUi();
  var sheet = SpreadsheetApp.getActiveSheet();

  var stateMap = {
    "alabama": "AL", "alaska": "AK", "arizona": "AZ", "arkansas": "AR", "california": "CA",
    "colorado": "CO", "connecticut": "CT", "delaware": "DE", "florida": "FL", "georgia": "GA",
    "hawaii": "HI", "idaho": "ID", "illinois": "IL", "indiana": "IN", "iowa": "IA",
    "kansas": "KS", "kentucky": "KY", "louisiana": "LA", "maine": "ME", "maryland": "MD",
    "massachusetts": "MA", "michigan": "MI", "minnesota": "MN", "mississippi": "MS", "missouri": "MO",
    "montana": "MT", "nebraska": "NE", "nevada": "NV", "new hampshire": "NH", "new jersey": "NJ",
    "new mexico": "NM", "new york": "NY", "north carolina": "NC", "north dakota": "ND", "ohio": "OH",
    "oklahoma": "OK", "oregon": "OR", "pennsylvania": "PA", "rhode island": "RI", "south carolina": "SC",
    "south dakota": "SD", "tennessee": "TN", "texas": "TX", "utah": "UT", "vermont": "VT",
    "virginia": "VA", "washington": "WA", "west virginia": "WV", "wisconsin": "WI", "wyoming": "WY"
  };

  var dateResponse = ui.prompt('発送日確認', '発送日はいつですか？\n(例: 5月26日)', ui.ButtonSet.OK_CANCEL);
  if (dateResponse.getSelectedButton() !== ui.Button.OK) {
    return;
  }

  var shippingDate = dateResponse.getResponseText().trim();

  var defaultMgmtNum = sheet.getRange("Q1").getValue();
  var sameResponse = ui.alert(
    '管理番号確認',
    '件名の管理番号は同じですか？\n(現在の対象: ' + defaultMgmtNum + ')',
    ui.ButtonSet.YES_NO
  );

  var managementNum = "";

  if (sameResponse === ui.Button.YES) {
    managementNum = String(defaultMgmtNum || "").trim();
  } else if (sameResponse === ui.Button.NO) {
    var numResponse = ui.prompt('管理番号入力', '対象管理番号は？', ui.ButtonSet.OK_CANCEL);
    if (numResponse.getSelectedButton() !== ui.Button.OK) {
      return;
    }
    managementNum = numResponse.getResponseText().trim();
  } else {
    return;
  }

  var subjectToSearch = "RE: [EXTERNAL] Arrival Date/Time_Port of Arrival【" + managementNum + "】";
  var query = 'subject:"' + subjectToSearch + '"';

  try {
    var threads = GmailApp.search(query, 0, 1);

    if (threads.length === 0) {
      ui.alert("対象のメールが見つかりませんでした。\n検索条件: " + subjectToSearch);
      return;
    }

    var messages = threads[0].getMessages();
    var latestMessage = messages[messages.length - 1];
    var body = latestMessage.getPlainBody();

    if (body.indexOf("貨物" + managementNum) === -1) {
      ui.alert("注意: メール本文に「貨物" + managementNum + "」の記載が見つかりませんでした。処理を続行します。");
    }

    var targetBlock = "";
    var blockRegex = new RegExp(
      escapeRegExp_(shippingDate) + ".*ご集荷の場合：([\\s\\S]*?)(?=\\d+月\\d+日.*ご集荷の場合：|$)",
      "i"
    );

    var blockMatch = body.match(blockRegex);

    if (blockMatch) {
      targetBlock = blockMatch[1];
    } else {
      targetBlock = body;
      ui.alert("指定日のブロックが見つからなかったため、メール本文全体から取得を試みます。");
    }

    var city = "取得失敗";
    var stateAbbr = "取得失敗";
    var arrivalDate = "取得失敗";
    var arrivalTime = "取得失敗";

    // 新形式: Port of Arrival | Indianapolis , IN
    var portRegex = /Port of Arrival\s*\|\s*([A-Za-z\s.'-]+?)\s*,\s*([A-Z]{2})/i;
    var portMatch = targetBlock.match(portRegex) || body.match(portRegex);

    if (portMatch) {
      city = portMatch[1].trim();
      stateAbbr = portMatch[2].trim().toUpperCase();
    } else {
      // 旧形式: 到着空港：... at Indianapolis (State: Indiana)
      var airportRegex = /到着空港：.*?at\s+([A-Za-z\s.'-]+?)\s*\(/i;
      var airportMatch = targetBlock.match(airportRegex) || body.match(airportRegex);

      if (airportMatch) {
        city = airportMatch[1].trim();
      }

      var stateRegex = /\(State:\s*([A-Za-z\s]+)\)/i;
      var stateMatch = targetBlock.match(stateRegex) || body.match(stateRegex);

      if (stateMatch) {
        var stateFullName = stateMatch[1].trim().toLowerCase();
        stateAbbr = stateMap[stateFullName] || "変換失敗(" + stateMatch[1].trim() + ")";
      }
    }

    var arrivalDateRegex = /フライト到着日\(現地時間\)：\s*(.*)/;
    var arrivalDateMatch = targetBlock.match(arrivalDateRegex) || body.match(arrivalDateRegex);

    if (arrivalDateMatch) {
      arrivalDate = arrivalDateMatch[1].trim();
    }

    var arrivalTimeRegex = /フライト到着時間\(現地時間\/24時間表示\)：\s*(.*)/;
    var arrivalTimeMatch = targetBlock.match(arrivalTimeRegex) || body.match(arrivalTimeRegex);

    if (arrivalTimeMatch) {
      arrivalTime = arrivalTimeMatch[1].trim();
    }

    sheet.getRange("Q3").setValue(city);
    sheet.getRange("Q4").setValue(stateAbbr);
    sheet.getRange("Q5").setValue(arrivalDate);
    sheet.getRange("Q6").setValue(arrivalTime);

    ui.alert(
      "メール情報の取得が完了しました。\n\n" +
      "到着空港 都市名 (Q3): " + city + "\n" +
      "州略称 (Q4): " + stateAbbr + "\n" +
      "到着日 (Q5): " + arrivalDate + "\n" +
      "到着時間 (Q6): " + arrivalTime
    );

  } catch (e) {
    ui.alert("エラーが発生しました: " + e.message);
  }
}

function escapeRegExp_(text) {
  return String(text || "").replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
}

function updateFbaShippingList() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('FBA送料一覧');
  const externalId = '1_PqYPsK8GxSuFJ_Fq_FheExl8-xjHoq2JnpKIZ9zTus';
  const externalSheet = SpreadsheetApp.openById(externalId).getSheetByName('進捗管理');

  const data = sheet.getDataRange().getValues();
  const extData = externalSheet.getDataRange().getValues();
  const extMap = new Map();

  // 外部シート「進捗管理」のデータをマップ化 (A列:管理番号, N列:FBA納品番号, AQ列:追跡番号, AS列:FDA)
  for (let i = 1; i < extData.length; i++) {
    const mgmtNum = String(extData[i][0]).trim(); // A列
    if (mgmtNum) {
      extMap.set(mgmtNum, {
        n: extData[i][13],  // N列 (FBA納品番号)
        aq: extData[i][42], // AQ列 (追跡番号)
        as: extData[i][44]  // AS列 (エントリー番号)
      });
    }
  }

  let updated = false;
  // FBA送料一覧の各行をループ (B列:管理番号, M列:追跡番号, N列:納品プラン番号, O列:FDA)
  for (let i = 1; i < data.length; i++) {
    const mgmtNum = String(data[i][1]).trim(); // B列
    if (extMap.has(mgmtNum)) {
      const info = extMap.get(mgmtNum);
      
      // M列(追跡番号)が空なら外部のAQ列を転記
      if (!data[i][12] && info.aq) {
        data[i][12] = info.aq;
        updated = true;
      }
      // N列(納品プラン番号)が空なら外部のN列を転記
      if (!data[i][13] && info.n) {
        data[i][13] = info.n;
        updated = true;
      }
      // O列(FDA)が空なら外部のAS列を転記
      if (!data[i][14] && info.as) {
        data[i][14] = info.as;
        updated = true;
      }
    }
  }

  if (updated) {
    sheet.getRange(1, 1, data.length, data[0].length).setValues(data);
    SpreadsheetApp.getUi().alert('更新が完了しました。');
  } else {
    SpreadsheetApp.getUi().alert('更新が必要な空欄、または一致する管理番号はありませんでした。');
  }
}

/**
 * 日本Amazon評価依頼システム：トリガー修復用
 *
 * 目的：
 * - 存在しない scheduledXXX 関数を補完する
 * - 壊れている既存トリガーを削除する
 * - 正しい時間主導トリガーを再作成する
 *
 * 注意：
 * - sendApprovedReviewRequests は自動トリガー化しません。
 * - 本送信は手動確認後に行う前提です。
 */

// ============================================================
// 時間主導トリガー用ラッパー関数
// ============================================================

function scheduledFetchRecentOrders() {
  return fetchRecentFbaShippedOrders();
}

function scheduledFetchOrderItemsBatch() {
  return fetchOrderItemsForReviewRequests();
}

function scheduledImportRefundEmails() {
  return importRefundEmailsToExclusions();
}

function scheduledRunDryRunBatch() {
  return runReviewRequestDryRun();
}

function scheduledBuildCandidates() {
  return buildReviewRequestCandidates();
}

// ============================================================
// トリガー再作成
// ============================================================

function resetReviewRequestJpTriggers() {
  const targetFunctions = [
    'scheduledFetchRecentOrders',
    'scheduledFetchOrderItemsBatch',
    'scheduledImportRefundEmails',
    'scheduledRunDryRunBatch',
    'scheduledBuildCandidates',
  ];

  // 既存の対象トリガーを削除
  const triggers = ScriptApp.getProjectTriggers();

  triggers.forEach(trigger => {
    const handler = trigger.getHandlerFunction();
    if (targetFunctions.includes(handler)) {
      ScriptApp.deleteTrigger(trigger);
      Logger.log('[delete trigger] ' + handler);
    }
  });

  // 新規トリガー作成
  createDailyTrigger_('scheduledFetchRecentOrders', 1, 5);
  createDailyTrigger_('scheduledFetchOrderItemsBatch', 2, 5);
  createDailyTrigger_('scheduledImportRefundEmails', 3, 5);
  createDailyTrigger_('scheduledRunDryRunBatch', 5, 5);
  createDailyTrigger_('scheduledBuildCandidates', 6, 5);

  Logger.log('日本Amazon評価依頼トリガーを再作成しました。');
  Logger.log('01:05 scheduledFetchRecentOrders');
  Logger.log('02:05 scheduledFetchOrderItemsBatch');
  Logger.log('03:05 scheduledImportRefundEmails');
  Logger.log('05:05 scheduledRunDryRunBatch');
  Logger.log('06:05 scheduledBuildCandidates');
}

function createDailyTrigger_(functionName, hour, minute) {
  ScriptApp.newTrigger(functionName)
    .timeBased()
    .everyDays(1)
    .atHour(hour)
    .nearMinute(minute)
    .create();

  Logger.log('[create trigger] ' + functionName + ' / ' + hour + ':' + String(minute).padStart(2, '0'));
}

// ============================================================
// トリガー確認用
// ============================================================

function listReviewRequestJpTriggers() {
  const triggers = ScriptApp.getProjectTriggers();

  triggers.forEach(trigger => {
    Logger.log(
      'handler=' + trigger.getHandlerFunction() +
      ' / source=' + trigger.getEventType()
    );
  });
}

/**
 * 日本Amazon評価依頼システム：トリガー修復用
 *
 * 目的：
 * - トリガー画面に登録されている scheduledXXX 関数を用意する
 * - 壊れている既存トリガーを削除する
 * - 正しい時間主導トリガーを再作成する
 *
 * 注意：
 * - sendApprovedReviewRequests は自動トリガー化しません。
 * - 評価依頼の本送信は、候補確認後に手動実行してください。
 */

// ============================================================
// 時間主導トリガー用ラッパー関数
// ============================================================

function scheduledFetchRecentOrders() {
  return fetchRecentFbaShippedOrders();
}

function scheduledFetchOrderItemsBatch() {
  return fetchOrderItemsForReviewOrders();
}

function scheduledImportRefundEmails() {
  return importRefundEmailsToExclusions();
}

function scheduledRunDryRunBatch() {
  return runReviewRequestDryRun();
}

function scheduledBuildCandidates() {
  return buildReviewRequestCandidates();
}

// ============================================================
// トリガー再作成
// ============================================================

function resetReviewRequestJpTriggers() {
  const targetFunctions = [
    'scheduledFetchRecentOrders',
    'scheduledFetchOrderItemsBatch',
    'scheduledImportRefundEmails',
    'scheduledRunDryRunBatch',
    'scheduledBuildCandidates',
  ];

  const triggers = ScriptApp.getProjectTriggers();

  // 既存の対象トリガーを削除
  triggers.forEach(trigger => {
    const handler = trigger.getHandlerFunction();

    if (targetFunctions.includes(handler)) {
      ScriptApp.deleteTrigger(trigger);
      Logger.log('[delete trigger] ' + handler);
    }
  });

  // 新規トリガー作成
  createDailyTrigger_('scheduledFetchRecentOrders', 1, 5);
  createDailyTrigger_('scheduledFetchOrderItemsBatch', 2, 5);
  createDailyTrigger_('scheduledImportRefundEmails', 3, 5);
  createDailyTrigger_('scheduledRunDryRunBatch', 5, 5);
  createDailyTrigger_('scheduledBuildCandidates', 6, 5);

  Logger.log('日本Amazon評価依頼トリガーを再作成しました。');
  Logger.log('01:05 scheduledFetchRecentOrders');
  Logger.log('02:05 scheduledFetchOrderItemsBatch');
  Logger.log('03:05 scheduledImportRefundEmails');
  Logger.log('05:05 scheduledRunDryRunBatch');
  Logger.log('06:05 scheduledBuildCandidates');
}

function createDailyTrigger_(functionName, hour, minute) {
  ScriptApp.newTrigger(functionName)
    .timeBased()
    .everyDays(1)
    .atHour(hour)
    .nearMinute(minute)
    .create();

  Logger.log(
    '[create trigger] ' +
      functionName +
      ' / ' +
      hour +
      ':' +
      String(minute).padStart(2, '0')
  );
}

// ============================================================
// トリガー確認用
// ============================================================

function listReviewRequestJpTriggers() {
  const triggers = ScriptApp.getProjectTriggers();

  if (!triggers.length) {
    Logger.log('トリガーはありません。');
    return;
  }

  triggers.forEach(trigger => {
    Logger.log(
      'handler=' +
        trigger.getHandlerFunction() +
        ' / eventType=' +
        trigger.getEventType()
    );
  });
}