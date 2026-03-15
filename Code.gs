/**
 * 営業メール自動生成スクリプト
 * Google Apps Script (GAS) + OpenAI API (gpt-4o)
 */

function generateSalesEmails() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getActiveSheet();
  
  // スクリプトプロパティからAPIキー取得
  const apiKey = PropertiesService.getScriptProperties().getProperty('OPENAI_API_KEY');
  if (!apiKey) {
    Logger.log('エラー: スクリプトプロパティに OPENAI_API_KEY が設定されていません');
    return;
  }

  // データ範囲を取得（2行目から最終行まで）
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    Logger.log('データがありません（2行目以降にデータを入力してください）');
    return;
  }

  Logger.log(`処理開始: ${lastRow - 1} 件のデータを処理します`);

  for (let row = 2; row <= lastRow; row++) {
    const companyName = sheet.getRange(row, 1).getValue(); // A列: 会社名
    const url         = sheet.getRange(row, 2).getValue(); // B列: URL
    // const contactName = sheet.getRange(row, 3).getValue(); // C列: 担当者名（将来拡張用）

    // 会社名・URLが空の行はスキップ
    if (!companyName || !url) {
      const skipMsg = `行 ${row}: 会社名またはURLが空のためスキップ`;
      Logger.log(skipMsg);
      sheet.getRange(row, 5).setValue(skipMsg); // E列にログ
      continue;
    }

    try {
      Logger.log(`行 ${row}: ${companyName} の処理を開始`);

      const emailText = callOpenAI(apiKey, companyName, url);

      // D列にメール文を書き込み
      sheet.getRange(row, 4).setValue(emailText);

      // E列に成功ログを記録
      const successMsg = `✅ ${new Date().toLocaleString('ja-JP')} 生成成功`;
      sheet.getRange(row, 5).setValue(successMsg);

      Logger.log(`行 ${row}: 生成成功`);

    } catch (error) {
      // エラーが出ても止まらず次の行へ
      const errorMsg = `❌ ${new Date().toLocaleString('ja-JP')} エラー: ${error.message}`;
      sheet.getRange(row, 5).setValue(errorMsg);
      Logger.log(`行 ${row} でエラー発生: ${error.message}`);
    }

    // 1行処理ごとに1秒待機（API制限対策）
    Utilities.sleep(1000);
  }

  Logger.log('全行の処理が完了しました');
  SpreadsheetApp.getUi().alert('完了', `${lastRow - 1} 件の処理が完了しました。\nD列・E列を確認してください。`, SpreadsheetApp.getUi().ButtonSet.OK);
}


/**
 * OpenAI API を呼び出して営業メール文を生成する
 * @param {string} apiKey   - OpenAI APIキー
 * @param {string} company  - 会社名
 * @param {string} url      - 企業URL
 * @returns {string}        - 生成されたメール文
 */
function callOpenAI(apiKey, company, url) {
  const endpoint = 'https://api.openai.com/v1/chat/completions';

  const payload = {
    model: 'gpt-4o',
    messages: [
      {
        role: 'system',
        content: 'あなたはプロの営業担当者です。簡潔で効果的な営業メールを書きます。'
      },
      {
        role: 'user',
        content: `${company}（${url}）向けの営業メールを200字で書いてください`
      }
    ],
    max_tokens: 500,
    temperature: 0.7
  };

  const options = {
    method: 'post',
    contentType: 'application/json',
    headers: {
      'Authorization': `Bearer ${apiKey}`
    },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true // HTTPエラーでも例外を出さずレスポンスを返す
  };

  const response = UrlFetchApp.fetch(endpoint, options);
  const responseCode = response.getResponseCode();
  const responseText = response.getContentText();

  if (responseCode !== 200) {
    throw new Error(`OpenAI API エラー (HTTP ${responseCode}): ${responseText}`);
  }

  const json = JSON.parse(responseText);

  if (json.error) {
    throw new Error(`OpenAI API エラー: ${json.error.message}`);
  }

  return json.choices[0].message.content.trim();
}


/**
 * スプレッドシートにカスタムメニューを追加する
 */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('🤖 営業メール生成')
    .addItem('▶ 実行する', 'generateSalesEmails')
    .addItem('🗑 D・E列をクリア', 'clearOutputColumns')
    .addToUi();
}


/**
 * D列（メール文）とE列（ログ）をクリアする
 */
function clearOutputColumns() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const lastRow = sheet.getLastRow();
  if (lastRow >= 2) {
    sheet.getRange(2, 4, lastRow - 1, 2).clearContent();
    SpreadsheetApp.getUi().alert('D列・E列をクリアしました');
  }
}
