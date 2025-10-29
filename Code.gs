/**
 * 法律クイズキャンペーン - Google Apps Script
 * 弁護士事務所向けクイズキャンペーンサイト
 */

// 設定定数
const SPREADSHEET_ID = '1jiQqV8ZtS-VskievvoV9mjVuem-ntfAGtSUQuYUjkds';
const SHEET_NAME = 'Entries';
const QUIZ_VERSION = '2025-10-21-v1';

/**
 * ウェブアプリのメイン処理
 */
function doGet() {
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/**
 * HTMLファイルのインクルード用
 */
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

/**
 * 応募フォーム送信処理
 * @param {Object} formData - フォームデータ
 * @return {Object} レスポンス
 */
function submitEntry(formData) {
  try {
    // 入力値検証
    if (!formData.name || !formData.email || !formData.consent) {
      return {
        success: false,
        message: '必須項目を入力してください。'
      };
    }

    // メールアドレス形式チェック
    const emailRegex = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
    if (!emailRegex.test(formData.email)) {
      return {
        success: false,
        message: '正しいメールアドレスを入力してください。'
      };
    }

    // 同一メールアドレスの重複チェック（1日1回制限）
    if (checkDuplicateEntry(formData.email)) {
      return {
        success: false,
        message: '本日は既に応募済みです。1日1回まで応募できます。'
      };
    }

    // スプレッドシートにデータ保存
    const result = saveToSpreadsheet(formData);
    
    if (result.success) {
      return {
        success: true,
        message: '応募が完了しました。抽選結果は月末に発表いたします。'
      };
    } else {
      return {
        success: false,
        message: '応募の処理中にエラーが発生しました。しばらくしてから再度お試しください。'
      };
    }

  } catch (error) {
    console.error('応募処理エラー:', error);
    return {
      success: false,
      message: 'システムエラーが発生しました。管理者にお問い合わせください。'
    };
  }
}

/**
 * 同一メールアドレスの重複チェック
 * @param {string} email - メールアドレス
 * @return {boolean} 重複ありの場合true
 */
function checkDuplicateEntry(email) {
  try {
    const spreadsheet = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = spreadsheet.getSheetByName(SHEET_NAME);
    
    if (!sheet) {
      return false; // シートが存在しない場合は重複なし
    }

    const data = sheet.getDataRange().getValues();
    const today = new Date();
    const todayStr = Utilities.formatDate(today, 'JST', 'yyyy-MM-dd');

    // ヘッダー行をスキップしてチェック
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const emailCol = 3; // email列（0ベース）
      const timestampCol = 0; // timestamp列
      
      if (row[emailCol] === email && row[timestampCol]) {
        const entryDate = new Date(row[timestampCol]);
        const entryDateStr = Utilities.formatDate(entryDate, 'JST', 'yyyy-MM-dd');
        
        if (entryDateStr === todayStr) {
          return true; // 同日の応募あり
        }
      }
    }

    return false;
  } catch (error) {
    console.error('重複チェックエラー:', error);
    return false; // エラーの場合は重複なしとして処理
  }
}

/**
 * スプレッドシートにデータ保存
 * @param {Object} formData - フォームデータ
 * @return {Object} 保存結果
 */
function saveToSpreadsheet(formData) {
  try {
    const spreadsheet = SpreadsheetApp.openById(SPREADSHEET_ID);
    let sheet = spreadsheet.getSheetByName(SHEET_NAME);
    
    // シートが存在しない場合は作成
    if (!sheet) {
      sheet = spreadsheet.insertSheet(SHEET_NAME);
      
      // ヘッダー行を追加
      const headers = [
        'timestamp',
        'quiz_version', 
        'name',
        'email',
        'company',
        'ip',
        'user_agent',
        'consent_pp',
        'consent_terms',
        'answer1',
        'answer2',
        'answer3',
        'answer4'
      ];
      sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    }

    // 現在時刻を取得
    const now = new Date();
    
    // データ行を追加
    const newRow = [
      now,                                    // timestamp
      QUIZ_VERSION,                          // quiz_version
      formData.name,                         // name
      formData.email,                        // email
      formData.company || '',                // company
      formData.ip || '',                     // ip
      formData.userAgent || '',              // user_agent
      formData.consent ? 'Yes' : 'No',       // consent_pp
      formData.consent ? 'Yes' : 'No',       // consent_terms
      formData.answer1 === true ? '○' : formData.answer1 === false ? '×' : '',  // answer1
      formData.answer2 === true ? '○' : formData.answer2 === false ? '×' : '',  // answer2
      formData.answer3 === true ? '○' : formData.answer3 === false ? '×' : '',  // answer3
      formData.answer4 === true ? '○' : formData.answer4 === false ? '×' : ''   // answer4
    ];

    sheet.appendRow(newRow);
    
    return { success: true };
  } catch (error) {
    console.error('スプレッドシート保存エラー:', error);
    return { success: false, error: error.toString() };
  }
}

/**
 * 月末抽選機能（2名を抽選）
 * 手動実行またはトリガー設定で実行
 */
function pickWinnersThisMonth() {
  try {
    const spreadsheet = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = spreadsheet.getSheetByName(SHEET_NAME);
    
    if (!sheet) {
      console.log('シートが見つかりません');
      return;
    }

    const data = sheet.getDataRange().getValues();
    if (data.length <= 1) {
      console.log('応募データがありません');
      return;
    }

    // 今月の応募者を取得
    const now = new Date();
    const currentMonth = now.getMonth();
    const currentYear = now.getFullYear();
    
    const currentMonthEntries = [];
    
    // ヘッダー行をスキップして処理
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const timestamp = new Date(row[0]);
      
      if (timestamp.getMonth() === currentMonth && timestamp.getFullYear() === currentYear) {
        currentMonthEntries.push({
          rowIndex: i + 1, // 実際の行番号（1ベース）
          email: row[3],   // email列
          name: row[2]     // name列
        });
      }
    }

    if (currentMonthEntries.length === 0) {
      console.log('今月の応募者がいません');
      return;
    }

    // 重複するメールアドレスを除去（1人1回のみ抽選対象）
    const uniqueEntries = [];
    const seenEmails = new Set();
    
    for (const entry of currentMonthEntries) {
      if (!seenEmails.has(entry.email)) {
        uniqueEntries.push(entry);
        seenEmails.add(entry.email);
      }
    }

    if (uniqueEntries.length < 2) {
      console.log('抽選対象者が2名未満です（対象者数: ' + uniqueEntries.length + '名）');
      return;
    }

    // 2名をランダム抽選
    const winners = [];
    const shuffled = uniqueEntries.sort(() => Math.random() - 0.5);
    
    for (let i = 0; i < Math.min(2, shuffled.length); i++) {
      winners.push(shuffled[i]);
    }

    // 抽選結果をログ出力
    console.log('=== 抽選結果 ===');
    console.log('抽選対象者数: ' + uniqueEntries.length + '名');
    console.log('当選者数: ' + winners.length + '名');
    
    winners.forEach((winner, index) => {
      console.log(`${index + 1}等: ${winner.name} (${winner.email})`);
    });

    // 抽選結果をスプレッドシートに記録（別シートに保存）
    let resultSheet = spreadsheet.getSheetByName('抽選結果');
    if (!resultSheet) {
      resultSheet = spreadsheet.insertSheet('抽選結果');
      resultSheet.getRange(1, 1, 1, 4).setValues([['抽選日', '当選者名', 'メールアドレス', '備考']]);
    }

    const resultRows = [];
    const resultDate = Utilities.formatDate(now, 'JST', 'yyyy-MM-dd HH:mm:ss');
    
    winners.forEach((winner, index) => {
      resultRows.push([resultDate, winner.name, winner.email, `${index + 1}等`]);
    });

    resultSheet.getRange(resultSheet.getLastRow() + 1, 1, resultRows.length, 4).setValues(resultRows);

    return {
      success: true,
      totalEntries: uniqueEntries.length,
      winners: winners.map(w => ({ name: w.name, email: w.email }))
    };

  } catch (error) {
    console.error('抽選処理エラー:', error);
    return {
      success: false,
      error: error.toString()
    };
  }
}

/**
 * テスト用：応募データの確認
 */
function getEntryStats() {
  try {
    const spreadsheet = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = spreadsheet.getSheetByName(SHEET_NAME);
    
    if (!sheet) {
      return { total: 0, today: 0 };
    }

    const data = sheet.getDataRange().getValues();
    const total = data.length - 1; // ヘッダー行を除く
    
    const today = new Date();
    const todayStr = Utilities.formatDate(today, 'JST', 'yyyy-MM-dd');
    let todayCount = 0;
    
    for (let i = 1; i < data.length; i++) {
      const timestamp = new Date(data[i][0]);
      const dateStr = Utilities.formatDate(timestamp, 'JST', 'yyyy-MM-dd');
      if (dateStr === todayStr) {
        todayCount++;
      }
    }

    return {
      total: total,
      today: todayCount
    };
  } catch (error) {
    console.error('統計取得エラー:', error);
    return { total: 0, today: 0 };
  }
}
