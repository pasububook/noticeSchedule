/**
 * 指定したGoogle Spreadsheetの「提出物」シートから、指定期間内のタスクを取得します。
 *
 * @param {Date} startDate - 検索範囲の開始日。
 * @param {Date} endDate - 検索範囲の終了日。
 * @param {string} spreadsheetId - 対象のGoogle SpreadsheetのID。
 * @return {Array<[Date, string]>} 該当する日付とタスクの二次元配列。例: [[new Date("2025/06/15"), "タスクA"], [new Date("2025/06/15"), "タスクB"]]
 */
function getTasksByDateRange(startDate, endDate, spreadsheetId) {
  // 引数の妥当性をチェック
  if (!(startDate instanceof Date) || !(endDate instanceof Date) || typeof spreadsheetId !== 'string' || spreadsheetId.trim() === '') {
    throw new Error('引数が正しくありません。startDateとendDateはDateオブジェクト、spreadsheetIdは有効な文字列である必要があります。');
  }

  try {
    const spreadsheet = SpreadsheetApp.openById(spreadsheetId);
    const sheet = spreadsheet.getSheetByName('提出物');

    if (!sheet) {
      // シートがない場合はエラーがわかるようにメッセージを投げる
      throw new Error("'提出物' という名前のシートが見つかりません。");
    }

    const dataRange = sheet.getDataRange();
    const values = dataRange.getValues();
    const results = []; // 戻り値となる二次元配列

    // 日付の比較のために、時刻情報をリセット
    const start = new Date(startDate.getFullYear(), startDate.getMonth(), startDate.getDate());
    const end = new Date(endDate.getFullYear(), endDate.getMonth(), endDate.getDate());

    // 1行目はヘッダーとしてスキップし、2行目からループ
    for (let i = 1; i < values.length; i++) {
      const row = values[i];
      const dateValue = row[0]; // A列の日付 (Dateオブジェクト)
      const taskValue = row[2]; // B列のタスク

      // A列が有効な日付で、B列に値がある場合のみ処理
      if (dateValue instanceof Date && taskValue) {
        const currentDate = new Date(dateValue.getFullYear(), dateValue.getMonth(), dateValue.getDate());

        // 日付が指定された範囲内にあるかチェック
        if (currentDate >= start && currentDate <= end) {
          let taskItems = [];

          // B列の値を指定のルールで配列に変換
          if (String(taskValue).includes('\n')) {
            taskItems = taskValue.split('\n');
          } else if (String(taskValue).includes(',')) {
            taskItems = taskValue.split(',');
          } else {
            taskItems = [taskValue];
          }
          
          // 各タスクを [日付, タスク] の形式で結果配列に追加
          taskItems.forEach(item => {
            const trimmedItem = String(item).trim();
            if (trimmedItem) { // 空の項目は除外
              results.push([dateValue, trimmedItem]);
            }
          });
        }
      }
    }
    
    // 該当するタスクがない場合は空の配列を返す
    return results;

  } catch (e) {
    Logger.log(e.toString());
    // エラーが発生した場合は、呼び出し元でエラーをハンドリングできるように再スローする
    throw e;
  }
}

/**
 * 与えられた二次元配列を処理し、日付をキーとして同じ日付のテキスト要素を結合し、
 * 最終的に日付と結合されたテキスト要素が交互に並ぶフラットな配列を返します。
 *
 * @param {Array<Array<Date|string>>} tasks - 処理する二次元配列。
 * 各内部配列は、最初の要素がDateオブジェクト、2番目の要素が文字列であると想定されます。
 * @return {Array<string>} 処理されたフラットな配列。
 */
function processAndFlattenArray(tasks) {
  const groupedTasks = new Map();

  tasks.forEach(row => {
    const date = row[0];
    const text = row[1];

    // 日付部分のみを文字列として取得（例: '2025-06-11'）
    const dateKey = Utilities.formatDate(date, 'JST', 'yyyy-MM-dd');

    if (!groupedTasks.has(dateKey)) {
      groupedTasks.set(dateKey, []);
    }
    groupedTasks.get(dateKey).push(text);
  });

  const result = [];
  // 日付キーをソート
  const sortedDateKeys = Array.from(groupedTasks.keys()).sort();

  sortedDateKeys.forEach(dateKey => {
    // 表示用の日付フォーマットを作成（例: '6月11日'）
    // dateKeyは'yyyy-MM-dd'形式なので、Dateオブジェクトに変換してフォーマット
    const displayDate = Utilities.formatDate(new Date(dateKey), 'JST', 'M月dd日');
    result.push(displayDate);
    result.push(...groupedTasks.get(dateKey));
  });

  return result;
}

/**
 * 指定された日付の isShortened が True かどうかを判定する関数
 * @param {Date} targetDate - 判定したい日付
 * @param {string} spreadsheetId - スプレッドシートのID
 * @return {boolean} - isShortenedがTrueならtrue、それ以外（または日付がない場合）はfalse
 */
function isShortened(targetDate, spreadsheetId) {
  try {
    const spreadsheet = SpreadsheetApp.openById(spreadsheetId);
    // シート名は元のコードに合わせて '提出物' としています。
    // 実際のシート名と異なる場合は適宜変更してください。
    const sheet = spreadsheet.getSheetByName('提出物');

    if (!sheet) {
      throw new Error("'提出物' という名前のシートが見つかりません。");
    }

    const dataRange = sheet.getDataRange();
    const values = dataRange.getValues();

    // 指定された日付の年月日を取得（比較用）
    const targetYear = targetDate.getFullYear();
    const targetMonth = targetDate.getMonth();
    const targetDay = targetDate.getDate();

    // 1行目はヘッダーとしてスキップし、2行目からループ
    for (let i = 1; i < values.length; i++) {
      const row = values[i];
      const dateValue = row[0]; // A列の日付

      // A列が有効な日付オブジェクトである場合のみ処理
      if (dateValue instanceof Date) {
        
        // 年月日が完全に一致するかチェック
        if (dateValue.getFullYear() === targetYear && 
            dateValue.getMonth() === targetMonth && 
            dateValue.getDate() === targetDay) {
          
          const isShortenedValue = row[1]; // B列の値 (インデックスは1)
          
          // 値が boolean の true、または文字列の 'TRUE' の場合に true を返す
          if (isShortenedValue === true || String(isShortenedValue).toUpperCase() === 'TRUE') {
            return true;
          } else {
            return false;
          }
        }
      }
    }
    
    // 指定された日付がシート内に見つからなかった場合は false を返す
    return false;

  } catch (e) {
    Logger.log(e.toString());
    // エラーが発生した場合は再スロー
    throw e;
  }
}
