function main() {
  const scriptProperties = PropertiesService.getScriptProperties();
  const GOOGLE_CALENDAR_ID = scriptProperties.getProperty("GOOGLE_CALENDAR_ID");
  const GOOGLE_CHAT_WEBHOOK_URL = scriptProperties.getProperty("GOOGLE_CHAT_WEBHOOK_URL");
  const TASKS_SPREADSHEET_ID =scriptProperties.getProperty("TASKS_SPREADSHEET_ID")

  // 時間割を取得
  const events_date = getTomorrowDate()
  events_data = getEventsForDate(events_date, GOOGLE_CALENDAR_ID)

  output_text_event = ""
  for (let i = 0; i < events_data.length; i++){
    table = `\n${i+1}) ${events_data[i][0]}`;
    output_text_event += table;
  }

  const output_text_date = getTomorrowDate_ja();

  // 提出物を取得
  const today = new Date();
  const startDate = new Date();
  startDate.setDate(today.getDate() + 1);  // 開始日を「明日の日付」に設定
  const endDate = new Date();
  endDate.setDate(today.getDate() + 4);  // 1. 開始日を今日から4日目に設定

  tasks = getTasksByDateRange(startDate, endDate, TASKS_SPREADSHEET_ID)
  tasks_array = processAndFlattenArray(tasks)
  console.log(tasks_array)

  output_tasks = ""
  for (let i = 0; i < tasks_array.length; i++){
    if (i == 0) {
      output_tasks += `${tasks_array[i]}`;
    } else {
      output_tasks += `\n${tasks_array[i]}`;
    }
  }

  // チャットに送信
  console.log("output_tasks:" + output_tasks)
  if (output_text_event && tasks_array.length > 0){
    const output_text = `*予定*\n${output_tasks}`
    google_chat_webhook(output_text, GOOGLE_CHAT_WEBHOOK_URL);
    console.log("message:" + output_tasks)
  }
}

/**
 * 明日の日付を "YYYY-MM-DD" 形式で取得する関数
 * @return {string} 明日の日付（例: '2025-04-19'）
 */
function getTomorrowDate() {
  const today = new Date();
  today.setDate(today.getDate() + 1);

  const year = today.getFullYear();
  const month = ('0' + (today.getMonth() + 1)).slice(-2);
  const day = ('0' + today.getDate()).slice(-2);

  return `${year}-${month}-${day}`;
}

/**
 * 明日の日付を "M月D日 ()" 形式で取得する関数
 * @return {string} 明日の日付（例: '04月19日 (金)'）
 */
function getTomorrowDate_ja() {
  const today = new Date();
  const tomorrow = new Date(today);
  tomorrow.setDate(today.getDate() + 1);

  const month = Utilities.formatDate(tomorrow, Session.getScriptTimeZone(), 'M'); // 'MM' を 'M' に変更
  const day = Utilities.formatDate(tomorrow, Session.getScriptTimeZone(), 'dd');

  const days = ["日", "月", "火", "水", "木", "金", "土"]

  const formattedDate = `${month}月${day}日 (${days[tomorrow.getDay()]})`;

  return formattedDate;
}
