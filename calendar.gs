/**
 * Google Calendar のイベントを取得・整形するクラス
 */
class CalendarEventFetcher {
  constructor(date, calendarId) {
    // 指定された日付（開始と終了の範囲を決定するために使用）
    this.date = new Date(date);
    // 指定されたカレンダーID
    this.calendarId = calendarId;
    // カレンダーサービスの取得
    this.calendar = CalendarApp.getCalendarById(calendarId);
  }

  /**
   * 指定した日付のイベントをすべて取得する
   * @return {Array} 2次元配列 [タイトル, 説明] の形式
   */
  getSortedEvents() {
    // その日の 0:00 から 23:59:59 までを範囲として設定
    const startTime = new Date(this.date);
    startTime.setHours(0, 0, 0, 0);
    const endTime = new Date(this.date);
    endTime.setHours(23, 59, 59, 999);

    // イベントを取得
    const events = this.calendar.getEvents(startTime, endTime);

    // 開始時刻 → 終了時刻の順にソート
    events.sort((a, b) => {
      const startDiff = a.getStartTime() - b.getStartTime();
      if (startDiff !== 0) return startDiff;

      const endDiff = a.getEndTime() - b.getEndTime();
      return endDiff;
    });

    // タイトルと説明を2次元配列として格納
    return events.map(event => [event.getTitle(), event.getDescription()]);
  }
}

/**
 * 公開関数：指定した日付とカレンダーIDに基づき、イベントを取得する
 * @param {string} date - 対象日付（例: '2025-04-18'）
 * @param {string} calendarId - 対象のGoogleカレンダーID
 * @return {Array} 2次元配列 [タイトル, 説明] の形式
 */
function getEventsForDate(date, calendarId) {
  // イベント取得クラスのインスタンスを生成
  const fetcher = new CalendarEventFetcher(date, calendarId);
  // ソートされたイベントを取得して返す
  return fetcher.getSortedEvents();
}
