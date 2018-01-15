this.config = {};

/*
 * 毎日1回実行され、その日の予定を通知する
 */
function dayMain() {
  // configデータを取得する
  initConfigData();
  
  // 以前のトリガーを削除する
  deleteDayTrigger();
  
  // Slackに投稿する内容
  var contents = "今日の予定\n";
  contents += "```\n";

  // 当日のイベントを取得する
  var result = getDayEvents(this.config.calendar_id);

  // イベントがあればそれを記載する
  if (result != "") {
    contents += result;
  } else {
    // イベントが無いければない旨を記載する
    contents += "今日の予定は何もありません";
  }
  
  contents += "```\n";
  
  // Slackに投稿する
  const payload = {
    "text": contents,
    "channel": this.config.channel,
    "username": this.config.username,
    "icon_url": this.config.icon_url,
  };
  
  postSlack(payload);
}

/*
 * 週に一度実行され、翌週の予定を全て表示する
 */
function weekMain() {
  // configデータを設定する
  initConfigData();
  
  // 以前のトリガーを削除する
  deleteWeekTrigger();
  
  // Slackに投稿する内容
  var contents = "来週の予定\n";

  // イベントの有無を判定するフラグ
  var flag = false;

  // 次に調べる日を格納する変数
  // (Moment.jsをライブラリに登録しておく必要あり)
  const next_date = Moment.moment();
  for (var i=1; i<=7; i++) {
    // 1日ずらす
    next_date.add(1, 'days');

    // 指定した日のイベントを取得する
    var tmp = getDayEvents(this.config.calendar_id, next_date);

    // イベントがあればそれを記載する
    if (tmp) {
      contents += "> " + next_date.format("MM/DD") + "(" + getDayJP(next_date.day()) + ")\n";
      contents += "```\n";
      contents += tmp;
      contents += "```\n";
      
      // イベント存在フラグを立てる
      flag = true;
    }
  }

  // 何もイベントがなければその旨を記載
  if (!flag) {
    contents += "```\n";
    contents += "来週の予定は何もありません";
    contents += "```\n";
  }
  
  // Slackに投稿する
  const payload = {
    "text": contents,
    "channel": this.config.channel,
    "username": this.config.username,
    "icon_url": this.config.icon_url,
  };
  
  postSlack(payload);
}

/*
 * 指定された日（指定されなければ当日）のイベントを全て取得する
 */
function getDayEvents(calendar_id, d) {
  const date = d || Moment.moment();

  // イベントの内容を文字列化して返す変数
  var contents = "";

  // カレンダーの取得とイベントの取得
  const calendar = CalendarApp.getCalendarById(calendar_id);
  const events = calendar.getEventsForDay(date.toDate());
  
  for (var i=0; i<events.length; i++) {
    // 終日イベント
    if (events[i].isAllDayEvent()) {
      var tmp = date.format("MM/DD") + " ";
      tmp += events[i].getTitle()+"\n";
      contents = tmp + contents;
    } else {
      // 時間指定イベント
      var tmp = Utilities.formatDate(events[i].getStartTime(), "GMT+0900", "MM/dd HH:mm");
      tmp += Utilities.formatDate(events[i].getEndTime(), "GMT+0900", "-HH:mm ");
      tmp += events[i].getTitle()+"\n";
      contents = contents + tmp;
    }
  }
  
  return contents;
}

/*
 * Slackに投稿する
 */
function postSlack(payload) {
  const options = {
    "method": "post",
    "contentType": "application/json",
    "payload": JSON.stringify(payload)
  };
  
  Logger.log(options);
  
  UrlFetchApp.fetch(this.config.post_slack_url, options);
}

/*
 * 毎日投稿する時間をセットする
 * （GASのトリガーは分単位で指定できないため、この関数で詳細時間をセットする）
 */
function setDayTrigger() {
  // configデータを設定する
  initConfigData();
  
  // 発火時間を設定する
  var triggerDate = new Date();
  triggerDate.setDate(triggerDate.getDate()+1);
  triggerDate.setHours(this.config.day_trigger_hour);
  triggerDate.setMinutes(this.config.day_trigger_minute);
  
  // 時間指定してトリガーをセットする
  ScriptApp.newTrigger("dayMain").timeBased().at(triggerDate).create();
}

/*
 * 週１で投稿する時間をセットする
 * (GASのトリガーは時間指定できないため、1時間前にセットしておき、この関数で詳細時間をセットする)
 */
function setWeekTrigger() {
  // configデータを設定する
  initConfigData();
  
  // 発火時間を設定する
  var triggerDate = new Date();
  triggerDate.setHours(this.config.week_trigger_hour);
  triggerDate.setMinutes(this.config.week_trigger_minute);

  // 時間指定してトリガーをセットする
  ScriptApp.newTrigger("weekMain").timeBased().at(triggerDate).create();
}

/*
 * 設定した毎日のトリガーを削除する
 */
function deleteDayTrigger() {
  // トリガーの取得
  const triggers = ScriptApp.getProjectTriggers();
  for (var i=0; i<triggers.length; i++) {
    // 毎日投稿するトリガーであれば削除
    if (triggers[i].getHandlerFunction() == "dayMain"){
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }
}

/*
 * 設定した週1のトリガーを削除する
 */
function deleteWeekTrigger() {
  // トリガーの取得
  const triggers = ScriptApp.getProjectTriggers();
  for (var i=0; i<triggers.length; i++) {
    // 週1で投稿するトリガーなら削除
    if (triggers[i].getHandlerFunction() == "weekMain"){
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }
}

/*
 * 日本語名の曜日を取得する
 */
function getDayJP(num) {
  switch(num) {
    case 0:
      return "日";
    case 1:
      return "月";
    case 2:
      return "火";
    case 3:
      return "水";
    case 4:
      return "木";
    case 5:
      return "金";
    case 6:
      return "土";
  }
}

/*
 * スプレッドシートから設定データを取得する
 */
function initConfigData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("config");
  
  const last_row = sheet.getLastRow();
  for (var i=1; i<=last_row; i++) {
    key = sheet.getRange(i, 1).getValue();
    value = sheet.getRange(i, 2).getValue();
    
    this.config[key] = value;
  }
}