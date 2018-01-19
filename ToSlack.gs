this.config = {};

/** 以下、定期通知の関数 **/

/**
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
  var payload = {
    "text": contents,
    "channel": this.config.channel,
    "username": this.config.username,
    "icon_url": this.config.icon_url,
  };
  
  postSlack(payload);
}

/**
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
  var next_date = Moment.moment();
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
  var payload = {
    "text": contents,
    "channel": this.config.channel,
    "username": this.config.username,
    "icon_url": this.config.icon_url,
  };
  
  postSlack(payload);
}

/**
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

/**
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

/**
 * 設定した毎日のトリガーを削除する
 */
function deleteDayTrigger() {
  // トリガーの取得
  var triggers = ScriptApp.getProjectTriggers();
  for (var i=0; i<triggers.length; i++) {
    // 毎日投稿するトリガーであれば削除
    if (triggers[i].getHandlerFunction() == "dayMain"){
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }
}

/**
 * 設定した週1のトリガーを削除する
 */
function deleteWeekTrigger() {
  // トリガーの取得
  var triggers = ScriptApp.getProjectTriggers();
  for (var i=0; i<triggers.length; i++) {
    // 週1で投稿するトリガーなら削除
    if (triggers[i].getHandlerFunction() == "weekMain"){
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }
}

/** 以下、検知型通知の関数 **/

/**
 * イベントの新規登録・削除を通知する
 * SpreadSheetをデータベースとし、Calendarのイベントと比較することで差分を検知する
 * トリガーにて1分毎に実行するように指定（トリガーの上限時間に引っかからないように注意する）
 */
function notifyEvents() {
  // configデータを設定する
  initConfigData();

  // SpreadSheetに記録しているイベントを取得する
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("event");
  var last_row = sheet.getLastRow();
  
  var saved_events = {};
  for (var i=3; i<=last_row; i++) {
    var id = sheet.getRange(i, 1).getValue();
    var title = sheet.getRange(i, 2).getValue();
    var start_time = Moment.moment(sheet.getRange(i, 3).getValue());
    var end_time = Moment.moment(sheet.getRange(i, 4).getValue());
    var is_allday = (sheet.getRange(i, 5).getValue() ? true : false);
    
    if (Moment.moment().isAfter(end_time)) {
      sheet.deleteRow(i);
      i--;
      last_row--;
      continue;
    } 
    
    if (!saved_events[id]) {
      saved_events[id] = {
        title: title,
        start_time: start_time,
        end_time: end_time,
        is_allday: is_allday,
        is_deleted: true,
        row: i,
      }
    } else {
      Logger.log('重複したイベントが見つかりました。');
      return;
    }
  }
  
  var start_date = Moment.moment();
  var end_date = start_date.clone().add(1, 'years').endOf('day');
  
  var calendar = CalendarApp.getCalendarById(this.config.calendar_id);
  var events = calendar.getEvents(start_date.toDate(), end_date.toDate());
  
  var contents = '';
  for (var i=0; i<events.length; i++) {
    var event_id = events[i].getId();
    
    // 保存されていないイベント（=新規イベント）があった場合の処理
    if (!saved_events[event_id]) {
      var st = Moment.moment(events[i].getStartTime()).format('YYYY-MM-DD HH:mm');
      var et = Moment.moment(events[i].getEndTime()).format('YYYY-MM-DD HH:mm');

      saved_events[event_id] = {
        title: events[i].getTitle(),
        start_time: st,
        end_time: et,
        is_allday: events[i].isAllDayEvent(),
        is_deleted: false,
      };

      last_row++;
      
      // SpreadSheetに新規追加する
      sheet.getRange(last_row, 1).setValue(event_id);
      sheet.getRange(last_row, 2).setValue(events[i].getTitle()); 
      sheet.getRange(last_row, 3).setValue(st); 
      sheet.getRange(last_row, 4).setValue(et); 
      
      contents += '```\n';
      // 終日イベントの場合
      if (events[i].isAllDayEvent()) {
        sheet.getRange(last_row, 5).setValue('◯');
        contents += Moment.moment(events[i].getStartTime()).format('MM/DD');
      } else {
        // 時間指定イベントの場合
        contents += Moment.moment(events[i].getStartTime()).format('MM/DD HH:mm') + '-' +
          Moment.moment(events[i].getEndTime()).format('HH:mm');
      }
      
      contents += ' ' + events[i].getTitle() + '\n';
      contents += '```\n';
    } else {
      saved_events[event_id].is_deleted = false;
    }
  }
  
  // 削除されたイベントの通知
  var removed_contents = '';
  for (var id in saved_events) {
    var event = saved_events[id];
    
    if (event.is_deleted) {
      removed_contents += '```\n';
      if (event.is_allday) {
        removed_contents += Moment.moment(event.start_time).format('MM/DD');
      } else {
        removed_contents += Moment.moment(event.start_time).format('MM/DD HH:mm') + '-' +
          Moment.moment(event.end_time).format('HH:mm');
      }
      removed_contents += ' ' + event.title + '\n';
      removed_contents += '```\n';
      
      sheet.deleteRow(event.row);
    }
  }
  
  if (removed_contents) {
    removed_contents = '予定がキャンセルされました。\n' + removed_contents;
    
    // Slackに投稿する
    var payload = {
      "text": removed_contents,
      "channel": this.config.channel,
      "username": this.config.username,
      "icon_url": this.config.icon_url,
    };
    
    postSlack(payload);
  }    

  // 新規イベントがあれば内容を整えてSlackに投稿する
  if (contents) {
    contents = '予定が追加されました。\n' + contents;
    
    // Slackに投稿する
    var payload = {
      "text": contents,
      "channel": this.config.channel,
      "username": this.config.username,
      "icon_url": this.config.icon_url,
    };
    
    postSlack(payload);
  }
}

/** 以下、Util系関数 **/

/**
 * 指定された日（指定されなければ当日）のイベントを全て取得する
 */
function getDayEvents(calendar_id, d) {
  var date = d || Moment.moment();

  // イベントの内容を文字列化して返す変数
  var contents = "";

  // カレンダーの取得とイベントの取得
  var calendar = CalendarApp.getCalendarById(calendar_id);
  var events = calendar.getEventsForDay(date.toDate());
  
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

/**
 * Slackに投稿する
 */
function postSlack(payload) {
  var options = {
    "method": "post",
    "contentType": "application/json",
    "payload": JSON.stringify(payload)
  };
  
  Logger.log(options);
  
  UrlFetchApp.fetch(this.config.post_slack_url, options);
}

/**
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

/**
 * スプレッドシートから設定データを取得する
 */
function initConfigData() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("config");
  
  var last_row = sheet.getLastRow();
  for (var i=1; i<=last_row; i++) {
    key = sheet.getRange(i, 1).getValue();
    value = sheet.getRange(i, 2).getValue();
    
    this.config[key] = value;
  }
}

/**
 * SpreadSheetにをログを出力する
 */
function log(type, str) {
}