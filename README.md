# slack-scheduler

## 概要

Google Spread Sheetからスクリプトエディタを開き、トリガーで通知したいイベントの関数を指定します。

通知イベント関数は以下の種類です。
- 日毎通知（dayMain()）
    - 指定した時間にその日の予定をすべて通知します。
    - setDayTrigger()をトリガーに指定すると分刻みで時間を指定することが可能です。
- 週毎通知（weekMain()）
    - 指定した日時に翌日から1週間分の予定をすべて通知します。
    - setWeekTrigger()をトリガーに指定すると分刻みで時間を指定することができます。
- リアルタイム通知(notifyEvents())
    - Google Calendarに新規登録・削除・~~変更~~した予定を通知します。
    - トリガーで検知時間間隔を設定するとその時間刻みでGoogle Calendarから情報を取得し、変更があれば通知します。

## 使い方

### 前準備（Spread Sheet）

1. Google Spread Sheetを作成します。
1. eventシートとconfigシートを作成します。
    - 必ず「event」と「config」で作成する必要があります。
1. configシートに以下の情報を追加します。
    - A列に以下のキー名を列挙します。
        - calendar_id
        - channel
        - username
        - icon_url
        - post_slack_url
        - day_trigger_hour
        - day_trigger_minute
        - week_trigger_hour
        - week_trigger_minute
    - B列にそれぞれのキーに対応する値を入れます。

| キー名 | 概要 | 値の例 |
| :--: | :-- | :--: |
| calendar_id | Google CalendarのID | xxxx@google.com |
| channel | Slackのチャンネル | #schedule |
| username | 通知するbotのユーザ名 | schedule_bot |
| icon_url | 通知するbotの画像 | http://xxxxx |
| post_slack_url | slackのwebhook URL | https://webhook.slack.com/xxxxxxx |
| day_trigger_hour | 日毎通知の時間 | 8 |
| day_trigger_minute | 日毎通知の分 | 30 |
| week_trigger_hour | 週毎通知の時間 | 8 |
| week_trigger_minute | 日毎通知の時間 | 0 |


### 実行（スクリプトエディタ）

1. Spread Sheetのメニューからスクリプトエディタを開きます。
1. ソースコードをエディタ内に記述します。
1. メニューからトリガー設定を開き、呼び出したい関数を設定します。
