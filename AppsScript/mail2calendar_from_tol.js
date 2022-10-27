const QUERY_RESERVE = 'subject:[tol] 「パーソナルジム Epice【エピス】」への予約を承認しました label:00_Task-02_tol label:inbox'
const LOCATION = ''
const TITLE = 'ジム'
const LABEL = '00_Task/02_tol'
const DATE_PREFIX = '予約日時：'

//
// メイン処理
//
function main() {
  pickUpMessage(QUERY_RESERVE, function (message) {
    parseReserve(message);
  });
}

//
// メール検索処理（コールバックにして他のパースもできるように）
//
function pickUpMessage(query, callback) {
  const threads = GmailApp.search(query, 0, 5);
  for (var x in threads) {
    var thread = threads[x]
    //解析処理
    for (var message of thread.getMessages()) {
      callback(message)
    }
    //アーカイブして処理対象外にする
    thread.moveToArchive();
  }
}

//
// メール解析
//
function parseReserve(message) {
  const strDate = message.getDate();
  const strMessage = message.getPlainBody();

  const regexp = RegExp(DATE_PREFIX + '.*', 'gi');

  const result = strMessage.match(regexp);
  if (result == null) {
    console.log("This message doesn't have info.");
    return;
  }
  const parsedDate = result[0].replace(DATE_PREFIX, '');
  console.log(parsedDate)

  const year = (new Date().getFullYear()); // 年を持っていないので処理年に
  const month = parsedDate.match(/[0-9]{1,2}月/i)[0].replace('月', ''); //１桁の月は0なし
  const dayOfMonth = parsedDate.match(/[0-9]{1,2}日/i)[0].replace('日', '');
  const startTimeHour = parsedDate.match(/[0-9]{1,2}:/i)[0].replace(':', '');
  const startTimeMinutes = parsedDate.match(/[0-9]{1,2}~/i)[0].replace('~', '');
  const endTimeHour = parsedDate.match(/~[0-9]{1,2}/i)[0].replace('~', '');
  const endTimeMinutes = parsedDate.match(/[0-9]{1,2}$/i)[0]

  //カレンダー
  createEvent(TITLE, "mailDate: " + strDate,
    LOCATION, year, month, dayOfMonth, startTimeHour, startTimeMinutes, endTimeHour, endTimeMinutes);
}


//
// カレンダー登録
//
function createEvent(title, description, location, year, month, dayOfMonth,
  startTimeHour, startTimeMinutes, endTimeHour, endTimeMinutes) {

  const calendar = CalendarApp.getDefaultCalendar();
  const startTime = new Date(year, month - 1, dayOfMonth, startTimeHour, startTimeMinutes, 0);
  const endTime = new Date(year, month - 1, dayOfMonth, endTimeHour, endTimeMinutes, 0);
  const option = {
    description: description,
    location: location,
  }

  console.log("start time: " + startTime);
  console.log("end time: " + endTime);
  var calendarEvent = calendar.createEvent(title, startTime, endTime, option);
  calendarEvent.setColor(CalendarApp.EventColor.ORANGE)
}

//
// パーステスト
//
function test() {
  console.log('8月11日(土) 11:00~11:50'.match(/[0-9]{1,2}月/i)[0].replace('月', ''))
  console.log('10月11日(土) 11:00~11:50'.match(/[0-9]{1,2}月/i)[0].replace('月', ''))
  console.log('10月09日(土) 11:00~11:50'.match(/[0-9]{1,2}日/i)[0].replace('日', ''))
  console.log('10月9日(土) 11:00~11:50'.match(/[0-9]{1,2}日/i)[0].replace('日', ''))
  console.log('10月10日(土) 11:00~11:50'.match(/[0-9]{1,2}日/i)[0].replace('日', ''))

  //開始時刻（時）
  console.log('10月22日(土) 11:00~11:50'.match(/[0-9]{1,2}:/i)[0].replace(':', ''))
  console.log('10月22日(土) 9:00~9:50'.match(/[0-9]{1,2}:/i)[0].replace(':', ''))
  //開始時刻（分）
  console.log('10月22日(土) 11:50~12:50'.match(/[0-9]{1,2}~/i)[0].replace('~', ''))
  console.log('10月22日(土) 9:15~10:15'.match(/[0-9]{1,2}~/i)[0].replace('~', ''))

  //終了時刻（時）
  console.log('10月22日(土) 9:15~10:15'.match(/~[0-9]{1,2}/i)[0].replace('~', ''))
  console.log('10月22日(土) 7:15~7:15'.match(/~[0-9]{1,2}/i)[0].replace('~', ''))
  
  //終了時刻（分）
  console.log('10月22日(土) 7:15~7:3'.match(/[0-9]{1,2}$/i)[0])
  console.log('10月22日(土) 7:15~7:30'.match(/[0-9]{1,2}$/i)[0])
}

