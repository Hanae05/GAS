// スプレッドシートの行列情報
var StartRow = 2;         // 開始行
var NameCol = 1;          // 名前列
var BirthdayCol = 2;      // 誕生日列

//------------------------------
// main処理
//------------------------------
function main() {
  // 今日の日付を取得
  var nowDate = new Date();
  
  // 今日の曜日を取得
  var todayDayOfWeek = nowDate.getDay();
  Logger.log(todayDayOfWeek);
  
  // 誕生日リストのスプレッドシートを参照
  var spreadsheet = SpreadsheetApp.openById('13g-BC6UvpVfQIH1B_neBHvpVcnSb_MlXhVzx-DEFXJ0');
  
  // シートにアクセス
  var sheet = spreadsheet.getSheetByName('シート1');
  
  // データ取得
  var data = sheet.getDataRange().getValues();
  
  for (var i = 0; i < data.length; i++) {
    if (i < StartRow) {
      continue;
    }
    
    // 名前を取得
    var name = data[i][NameCol];
    
    // 誕生日を取得
    var birthday = data[i][BirthdayCol];
    
    // 誕生日を日付でパースし、日付が正しく入力されているかチェック
    var isDate = !isNaN(Date.parse(birthday));
    
    // 誕生日入力済みであることをチェック
    if (isDate === true) {
      // 投稿処理
      postProc(name, birthday, nowDate); 
    }
  }
}

//------------------------------
// 投稿処理
//------------------------------

function postProc(name, birthday, nowDate)
{
  var todayDayOfWeek = nowDate.getDay();
  console.log(todayDayOfWeek);
  var message = "";

  // 日曜日仮
  if (todayDayOfWeek === 6){
    
    // 1日後(月曜日)
    var dtMon = new Date();
    dtMon.setDate(dtMon.getDate() + 1);
    
    // 2日後(火曜日)
    var dtTue = new Date();
    dtTue.setDate(dtTue.getDate() + 2);
    
    // 3日後(水曜日)
    var dtWed = new Date();
    dtWed.setDate(dtWed.getDate() + 3);
    
    //4日後(木曜日)
    var dtThu = new Date();
    dtThu.setDate(dtThu.getDate() + 4);
    
    //5日後(金曜日)
    var dtFri = new Date();
    dtFri.setDate(dtFri.getDate() + 5)
    
    //6日後(土曜日)
    var dtSat = new Date();
    dtSat.setDate(dtSat.getDate() + 6);
    
    //7日後(日曜日)
    var dtSun = new Date();
    dtSun.setDate(dtSun.getDate() + 7);
    
    
    if (isSend(dtMon, birthday) === true) {
      message = createNotice(name, "明日");
    }
    else if (isSend(dtTue, birthday) === true) {
      message = createNotice(name, "来週の火曜日");
    }
    else if (isSend(dtWed, birthday) === true) {
      message = createNotice(name, "来週の水曜日");
    }
    else if (isSend(dtThu, birthday) === true) {
      message = createNotice(name, "来週の木曜日");
    }
    else if (isSend(dtFri, birthday) === true) {
      message = createNotice(name, "来週の金曜日");
    }
    else if (isSend(dtSat, birthday) === true) {
      message = createNotice(name, "来週の土曜日");
    }
    else if (isSend(dtSun, birthday) === true) {
      message = createNotice(name, "来週の日曜日");
    }
    else if (isSend(nowDate, birthday) === true) {
      message = createMessage(name, "今日");
    }
  }
  else
  {
    if (isSend(nowDate, birthday) === true) {
      message = createMessage(name, "今日");
    }      
  }

  // メッセージが生成されていれば、Slackに投稿  
  if (message != "") {
    // おめでとう画像を取得
    var fileId = getRandomImageID();
        
    // Slackへ投稿
    postSlack(message, fileId);
  }
}

//------------------------------
// 送信有無取得
//------------------------------
function isSend(compareDate, birthdayDate)
{
  // 月、日が一致するかをチェックする
  if ((compareDate.getMonth() === birthdayDate.getMonth()) && 
    (compareDate.getDate() === birthdayDate.getDate()))
    {
      return true;
    }
  
  return false;
}

//------------------------------
// メッセージ作成
//------------------------------

//今日が誕生日の場合
function createMessage(name, dayMessage)
{
  var message = dayMessage + 'は, `' + name + '`さんの誕生日です :birthday:' + '\r\n' + 
    'おめでとうございます :tada: :tada: :tada:'
    
  return message;  
}


//来週が誕生日の場合
function createNotice(name, dayMessage)
{
  var message = dayMessage + 'は, `' + name + '`さんの誕生日です :birthday:' + '\r\n' + 
    'みんなでお祝いしましょう :tada: '
    
  return message;
}

//------------------------------
// 誕生日おめでとう画像のIDを取得
//------------------------------
function getRandomImageID() {
  var imageFileArray = [
    "16tpMv5UsCh83TTzFnCKZMYtmA-jej0DC",
    "1QoLp1SeNcD_YpnwxF5PWiYlJbR5PNqZa",
    "1pYL_I2XbfAIAAiQFSDeoEbdvUNavTPrY",
    "1uf7Lx8pFQNWZveO0KVZlUoRG0vGEIOht",
    "1fQsdr-VSpQUKb7YU8DLRk8f2b-hmQedy",
    "1BIoZ9tvZlMd-I9wbxCI6DL8ohd26rFR3",
    "1P0JnOTQVGmpOClERd8um5AZKqQQ9DvHp",
    "1eNIjKm7983mmv-PNZ6jE_WqR5InsizxY",
  ];

  // 0～7の値を取得
  var randomValue = Math.floor( Math.random() * 8 );
  
  return imageFileArray[randomValue];
}

//------------------------------
// Slackへ投稿
//------------------------------
function postSlack(text, fileId) {
  
  var payload = {
    text: text,
    attachments: [
        {
            color: "#2eb886",
            pretext: "",
            image_url: "http://drive.google.com/uc?export=download&id=" + fileId,
        }      
      ]
  };
 
  // Incoming WebhooksのURL
  var url = "https://hooks.slack.com/services/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";
  
  var options = {
    "method" : "POST",
    "headers": {"Content-type": "application/json"},
    "payload": JSON.stringify(payload)
  };
  UrlFetchApp.fetch(url, options); 
}

