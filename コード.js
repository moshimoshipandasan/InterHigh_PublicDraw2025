function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('インターハイ抽選（剣道競技）')
//      .addItem('1.男女の抽選順', 'randomizeCells1_2')
//      .addItem('1.団体抽選の開始', 'drandomizeCells')
      .addItem('未抽選番号のみ抽選', 'randomizeRemainingNumbers')
      .addItem('左サイドの抽選(1-24)', 'randomizeCells1_24')
      .addItem('右サイドの抽選(25-48)', 'randomizeCells25_48')
      // .addItem('第1ゾーン抽選(1-12)', 'randomizeZone1')
      // .addItem('第2ゾーン抽選(13-24)', 'randomizeZone2')
      // .addItem('第3ゾーン抽選(25-36)', 'randomizeZone3')
      // .addItem('第4ゾーン抽選(37-48)', 'randomizeZone4')
      .addItem('Aリーグ抽選(1-3)', 'randomizeALeague')
      .addItem('Pリーグ抽選(46-48)', 'randomizePLeague')
      .addItem('H&Iリーグ抽選(22-27)', 'randomizeHILeague')
      .addItem('Hリーグ抽選(22-24)', 'randomizeHLeague')
      .addItem('Iリーグ抽選(25-27)', 'randomizeILeague')
//      .addItem('3.開催地抽選済番号オープン', 'resetBackground')
      .addItem('デモ', 'demo')
      .addToUi();
}

function demo() {
  // 抽選シートをアクティブにする
  var ss = SpreadsheetApp.getActive();
  var sheet = ss.getSheetByName("抽選");
  sheet.activate();
  
  var numRows = 6;
  var numCols = 8;
  var data = [];
  var rangeToFill = sheet.getRange("A1:H6");
  rangeToFill.setBackground("#dcdcdc");
  rangeToFill.setFontColor("black");
  
  // 1から48までの乱数を生成して、配列に追加する
  var nums = [];
  for (var i = 1; i <= 48; i++) {
    nums.push(i);
  }
  for (var i = nums.length - 1; i > 0; i--) {
    var j = Math.floor(Math.random() * (i + 1));
    var temp = nums[i];
    nums[i] = nums[j];
    nums[j] = temp;
  }
  for (var i = 0; i < numRows; i++) {
    var row = [];
    for (var j = 0; j < numCols; j++) {
      row.push(nums[i * numCols + j]);
    }
    data.push(row);
  }
  
  var range = sheet.getRange(1, 1, numRows, numCols);
  range.setValues(data);  
}

function randomizeCells1_2() {
  // 抽選シートをアクティブにする
  var ss = SpreadsheetApp.getActive();
  var sheet = ss.getSheetByName("抽選");
  sheet.activate();
  
  var numRows = 1;
  var numCols = 2;
  var data = [];
  var rangeToFill0 = sheet.getRange("A1:H6");
  var rangeToFill = sheet.getRange("A1:B1");
  rangeToFill0.clearContent();
  rangeToFill0.setBackground("white");
  rangeToFill.setBackground("green");
  rangeToFill.setFontColor("green");

  // 1から24までの乱数を生成して、配列に追加する
  var nums = [];
  for (var i = 1; i <= 2; i++) {
    nums.push(i);
  }
  for (var i = nums.length - 1; i > 0; i--) {
    var j = Math.floor(Math.random() * (i + 1));
    var temp = nums[i];
    nums[i] = nums[j];
    nums[j] = temp;
  }
  for (var i = 0; i < numRows; i++) {
    var row = [];
    for (var j = 0; j < numCols; j++) {
//      row.push(nums[i * numCols + j]);
      if(nums[i * numCols + j] == 1){row.push("男")}else{row.push("女")};
    }
    data.push(row);
  }
  
  var range = sheet.getRange(1, 1, numRows, numCols);
  range.setValues(data);  
}

function randomizeCells1_24() {
  // 抽選シートをアクティブにする
  var ss = SpreadsheetApp.getActive();
  var sheet = ss.getSheetByName("抽選");
  sheet.activate();
  
  var numRows = 4;
  var numCols = 6;
  var data = [];
  var rangeToFill0 = sheet.getRange("A1:H6");
  rangeToFill0.clearContent();
  rangeToFill0.setBackground("white");

  // 抽選順一覧シートからA列の値を取得
  var targetSheet = SpreadsheetApp.getActive().getSheetByName("抽選順一覧");
  var lotteryValues = targetSheet.getRange("A:A").getValues();
  
  // 抽選済みの番号を配列に格納
  var usedNumbers = [];
  for (var i = 0; i < lotteryValues.length; i++) {
    if (lotteryValues[i][0] !== "" && !isNaN(lotteryValues[i][0])) {
      usedNumbers.push(parseInt(lotteryValues[i][0]));
    }
  }
  
  // 1から24までの未使用の番号を生成
  var availableNumbers = [];
  for (var i = 1; i <= 24; i++) {
    if (usedNumbers.indexOf(i) === -1) {
      availableNumbers.push(i);
    }
  }
  
  // 利用可能な番号をシャッフル
  for (var i = availableNumbers.length - 1; i > 0; i--) {
    var j = Math.floor(Math.random() * (i + 1));
    var temp = availableNumbers[i];
    availableNumbers[i] = availableNumbers[j];
    availableNumbers[j] = temp;
  }
  
  // シャッフルした番号を2次元配列に格納
  for (var i = 0; i < numRows; i++) {
    var row = [];
    for (var j = 0; j < numCols; j++) {
      if (i * numCols + j < availableNumbers.length) {
        row.push(availableNumbers[i * numCols + j]);
      } else {
        row.push(""); // 利用可能な番号がなければ空白を入れる
      }
    }
    data.push(row);
  }
  
  // スプレッドシートに値をセット
  var range = sheet.getRange(1, 1, numRows, numCols);
  range.setValues(data);
  
  // 数字が入っているセルのみ背景色と文字色を青に設定
  for (var i = 0; i < numRows; i++) {
    for (var j = 0; j < numCols; j++) {
      if (data[i][j] !== "") {
        sheet.getRange(i + 1, j + 1).setBackground("blue").setFontColor("blue");
      }
    }
  }
  
  // 確認メッセージは表示しない
}

function randomizeCells25_48() {
  // 抽選シートをアクティブにする
  var ss = SpreadsheetApp.getActive();
  var sheet = ss.getSheetByName("抽選");
  sheet.activate();
  
  var numRows = 4;
  var numCols = 6;
  var data = [];
  var rangeToFill0 = sheet.getRange("A1:H6");
  rangeToFill0.clearContent();
  rangeToFill0.setBackground("white");

  // 抽選順一覧シートからA列の値を取得
  var targetSheet = SpreadsheetApp.getActive().getSheetByName("抽選順一覧");
  var lotteryValues = targetSheet.getRange("A:A").getValues();
  
  // 抽選済みの番号を配列に格納
  var usedNumbers = [];
  for (var i = 0; i < lotteryValues.length; i++) {
    if (lotteryValues[i][0] !== "" && !isNaN(lotteryValues[i][0])) {
      usedNumbers.push(parseInt(lotteryValues[i][0]));
    }
  }
  
  // 25から48までの未使用の番号を生成
  var availableNumbers = [];
  for (var i = 25; i <= 48; i++) {
    if (usedNumbers.indexOf(i) === -1) {
      availableNumbers.push(i);
    }
  }
  
  // 利用可能な番号をシャッフル
  for (var i = availableNumbers.length - 1; i > 0; i--) {
    var j = Math.floor(Math.random() * (i + 1));
    var temp = availableNumbers[i];
    availableNumbers[i] = availableNumbers[j];
    availableNumbers[j] = temp;
  }
  
  // シャッフルした番号を2次元配列に格納
  for (var i = 0; i < numRows; i++) {
    var row = [];
    for (var j = 0; j < numCols; j++) {
      if (i * numCols + j < availableNumbers.length) {
        row.push(availableNumbers[i * numCols + j]);
      } else {
        row.push(""); // 利用可能な番号がなければ空白を入れる
      }
    }
    data.push(row);
  }
  
  // スプレッドシートに値をセット
  var range = sheet.getRange(1, 1, numRows, numCols);
  range.setValues(data);
  
  // 数字が入っているセルのみ背景色と文字色を青に設定
  for (var i = 0; i < numRows; i++) {
    for (var j = 0; j < numCols; j++) {
      if (data[i][j] !== "") {
        sheet.getRange(i + 1, j + 1).setBackground("blue").setFontColor("blue");
      }
    }
  }
  
  // 確認メッセージは表示しない
}

function randomizeZone1() {
  // 抽選シートをアクティブにする
  var ss = SpreadsheetApp.getActive();
  var sheet = ss.getSheetByName("抽選");
  sheet.activate();
  
  var numRows = 3;
  var numCols = 4;
  var data = [];
  var rangeToFill0 = sheet.getRange("A1:H6");
  rangeToFill0.clearContent();
  rangeToFill0.setBackground("white");

  // 抽選順一覧シートからA列の値を取得
  var targetSheet = SpreadsheetApp.getActive().getSheetByName("抽選順一覧");
  var lotteryValues = targetSheet.getRange("A:A").getValues();
  
  // 抽選済みの番号を配列に格納
  var usedNumbers = [];
  for (var i = 0; i < lotteryValues.length; i++) {
    if (lotteryValues[i][0] !== "" && !isNaN(lotteryValues[i][0])) {
      usedNumbers.push(parseInt(lotteryValues[i][0]));
    }
  }
  
  // 1から12までの未使用の番号を生成
  var availableNumbers = [];
  for (var i = 1; i <= 12; i++) {
    if (usedNumbers.indexOf(i) === -1) {
      availableNumbers.push(i);
    }
  }
  
  // 利用可能な番号をシャッフル
  for (var i = availableNumbers.length - 1; i > 0; i--) {
    var j = Math.floor(Math.random() * (i + 1));
    var temp = availableNumbers[i];
    availableNumbers[i] = availableNumbers[j];
    availableNumbers[j] = temp;
  }
  
  // シャッフルした番号を2次元配列に格納
  for (var i = 0; i < numRows; i++) {
    var row = [];
    for (var j = 0; j < numCols; j++) {
      if (i * numCols + j < availableNumbers.length) {
        row.push(availableNumbers[i * numCols + j]);
      } else {
        row.push(""); // 利用可能な番号がなければ空白を入れる
      }
    }
    data.push(row);
  }
  
  // スプレッドシートに値をセット
  var range = sheet.getRange(1, 1, numRows, numCols);
  range.setValues(data);
  
  // 数字が入っているセルのみ背景色と文字色を青に設定
  for (var i = 0; i < numRows; i++) {
    for (var j = 0; j < numCols; j++) {
      if (data[i][j] !== "") {
        sheet.getRange(i + 1, j + 1).setBackground("blue").setFontColor("blue");
      }
    }
  }
  
  // 確認メッセージは表示しない
}

function randomizeZone2() {
  // 抽選シートをアクティブにする
  var ss = SpreadsheetApp.getActive();
  var sheet = ss.getSheetByName("抽選");
  sheet.activate();
  
  var numRows = 3;
  var numCols = 4;
  var data = [];
  var rangeToFill0 = sheet.getRange("A1:H6");
  rangeToFill0.clearContent();
  rangeToFill0.setBackground("white");

  // 抽選順一覧シートからA列の値を取得
  var targetSheet = SpreadsheetApp.getActive().getSheetByName("抽選順一覧");
  var lotteryValues = targetSheet.getRange("A:A").getValues();
  
  // 抽選済みの番号を配列に格納
  var usedNumbers = [];
  for (var i = 0; i < lotteryValues.length; i++) {
    if (lotteryValues[i][0] !== "" && !isNaN(lotteryValues[i][0])) {
      usedNumbers.push(parseInt(lotteryValues[i][0]));
    }
  }
  
  // 13から24までの未使用の番号を生成
  var availableNumbers = [];
  for (var i = 13; i <= 24; i++) {
    if (usedNumbers.indexOf(i) === -1) {
      availableNumbers.push(i);
    }
  }
  
  // 利用可能な番号をシャッフル
  for (var i = availableNumbers.length - 1; i > 0; i--) {
    var j = Math.floor(Math.random() * (i + 1));
    var temp = availableNumbers[i];
    availableNumbers[i] = availableNumbers[j];
    availableNumbers[j] = temp;
  }
  
  // シャッフルした番号を2次元配列に格納
  for (var i = 0; i < numRows; i++) {
    var row = [];
    for (var j = 0; j < numCols; j++) {
      if (i * numCols + j < availableNumbers.length) {
        row.push(availableNumbers[i * numCols + j]);
      } else {
        row.push(""); // 利用可能な番号がなければ空白を入れる
      }
    }
    data.push(row);
  }
  
  // スプレッドシートに値をセット
  var range = sheet.getRange(1, 1, numRows, numCols);
  range.setValues(data);
  
  // 数字が入っているセルのみ背景色と文字色を青に設定
  for (var i = 0; i < numRows; i++) {
    for (var j = 0; j < numCols; j++) {
      if (data[i][j] !== "") {
        sheet.getRange(i + 1, j + 1).setBackground("blue").setFontColor("blue");
      }
    }
  }
  
  // 確認メッセージは表示しない
}

function randomizeZone3() {
  // 抽選シートをアクティブにする
  var ss = SpreadsheetApp.getActive();
  var sheet = ss.getSheetByName("抽選");
  sheet.activate();
  
  var numRows = 3;
  var numCols = 4;
  var data = [];
  var rangeToFill0 = sheet.getRange("A1:H6");
  rangeToFill0.clearContent();
  rangeToFill0.setBackground("white");

  // 抽選順一覧シートからA列の値を取得
  var targetSheet = SpreadsheetApp.getActive().getSheetByName("抽選順一覧");
  var lotteryValues = targetSheet.getRange("A:A").getValues();
  
  // 抽選済みの番号を配列に格納
  var usedNumbers = [];
  for (var i = 0; i < lotteryValues.length; i++) {
    if (lotteryValues[i][0] !== "" && !isNaN(lotteryValues[i][0])) {
      usedNumbers.push(parseInt(lotteryValues[i][0]));
    }
  }
  
  // 25から36までの未使用の番号を生成
  var availableNumbers = [];
  for (var i = 25; i <= 36; i++) {
    if (usedNumbers.indexOf(i) === -1) {
      availableNumbers.push(i);
    }
  }
  
  // 利用可能な番号をシャッフル
  for (var i = availableNumbers.length - 1; i > 0; i--) {
    var j = Math.floor(Math.random() * (i + 1));
    var temp = availableNumbers[i];
    availableNumbers[i] = availableNumbers[j];
    availableNumbers[j] = temp;
  }
  
  // シャッフルした番号を2次元配列に格納
  for (var i = 0; i < numRows; i++) {
    var row = [];
    for (var j = 0; j < numCols; j++) {
      if (i * numCols + j < availableNumbers.length) {
        row.push(availableNumbers[i * numCols + j]);
      } else {
        row.push(""); // 利用可能な番号がなければ空白を入れる
      }
    }
    data.push(row);
  }
  
  // スプレッドシートに値をセット
  var range = sheet.getRange(1, 1, numRows, numCols);
  range.setValues(data);
  
  // 数字が入っているセルのみ背景色と文字色を青に設定
  for (var i = 0; i < numRows; i++) {
    for (var j = 0; j < numCols; j++) {
      if (data[i][j] !== "") {
        sheet.getRange(i + 1, j + 1).setBackground("blue").setFontColor("blue");
      }
    }
  }
  
  // 確認メッセージは表示しない
}

function randomizeZone4() {
  // 抽選シートをアクティブにする
  var ss = SpreadsheetApp.getActive();
  var sheet = ss.getSheetByName("抽選");
  sheet.activate();
  
  var numRows = 3;
  var numCols = 4;
  var data = [];
  var rangeToFill0 = sheet.getRange("A1:H6");
  rangeToFill0.clearContent();
  rangeToFill0.setBackground("white");

  // 抽選順一覧シートからA列の値を取得
  var targetSheet = SpreadsheetApp.getActive().getSheetByName("抽選順一覧");
  var lotteryValues = targetSheet.getRange("A:A").getValues();
  
  // 抽選済みの番号を配列に格納
  var usedNumbers = [];
  for (var i = 0; i < lotteryValues.length; i++) {
    if (lotteryValues[i][0] !== "" && !isNaN(lotteryValues[i][0])) {
      usedNumbers.push(parseInt(lotteryValues[i][0]));
    }
  }
  
  // 37から48までの未使用の番号を生成
  var availableNumbers = [];
  for (var i = 37; i <= 48; i++) {
    if (usedNumbers.indexOf(i) === -1) {
      availableNumbers.push(i);
    }
  }
  
  // 利用可能な番号をシャッフル
  for (var i = availableNumbers.length - 1; i > 0; i--) {
    var j = Math.floor(Math.random() * (i + 1));
    var temp = availableNumbers[i];
    availableNumbers[i] = availableNumbers[j];
    availableNumbers[j] = temp;
  }
  
  // シャッフルした番号を2次元配列に格納
  for (var i = 0; i < numRows; i++) {
    var row = [];
    for (var j = 0; j < numCols; j++) {
      if (i * numCols + j < availableNumbers.length) {
        row.push(availableNumbers[i * numCols + j]);
      } else {
        row.push(""); // 利用可能な番号がなければ空白を入れる
      }
    }
    data.push(row);
  }
  
  // スプレッドシートに値をセット
  var range = sheet.getRange(1, 1, numRows, numCols);
  range.setValues(data);
  
  // 数字が入っているセルのみ背景色と文字色を青に設定
  for (var i = 0; i < numRows; i++) {
    for (var j = 0; j < numCols; j++) {
      if (data[i][j] !== "") {
        sheet.getRange(i + 1, j + 1).setBackground("blue").setFontColor("blue");
      }
    }
  }
  
  // 確認メッセージは表示しない
}

function zrandomizeCells() {
  // 抽選シートをアクティブにする
  var ss = SpreadsheetApp.getActive();
  var sheet = ss.getSheetByName("抽選");
  sheet.activate();
  
  var numRows = 6;
  var numCols = 8;
  var data = [];
  var rangeToFill = sheet.getRange("A1:H6");
  rangeToFill.setBackground("red");
  rangeToFill.setFontColor("red");

  // 1から48までの乱数を生成して、配列に追加する
  var nums = [];
  for (var i = 1; i <= 48; i++) {
    nums.push(i);
  }
  for (var i = nums.length - 1; i > 0; i--) {
    var j = Math.floor(Math.random() * (i + 1));
    var temp = nums[i];
    nums[i] = nums[j];
    nums[j] = temp;
  }

  for (var i = 0; i < numRows; i++) {
    var row = [];
    for (var j = 0; j < numCols; j++) {
      row.push(nums[i * numCols + j]);
    }
    data.push(row);
  }

  var range = sheet.getRange(1, 1, numRows, numCols);
  range.setValues(data);
}

function clearCellValueAndBackgroundColor() {
  var sheet = SpreadsheetApp.getActiveSheet();
  var range = sheet.getActiveRange(); // アクティブなセルの範囲を取得
  var value = range.getValue(); // セルの値を取得

  range.setBackground("white"); // セルの背景色を白にする
  return value;
}

function writeValueToLotteryList() {
  var sheet = SpreadsheetApp.getActiveSheet();
  var range = sheet.getActiveRange(); // アクティブなセルの範囲を取得
  var value = range.getValue(); // セルの値を取得

  range.setBackground("white"); // セルの背景色を白にする

  var targetSheet = SpreadsheetApp.getActive().getSheetByName("抽選順一覧"); // ターゲットシートを取得
  var lastRow = targetSheet.getRange("A:A").getValues().filter(String).length; // A列にデータが入っている最終行を取得
  targetSheet.getRange("A" + (lastRow + 1)).setValue(value); // 最終行の次の行に値を書き込む
//  Logger.log("Value " + value + " written to cell A" + (lastRow + 1) + " in sheet '抽選順一覧'.");
}

/**
 * 抽選順一覧シートのA列に含まれない数字を抽選シートのA1からH6の範囲内に
 * 1から48の重複しないランダムな数字を埋め込む関数
 */
function randomizeRemainingNumbers() {
  // 抽選シートをアクティブにする
  var ss = SpreadsheetApp.getActive();
  var sheet = ss.getSheetByName("抽選");
  sheet.activate();
  
  var numRows = 6;
  var numCols = 8;
  var data = [];
  var rangeToFill = sheet.getRange("A1:H6");
  rangeToFill.clearContent();
  // 全てのセルの背景色をリセット
  rangeToFill.setBackground("white");
  
  // 抽選順一覧シートからA列の値を取得
  var targetSheet = SpreadsheetApp.getActive().getSheetByName("抽選順一覧");
  var lotteryValues = targetSheet.getRange("A:A").getValues();
  
  // 抽選済みの番号を配列に格納
  var usedNumbers = [];
  for (var i = 0; i < lotteryValues.length; i++) {
    if (lotteryValues[i][0] !== "" && !isNaN(lotteryValues[i][0])) {
      usedNumbers.push(parseInt(lotteryValues[i][0]));
    }
  }
  
  // 1から48までの未使用の番号を生成
  var availableNumbers = [];
  for (var i = 1; i <= 48; i++) {
    if (usedNumbers.indexOf(i) === -1) {
      availableNumbers.push(i);
    }
  }
  
  // 利用可能な番号をシャッフル
  for (var i = availableNumbers.length - 1; i > 0; i--) {
    var j = Math.floor(Math.random() * (i + 1));
    var temp = availableNumbers[i];
    availableNumbers[i] = availableNumbers[j];
    availableNumbers[j] = temp;
  }
  
  // シャッフルした番号を2次元配列に格納
  for (var i = 0; i < numRows; i++) {
    var row = [];
    for (var j = 0; j < numCols; j++) {
      if (i * numCols + j < availableNumbers.length) {
        row.push(availableNumbers[i * numCols + j]);
      } else {
        row.push(""); // 利用可能な番号がなければ空白を入れる
      }
    }
    data.push(row);
  }
  
  // スプレッドシートに値をセット
  var range = sheet.getRange(1, 1, numRows, numCols);
  range.setValues(data);
  
  // 数字が入っているセルのみ背景色と文字色を青に設定
  for (var i = 0; i < numRows; i++) {
    for (var j = 0; j < numCols; j++) {
      if (data[i][j] !== "") {
        sheet.getRange(i + 1, j + 1).setBackground("blue").setFontColor("blue");
      }
    }
  }
  
  // 確認メッセージは表示しない
}

/**
 * Aリーグ(1-3)の抽選を行う関数
 */
function randomizeALeague() {
  // 抽選シートをアクティブにする
  var ss = SpreadsheetApp.getActive();
  var sheet = ss.getSheetByName("抽選");
  sheet.activate();
  
  var numRows = 1;
  var numCols = 3;
  var data = [];
  var rangeToFill0 = sheet.getRange("A1:H6");
  rangeToFill0.clearContent();
  rangeToFill0.setBackground("white");

  // 抽選順一覧シートからA列の値を取得
  var targetSheet = SpreadsheetApp.getActive().getSheetByName("抽選順一覧");
  var lotteryValues = targetSheet.getRange("A:A").getValues();
  
  // 抽選済みの番号を配列に格納
  var usedNumbers = [];
  for (var i = 0; i < lotteryValues.length; i++) {
    if (lotteryValues[i][0] !== "" && !isNaN(lotteryValues[i][0])) {
      usedNumbers.push(parseInt(lotteryValues[i][0]));
    }
  }
  
  // 1から3までの未使用の番号を生成
  var availableNumbers = [];
  for (var i = 1; i <= 3; i++) {
    if (usedNumbers.indexOf(i) === -1) {
      availableNumbers.push(i);
    }
  }
  
  // 利用可能な番号をシャッフル
  for (var i = availableNumbers.length - 1; i > 0; i--) {
    var j = Math.floor(Math.random() * (i + 1));
    var temp = availableNumbers[i];
    availableNumbers[i] = availableNumbers[j];
    availableNumbers[j] = temp;
  }
  
  // シャッフルした番号を2次元配列に格納
  var row = [];
  for (var j = 0; j < numCols; j++) {
    if (j < availableNumbers.length) {
      row.push(availableNumbers[j]);
    } else {
      row.push(""); // 利用可能な番号がなければ空白を入れる
    }
  }
  data.push(row);
  
  // スプレッドシートに値をセット
  var range = sheet.getRange(1, 1, numRows, numCols);
  range.setValues(data);
  
  // 数字が入っているセルのみ背景色と文字色を青に設定
  for (var j = 0; j < numCols; j++) {
    if (data[0][j] !== "") {
      sheet.getRange(1, j + 1).setBackground("blue").setFontColor("blue");
    }
  }
}

/**
 * Pリーグ(46-48)の抽選を行う関数
 */
function randomizePLeague() {
  // 抽選シートをアクティブにする
  var ss = SpreadsheetApp.getActive();
  var sheet = ss.getSheetByName("抽選");
  sheet.activate();
  
  var numRows = 1;
  var numCols = 3;
  var data = [];
  var rangeToFill0 = sheet.getRange("A1:H6");
  rangeToFill0.clearContent();
  rangeToFill0.setBackground("white");

  // 抽選順一覧シートからA列の値を取得
  var targetSheet = SpreadsheetApp.getActive().getSheetByName("抽選順一覧");
  var lotteryValues = targetSheet.getRange("A:A").getValues();
  
  // 抽選済みの番号を配列に格納
  var usedNumbers = [];
  for (var i = 0; i < lotteryValues.length; i++) {
    if (lotteryValues[i][0] !== "" && !isNaN(lotteryValues[i][0])) {
      usedNumbers.push(parseInt(lotteryValues[i][0]));
    }
  }
  
  // 46から48までの未使用の番号を生成
  var availableNumbers = [];
  for (var i = 46; i <= 48; i++) {
    if (usedNumbers.indexOf(i) === -1) {
      availableNumbers.push(i);
    }
  }
  
  // 利用可能な番号をシャッフル
  for (var i = availableNumbers.length - 1; i > 0; i--) {
    var j = Math.floor(Math.random() * (i + 1));
    var temp = availableNumbers[i];
    availableNumbers[i] = availableNumbers[j];
    availableNumbers[j] = temp;
  }
  
  // シャッフルした番号を2次元配列に格納
  var row = [];
  for (var j = 0; j < numCols; j++) {
    if (j < availableNumbers.length) {
      row.push(availableNumbers[j]);
    } else {
      row.push(""); // 利用可能な番号がなければ空白を入れる
    }
  }
  data.push(row);
  
  // スプレッドシートに値をセット
  var range = sheet.getRange(1, 1, numRows, numCols);
  range.setValues(data);
  
  // 数字が入っているセルのみ背景色と文字色を青に設定
  for (var j = 0; j < numCols; j++) {
    if (data[0][j] !== "") {
      sheet.getRange(1, j + 1).setBackground("blue").setFontColor("blue");
    }
  }
}

/**
 * Hリーグ(22-24)の抽選を行う関数
 */
function randomizeHLeague() {
  // 抽選シートをアクティブにする
  var ss = SpreadsheetApp.getActive();
  var sheet = ss.getSheetByName("抽選");
  sheet.activate();
  
  var numRows = 1;
  var numCols = 3;
  var data = [];
  var rangeToFill0 = sheet.getRange("A1:H6");
  rangeToFill0.clearContent();
  rangeToFill0.setBackground("white");

  // 抽選順一覧シートからA列の値を取得
  var targetSheet = SpreadsheetApp.getActive().getSheetByName("抽選順一覧");
  var lotteryValues = targetSheet.getRange("A:A").getValues();
  
  // 抽選済みの番号を配列に格納
  var usedNumbers = [];
  for (var i = 0; i < lotteryValues.length; i++) {
    if (lotteryValues[i][0] !== "" && !isNaN(lotteryValues[i][0])) {
      usedNumbers.push(parseInt(lotteryValues[i][0]));
    }
  }
  
  // 22から24までの未使用の番号を生成
  var availableNumbers = [];
  for (var i = 22; i <= 24; i++) {
    if (usedNumbers.indexOf(i) === -1) {
      availableNumbers.push(i);
    }
  }
  
  // 利用可能な番号をシャッフル
  for (var i = availableNumbers.length - 1; i > 0; i--) {
    var j = Math.floor(Math.random() * (i + 1));
    var temp = availableNumbers[i];
    availableNumbers[i] = availableNumbers[j];
    availableNumbers[j] = temp;
  }
  
  // シャッフルした番号を2次元配列に格納
  var row = [];
  for (var j = 0; j < numCols; j++) {
    if (j < availableNumbers.length) {
      row.push(availableNumbers[j]);
    } else {
      row.push(""); // 利用可能な番号がなければ空白を入れる
    }
  }
  data.push(row);
  
  // スプレッドシートに値をセット
  var range = sheet.getRange(1, 1, numRows, numCols);
  range.setValues(data);
  
  // 数字が入っているセルのみ背景色と文字色を青に設定
  for (var j = 0; j < numCols; j++) {
    if (data[0][j] !== "") {
      sheet.getRange(1, j + 1).setBackground("blue").setFontColor("blue");
    }
  }
}

/**
 * Iリーグ(25-27)の抽選を行う関数
 */
function randomizeILeague() {
  // 抽選シートをアクティブにする
  var ss = SpreadsheetApp.getActive();
  var sheet = ss.getSheetByName("抽選");
  sheet.activate();
  
  var numRows = 1;
  var numCols = 3;
  var data = [];
  var rangeToFill0 = sheet.getRange("A1:H6");
  rangeToFill0.clearContent();
  rangeToFill0.setBackground("white");

  // 抽選順一覧シートからA列の値を取得
  var targetSheet = SpreadsheetApp.getActive().getSheetByName("抽選順一覧");
  var lotteryValues = targetSheet.getRange("A:A").getValues();
  
  // 抽選済みの番号を配列に格納
  var usedNumbers = [];
  for (var i = 0; i < lotteryValues.length; i++) {
    if (lotteryValues[i][0] !== "" && !isNaN(lotteryValues[i][0])) {
      usedNumbers.push(parseInt(lotteryValues[i][0]));
    }
  }
  
  // 25から27までの未使用の番号を生成
  var availableNumbers = [];
  for (var i = 25; i <= 27; i++) {
    if (usedNumbers.indexOf(i) === -1) {
      availableNumbers.push(i);
    }
  }
  
  // 利用可能な番号をシャッフル
  for (var i = availableNumbers.length - 1; i > 0; i--) {
    var j = Math.floor(Math.random() * (i + 1));
    var temp = availableNumbers[i];
    availableNumbers[i] = availableNumbers[j];
    availableNumbers[j] = temp;
  }
  
  // シャッフルした番号を2次元配列に格納
  var row = [];
  for (var j = 0; j < numCols; j++) {
    if (j < availableNumbers.length) {
      row.push(availableNumbers[j]);
    } else {
      row.push(""); // 利用可能な番号がなければ空白を入れる
    }
  }
  data.push(row);
  
  // スプレッドシートに値をセット
  var range = sheet.getRange(1, 1, numRows, numCols);
  range.setValues(data);
  
  // 数字が入っているセルのみ背景色と文字色を青に設定
  for (var j = 0; j < numCols; j++) {
    if (data[0][j] !== "") {
      sheet.getRange(1, j + 1).setBackground("blue").setFontColor("blue");
    }
  }
}

/**
 * H&Iリーグ(22-27)の抽選を行う関数
 */
function randomizeHILeague() {
  // 抽選シートをアクティブにする
  var ss = SpreadsheetApp.getActive();
  var sheet = ss.getSheetByName("抽選");
  sheet.activate();
  
  var numRows = 2;
  var numCols = 3;
  var data = [];
  var rangeToFill0 = sheet.getRange("A1:H6");
  rangeToFill0.clearContent();
  rangeToFill0.setBackground("white");

  // 抽選順一覧シートからA列の値を取得
  var targetSheet = SpreadsheetApp.getActive().getSheetByName("抽選順一覧");
  var lotteryValues = targetSheet.getRange("A:A").getValues();
  
  // 抽選済みの番号を配列に格納
  var usedNumbers = [];
  for (var i = 0; i < lotteryValues.length; i++) {
    if (lotteryValues[i][0] !== "" && !isNaN(lotteryValues[i][0])) {
      usedNumbers.push(parseInt(lotteryValues[i][0]));
    }
  }
  
  // 22から27までの未使用の番号を生成
  var availableNumbers = [];
  for (var i = 22; i <= 27; i++) {
    if (usedNumbers.indexOf(i) === -1) {
      availableNumbers.push(i);
    }
  }
  
  // 利用可能な番号をシャッフル
  for (var i = availableNumbers.length - 1; i > 0; i--) {
    var j = Math.floor(Math.random() * (i + 1));
    var temp = availableNumbers[i];
    availableNumbers[i] = availableNumbers[j];
    availableNumbers[j] = temp;
  }
  
  // シャッフルした番号を2次元配列に格納
  for (var i = 0; i < numRows; i++) {
    var row = [];
    for (var j = 0; j < numCols; j++) {
      if (i * numCols + j < availableNumbers.length) {
        row.push(availableNumbers[i * numCols + j]);
      } else {
        row.push(""); // 利用可能な番号がなければ空白を入れる
      }
    }
    data.push(row);
  }
  
  // スプレッドシートに値をセット
  var range = sheet.getRange(1, 1, numRows, numCols);
  range.setValues(data);
  
  // 数字が入っているセルのみ背景色と文字色を青に設定
  for (var i = 0; i < numRows; i++) {
    for (var j = 0; j < numCols; j++) {
      if (data[i][j] !== "") {
        sheet.getRange(i + 1, j + 1).setBackground("blue").setFontColor("blue");
      }
    }
  }
}
