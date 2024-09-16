//B1にスケジュールの開始日
//B２にスケジュールの終了日
//A列からD列に掛けて３〜５行目を結合し、タスク名、開始日、終了日、担当者を記載する
//6行目以降から記載をしていく
function onEdit(e) {
    if (!e || !e.range) {
      Logger.log('イベントオブジェクト e が未定義です。');
      return;
    }
  
    var sheet = e.range.getSheet();  // 変更が行われたシートを取得
    
    if (sheet.getName() == 'スケジュール') {
      var editedRow = e.range.getRow();  // 編集された行番号を取得
      var editedColumn = e.range.getColumn();  // 編集された列番号を取得
      Logger.log('編集された行: ' + editedRow);

      if (editedRow == 1 && editedColumn == 2 || editedRow == 2 && editedColumn == 2) {
        resetCells();
        updateSummaryRows();
        var lastRow = sheet.getMaxRows();
        for(var row_num = 6; row_num <= lastRow; row_num++){
          updateRowForGanttChart(row_num); 
        }
      }
      
      if (editedRow >= 6) {  // 6行目以降がタスク情報のエリア
        Logger.log('ガントチャートの更新を開始します。');
        updateRowForGanttChart(editedRow);  // 編集された行のみ処理を実行
      }
    } else {
      Logger.log('スケジュールシートではありません。');
    }
  }

function resetCells() {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('スケジュール');
    var lastRow = sheet.getMaxRows();  // 最終行を取得
    var lastColumn = sheet.getMaxColumns();  // 最終列を取得
  
    // 6行目以降、E列（5列目）から最後の列までの範囲を取得
    var range = sheet.getRange(1, 5, lastRow, lastColumn - 4);
  
    // セルの結合を解除
    range.breakApart();
    range.clearContent();
    
    // 背景色を白にリセット
    range.setBackground('#FFFFFF');

    
}

function updateRowForGanttChart(editedRow) {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('スケジュール');

    // 開始日と終了日の取得
    var startDate = new Date(sheet.getRange('B1').getValue());
    var endDate = new Date(sheet.getRange('B2').getValue());

    // 日付リストの作成
    var dateList = [];
    for (var d = new Date(startDate); d <= endDate; d.setDate(d.getDate() + 1)) {
        dateList.push(new Date(d));
    }

    // 開始日と終了日の範囲に基づいて祝日リストを取得
    var holidays = getHolidaysInRange(startDate, endDate);

    // タスク情報を取得
    var task = sheet.getRange(editedRow, 1, 1, 4).getValues()[0];  // 編集された行のタスク情報を取得
    if (task[0] === '') return;  // タスク名が空の場合は処理しない

    var taskStart = new Date(task[1]);
    var taskEnd = new Date(task[2]);
    var assignee = task[3];

    // 担当者ごとの色設定
    var taskColor = '#0000FF';  // デフォルトは青
    if (assignee == '自分の名前を記載') {
        taskColor = '#008000';  // 緑色
    } else if (assignee == 'リリース') {
        taskColor = '#E06666';  // 赤色
    }

    // 背景色を編集するための範囲を取得
    var range = sheet.getRange(editedRow, 5, 1, dateList.length);
    var backgrounds = range.getBackgrounds()[0];

    // 背景色をリセットするが、土日と祝日にはリセットしない
    for (var col = 0; col < dateList.length; col++) {
        var date = dateList[col];
        if (isWeekend(date) || isHoliday(date, holidays)) {
        continue;  // 土日や祝日はリセットしない
        }
        backgrounds[col] = '#FFFFFF';  // 土日・祝日以外は白にリセット
        if (editedRow % 2 == 0) {
        backgrounds[col] = '#F0F0F0';  // 白にリセット
        }
    }

    // 日付リストに基づき、タスク期間に色を適用（土日・祝日は除外）
    for (var col = 0; col < dateList.length; col++) {
        var date = dateList[col];
        
        // 土日・祝日は色を変更しない（既に除外されているためスキップ）
        if (isWeekend(date) || isHoliday(date, holidays)) {
        continue;
        }
        
        if (date >= taskStart && date <= taskEnd) {
        backgrounds[col] = taskColor;
        }
    }

    // 最終的な色をスプレッドシートに反映
    range.setBackgrounds([backgrounds]);
}

// 土日かどうかを判定する関数
function isWeekend(date) {
    var day = date.getDay();
    return day === 0 || day === 6;  // 0: 日曜日, 6: 土曜日
}

// 祝日リストを取得する関数
function getHolidaysInRange(startDate, endDate) {
// 日本の祝日を取得（Google Apps Scriptのカレンダーサービスを使って範囲指定で取得）
    var calendar = CalendarApp.getCalendarById('ja.japanese#holiday@group.v.calendar.google.com');
    var holidays = calendar.getEvents(startDate, endDate);  // 指定範囲の祝日を取得
    var holidayDates = holidays.map(function(event) {
        return event.getStartTime();
    });
    return holidayDates;
}

// 祝日かどうかを判定する関数
function isHoliday(date, holidays) {
for (var i = 0; i < holidays.length; i++) {
    if (date.getFullYear() === holidays[i].getFullYear() &&
        date.getMonth() === holidays[i].getMonth() &&
        date.getDate() === holidays[i].getDate()) {
    return true;
    }
}
return false;
}

function updateSummaryRows() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('スケジュール');
  var calendarId = 'ja.japanese#holiday@group.v.calendar.google.com'; 
  var startDate = new Date(sheet.getRange('B1').getValue());
  var endDate = new Date(sheet.getRange('B2').getValue());
  
  var dateList = [];
  for (var d = new Date(startDate); d <= endDate; d.setDate(d.getDate() + 1)) {
    dateList.push(new Date(d));
  }
  
  var monthColors = ['#8b4513', '#556b2f', '#483d8b', '#006400', '#8b0000', '#8b008b'];
  var currentMonth = '';
  var monthColorIndex = 0;
  var mergeStartColumn = 5;
  
  for (var i = 0; i < dateList.length; i++) {
    var date = dateList[i];
    var month = Utilities.formatDate(date, 'JST', 'MM');

    sheet.getRange(4, i + 5).setValue(Utilities.formatDate(date, 'JST', 'dd'));
    sheet.getRange(5, i + 5).setValue(['日', '月', '火', '水', '木', '金', '土'][date.getDay()]);
    
    if (month != currentMonth) {
      if (i > 0) {
        var endColumn = i + 5 - 1;
        var numColumns = endColumn - mergeStartColumn + 1;
        sheet.getRange(3, mergeStartColumn, 1, numColumns).merge()
          .setBackground(monthColors[monthColorIndex % monthColors.length])
          .setFontColor('#FFFFFF')
          .setHorizontalAlignment('center')
          .setValue(currentMonth + '月');
        
        monthColorIndex++;
      }
      currentMonth = month;
      mergeStartColumn = i + 5;
    }
  }
  
  var endColumn = dateList.length + 5 - 1;
  var numColumns = endColumn - mergeStartColumn + 1;
  sheet.getRange(3, mergeStartColumn, 1, numColumns).merge()
    .setBackground(monthColors[monthColorIndex % monthColors.length])
    .setFontColor('#FFFFFF')
    .setHorizontalAlignment('center')
    .setValue(currentMonth + '月');
  
  var lastRow = sheet.getMaxRows();
  var range = sheet.getRange(6, 5, lastRow - 5, dateList.length);
  var values = range.getValues();
  var backgrounds = range.getBackgrounds();
  
  var holidays = CalendarApp.getCalendarById(calendarId).getEvents(startDate, endDate);
  var holidayDates = holidays.map(function(event) {
    return Utilities.formatDate(event.getAllDayStartDate(), 'JST', 'yyyyMMdd');
  });
  
  for (var row = 0; row < values.length; row++) {
    for (var col = 0; col < dateList.length; col++) {
      var date = dateList[col];
      var dateFormatted = Utilities.formatDate(date, 'JST', 'yyyyMMdd');
      var isWeekend = date.getDay() == 6 || date.getDay() == 0;
      var isHoliday = holidayDates.indexOf(dateFormatted) !== -1;

      if (isWeekend || isHoliday) {
        backgrounds[row][col] = (date.getDay() == 6) ? '#ADD8E6' : '#F4CCCC';
      } else {
        backgrounds[row][col] = (row % 2 == 0) ? '#F0F0F0' : '#FFFFFF';
      }
    }
  }

  range.setBackgrounds(backgrounds);
}

// スプレッドシートが開かれたときに実行する
function onOpen() {
  updateSummarytodayRows();  // ロード時に行と列を更新する関数を呼び出す
}

function updateSummarytodayRows() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('スケジュール');
  var startDate = new Date(sheet.getRange('B1').getValue());  // 開始日
  var endDate = new Date(sheet.getRange('B2').getValue());    // 終了日
  var today = new Date();  // 本日の日付
  
  // 日付リストの作成
  var dateList = [];
  for (var d = new Date(startDate); d <= endDate; d.setDate(d.getDate() + 1)) {
    dateList.push(new Date(d));
  }
  
  // 日付と曜日の設定、および当日の背景色設定
  for (var i = 0; i < dateList.length; i++) {
    var date = dateList[i];

    // 4行目に日付、5行目に曜日を設定
    sheet.getRange(4, i + 5).setValue(Utilities.formatDate(date, 'JST', 'dd'));
    sheet.getRange(5, i + 5).setValue(['日', '月', '火', '水', '木', '金', '土'][date.getDay()]);
    
    // 当日判定：今日の日付と一致する場合に色を設定
    if (Utilities.formatDate(today, 'JST', 'yyyyMMdd') === Utilities.formatDate(date, 'JST', 'yyyyMMdd')) {
      sheet.getRange(4, i + 5).setBackground('#FFFF00');  // 4行目の当日セルに黄色
      sheet.getRange(5, i + 5).setBackground('#FFFF00');  // 5行目の当日セルに黄色
    } else {
      // 当日でない場合は背景色をリセット
      sheet.getRange(4, i + 5).setBackground('#FFFFFF');
      sheet.getRange(5, i + 5).setBackground('#FFFFFF');
    }
  }
}

function onEdit(e) {
    if (!e || !e.range) {
      Logger.log('イベントオブジェクト e が未定義です。');
      return;
    }
  
    var sheet = e.range.getSheet();  // 変更が行われたシートを取得
    
    if (sheet.getName() == 'スケジュール') {
      var editedRow = e.range.getRow();  // 編集された行番号を取得
      var editedColumn = e.range.getColumn();  // 編集された列番号を取得
      Logger.log('編集された行: ' + editedRow);

      if (editedRow == 1 && editedColumn == 2 || editedRow == 2 && editedColumn == 2) {
        resetCells();
        updateSummaryRows();
        var lastRow = sheet.getMaxRows();
        for(var row_num = 6; row_num <= lastRow; row_num++){
          updateRowForGanttChart(row_num); 
        }
      }
      
      if (editedRow >= 6) {  // 6行目以降がタスク情報のエリア
        Logger.log('ガントチャートの更新を開始します。');
        updateRowForGanttChart(editedRow);  // 編集された行のみ処理を実行
      }
    } else {
      Logger.log('スケジュールシートではありません。');
    }
  }

function resetCells() {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('スケジュール');
    var lastRow = sheet.getMaxRows();  // 最終行を取得
    var lastColumn = sheet.getMaxColumns();  // 最終列を取得
  
    // 6行目以降、E列（5列目）から最後の列までの範囲を取得
    var range = sheet.getRange(1, 5, lastRow, lastColumn - 4);
  
    // セルの結合を解除
    range.breakApart();
    range.clearContent();
    
    // 背景色を白にリセット
    range.setBackground('#FFFFFF');

    
}

function updateRowForGanttChart(editedRow) {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('スケジュール');

    // 開始日と終了日の取得
    var startDate = new Date(sheet.getRange('B1').getValue());
    var endDate = new Date(sheet.getRange('B2').getValue());

    // 日付リストの作成
    var dateList = [];
    for (var d = new Date(startDate); d <= endDate; d.setDate(d.getDate() + 1)) {
        dateList.push(new Date(d));
    }

    // 開始日と終了日の範囲に基づいて祝日リストを取得
    var holidays = getHolidaysInRange(startDate, endDate);

    // タスク情報を取得
    var task = sheet.getRange(editedRow, 1, 1, 4).getValues()[0];  // 編集された行のタスク情報を取得
    if (task[0] === '') return;  // タスク名が空の場合は処理しない

    var taskStart = new Date(task[1]);
    var taskEnd = new Date(task[2]);
    var assignee = task[3];

    // 担当者ごとの色設定
    var taskColor = '#0000FF';  // デフォルトは青
    if (assignee == '自分の名前を記載') {
        taskColor = '#008000';  // 緑色
    } else if (assignee == 'リリース') {
        taskColor = '#E06666';  // 赤色
    }

    // 背景色を編集するための範囲を取得
    var range = sheet.getRange(editedRow, 5, 1, dateList.length);
    var backgrounds = range.getBackgrounds()[0];

    // 背景色をリセットするが、土日と祝日にはリセットしない
    for (var col = 0; col < dateList.length; col++) {
        var date = dateList[col];
        if (isWeekend(date) || isHoliday(date, holidays)) {
        continue;  // 土日や祝日はリセットしない
        }
        backgrounds[col] = '#FFFFFF';  // 土日・祝日以外は白にリセット
        if (editedRow % 2 == 0) {
        backgrounds[col] = '#F0F0F0';  // 白にリセット
        }
    }

    // 日付リストに基づき、タスク期間に色を適用（土日・祝日は除外）
    for (var col = 0; col < dateList.length; col++) {
        var date = dateList[col];
        
        // 土日・祝日は色を変更しない（既に除外されているためスキップ）
        if (isWeekend(date) || isHoliday(date, holidays)) {
        continue;
        }
        
        if (date >= taskStart && date <= taskEnd) {
        backgrounds[col] = taskColor;
        }
    }

    // 最終的な色をスプレッドシートに反映
    range.setBackgrounds([backgrounds]);
}

// 土日かどうかを判定する関数
function isWeekend(date) {
    var day = date.getDay();
    return day === 0 || day === 6;  // 0: 日曜日, 6: 土曜日
}

// 祝日リストを取得する関数
function getHolidaysInRange(startDate, endDate) {
// 日本の祝日を取得（Google Apps Scriptのカレンダーサービスを使って範囲指定で取得）
    var calendar = CalendarApp.getCalendarById('ja.japanese#holiday@group.v.calendar.google.com');
    var holidays = calendar.getEvents(startDate, endDate);  // 指定範囲の祝日を取得
    var holidayDates = holidays.map(function(event) {
        return event.getStartTime();
    });
    return holidayDates;
}

// 祝日かどうかを判定する関数
function isHoliday(date, holidays) {
for (var i = 0; i < holidays.length; i++) {
    if (date.getFullYear() === holidays[i].getFullYear() &&
        date.getMonth() === holidays[i].getMonth() &&
        date.getDate() === holidays[i].getDate()) {
    return true;
    }
}
return false;
}

function updateSummaryRows() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('スケジュール');
  var calendarId = 'ja.japanese#holiday@group.v.calendar.google.com'; 
  var startDate = new Date(sheet.getRange('B1').getValue());
  var endDate = new Date(sheet.getRange('B2').getValue());
  
  var dateList = [];
  for (var d = new Date(startDate); d <= endDate; d.setDate(d.getDate() + 1)) {
    dateList.push(new Date(d));
  }
  
  var monthColors = ['#8b4513', '#556b2f', '#483d8b', '#006400', '#8b0000', '#8b008b'];
  var currentMonth = '';
  var monthColorIndex = 0;
  var mergeStartColumn = 5;
  
  for (var i = 0; i < dateList.length; i++) {
    var date = dateList[i];
    var month = Utilities.formatDate(date, 'JST', 'MM');

    sheet.getRange(4, i + 5).setValue(Utilities.formatDate(date, 'JST', 'dd'));
    sheet.getRange(5, i + 5).setValue(['日', '月', '火', '水', '木', '金', '土'][date.getDay()]);
    
    if (month != currentMonth) {
      if (i > 0) {
        var endColumn = i + 5 - 1;
        var numColumns = endColumn - mergeStartColumn + 1;
        sheet.getRange(3, mergeStartColumn, 1, numColumns).merge()
          .setBackground(monthColors[monthColorIndex % monthColors.length])
          .setFontColor('#FFFFFF')
          .setHorizontalAlignment('center')
          .setValue(currentMonth + '月');
        
        monthColorIndex++;
      }
      currentMonth = month;
      mergeStartColumn = i + 5;
    }
  }
  
  var endColumn = dateList.length + 5 - 1;
  var numColumns = endColumn - mergeStartColumn + 1;
  sheet.getRange(3, mergeStartColumn, 1, numColumns).merge()
    .setBackground(monthColors[monthColorIndex % monthColors.length])
    .setFontColor('#FFFFFF')
    .setHorizontalAlignment('center')
    .setValue(currentMonth + '月');
  
  var lastRow = sheet.getMaxRows();
  var range = sheet.getRange(6, 5, lastRow - 5, dateList.length);
  var values = range.getValues();
  var backgrounds = range.getBackgrounds();
  
  var holidays = CalendarApp.getCalendarById(calendarId).getEvents(startDate, endDate);
  var holidayDates = holidays.map(function(event) {
    return Utilities.formatDate(event.getAllDayStartDate(), 'JST', 'yyyyMMdd');
  });
  
  for (var row = 0; row < values.length; row++) {
    for (var col = 0; col < dateList.length; col++) {
      var date = dateList[col];
      var dateFormatted = Utilities.formatDate(date, 'JST', 'yyyyMMdd');
      var isWeekend = date.getDay() == 6 || date.getDay() == 0;
      var isHoliday = holidayDates.indexOf(dateFormatted) !== -1;

      if (isWeekend || isHoliday) {
        backgrounds[row][col] = (date.getDay() == 6) ? '#ADD8E6' : '#F4CCCC';
      } else {
        backgrounds[row][col] = (row % 2 == 0) ? '#F0F0F0' : '#FFFFFF';
      }
    }
  }

  range.setBackgrounds(backgrounds);
}

// スプレッドシートが開かれたときに実行する
function onOpen() {
  updateSummarytodayRows();  // ロード時に行と列を更新する関数を呼び出す
}

function updateSummarytodayRows() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('スケジュール');
  var startDate = new Date(sheet.getRange('B1').getValue());  // 開始日
  var endDate = new Date(sheet.getRange('B2').getValue());    // 終了日
  var today = new Date();  // 本日の日付
  
  // 日付リストの作成
  var dateList = [];
  for (var d = new Date(startDate); d <= endDate; d.setDate(d.getDate() + 1)) {
    dateList.push(new Date(d));
  }
  
  // 日付と曜日の設定、および当日の背景色設定
  for (var i = 0; i < dateList.length; i++) {
    var date = dateList[i];

    // 4行目に日付、5行目に曜日を設定
    sheet.getRange(4, i + 5).setValue(Utilities.formatDate(date, 'JST', 'dd'));
    sheet.getRange(5, i + 5).setValue(['日', '月', '火', '水', '木', '金', '土'][date.getDay()]);
    
    // 当日判定：今日の日付と一致する場合に色を設定
    if (Utilities.formatDate(today, 'JST', 'yyyyMMdd') === Utilities.formatDate(date, 'JST', 'yyyyMMdd')) {
      sheet.getRange(4, i + 5).setBackground('#FFFF00');  // 4行目の当日セルに黄色
      sheet.getRange(5, i + 5).setBackground('#FFFF00');  // 5行目の当日セルに黄色
    } else {
      // 当日でない場合は背景色をリセット
      sheet.getRange(4, i + 5).setBackground('#FFFFFF');
      sheet.getRange(5, i + 5).setBackground('#FFFFFF');
    }
  }
}










