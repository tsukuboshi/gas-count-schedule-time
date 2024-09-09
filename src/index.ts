/**
 * Copyright 2023 tsukuboshi
 *
 * Licensed under the Apache License, Version 2.0 (the "License");
 * you may not use this file except in compliance with the License.
 * You may obtain a copy of the License at
 *
 *       http://www.apache.org/licenses/LICENSE-2.0
 *
 * Unless required by applicable law or agreed to in writing, software
 * distributed under the License is distributed on an "AS IS" BASIS,
 * WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
 * See the License for the specific language governing permissions and
 * limitations under the License.
 */

import { numberOfCalenders } from './env';

// eslint-disable-next-line @typescript-eslint/no-unused-vars
function onOpen() {
  // UIにカスタムメニューを追加
  addCustomMenuToUi();
  // シート1にプロパティを書き込む
  writePropertyToSheet1();
}

// eslint-disable-next-line @typescript-eslint/no-unused-vars
function main(): void {
  // スプレッドシートをIDで読み込む
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = spreadsheet.getActiveSheet();

  // B列のセルよりカレンダーID一覧を配列で取得
  const calendarIds = sheet
    .getRange('A2:A4')
    .getValues()
    .flat()
    .filter(id => id !== '');

  // C列のセルよりカラー名一覧を配列で取得
  const colorNames = sheet.getRange('C2:C12').getValues().flat();

  // D列のセルよりラベル名一覧を配列で取得
  const labelNames = sheet.getRange('D2:D12').getValues().flat();

  const colorMap: { [key: string]: string } = {
    '0': 'デフォルト', // カラーが設定されていない場合の色番号は'0'とする
  };

  // 色番号とラベル名のペアをColorMapに追加
  for (let i = 0; i < colorNames.length; i++) {
    colorMap[String(i + 1)] = colorNames[i];
  }

  // カラー名とラベル名のペアをLabelMapに追加
  const labelMap: { [key: string]: string } = {};
  for (let i = 0; i < colorNames.length; i++) {
    labelMap[colorNames[i]] = labelNames[i];
  }

  // 現在の年と月を取得または指定された年月を使用
  const { year: currentYear, month: currentMonth } = getCurrentYearMonth(sheet);

  // 現在の月のシートを取得または作成
  const currentMonthSheet = getOrCreateCurrentMonthSheet(
    spreadsheet,
    currentYear,
    currentMonth
  );

  // 月初から月末までの開始時刻と終了時刻を取得
  const { startTime, endTime } = getStartTimeAndEndTime(
    currentYear,
    currentMonth - 1
  ); // monthは0から始まるため、1を引く

  // 全カレンダーのイベントを保持する配列
  const allEvents: GoogleAppsScript.Calendar.CalendarEvent[] = [];

  // 各イベントに対応するカレンダー名を保持する配列
  const calendarNames: string[] = []; // この行を追加

  // 各カレンダーからイベントを取得し、allEventsに統合する
  calendarIds.forEach(calendarId => {
    const calendar = getCalendar(calendarId);
    const events = getEventsfromCalendar(calendar, startTime, endTime);
    const calendarName = calendar.getName(); // カレンダー名を取得
    events.forEach(event => {
      allEvents.push(event);
      calendarNames.push(calendarName); // 各イベントに対応するカレンダー名を保存
    });
  });

  // ヘッダー行をシートに書き込む関数
  writeHeaderColumn(currentMonthSheet);

  // イベントをシートに書き込む
  writeEventsToSheet(
    allEvents,
    currentMonthSheet,
    calendarNames,
    colorMap,
    labelMap
  );

  // 色番号ごとの工数を計算
  const workHoursMap = calculateWorkHoursByColor(allEvents);

  // 計算した工数をシートに記載
  const dataRange = writeSummaryToSheet(
    currentMonthSheet,
    workHoursMap,
    colorMap,
    labelMap
  );

  // チャートを作成または更新
  createOrUpdateChart(currentMonthSheet, dataRange);
}

// カスタムメニューをUIに追加する関数
function addCustomMenuToUi() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('カスタムメニュー').addItem('工数計算', 'main').addToUi();
}

// プロパティをデフォルトシートに書き込む関数
function writePropertyToSheet1(): void {
  // スプレッドシートをIDで読み込む
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = spreadsheet.getActiveSheet();

  // シート1名を変更
  const sheetName = 'プロパティ';
  sheet.setName(sheetName);

  // ヘッダーを生成
  const calenderHeader = ['Calender IDs'];
  const colorHeaders = [
    'Lavender',
    'Sage',
    'Grape',
    'Flamingo',
    'Banana',
    'Tangerine',
    'Peacock',
    'Graphite',
    'Blueberry',
    'Basil',
    'Tomato',
  ];
  const timeHeaders = ['Year', 'Month'];

  // ヘッダー行を設定
  const headers = [calenderHeader, colorHeaders, timeHeaders];
  const headerRange = sheet.getRange(1, 1, 1, headers.length);
  headerRange.setValues([headers]);
  headerRange.setFontWeight('bold');
  headerRange.setBackground('lightblue');

  // すべてのセルの範囲を取得
  const allRange = sheet.getRange(1, 1, numberOfCalenders + 1, headers.length);

  // すべてのセルに枠線を設定
  allRange.setBorder(
    true,
    true,
    true,
    true,
    true,
    true,
    'black',
    SpreadsheetApp.BorderStyle.SOLID
  );

  // 列の幅を自動調整
  sheet.autoResizeColumns(1, headers.length);

  console.log('プロパティシートが作成されました。');
}

// 現在の月のシートを取得または作成する関数
function getOrCreateCurrentMonthSheet(
  spreadsheet: GoogleAppsScript.Spreadsheet.Spreadsheet,
  year: number,
  month: number
): GoogleAppsScript.Spreadsheet.Sheet {
  const sheetName = `${year}年${month}月`;

  // 指定された名前のシートを取得
  let sheet = spreadsheet.getSheetByName(sheetName);

  // シートが存在しない場合は新たに作成
  if (!sheet) {
    sheet = spreadsheet.insertSheet(sheetName);
    if (!sheet) {
      throw new Error(`シート '${sheetName}' の作成に失敗しました。`);
    }
  }
  // 既存または新規作成されたシートを返す
  return sheet;
}

// カレンダーをIDで取得する関数
function getCalendar(calendarId: string): GoogleAppsScript.Calendar.Calendar {
  const calendar = CalendarApp.getCalendarById(calendarId);
  if (!calendar) {
    throw new Error('指定されたIDのカレンダーが見つかりません。');
  }
  return calendar;
}
// 月初から月末までの開始時刻と終了時刻を取得する関数
function getStartTimeAndEndTime(
  year: number,
  month: number
): { startTime: Date; endTime: Date } {
  // 月初
  const startTime = new Date(year, month, 1);

  // 月末
  const endTime = new Date(year, month + 1, 0);
  return { startTime, endTime };
}

// カレンダーからイベントを取得する関数
function getEventsfromCalendar(
  calendar: GoogleAppsScript.Calendar.Calendar,
  startTime: Date,
  endTime: Date
): GoogleAppsScript.Calendar.CalendarEvent[] {
  const events = calendar.getEvents(startTime, endTime);
  if (!events) {
    throw new Error('指定された期間にイベントが見つかりませんでした。');
  }
  return events;
}

// ヘッダー行をシートに書き込む関数
function writeHeaderColumn(sheet: GoogleAppsScript.Spreadsheet.Sheet): void {
  // ヘッダー行を読み込み
  const headers = [
    'カレンダー名',
    'タイトル',
    '色',
    'ラベル名',
    'ステータス',
    '開始日時',
    '工数(h)',
    '',
    'ラベル名一覧',
    '工数合計(h)',
    '工数割合(%)',
  ];

  const headerRange = sheet.getRange(1, 1, 1, headers.length);

  headerRange.setValues([headers]);
  headerRange.setFontWeight('bold');
  headerRange.setBackground('orange');
  headerRange.setBorder(
    true,
    true,
    true,
    true,
    true,
    true,
    'black',
    SpreadsheetApp.BorderStyle.SOLID
  );
}

// イベントをシートに書き込む関数
function writeEventsToSheet(
  events: GoogleAppsScript.Calendar.CalendarEvent[],
  sheet: GoogleAppsScript.Spreadsheet.Sheet,
  calendarNames: string[],
  colorMap: { [key: string]: string },
  labelMap: { [key: string]: string }
): void {
  // イベントデータを2次元配列に変換
  const eventData = events.map((event, index) => {
    // イベントの詳細情報を取得
    const { title, color, status, startTime, durationHours } = getEventContent(event);
    // イベントに対応するカレンダー名を取得
    const calendarName = calendarNames[index];
    // イベントの色番号を色名に変換
    const colorName = colorMap[color];
    // イベントの色名をラベル名に変換
    const labelName = labelMap[colorName];

    // イベント情報を配列として返す
    return [
      calendarName,
      title,
      colorName,
      labelName,
      status,
      startTime,
      durationHours.toString()
    ];
  });

  // イベントデータが存在する場合のみシートに書き込み
  if (eventData.length > 0) {
    // 書き込む範囲を取得（2行目から開始、A列からG列まで）
    const range = sheet.getRange(2, 1, eventData.length, 7);
    // データを一括で書き込み
    range.setValues(eventData);
  }
}

// ステータスがOWNERまたはYESのみの予定の内、色番号が同じ予定について、工数の合計を計算する関数
function calculateWorkHoursByColor(
  events: GoogleAppsScript.Calendar.CalendarEvent[]
): Map<string, number> {
  // 色番号ごとの工数を計算するMap
  const workHoursMap = new Map<string, number>();

  // イベントの情報を取得し、色番号ごとに工数を計算
  events.forEach(event => {
    // イベントのステータスを取得(nullの場合は空文字に変換)
    const statusObj = event.getMyStatus();
    let status = statusObj ? statusObj.toString() : '';
    //  ステータスが設定されていない場合の状態は'ETC'とする
    if (status === '') {
      status = 'ETC';
    }
    // ステータスがOWNER, YESの予定のみを対象とする
    if (status === 'OWNER' || status === 'YES') {
      // イベントの情報を取得
      const eventContent = getEventContent(event);
      let color = eventContent.color;
      const durationHours = eventContent.durationHours;
      // カラーが設定されていない場合の色番号は'0'とする
      if (color === '') {
        color = '0';
      }
      // 色番号ごとに工数を計算
      const currentHours = workHoursMap.get(color);
      if (!currentHours) {
        // Mapに色番号が存在しない場合は新たに追加
        workHoursMap.set(color, durationHours);
      } else {
        // Mapに色番号が存在する場合は工数を加算
        workHoursMap.set(color, currentHours + durationHours);
      }
    }
  });
  return workHoursMap;
}

// 集計結果ををシートに記載する関数
function writeSummaryToSheet(
  sheet: GoogleAppsScript.Spreadsheet.Sheet,
  workHoursMap: Map<string, number>,
  colorMap: { [key: string]: string },
  labelMap: { [key: string]: string }
): GoogleAppsScript.Spreadsheet.Range {
  // デフォルトカラー（'0'）の工数を削除
  workHoursMap.delete('0');

  // 全工数の合計を計算
  const totalWorkHours = Array.from(workHoursMap.values()).reduce(
    (sum, hours) => sum + hours,
    0
  );

  // サマリーデータの生成
  const summaryData = Object.entries(labelMap)
    // ラベル名が存在する項目のみをフィルタリング
    .filter(([, labelName]) => labelName)
    // 各ラベルの情報を配列に変換
    .map(([colorName, labelName]) => {
      // 色名に対応する色番号を取得
      const color = Object.keys(colorMap).find(
        key => colorMap[key] === colorName
      );
      // 色番号に対応する工数を取得（存在しない場合は0）
      const hours = color ? workHoursMap.get(color) || 0 : 0;
      // 工数の割合を計算（小数点以下1桁まで）
      const workPercentage = hours
        ? ((hours / totalWorkHours) * 100).toFixed(1)
        : '0';
      // [ラベル名, 工数, 割合] の形式で返す
      return [labelName, hours, workPercentage];
    });

  // データが存在する場合のみシートに書き込み
  if (summaryData.length > 0) {
    // 書き込む範囲を取得（2行目、I列から開始）
    const range = sheet.getRange(2, 9, summaryData.length, 3);
    // データを一括で書き込み
    range.setValues(summaryData);
  }

  // ラベル名と工数のデータ範囲（I列とJ列）を返す
  return sheet.getRange(2, 9, summaryData.length, 2);
}

//
function createOrUpdateChart(
  sheet: GoogleAppsScript.Spreadsheet.Sheet,
  dataRange: GoogleAppsScript.Spreadsheet.Range
): void {
  // チャートの種類を円グラフに指定
  const chartType = Charts.ChartType.PIE;

  // チャートの位置をM列（13列目）に設定
  const column = 13;

  // シートから既存のチャートを取得
  const charts = sheet.getCharts();

  // 既存のチャートがあれば更新、なければ新しく作成
  if (charts.length > 0) {
    // 既存のチャートを更新
    const chart = charts[0];
    const updatedChart = chart
      .modify()
      .addRange(dataRange)
      .setPosition(1, column, 0, 0)
      .build();
    sheet.updateChart(updatedChart);
  } else {
    // 新しいチャートを作成
    const newChart = sheet
      .newChart()
      .setChartType(chartType)
      .addRange(dataRange)
      .setPosition(1, column, 0, 0)
      .build();
    sheet.insertChart(newChart);
  }
}

// 現在の年と月を取得する関数
function getCurrentYearMonth(sheet: GoogleAppsScript.Spreadsheet.Sheet): {
  year: number;
  month: number;
} {
  // セルG2とG3の値を取得
  const yearValue = sheet.getRange('G2').getValue();
  const monthValue = sheet.getRange('G3').getValue();

  // セルG2とG3に値が設定されている場合はそれを使用、そうでない場合は現在の年月を使用
  const year = yearValue ? Number(yearValue) : undefined;
  const month = monthValue ? Number(monthValue) : undefined;

  if (year !== undefined && month !== undefined) {
    // セルG2とG3に値が設定されている場合はそれを使用
    return { year, month };
  } else {
    // セルG2とG3に値が設定されていない場合は現在の年月を返す
    const now = new Date();
    const currentYear = now.getFullYear();
    const currentMonth = now.getMonth() + 1; // getMonth() は0から始まるため、1を足す
    return { year: currentYear, month: currentMonth };
  }
}

// イベントの情報を取得する関数
function getEventContent(event: GoogleAppsScript.Calendar.CalendarEvent): {
  title: string;
  color: string;
  status: string;
  startTime: GoogleAppsScript.Base.Date;
  durationHours: number;
} {
  // イベントのタイトルを取得
  const title = event.getTitle();
  // イベントの色番号を取得
  let color = event.getColor();
  // カラーが設定されていない場合の色番号は'0'とする
  if (color === '') {
    color = '0';
  }
  // イベントのステータスを取得
  const statusObj = event.getMyStatus();
  //  ステータスが設定されていない場合の状態は'ETC'とする
  let status = statusObj ? statusObj.toString() : '';
  if (status === '') {
    status = 'ETC';
  }
  // イベントの開始日時を取得
  const startTime = event.getStartTime();
  // イベントの終了日時を取得
  const endTime = event.getEndTime();
  // イベントの工数を計算(ミリ秒を時間に変換)
  const durationHours =
    (endTime.getTime() - startTime.getTime()) / (1000 * 60 * 60);
  return { title, color, status, startTime, durationHours };
}
