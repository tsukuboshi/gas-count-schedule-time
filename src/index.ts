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

  // セルを読み込み
  const cella1 = sheet.getRange('A1');
  const cella2 = sheet.getRange('A2');
  const cella3 = sheet.getRange('A3');
  const cella4 = sheet.getRange('A4');
  const cellc1 = sheet.getRange('C1');
  const cellc2 = sheet.getRange('C2');
  const cellc3 = sheet.getRange('C3');
  const cellc4 = sheet.getRange('C4');
  const cellc5 = sheet.getRange('C5');
  const cellc6 = sheet.getRange('C6');
  const cellc7 = sheet.getRange('C7');
  const cellc8 = sheet.getRange('C8');
  const cellc9 = sheet.getRange('C9');
  const cellc10 = sheet.getRange('C10');
  const cellc11 = sheet.getRange('C11');
  const cellc12 = sheet.getRange('C12');
  const celld1 = sheet.getRange('D1');
  const celld2 = sheet.getRange('D2');
  const celld3 = sheet.getRange('D3');
  const celld4 = sheet.getRange('D4');
  const celld5 = sheet.getRange('D5');
  const celld6 = sheet.getRange('D6');
  const celld7 = sheet.getRange('D7');
  const celld8 = sheet.getRange('D8');
  const celld9 = sheet.getRange('D9');
  const celld10 = sheet.getRange('D10');
  const celld11 = sheet.getRange('D11');
  const celld12 = sheet.getRange('D12');
  const cellf1 = sheet.getRange('F1');
  const cellf2 = sheet.getRange('F2');
  const cellf3 = sheet.getRange('F3');
  const cellg1 = sheet.getRange('G1');
  const cellg2 = sheet.getRange('G2');
  const cellg3 = sheet.getRange('G3');

  // セルに値を設定
  cella1.setValue('カレンダーID一覧');
  cellc1.setValue('カレンダー色');
  cellc2.setValue('ラベンダー');
  cellc3.setValue('セージ');
  cellc4.setValue('ブドウ');
  cellc5.setValue('フラミンゴ');
  cellc6.setValue('バナナ');
  cellc7.setValue('ミカン');
  cellc8.setValue('ピーコック');
  cellc9.setValue('グラファイト');
  cellc10.setValue('ブルーベリー');
  cellc11.setValue('バジル');
  cellc12.setValue('トマト');
  celld1.setValue('対応するラベル名');
  cellf1.setValue('対象年月');
  cellf2.setValue('対象年');
  cellf3.setValue('対象月');
  cellg1.setValue('値');

  // ヘッダー行のセルを配列にまとめる
  const headerCells = [cella1, cellc1, celld1, cellf1, cellg1];

  headerCells.forEach(cell => {
    // ヘッダー行のフォントを太字に設定
    cell.setFontWeight('bold');
    // ヘッダー行の背景色を設定
    cell.setBackground('lightblue');
  });

  // 全てのセルを配列にまとめる
  const cells = [
    cella1,
    cella2,
    cella3,
    cella4,
    cellc1,
    cellc2,
    cellc3,
    cellc4,
    cellc5,
    cellc6,
    cellc7,
    cellc8,
    cellc9,
    cellc10,
    cellc11,
    cellc12,
    celld1,
    celld2,
    celld3,
    celld4,
    celld5,
    celld6,
    celld7,
    celld8,
    celld9,
    celld10,
    celld11,
    celld12,
    cellf1,
    cellf2,
    cellf3,
    cellg1,
    cellg2,
    cellg3,
  ];

  cells.forEach(cell => {
    // ヘッダー行の枠線を設定
    cell.setBorder(
      true,
      true,
      true,
      true,
      true,
      true,
      'black',
      SpreadsheetApp.BorderStyle.SOLID
    );
  });
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
  const cella1 = sheet.getRange('A1');
  const cellb1 = sheet.getRange('B1');
  const cellc1 = sheet.getRange('C1');
  const celld1 = sheet.getRange('D1');
  const celle1 = sheet.getRange('E1');
  const cellf1 = sheet.getRange('F1');
  const cellg1 = sheet.getRange('G1');
  const celli1 = sheet.getRange('I1');
  const cellj1 = sheet.getRange('J1');
  const cellk1 = sheet.getRange('K1');

  // ヘッダー行に値を設定
  cella1.setValue('カレンダー名');
  cellb1.setValue('タイトル');
  cellc1.setValue('色');
  celld1.setValue('ラベル名');
  celle1.setValue('ステータス');
  cellf1.setValue('開始日時');
  cellg1.setValue('工数(h)');
  celli1.setValue('ラベル名一覧');
  cellj1.setValue('工数合計(h)');
  cellk1.setValue('工数割合(%)');

  // ヘッダー行のセルを配列にまとめる
  const cells = [
    cella1,
    cellb1,
    cellc1,
    celld1,
    celle1,
    cellf1,
    cellg1,
    celli1,
    cellj1,
    cellk1,
  ];

  cells.forEach(cell => {
    // ヘッダー行のフォントを太字に設定
    cell.setFontWeight('bold');
    // ヘッダー行の背景色を設定
    cell.setBackground('orange');
    // ヘッダー行の枠線を設定
    cell.setBorder(
      true,
      true,
      true,
      true,
      true,
      true,
      'black',
      SpreadsheetApp.BorderStyle.SOLID
    );
  });
}

// イベントをシートに書き込む関数
function writeEventsToSheet(
  events: GoogleAppsScript.Calendar.CalendarEvent[],
  sheet: GoogleAppsScript.Spreadsheet.Sheet,
  calendarNames: string[],
  colorMap: { [key: string]: string },
  labelMap: { [key: string]: string }
): void {
  for (let i = 0; i < events.length; i++) {
    // イベントの情報を取得
    const event = events[i];
    const { title, color, status, startTime, durationHours } =
      getEventContent(event);

    // イベントに対応するカレンダー名を取得
    const calendarName = calendarNames[i];

    // イベントの色番号を色名に変換
    const colorName = colorMap[color];

    // イベントの色名をラベル名に変換
    const labelName = labelMap[colorName];

    // ヘッダー行を考慮して2行目から開始
    const row = i + 2;

    // イベントの情報をシートに書き込む
    sheet.getRange('A' + row).setValue(calendarName);
    sheet.getRange('B' + row).setValue(title);
    sheet.getRange('C' + row).setValue(colorName);
    sheet.getRange('D' + row).setValue(labelName);
    sheet.getRange('E' + row).setValue(status);
    sheet.getRange('F' + row).setValue(startTime);
    sheet.getRange('G' + row).setValue(durationHours.toString());
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
  // デフォルトカラーの色番号（'0'）の工数をMapから削除する
  workHoursMap.delete('0');

  // 全工数の合計を計算する
  let totalWorkHours = 0;
  workHoursMap.forEach(hours => (totalWorkHours += hours));

  // ヘッダー行を考慮して2行目から開始
  let row = 2;

  // labelMapを使用してラベル名一覧を順番にシートに書き込む
  Object.keys(labelMap).forEach(colorName => {
    // カラー名に対応するラベル名を取得
    const labelName = labelMap[colorName];

    // ラベル名が空文字でない場合のみ処理を行う
    if (labelName) {
      // 色名に対応する色番号を取得
      const color = Object.keys(colorMap).find(
        key => colorMap[key] === colorName
      );
      // 色番号に対応する工数を取得(存在しない場合は0)
      const hours = color ? workHoursMap.get(color) : 0;
      // イベントの工数の割合を計算(存在しない場合は0)
      const workPercentage = hours
        ? ((hours / totalWorkHours) * 100).toFixed(1)
        : 0;

      // イベントの情報をシートに書き込む
      sheet.getRange('I' + row).setValue(labelName);
      sheet.getRange('J' + row).setValue(hours);
      sheet.getRange('K' + row).setValue(workPercentage);
      row++;
    }
  });

  // ラベル名一覧と工数合計が記載されている範囲を指定
  const dataRange = sheet.getRange('I2:J' + (row - 1));
  return dataRange;
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
