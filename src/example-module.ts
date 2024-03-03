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

export function main(
  calendarIds: string[],
  LabelMap: { [key: string]: string }
): void {
  const colorMap: { [key: string]: string } = {
    '0': 'デフォルト', // カラーが設定されていない場合の色番号は'0とする
    '1': 'ラベンダー',
    '2': 'セージ',
    '3': 'ブドウ',
    '4': 'フラミンゴ',
    '5': 'バナナ',
    '6': 'ミカン',
    '7': 'ピーコック',
    '8': 'グラファイト',
    '9': 'ブルーベリー',
    '10': 'バジル',
    '11': 'トマト',
  };

  // スプレッドシートをIDで読み込む
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();

  // 現在の年と月を取得または指定された年月を使用
  const { year: currentYear, month: currentMonth } = getCurrentYearMonth();

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
  writeEventsToSheetByColor(
    allEvents,
    currentMonthSheet,
    calendarNames,
    colorMap,
    LabelMap
  );

  // 色番号ごとの工数を計算
  const workHoursMap = calculateWorkHoursByColor(allEvents);

  // 計算した工数をシートに記載
  writeWorkHoursToSheet(currentMonthSheet, workHoursMap, colorMap, LabelMap);
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
function writeEventsToSheetByColor(
  events: GoogleAppsScript.Calendar.CalendarEvent[],
  sheet: GoogleAppsScript.Spreadsheet.Sheet,
  calendarNames: string[],
  colorMap: { [key: string]: string },
  LabelMap: { [key: string]: string }
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
    const labelName = LabelMap[colorName];

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
    // ステータスがOWNER, YES, ETCの予定のみを対象とする
    if (status === 'OWNER' || status === 'YES' || status === 'ETC') {
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

// スプレッドシートのG列に色番号、H列に計算した工数合計を記載する関数
function writeWorkHoursToSheet(
  sheet: GoogleAppsScript.Spreadsheet.Sheet,
  workHoursMap: Map<string, number>,
  colorMap: { [key: string]: string },
  LabelMap: { [key: string]: string }
): void {
  // ヘッダー行を考慮して2行目から開始
  let row = 2;

  // 色番号ごとの工数をシートに書き込む
  workHoursMap.forEach((hours, color) => {
    // イベントの色番号を色名に変換
    const colorName = colorMap[color];

    // イベントの色名をラベル名に変換
    const labelName = LabelMap[colorName];

    // イベントの情報をシートに書き込む
    sheet.getRange('I' + row).setValue(labelName);
    sheet.getRange('J' + row).setValue(hours);
    row++;
  });
}

// 現在の年と月を取得する関数
function getCurrentYearMonth(): { year: number; month: number } {
  // スクリプトのプロパティから年と月を取得
  const scriptProperties = PropertiesService.getScriptProperties();
  const propYear = scriptProperties.getProperty('YEAR');
  const propMonth = scriptProperties.getProperty('MONTH');

  // 環境変数が設定されている場合はそれを使用、そうでない場合は現在の年月を使用
  const envYear = propYear ? Number(propYear) : undefined;
  const envMonth = propMonth ? Number(propMonth) : undefined;

  if (envYear !== undefined && envMonth !== undefined) {
    // スクリプトのプロパティが設定されている場合はそれを使用
    return { year: envYear, month: envMonth };
  } else {
    // 引数が与えられていない場合は現在の年月を返す
    const now = new Date();
    const year = now.getFullYear();
    const month = now.getMonth() + 1; // getMonth() は0から始まるため、1を足す
    return { year, month };
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
