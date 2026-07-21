/**
 * 社用車予約管理システム
 * Code.gs - サーバーサイドロジック
 */

// ==================== 定数 ====================
const SHEET_VEHICLE = '車両マスタ';
const SHEET_UNAVAILABLE = '使用不可期間';
const SHEET_RESERVATION = '予約データ';

const VEHICLE_HEADERS = ['車両ID', '車種名', 'ナンバープレート', '定員', '利用可能開始時刻', '利用可能終了時刻', '備考'];
const UNAVAILABLE_HEADERS = ['ID', '車両ID', '開始日時', '終了日時', '理由'];
const RESERVATION_HEADERS = ['予約ID', '車両ID', '開始日時', '終了日時', '利用者名', '利用目的', '行き先', '登録日時', 'シリーズID'];

// ==================== 初期化 ====================

/**
 * 実行コンテキストによらず、対象のスプレッドシートを確実に取得する。
 * スプレッドシートIDはハードコードせず、PropertiesService(スクリプトプロパティ)に
 * 自動保存されたものを使う。IDはこのスクリプトが最初にonOpen()で実行された時点
 * (＝スプレッドシートを開いたとき)に自動的に保存される。
 */
function getSpreadsheet_() {
  const props = PropertiesService.getScriptProperties();
  let id = props.getProperty('SPREADSHEET_ID');

  if (!id) {
    // まだ保存されていない場合、コンテナバインドの文脈(サイドバー・メニュー経由)であれば
    // アクティブなスプレッドシートから取得して保存する
    const active = SpreadsheetApp.getActiveSpreadsheet();
    if (active) {
      id = active.getId();
      props.setProperty('SPREADSHEET_ID', id);
    } else {
      throw new Error(
        'スプレッドシートIDが未設定です。一度スプレッドシートを開いて(メニューが表示されるまで待って)から、' +
        'もう一度お試しください。'
      );
    }
  }
  return SpreadsheetApp.openById(id);
}

/**
 * スプレッドシートを開いたときにメニューを追加し、同時にスプレッドシートIDを
 * スクリプトプロパティに保存する(Webアプリ経由でも参照できるようにするため)
 */
function onOpen() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  PropertiesService.getScriptProperties().setProperty('SPREADSHEET_ID', ss.getId());

  SpreadsheetApp.getUi()
    .createMenu('社用車予約管理')
    .addItem('初期セットアップ', 'setupSheets')
    .addItem('マスタ画面を開く', 'showMasterSidebar')
    .addToUi();
}

/**
 * 必要なシートが存在しない場合に作成し、ヘッダーを設定する。
 * また、初期状態でスプレッドシートに残っている不要なシート(既定の「シート1」など)を
 * 自動的に削除し、車両マスタ・使用不可期間・予約データの3シートだけが残るようにする。
 *
 * Apps Scriptエディタの関数選択プルダウンからこの関数を直接実行しても動作する
 * (スプレッドシートのメニュー経由でなくてもよい)。GitHubからコードを取得して
 * ゼロから再現する場合は、この関数を実行するだけでセットアップが完了する。
 */
function setupSheets() {
  const ss = getSpreadsheet_();
  ensureSheet_(ss, SHEET_VEHICLE, VEHICLE_HEADERS);
  ensureSheet_(ss, SHEET_UNAVAILABLE, UNAVAILABLE_HEADERS);
  ensureSheet_(ss, SHEET_RESERVATION, RESERVATION_HEADERS);

  // 上記3シート以外(新規スプレッドシート作成時の既定シートなど)は不要なので削除する
  const requiredNames = [SHEET_VEHICLE, SHEET_UNAVAILABLE, SHEET_RESERVATION];
  ss.getSheets().forEach(function (sheet) {
    if (requiredNames.indexOf(sheet.getName()) === -1) {
      ss.deleteSheet(sheet);
    }
  });

  Logger.log('初期セットアップが完了しました。「車両マスタ」「使用不可期間」「予約データ」の3シートを作成し、不要なシートを削除しました。');

  // スプレッドシートのメニューから実行した場合のみアラートを表示する。
  // Apps Scriptエディタから直接実行した場合はUIが存在せずエラーになるため、その場合は何もしない
  // (実行結果は実行ログ、または左メニューの「実行数」から確認できる)。
  try {
    SpreadsheetApp.getUi().alert('初期セットアップが完了しました。「車両マスタ」「使用不可期間」「予約データ」の3シートを作成しました。');
  } catch (e) {
    // no-op: エディタから直接実行した場合
  }
}

function ensureSheet_(ss, name, headers) {
  let sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
  }
  if (sheet.getLastRow() === 0) {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function getSheet_(name) {
  const ss = getSpreadsheet_();
  const sheet = ss.getSheetByName(name);
  if (!sheet) {
    throw new Error('シート「' + name + '」が見つかりません。先に初期セットアップを実行してください。');
  }
  return sheet;
}

// ==================== Web公開エントリーポイント ====================

function doGet(e) {
  const page = (e && e.parameter && e.parameter.page) || 'reservation';
  let template;
  if (page === 'master') {
    template = HtmlService.createTemplateFromFile('マスタ画面');
  } else if (page === 'register') {
    template = HtmlService.createTemplateFromFile('予約登録画面');
    template.prefill = {
      vehicleId: (e && e.parameter && e.parameter.vehicleId) || '',
      startDate: (e && e.parameter && e.parameter.startDate) || '',
      startTime: (e && e.parameter && e.parameter.startTime) || '',
      endDate: (e && e.parameter && e.parameter.endDate) || '',
      endTime: (e && e.parameter && e.parameter.endTime) || ''
    };
  } else {
    // 'reservation' およびページ指定なし(デフォルト)の場合
    template = HtmlService.createTemplateFromFile('予約確認画面');
  }
  // 相対URL(?page=xxx)はWebアプリがiframe内に埋め込まれる構造のため誤解決される。
  // 各画面のナビゲーションリンクは、この絶対URLを基準に組み立てる。
  template.webAppUrl = ScriptApp.getService().getUrl();
  return template.evaluate()
    .setTitle('社用車予約管理')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

function showMasterSidebar() {
  const template = HtmlService.createTemplateFromFile('マスタ画面');
  template.webAppUrl = ScriptApp.getService().getUrl();
  const html = template.evaluate().setTitle('車両マスタ管理');
  SpreadsheetApp.getUi().showSidebar(html);
}

// ==================== 車両ID採番 ====================

function generateId_(sheet, idColIndex) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return 1;
  const ids = sheet.getRange(2, idColIndex, lastRow - 1, 1).getValues().flat().filter(function (v) { return v !== ''; });
  if (ids.length === 0) return 1;
  return Math.max.apply(null, ids.map(Number)) + 1;
}

/**
 * スプレッドシートが時刻の入力をDate型として自動認識した場合に "HH:mm" 文字列へ変換する。
 * google.script.run はDate型(特に1899年などの特殊な日付)の返却でシリアライズに失敗し、
 * 応答が返らないまま止まることがあるため、クライアントに返す前に必ず文字列化する。
 */
function formatTimeCell_(value) {
  if (value instanceof Date) {
    return Utilities.formatDate(value, Session.getScriptTimeZone(), 'HH:mm');
  }
  return value;
}

/**
 * 日付+時刻(使用不可期間・予約の開始/終了日時など)をDate型のまま返さず
 * ISO 8601文字列に変換する。google.script.runでのDate型の返却は
 * シリアライズに失敗し応答が返らなくなることがあるため、必ず文字列化してから返す。
 */
function formatDateTimeCell_(value) {
  if (value instanceof Date) {
    return Utilities.formatDate(value, Session.getScriptTimeZone(), "yyyy-MM-dd'T'HH:mm:ss");
  }
  return value;
}

// ==================== 車両マスタ CRUD ====================

/**
 * 車両一覧を取得
 * @return {Array<Object>}
 */
function getVehicles() {
  const sheet = getSheet_(SHEET_VEHICLE);
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  const values = sheet.getRange(2, 1, lastRow - 1, VEHICLE_HEADERS.length).getValues();
  return values.map(function (row) {
    return {
      vehicleId: row[0],
      name: row[1],
      plateNumber: row[2],
      capacity: row[3],
      availableStart: formatTimeCell_(row[4]),
      availableEnd: formatTimeCell_(row[5]),
      note: row[6]
    };
  });
}

/**
 * 車両を新規登録または更新する
 * @param {Object} vehicle {vehicleId, name, plateNumber, capacity, availableStart, availableEnd, note}
 * @return {Object} 登録後の車両データ
 */
function saveVehicle(vehicle) {
  const sheet = getSheet_(SHEET_VEHICLE);
  const lastRow = sheet.getLastRow();

  if (vehicle.vehicleId) {
    // 更新
    const rowIndex = findRowById_(sheet, lastRow, Number(vehicle.vehicleId));
    if (rowIndex === -1) {
      throw new Error('更新対象の車両が見つかりません。(ID: ' + vehicle.vehicleId + ')');
    }
    sheet.getRange(rowIndex, 2, 1, 6).setValues([[
      vehicle.name, vehicle.plateNumber, vehicle.capacity,
      vehicle.availableStart, vehicle.availableEnd, vehicle.note
    ]]);
    return vehicle;
  } else {
    // 新規登録
    const newId = generateId_(sheet, 1);
    sheet.appendRow([
      newId, vehicle.name, vehicle.plateNumber, vehicle.capacity,
      vehicle.availableStart, vehicle.availableEnd, vehicle.note
    ]);
    vehicle.vehicleId = newId;
    return vehicle;
  }
}

/**
 * 車両を削除する(関連する使用不可期間・予約データも削除)
 * @param {number} vehicleId
 */
function deleteVehicle(vehicleId) {
  const sheet = getSheet_(SHEET_VEHICLE);
  const rowIndex = findRowById_(sheet, sheet.getLastRow(), Number(vehicleId));
  if (rowIndex === -1) {
    throw new Error('削除対象の車両が見つかりません。(ID: ' + vehicleId + ')');
  }
  sheet.deleteRow(rowIndex);

  deleteRowsByVehicleId_(getSheet_(SHEET_UNAVAILABLE), 2, Number(vehicleId));
  deleteRowsByVehicleId_(getSheet_(SHEET_RESERVATION), 2, Number(vehicleId));
}

function findRowById_(sheet, lastRow, id) {
  if (lastRow < 2) return -1;
  const ids = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
  for (let i = 0; i < ids.length; i++) {
    if (Number(ids[i][0]) === id) return i + 2;
  }
  return -1;
}

function deleteRowsByVehicleId_(sheet, vehicleIdColIndex, vehicleId) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return;
  const values = sheet.getRange(2, vehicleIdColIndex, lastRow - 1, 1).getValues();
  for (let i = values.length - 1; i >= 0; i--) {
    if (Number(values[i][0]) === vehicleId) {
      sheet.deleteRow(i + 2);
    }
  }
}

// ==================== 使用不可期間 CRUD ====================

/**
 * 使用不可期間の一覧を取得(車両ID指定で絞り込み可)
 * @param {number} [vehicleId]
 * @return {Array<Object>}
 */
function getUnavailablePeriods(vehicleId) {
  const sheet = getSheet_(SHEET_UNAVAILABLE);
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  const values = sheet.getRange(2, 1, lastRow - 1, UNAVAILABLE_HEADERS.length).getValues();
  const list = values.map(function (row) {
    return {
      id: row[0],
      vehicleId: row[1],
      start: formatDateTimeCell_(row[2]),
      end: formatDateTimeCell_(row[3]),
      reason: row[4]
    };
  });
  if (vehicleId === undefined || vehicleId === null || vehicleId === '') return list;
  return list.filter(function (item) { return Number(item.vehicleId) === Number(vehicleId); });
}

/**
 * 使用不可期間を登録する
 * @param {Object} period {vehicleId, start, end, reason} start/endはISO文字列
 */
function saveUnavailablePeriod(period) {
  const sheet = getSheet_(SHEET_UNAVAILABLE);
  const newId = generateId_(sheet, 1);
  sheet.appendRow([newId, period.vehicleId, new Date(period.start), new Date(period.end), period.reason]);
  period.id = newId;
  return period;
}

/**
 * 使用不可期間を削除する
 * @param {number} id
 */
function deleteUnavailablePeriod(id) {
  const sheet = getSheet_(SHEET_UNAVAILABLE);
  const rowIndex = findRowById_(sheet, sheet.getLastRow(), Number(id));
  if (rowIndex === -1) {
    throw new Error('削除対象の使用不可期間が見つかりません。(ID: ' + id + ')');
  }
  sheet.deleteRow(rowIndex);
}

// ==================== 予約データ ====================

/**
 * 指定年月に重なる予約を、全車両分まとめて取得する(予約確認画面の全車両表示用)
 * @param {number} year
 * @param {number} month 1-12
 * @return {Array<Object>}
 */
function getReservationsForMonthAllVehicles(year, month) {
  const sheet = getSheet_(SHEET_RESERVATION);
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  const values = sheet.getRange(2, 1, lastRow - 1, RESERVATION_HEADERS.length).getValues();

  const monthStart = new Date(Number(year), Number(month) - 1, 1, 0, 0, 0);
  const monthEnd = new Date(Number(year), Number(month), 0, 23, 59, 59);
  const list = [];

  values.forEach(function (row) {
    const start = row[2] instanceof Date ? row[2] : new Date(row[2]);
    const end = row[3] instanceof Date ? row[3] : new Date(row[3]);
    if (end < monthStart || start > monthEnd) return;

    list.push({
      reservationId: row[0],
      vehicleId: row[1],
      start: formatDateTimeCell_(start),
      end: formatDateTimeCell_(end),
      userName: row[4],
      purpose: row[5],
      destination: row[6],
      seriesId: row[8] || ''
    });
  });

  return list;
}

/**
 * 指定車両・指定年月に重なる予約一覧を取得する(予約確認画面用)
 * 日をまたぐ予約にも対応するため、日付単体の一致ではなく「期間が月と重なるか」で判定する。
 * @param {number} vehicleId
 * @param {number} year
 * @param {number} month 1-12
 * @return {Array<Object>} 各要素の start/end はISO文字列("yyyy-MM-ddTHH:mm:ss")
 */
function getReservationsForMonth(vehicleId, year, month) {
  const sheet = getSheet_(SHEET_RESERVATION);
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  const values = sheet.getRange(2, 1, lastRow - 1, RESERVATION_HEADERS.length).getValues();

  const monthStart = new Date(Number(year), Number(month) - 1, 1, 0, 0, 0);
  const monthEnd = new Date(Number(year), Number(month), 0, 23, 59, 59);
  const list = [];

  values.forEach(function (row) {
    if (Number(row[1]) !== Number(vehicleId)) return;
    const start = row[2] instanceof Date ? row[2] : new Date(row[2]);
    const end = row[3] instanceof Date ? row[3] : new Date(row[3]);
    if (end < monthStart || start > monthEnd) return; // 月と重ならない予約は除外

    list.push({
      reservationId: row[0],
      vehicleId: row[1],
      start: formatDateTimeCell_(start),
      end: formatDateTimeCell_(end),
      userName: row[4],
      purpose: row[5],
      destination: row[6],
      seriesId: row[8] || ''
    });
  });

  return list;
}

/**
 * 指定車両・指定期間と重なる予約を取得する(内部処理用。saveReservationの重複チェックで使用)
 */
function getReservationsOverlapping_(vehicleId, start, end) {
  const sheet = getSheet_(SHEET_RESERVATION);
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  const values = sheet.getRange(2, 1, lastRow - 1, RESERVATION_HEADERS.length).getValues();
  const result = [];
  values.forEach(function (row) {
    if (Number(row[1]) !== Number(vehicleId)) return;
    const rStart = row[2] instanceof Date ? row[2] : new Date(row[2]);
    const rEnd = row[3] instanceof Date ? row[3] : new Date(row[3]);
    if (rStart < end && rEnd > start) {
      result.push({ reservationId: row[0], start: rStart, end: rEnd });
    }
  });
  return result;
}

function timeToMinutes_(t) {
  const parts = String(t).split(':');
  return Number(parts[0]) * 60 + Number(parts[1]);
}

/**
 * 予約を登録する。以下を順にチェックし、問題なければ登録する。
 * 1. 必須項目・開始日時 < 終了日時
 * 2. 車両に利用可能時間の制限がある場合、日をまたぐ予約は不可。同日内なら範囲内かチェック
 * 3. 使用不可期間(車検・点検など)と重複していないか
 * 4. 既存の予約と重複していないか
 * @param {Object} reservation {vehicleId, startDate, startTime, endDate, endTime, userName, purpose, destination}
 * @return {Object} 登録された予約データ
 */
function saveReservation(reservation) {
  if (!reservation.vehicleId || !reservation.startDate || !reservation.startTime ||
      !reservation.endDate || !reservation.endTime || !reservation.userName) {
    throw new Error('車両・利用開始日時・返却日時・利用者名は必須です。');
  }

  const start = new Date(reservation.startDate + 'T' + reservation.startTime + ':00');
  const end = new Date(reservation.endDate + 'T' + reservation.endTime + ':00');
  if (start >= end) {
    throw new Error('返却日時は利用開始日時より後にしてください。');
  }

  const vehicle = getVehicles().find(function (v) { return Number(v.vehicleId) === Number(reservation.vehicleId); });
  if (!vehicle) {
    throw new Error('指定された車両が見つかりません。');
  }

  const isMultiDay = reservation.startDate !== reservation.endDate;
  if (vehicle.availableStart && vehicle.availableEnd) {
    if (isMultiDay) {
      throw new Error(
        'この車両には利用可能時間の制限があるため、日をまたぐ予約はできません。' +
        'マスタ画面で「時間帯を制限しない」に設定した車両を選ぶか、制限を見直してください。'
      );
    }
    if (timeToMinutes_(reservation.startTime) < timeToMinutes_(vehicle.availableStart) ||
        timeToMinutes_(reservation.endTime) > timeToMinutes_(vehicle.availableEnd)) {
      throw new Error('この車両の利用可能時間(' + vehicle.availableStart + '〜' + vehicle.availableEnd + ')の範囲外です。');
    }
  }

  const periods = getUnavailablePeriods(reservation.vehicleId);
  const overlapPeriod = periods.some(function (p) {
    return new Date(p.start) < end && new Date(p.end) > start;
  });
  if (overlapPeriod) {
    throw new Error('この期間は使用不可期間(車検・点検など)と重複しています。');
  }

  const overlapReservations = getReservationsOverlapping_(reservation.vehicleId, start, end);
  if (overlapReservations.length > 0) {
    throw new Error('この期間はすでに他の予約と重複しています。画面を更新して空き状況を再度ご確認ください。');
  }

  const sheet = getSheet_(SHEET_RESERVATION);
  const newId = generateId_(sheet, 1);
  sheet.appendRow([
    newId, reservation.vehicleId, start, end,
    reservation.userName, reservation.purpose || '', reservation.destination || '',
    new Date(), ''
  ]);
  reservation.reservationId = newId;
  return reservation;
}

/**
 * 既存の予約を更新する(予約確認画面からの編集用)。バリデーション内容はsaveReservationと同様だが、
 * 重複チェックでは自分自身の予約は除外する。
 * @param {Object} reservation {reservationId, vehicleId, startDate, startTime, endDate, endTime, userName, purpose, destination}
 */
function updateReservation(reservation) {
  if (!reservation.reservationId) {
    throw new Error('予約IDが指定されていません。');
  }
  if (!reservation.vehicleId || !reservation.startDate || !reservation.startTime ||
      !reservation.endDate || !reservation.endTime || !reservation.userName) {
    throw new Error('車両・利用開始日時・返却日時・利用者名は必須です。');
  }

  const start = new Date(reservation.startDate + 'T' + reservation.startTime + ':00');
  const end = new Date(reservation.endDate + 'T' + reservation.endTime + ':00');
  if (start >= end) {
    throw new Error('返却日時は利用開始日時より後にしてください。');
  }

  const vehicle = getVehicles().find(function (v) { return Number(v.vehicleId) === Number(reservation.vehicleId); });
  if (!vehicle) {
    throw new Error('指定された車両が見つかりません。');
  }

  const isMultiDay = reservation.startDate !== reservation.endDate;
  if (vehicle.availableStart && vehicle.availableEnd) {
    if (isMultiDay) {
      throw new Error('この車両には利用可能時間の制限があるため、日をまたぐ予約はできません。');
    }
    if (timeToMinutes_(reservation.startTime) < timeToMinutes_(vehicle.availableStart) ||
        timeToMinutes_(reservation.endTime) > timeToMinutes_(vehicle.availableEnd)) {
      throw new Error('この車両の利用可能時間(' + vehicle.availableStart + '〜' + vehicle.availableEnd + ')の範囲外です。');
    }
  }

  const periods = getUnavailablePeriods(reservation.vehicleId);
  const overlapPeriod = periods.some(function (p) {
    return new Date(p.start) < end && new Date(p.end) > start;
  });
  if (overlapPeriod) {
    throw new Error('この期間は使用不可期間(車検・点検など)と重複しています。');
  }

  const overlapReservations = getReservationsOverlapping_(reservation.vehicleId, start, end)
    .filter(function (r) { return Number(r.reservationId) !== Number(reservation.reservationId); });
  if (overlapReservations.length > 0) {
    throw new Error('この期間はすでに他の予約と重複しています。画面を更新して空き状況を再度ご確認ください。');
  }

  const sheet = getSheet_(SHEET_RESERVATION);
  const rowIndex = findRowById_(sheet, sheet.getLastRow(), Number(reservation.reservationId));
  if (rowIndex === -1) {
    throw new Error('更新対象の予約が見つかりません。(ID: ' + reservation.reservationId + ')');
  }
  sheet.getRange(rowIndex, 2, 1, 6).setValues([[
    reservation.vehicleId, start, end,
    reservation.userName, reservation.purpose || '', reservation.destination || ''
  ]]);
  return reservation;
}

/**
 * 予約を削除する(予約確認画面からの削除用)
 * @param {number} reservationId
 */
function deleteReservation(reservationId) {
  const sheet = getSheet_(SHEET_RESERVATION);
  const rowIndex = findRowById_(sheet, sheet.getLastRow(), Number(reservationId));
  if (rowIndex === -1) {
    throw new Error('削除対象の予約が見つかりません。(ID: ' + reservationId + ')');
  }
  sheet.deleteRow(rowIndex);
}

// ==================== 繰り返し予約 ====================

/**
 * 指定期間内の日本の祝日を、'yyyy-MM-dd' 文字列のSetとして取得する。
 * Googleが提供する「日本の祝日」公開カレンダーを使うため、追加設定は不要。
 */
function getJapaneseHolidays_(startDate, endDate) {
  const calendar = CalendarApp.getCalendarById('ja.japanese#holiday@group.v.calendar.google.com');
  const events = calendar.getEvents(startDate, endDate);
  const tz = Session.getScriptTimeZone();
  const set = {};
  events.forEach(function (ev) {
    const dateStr = Utilities.formatDate(ev.getStartTime(), tz, 'yyyy-MM-dd');
    set[dateStr] = true;
  });
  return set;
}

/**
 * 繰り返しルールから、実際に予約する日付(Dateオブジェクト、日帰り)の一覧を生成する。
 * 終了日・回数は指定させず、開始日から1年後までを上限として自動生成する。
 * @param {Object} rule {startDate:'yyyy-MM-dd', repeatType:'daily'|'weekly'|'monthly', weekdaysOnly:boolean}
 * @return {Array<Date>}
 */
function buildRecurringDates_(rule) {
  const start = new Date(rule.startDate + 'T00:00:00');
  const limit = new Date(start.getFullYear() + 1, start.getMonth(), start.getDate());
  const dates = [];

  if (rule.repeatType === 'daily') {
    let holidays = {};
    if (rule.weekdaysOnly) {
      holidays = getJapaneseHolidays_(start, limit);
    }
    for (let d = new Date(start); d <= limit; d.setDate(d.getDate() + 1)) {
      const cur = new Date(d);
      if (rule.weekdaysOnly) {
        const dow = cur.getDay();
        const dateStr = Utilities.formatDate(cur, Session.getScriptTimeZone(), 'yyyy-MM-dd');
        if (dow === 0 || dow === 6 || holidays[dateStr]) continue; // 土日祝を除く
      }
      dates.push(cur);
    }
  } else if (rule.repeatType === 'weekly') {
    for (let d = new Date(start); d <= limit; d.setDate(d.getDate() + 7)) {
      dates.push(new Date(d));
    }
  } else if (rule.repeatType === 'monthly') {
    const dayOfMonth = start.getDate();
    let monthCursor = 0;
    while (true) {
      const expectedMonth = ((start.getMonth() + monthCursor) % 12 + 12) % 12;
      const candidate = new Date(start.getFullYear(), start.getMonth() + monthCursor, dayOfMonth);
      if (candidate > limit) break;
      // new Date()は存在しない日(例:2月31日)を翌月にロールオーバーするため、
      // 月がずれていたら「その月には該当日が存在しない」とみなしてスキップする
      if (candidate.getMonth() === expectedMonth) {
        dates.push(candidate);
      }
      monthCursor++;
    }
  } else {
    throw new Error('不正な繰り返し種別です: ' + rule.repeatType);
  }

  return dates;
}

/**
 * 繰り返し予約を登録する。各回ごとに車両の利用可能時間・使用不可期間・既存予約との
 * 重複をチェックし、重複する日はスキップして他の日程のみ登録する。
 * @param {Object} rule {
 *   vehicleId, startDate:'yyyy-MM-dd', startTime:'HH:mm', endTime:'HH:mm',
 *   userName, purpose, destination, repeatType:'daily'|'weekly'|'monthly', weekdaysOnly:boolean
 * }
 * @return {Object} {seriesId, createdCount, skipped: [{date, reason}]}
 */
function saveRecurringReservation(rule) {
  if (!rule.vehicleId || !rule.startDate || !rule.startTime || !rule.endTime || !rule.userName || !rule.repeatType) {
    throw new Error('車両・開始日・時間・利用者名・繰り返しの種類は必須です。');
  }
  if (timeToMinutes_(rule.startTime) >= timeToMinutes_(rule.endTime)) {
    throw new Error('終了時刻は開始時刻より後にしてください。');
  }

  const vehicle = getVehicles().find(function (v) { return Number(v.vehicleId) === Number(rule.vehicleId); });
  if (!vehicle) {
    throw new Error('指定された車両が見つかりません。');
  }
  if (vehicle.availableStart && vehicle.availableEnd) {
    if (timeToMinutes_(rule.startTime) < timeToMinutes_(vehicle.availableStart) ||
        timeToMinutes_(rule.endTime) > timeToMinutes_(vehicle.availableEnd)) {
      throw new Error('この車両の利用可能時間(' + vehicle.availableStart + '〜' + vehicle.availableEnd + ')の範囲外です。');
    }
  }

  const occurrenceDates = buildRecurringDates_(rule);
  if (occurrenceDates.length === 0) {
    throw new Error('条件に合致する日程がありませんでした。');
  }

  const tz = Session.getScriptTimeZone();
  const periods = getUnavailablePeriods(rule.vehicleId);
  const sheet = getSheet_(SHEET_RESERVATION);
  let nextId = generateId_(sheet, 1);
  const seriesId = Utilities.getUuid();
  const now = new Date();

  const rowsToAppend = [];
  const skipped = [];
  // 同じシリーズ内で登録予定の日程同士が重複することはないが、既存データとの重複は都度チェックする
  const alreadyReserved = []; // {start, end} 今回のバッチ内で確定した予約(既存データと合わせて重複判定に使う)

  occurrenceDates.forEach(function (dateObj) {
    const dateStr = Utilities.formatDate(dateObj, tz, 'yyyy-MM-dd');
    const start = new Date(dateStr + 'T' + rule.startTime + ':00');
    const end = new Date(dateStr + 'T' + rule.endTime + ':00');

    const overlapPeriod = periods.some(function (p) {
      return new Date(p.start) < end && new Date(p.end) > start;
    });
    if (overlapPeriod) {
      skipped.push({ date: dateStr, reason: '使用不可期間と重複' });
      return;
    }

    const overlapExisting = getReservationsOverlapping_(rule.vehicleId, start, end).length > 0;
    const overlapBatch = alreadyReserved.some(function (r) { return r.start < end && r.end > start; });
    if (overlapExisting || overlapBatch) {
      skipped.push({ date: dateStr, reason: '既存の予約と重複' });
      return;
    }

    rowsToAppend.push([
      nextId, rule.vehicleId, start, end,
      rule.userName, rule.purpose || '', rule.destination || '',
      now, seriesId
    ]);
    alreadyReserved.push({ start: start, end: end });
    nextId++;
  });

  if (rowsToAppend.length > 0) {
    const startRow = sheet.getLastRow() + 1;
    sheet.getRange(startRow, 1, rowsToAppend.length, RESERVATION_HEADERS.length).setValues(rowsToAppend);
  }

  return {
    seriesId: seriesId,
    createdCount: rowsToAppend.length,
    skipped: skipped
  };
}

/**
 * 繰り返し予約のシリーズをまとめて削除する
 * @param {string} seriesId
 * @return {number} 削除件数
 */
function deleteReservationSeries(seriesId) {
  if (!seriesId) {
    throw new Error('シリーズIDが指定されていません。');
  }
  const sheet = getSheet_(SHEET_RESERVATION);
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return 0;
  const values = sheet.getRange(2, 1, lastRow - 1, RESERVATION_HEADERS.length).getValues();
  let deletedCount = 0;
  for (let i = values.length - 1; i >= 0; i--) {
    if (values[i][8] === seriesId) {
      sheet.deleteRow(i + 2);
      deletedCount++;
    }
  }
  return deletedCount;
}