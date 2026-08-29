const NV11 = Object.freeze({
  TZ: 'Asia/Tokyo',
  SHEETS: { SETTINGS: '基本設定', RATES: '日別料金', RESERVATIONS: '予約一覧' },
  STATUS: { HOLD: '決済待ち', CONFIRMED: '予約確定', EXPIRED: '期限切れ', CANCELLED: 'キャンセル' },
  HOLD_MINUTES: 20
});

function doGet(e) {
  const params = (e && e.parameter) || {};
  if (params.page === 'success') return handleStripeReturn_(params);
  if (params.page === 'cancel') return renderMessage_('決済が完了していません', '予約はまだ確定していません。もう一度予約画面からお手続きください。', 'Payment not completed', 'Your reservation has not been confirmed.');
  return HtmlService.createTemplateFromFile('Booking')
    .evaluate()
    .setTitle('Nagashima Villa 11｜空室確認・予約')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/** Stripe webhook。決済後にゲストが画面を閉じても予約を確定します。 */
function doPost(e) {
  try {
    const props = PropertiesService.getScriptProperties();
    const expectedToken = props.getProperty('STRIPE_WEBHOOK_TOKEN');
    const actualToken = e && e.parameter && e.parameter.stripe_webhook;
    if (!expectedToken || actualToken !== expectedToken) throw new Error('Unauthorized');
    const event = JSON.parse((e.postData && e.postData.contents) || '{}');
    if (event.type === 'checkout.session.completed' && event.data && event.data.object) {
      const session = retrieveStripeSession_(String(event.data.object.id));
      if (session.payment_status === 'paid' && session.metadata && session.metadata.reservation_id) {
        confirmReservation_(String(session.metadata.reservation_id), String(session.id));
      }
    }
    return ContentService.createTextOutput('ok');
  } catch (error) {
    return ContentService.createTextOutput('error').setMimeType(ContentService.MimeType.TEXT);
  }
}

/** 初回のみ実行。必要なシートと自社予約カレンダーを作成します。 */
function setupNagashimaVilla11() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  if (!spreadsheet) throw new Error('予約管理用スプレッドシートから実行してください。');
  spreadsheet.setSpreadsheetTimeZone(NV11.TZ);

  const settings = getOrCreateSheet_(spreadsheet, NV11.SHEETS.SETTINGS);
  const rates = getOrCreateSheet_(spreadsheet, NV11.SHEETS.RATES);
  const reservations = getOrCreateSheet_(spreadsheet, NV11.SHEETS.RESERVATIONS);

  if (settings.getLastRow() === 0) {
    settings.getRange(1, 1, 16, 3).setValues([
      ['設定項目', '値', '説明'],
      ['施設名', 'Nagashima Villa 11', 'メール・決済画面に表示'],
      ['基本料金（5名まで）', 30000, '1泊・税込'],
      ['週末倍率', 1.2, '金・土・日曜日'],
      ['追加人数料金', 3000, '6人目以降・子どもを含む'],
      ['清掃料金', 3000, '1予約につき'],
      ['基本人数', 5, '基本料金に含む人数'],
      ['最大人数', 11, '大人・子どもの合計'],
      ['最低宿泊数', 1, '泊'],
      ['最大宿泊数', 14, '必要に応じて変更'],
      ['予約受付締切時刻', '12:00', '当日予約の締切'],
      ['販売日数', 365, '本日から何日先まで販売するか'],
      ['管理者メール', 'nagashima.villa.11@gmail.com', '予約通知先'],
      ['通貨', 'JPY', 'Stripe決済通貨'],
      ['仮押さえ時間（分）', NV11.HOLD_MINUTES, '決済開始からの保持時間'],
      ['WebサイトURL', 'https://nagashimavilla11.com/', '決済後の案内に使用']
    ]);
    settings.setFrozenRows(1).autoResizeColumns(1, 3);
  }

  if (rates.getLastRow() === 0) {
    rates.getRange(1, 1, 2, 7).setValues([
      ['日付', '上書き料金', '倍率', '販売停止', '最低泊数', '名称・メモ', '入力例'],
      [new Date(), '', '', false, '', '例：年末年始', '上書き料金か倍率のどちらかを入力']
    ]);
    rates.getRange('A:A').setNumberFormat('yyyy/mm/dd');
    rates.getRange('B:B').setNumberFormat('¥#,##0');
    rates.getRange('D2:D').insertCheckboxes();
    rates.setFrozenRows(1).autoResizeColumns(1, 7);
  }

  if (reservations.getLastRow() === 0) {
    reservations.getRange(1, 1, 1, 21).setValues([[
      '予約番号', '状態', '受付日時', 'チェックイン', 'チェックアウト', '泊数', '大人', '子ども', '合計人数',
      '宿泊料金', '追加人数料金', '清掃料金', '合計金額', '氏名', 'フリガナ', 'メール', '電話番号',
      '住所・国', '到着予定時刻', 'Stripe Session ID', 'カレンダーEvent ID'
    ]]);
    reservations.getRange('C:E').setNumberFormat('yyyy/mm/dd hh:mm');
    reservations.getRange('J:M').setNumberFormat('¥#,##0');
    reservations.setFrozenRows(1).autoResizeColumns(1, 21);
  }

  const props = PropertiesService.getScriptProperties();
  props.setProperty('SPREADSHEET_ID', spreadsheet.getId());
  if (!props.getProperty('STRIPE_WEBHOOK_TOKEN')) props.setProperty('STRIPE_WEBHOOK_TOKEN', Utilities.getUuid().replace(/-/g, ''));
  if (!props.getProperty('BOOKING_CALENDAR_ID')) {
    const calendar = CalendarApp.createCalendar('ナガシマヴィラ11｜自社予約・販売停止', { timeZone: NV11.TZ });
    props.setProperty('BOOKING_CALENDAR_ID', calendar.getId());
  }
  installCleanupTrigger_();
  SpreadsheetApp.getUi().alert('初期設定が完了しました。次にスクリプトプロパティを設定してください。');
}

function getInitialData() {
  cleanupExpiredHolds_();
  const settings = getSettings_();
  const today = today_();
  return {
    minDate: formatDate_(today),
    maxDate: formatDate_(addDays_(today, Number(settings['販売日数'] || 365))),
    maxGuests: Number(settings['最大人数'] || 11),
    baseGuests: Number(settings['基本人数'] || 5),
    basePrice: Number(settings['基本料金（5名まで）'] || 30000),
    extraGuestPrice: Number(settings['追加人数料金'] || 3000),
    cleaningFee: Number(settings['清掃料金'] || 3000)
  };
}

function calculateQuote(payload) {
  cleanupExpiredHolds_();
  return calculateQuote_(payload, false);
}

function createCheckout(payload) {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    cleanupExpiredHolds_();
    validateGuestDetails_(payload);
    const quote = calculateQuote_(payload, true);
    if (!quote.available) throw new Error(quote.message || '選択された日程は予約できません。');

    const reservationId = makeReservationId_();
    const now = new Date();
    const settings = getSettings_();
    const holdMinutes = Number(settings['仮押さえ時間（分）'] || NV11.HOLD_MINUTES);
    const calendar = getBookingCalendar_();
    const start = parseDate_(payload.checkIn);
    const end = parseDate_(payload.checkOut);
    const event = calendar.createAllDayEvent('【決済待ち】' + reservationId, start, end, {
      description: 'Stripe決済待ち。期限：' + Utilities.formatDate(new Date(now.getTime() + holdMinutes * 60000), NV11.TZ, 'yyyy/MM/dd HH:mm')
    });

    appendReservation_(reservationId, payload, quote, event.getId(), now);
    try {
      const checkout = createStripeCheckout_(reservationId, payload, quote);
      updateReservationField_(reservationId, 20, checkout.id);
      return { ok: true, reservationId: reservationId, checkoutUrl: checkout.url };
    } catch (error) {
      event.deleteEvent();
      updateReservationStatus_(reservationId, NV11.STATUS.CANCELLED);
      throw error;
    }
  } finally {
    lock.releaseLock();
  }
}

function handleStripeReturn_(params) {
  try {
    const sessionId = String(params.session_id || '');
    const reservationId = String(params.reservation_id || '');
    if (!sessionId || !reservationId) throw new Error('決済情報を確認できません。');
    const session = retrieveStripeSession_(sessionId);
    if (session.payment_status !== 'paid' || !session.metadata || session.metadata.reservation_id !== reservationId) {
      throw new Error('決済が完了していません。');
    }
    confirmReservation_(reservationId, sessionId);
    return renderMessage_('ご予約が確定しました', '予約番号：' + reservationId + '。確認メールをご確認ください。', 'Reservation confirmed', 'Confirmation number: ' + reservationId + '. Please check your email.');
  } catch (error) {
    return renderMessage_('確認処理に失敗しました', error.message + ' 管理者までお問い合わせください。', 'Confirmation error', 'Please contact the property.');
  }
}

function calculateQuote_(payload, finalCheck) {
  const settings = getSettings_();
  const checkIn = parseDate_(payload.checkIn);
  const checkOut = parseDate_(payload.checkOut);
  const today = today_();
  if (!checkIn || !checkOut || checkOut <= checkIn) throw new Error('チェックイン日とチェックアウト日をご確認ください。');
  if (checkIn < today) throw new Error('過去の日付は選択できません。');

  const cutoff = String(settings['予約受付締切時刻'] || '12:00').split(':');
  if (sameDate_(checkIn, today)) {
    const now = new Date();
    const cutoffDate = new Date(now.getFullYear(), now.getMonth(), now.getDate(), Number(cutoff[0] || 12), Number(cutoff[1] || 0));
    if (now >= cutoffDate) throw new Error('当日のご予約受付は12:00で終了します。');
  }

  const nights = dateDiff_(checkIn, checkOut);
  if (nights < Number(settings['最低宿泊数'] || 1)) throw new Error('最低宿泊数を満たしていません。');
  if (nights > Number(settings['最大宿泊数'] || 14)) throw new Error('一度に予約できる最大宿泊数を超えています。');

  const adults = Number(payload.adults || 0);
  const children = Number(payload.children || 0);
  const guests = adults + children;
  if (adults < 1 || guests > Number(settings['最大人数'] || 11)) throw new Error('宿泊人数をご確認ください。');

  const conflict = findConflict_(checkIn, checkOut);
  if (conflict) return { available: false, message: '選択された日程には予約済みまたは販売停止の日があります。別の日程をお選びください。' };

  const overrides = getRateOverrides_();
  const basePrice = Number(settings['基本料金（5名まで）'] || 30000);
  const weekendMultiplier = Number(settings['週末倍率'] || 1.2);
  const breakdown = [];
  let accommodation = 0;
  let requiredMinNights = Number(settings['最低宿泊数'] || 1);
  for (let date = new Date(checkIn); date < checkOut; date = addDays_(date, 1)) {
    const key = formatDate_(date);
    const rule = overrides[key] || {};
    if (rule.closed) return { available: false, message: key + 'は販売停止日です。' };
    requiredMinNights = Math.max(requiredMinNights, Number(rule.minNights || 1));
    const isWeekend = [0, 5, 6].indexOf(date.getDay()) >= 0;
    let price = Math.round(basePrice * (isWeekend ? weekendMultiplier : 1));
    if (rule.multiplier) price = Math.round(basePrice * Number(rule.multiplier));
    if (rule.price) price = Number(rule.price);
    accommodation += price;
    breakdown.push({ date: key, price: price, label: rule.note || (isWeekend ? '週末料金' : '基本料金') });
  }
  if (nights < requiredMinNights) throw new Error('選択された期間は最低' + requiredMinNights + '泊からご予約いただけます。');

  const extraGuests = Math.max(0, guests - Number(settings['基本人数'] || 5));
  const extraGuestFee = extraGuests * Number(settings['追加人数料金'] || 3000) * nights;
  const cleaningFee = Number(settings['清掃料金'] || 3000);
  return {
    available: true, nights: nights, guests: guests, accommodation: accommodation,
    extraGuests: extraGuests, extraGuestFee: extraGuestFee, cleaningFee: cleaningFee,
    total: accommodation + extraGuestFee + cleaningFee, breakdown: breakdown,
    finalChecked: Boolean(finalCheck)
  };
}

function findConflict_(start, end) {
  const events = getBookingCalendar_().getEvents(start, end);
  if (events.length) return true;
  const airbnbUrl = PropertiesService.getScriptProperties().getProperty('AIRBNB_ICAL_URL');
  if (!airbnbUrl) return false;
  try {
    const response = UrlFetchApp.fetch(airbnbUrl, { muteHttpExceptions: true });
    if (response.getResponseCode() < 200 || response.getResponseCode() >= 300) throw new Error('iCal HTTP error');
    const text = response.getContentText();
    if (text.indexOf('BEGIN:VCALENDAR') === -1) throw new Error('Invalid iCal');
    return parseIcalPeriods_(text).some(function(period) { return period.start < end && period.end > start; });
  } catch (error) {
    throw new Error('Airbnbの空室情報を確認できませんでした。時間をおいて再度お試しください。');
  }
}

function parseIcalPeriods_(text) {
  const unfolded = String(text || '').replace(/\r?\n[ \t]/g, '');
  const events = unfolded.split('BEGIN:VEVENT').slice(1);
  return events.map(function(block) {
    const startMatch = block.match(/DTSTART(?:;VALUE=DATE)?:([0-9]{8})/);
    const endMatch = block.match(/DTEND(?:;VALUE=DATE)?:([0-9]{8})/);
    return startMatch && endMatch ? { start: parseIcalDate_(startMatch[1]), end: parseIcalDate_(endMatch[1]) } : null;
  }).filter(Boolean);
}

function createStripeCheckout_(reservationId, payload, quote) {
  const props = PropertiesService.getScriptProperties();
  const secret = props.getProperty('STRIPE_SECRET_KEY');
  const webAppUrl = props.getProperty('WEB_APP_URL');
  if (!secret || !webAppUrl) throw new Error('StripeまたはWebアプリURLが設定されていません。');
  const success = webAppUrl + '?page=success&session_id={CHECKOUT_SESSION_ID}&reservation_id=' + encodeURIComponent(reservationId);
  const cancel = webAppUrl + '?page=cancel&reservation_id=' + encodeURIComponent(reservationId);
  const form = {
    mode: 'payment', success_url: success, cancel_url: cancel,
    customer_email: String(payload.email), locale: 'auto',
    'line_items[0][quantity]': '1',
    'line_items[0][price_data][currency]': 'jpy',
    'line_items[0][price_data][unit_amount]': String(quote.total),
    'line_items[0][price_data][product_data][name]': 'Nagashima Villa 11 宿泊予約',
    'line_items[0][price_data][product_data][description]': payload.checkIn + '〜' + payload.checkOut + '・' + quote.guests + '名・' + quote.nights + '泊',
    'metadata[reservation_id]': reservationId,
    'payment_intent_data[metadata][reservation_id]': reservationId
  };
  const response = UrlFetchApp.fetch('https://api.stripe.com/v1/checkout/sessions', {
    method: 'post', headers: { Authorization: 'Bearer ' + secret }, payload: form, muteHttpExceptions: true
  });
  const data = JSON.parse(response.getContentText());
  if (response.getResponseCode() >= 300 || !data.url) throw new Error((data.error && data.error.message) || '決済画面を作成できませんでした。');
  return data;
}

function retrieveStripeSession_(sessionId) {
  const secret = PropertiesService.getScriptProperties().getProperty('STRIPE_SECRET_KEY');
  const response = UrlFetchApp.fetch('https://api.stripe.com/v1/checkout/sessions/' + encodeURIComponent(sessionId), {
    headers: { Authorization: 'Bearer ' + secret }, muteHttpExceptions: true
  });
  const data = JSON.parse(response.getContentText());
  if (response.getResponseCode() >= 300) throw new Error('Stripe決済を確認できませんでした。');
  return data;
}

function confirmReservation_(reservationId, sessionId) {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const row = findReservationRow_(reservationId);
    if (!row) throw new Error('予約情報が見つかりません。');
    const sheet = getSpreadsheet_().getSheetByName(NV11.SHEETS.RESERVATIONS);
    const values = sheet.getRange(row, 1, 1, 21).getValues()[0];
    if (values[1] === NV11.STATUS.CONFIRMED) return;
    if (String(values[19]) !== sessionId) throw new Error('決済情報が予約と一致しません。');
    sheet.getRange(row, 2).setValue(NV11.STATUS.CONFIRMED);
    const calendar = getBookingCalendar_();
    let event = calendar.getEventById(String(values[20]));
    if (event) {
      event.setTitle('【自社予約】' + reservationId).setDescription('決済完了済み');
    } else {
      event = calendar.createAllDayEvent('【自社予約】' + reservationId, new Date(values[3]), new Date(values[4]), { description: '決済完了済み' });
      sheet.getRange(row, 21).setValue(event.getId());
    }
    sendConfirmationEmails_(values);
  } finally { lock.releaseLock(); }
}

function appendReservation_(id, payload, quote, eventId, now) {
  const sheet = getSpreadsheet_().getSheetByName(NV11.SHEETS.RESERVATIONS);
  sheet.appendRow([
    id, NV11.STATUS.HOLD, now, parseDate_(payload.checkIn), parseDate_(payload.checkOut), quote.nights,
    Number(payload.adults), Number(payload.children), quote.guests, quote.accommodation, quote.extraGuestFee,
    quote.cleaningFee, quote.total, safe_(payload.name), safe_(payload.nameKana), safe_(payload.email), safe_(payload.phone),
    safe_(payload.addressCountry), safe_(payload.arrivalTime), '', eventId
  ]);
}

function sendConfirmationEmails_(values) {
  const settings = getSettings_();
  const id = values[0];
  const bodyJa = values[13] + ' 様\n\nNagashima Villa 11のご予約が確定しました。\n\n予約番号：' + id + '\nチェックイン：' + formatDate_(values[3]) + '\nチェックアウト：' + formatDate_(values[4]) + '\n宿泊人数：' + values[8] + '名\n合計金額：¥' + Number(values[12]).toLocaleString('ja-JP') + '\n\nチェックイン方法はご宿泊日が近づきましたらご案内します。';
  const bodyEn = '\n\nYour reservation is confirmed.\nConfirmation number: ' + id + '\nCheck-in: ' + formatDate_(values[3]) + '\nCheck-out: ' + formatDate_(values[4]) + '\nGuests: ' + values[8] + '\nTotal: JPY ' + Number(values[12]).toLocaleString('en-US');
  MailApp.sendEmail(String(values[15]), '【Nagashima Villa 11】ご予約確定 / Reservation confirmed', bodyJa + bodyEn, { name: 'Nagashima Villa 11', replyTo: String(settings['管理者メール']) });
  MailApp.sendEmail(String(settings['管理者メール']), '【自社予約】' + id + '・' + formatDate_(values[3]) + 'から', '自社予約の決済が完了しました。\n予約番号：' + id + '\n代表者：' + values[13] + '\n人数：' + values[8] + '名\n合計：¥' + Number(values[12]).toLocaleString('ja-JP'));
}

function cleanupExpiredHolds_() {
  const sheet = getSpreadsheet_().getSheetByName(NV11.SHEETS.RESERVATIONS);
  if (!sheet || sheet.getLastRow() < 2) return;
  const settings = getSettings_();
  const limit = Number(settings['仮押さえ時間（分）'] || NV11.HOLD_MINUTES) * 60000;
  const rows = sheet.getRange(2, 1, sheet.getLastRow() - 1, 21).getValues();
  const calendar = getBookingCalendar_();
  rows.forEach(function(row, index) {
    if (row[1] === NV11.STATUS.HOLD && new Date() - new Date(row[2]) > limit) {
      sheet.getRange(index + 2, 2).setValue(NV11.STATUS.EXPIRED);
      const event = calendar.getEventById(String(row[20] || ''));
      if (event) event.deleteEvent();
    }
  });
}

function installCleanupTrigger_() {
  ScriptApp.getProjectTriggers().filter(function(t) { return t.getHandlerFunction() === 'cleanupExpiredHolds_'; }).forEach(ScriptApp.deleteTrigger);
  ScriptApp.newTrigger('cleanupExpiredHolds_').timeBased().everyMinutes(15).create();
}

function getRateOverrides_() {
  const sheet = getSpreadsheet_().getSheetByName(NV11.SHEETS.RATES);
  if (!sheet || sheet.getLastRow() < 2) return {};
  const result = {};
  sheet.getRange(2, 1, sheet.getLastRow() - 1, 6).getValues().forEach(function(row) {
    if (!row[0]) return;
    result[formatDate_(row[0])] = { price: Number(row[1]) || 0, multiplier: Number(row[2]) || 0, closed: row[3] === true, minNights: Number(row[4]) || 0, note: String(row[5] || '') };
  });
  return result;
}

function getSettings_() {
  const sheet = getSpreadsheet_().getSheetByName(NV11.SHEETS.SETTINGS);
  const values = sheet.getRange(2, 1, Math.max(sheet.getLastRow() - 1, 1), 2).getValues();
  return values.reduce(function(result, row) { if (row[0]) result[String(row[0])] = row[1]; return result; }, {});
}

function validateGuestDetails_(payload) {
  ['name', 'email', 'phone', 'addressCountry'].forEach(function(key) { if (!String(payload[key] || '').trim()) throw new Error('必須項目を入力してください。'); });
  if (!/^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(String(payload.email))) throw new Error('メールアドレスをご確認ください。');
  if (!payload.agree) throw new Error('キャンセルポリシーと利用規則への同意が必要です。');
}

function renderMessage_(jaTitle, jaBody, enTitle, enBody) {
  const template = HtmlService.createTemplateFromFile('Result');
  template.jaTitle = jaTitle; template.jaBody = jaBody; template.enTitle = enTitle; template.enBody = enBody;
  return template.evaluate().setTitle(jaTitle).setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL).addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function getSpreadsheet_() { return SpreadsheetApp.openById(PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID')); }
function getBookingCalendar_() { return CalendarApp.getCalendarById(PropertiesService.getScriptProperties().getProperty('BOOKING_CALENDAR_ID')); }
function getOrCreateSheet_(ss, name) { return ss.getSheetByName(name) || ss.insertSheet(name); }
function updateReservationStatus_(id, status) { const row = findReservationRow_(id); if (row) getSpreadsheet_().getSheetByName(NV11.SHEETS.RESERVATIONS).getRange(row, 2).setValue(status); }
function updateReservationField_(id, column, value) { const row = findReservationRow_(id); if (row) getSpreadsheet_().getSheetByName(NV11.SHEETS.RESERVATIONS).getRange(row, column).setValue(value); }
function findReservationRow_(id) { const sheet = getSpreadsheet_().getSheetByName(NV11.SHEETS.RESERVATIONS); const finder = sheet.getRange('A:A').createTextFinder(String(id)).matchEntireCell(true).findNext(); return finder ? finder.getRow() : null; }
function makeReservationId_() { return 'NV11-' + Utilities.formatDate(new Date(), NV11.TZ, 'yyyyMMdd') + '-' + Utilities.getUuid().slice(0, 6).toUpperCase(); }
function safe_(value) { const text = String(value || '').trim(); return /^[=+\-@]/.test(text) ? "'" + text : text; }
function today_() { const now = new Date(); return new Date(now.getFullYear(), now.getMonth(), now.getDate()); }
function parseDate_(value) { const match = String(value || '').match(/^(\d{4})-(\d{2})-(\d{2})$/); return match ? new Date(Number(match[1]), Number(match[2]) - 1, Number(match[3])) : null; }
function parseIcalDate_(value) { return new Date(Number(value.slice(0, 4)), Number(value.slice(4, 6)) - 1, Number(value.slice(6, 8))); }
function formatDate_(date) { return Utilities.formatDate(new Date(date), NV11.TZ, 'yyyy-MM-dd'); }
function addDays_(date, days) { const result = new Date(date); result.setDate(result.getDate() + days); return result; }
function dateDiff_(start, end) { return Math.round((Date.UTC(end.getFullYear(), end.getMonth(), end.getDate()) - Date.UTC(start.getFullYear(), start.getMonth(), start.getDate())) / 86400000); }
function sameDate_(a, b) { return a.getFullYear() === b.getFullYear() && a.getMonth() === b.getMonth() && a.getDate() === b.getDate(); }
