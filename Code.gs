// =====================================================================
// Учёт трат: многопользовательская версия
// Главная таблица (к которой привязан скрипт) хранит реестр:
//   лист _Users   - кто зарегистрирован и где лежит его таблица
//   лист _Invites - инвайт-коды
// У каждого пользователя своя Google-таблица с листами «Траты» и «Архив».
// =====================================================================

// Категории: единственный список на всё приложение, порядок = порядок в выпадающем списке.
// Добавлять можно смело. Старые названия не меняй напрямую: переименование делается через ALIASES,
// иначе старые траты пропадут из итогов
const ALLOWED = ['Продукты', 'Кафе и доставка', 'Транспорт', 'Жилье', 'Шопинг', 'Спорт',
  'Подписки', 'Развлечения', 'Учеба', 'Прочее'];
// Старое название -> новое. Старые траты со старым названием считаются в новой категории
const ALIASES = { 'Еда': 'Продукты' };
function canon_(c) {
  c = String(c);
  return ALIASES.hasOwnProperty(c) ? ALIASES[c] : c;
}
const INVITE_DAYS = 7;          // сколько дней живёт инвайт-код
const MAX_FAILS_PER_HOUR = 5;   // попыток ввести неверный код в час

function props_() { return PropertiesService.getScriptProperties(); }
function registry_() { return SpreadsheetApp.getActiveSpreadsheet(); }
function isOwner_(telegramId) { return String(telegramId) === props_().getProperty('OWNER_ID'); }

// ---------------------------------------------------------------------
// Проверка подписи Telegram: возвращает данные пользователя или null
// ---------------------------------------------------------------------
function verify_(initData) {
  const token = props_().getProperty('BOT_TOKEN');
  const params = {};
  initData.split('&').forEach(function (pair) {
    const i = pair.indexOf('=');
    params[decodeURIComponent(pair.slice(0, i))] = decodeURIComponent(pair.slice(i + 1));
  });
  const hash = params.hash;
  delete params.hash;
  const check = Object.keys(params).sort()
    .map(function (k) { return k + '=' + params[k]; }).join('\n');
  const secret = Utilities.computeHmacSha256Signature(token, 'WebAppData');
  const sig = Utilities.computeHmacSha256Signature(Utilities.newBlob(check).getBytes(), secret);
  const hex = sig.map(function (b) { return ('0' + (b & 0xff).toString(16)).slice(-2); }).join('');
  if (hex !== hash) return null;
  if (Date.now() / 1000 - Number(params.auth_date) > 86400) return null;
  return JSON.parse(params.user);
}

function out_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

// ---------------------------------------------------------------------
// Реестр пользователей
// ---------------------------------------------------------------------
// Колонки _Users: telegram_id | имя | spreadsheet_id | статус | дата регистрации
function usersSheet_() {
  const ss = registry_();
  let sh = ss.getSheetByName('_Users');
  if (!sh) {
    sh = ss.insertSheet('_Users');
    sh.getRange('A:A').setNumberFormat('@');
    sh.getRange('C:C').setNumberFormat('@');
    sh.appendRow(['telegram_id', 'имя', 'spreadsheet_id', 'статус', 'дата регистрации']);
  }
  return sh;
}

// Имя для списка: «Имя Фамилия (@ник)», если ника нет, то без него
// Есть ли в строке видимые символы (имена из пробелов и невидимых знаков не считаются)
function visible_(s) {
  return String(s || '').replace(/[\p{Z}\p{C}ㅤᅟᅠ⠀]/gu, '').length > 0;
}

function displayName_(tgUser) {
  const base = [tgUser.first_name, tgUser.last_name].filter(visible_).join(' ').trim() ||
    tgUser.username || String(tgUser.id);
  return tgUser.username && base !== tgUser.username
    ? base + ' (@' + tgUser.username + ')' : base;
}

// Колонки _Invites: код | создан | истекает | использован кем | когда
function invitesSheet_() {
  const ss = registry_();
  let sh = ss.getSheetByName('_Invites');
  if (!sh) {
    sh = ss.insertSheet('_Invites');
    sh.getRange('A:A').setNumberFormat('@');
    sh.getRange('D:D').setNumberFormat('@');
    sh.appendRow(['код', 'создан', 'истекает', 'использован кем', 'когда']);
  }
  return sh;
}

function allUsers_() {
  const rows = usersSheet_().getDataRange().getValues();
  const res = [];
  for (let i = 1; i < rows.length; i++) {
    if (!rows[i][0]) continue;
    res.push({
      row: i + 1,
      id: String(rows[i][0]),
      name: String(rows[i][1]),
      sheetId: String(rows[i][2]),
      status: String(rows[i][3]),
      since: rows[i][4]
    });
  }
  return res;
}

function getUser_(telegramId) {
  const id = String(telegramId);
  const users = allUsers_();
  for (let i = 0; i < users.length; i++) {
    if (users[i].id === id) return users[i];
  }
  return null;
}

// Открывает таблицу пользователя. Бросает ошибку, если доступа нет
function userSs_(user) { return SpreadsheetApp.openById(user.sheetId); }

// Создаёт в таблице нужные листы, если их нет
function ensureSheets_(ss) {
  let tr = ss.getSheetByName('Траты');
  if (!tr) {
    tr = ss.insertSheet('Траты');
    tr.appendRow(['Дата', 'Сумма', 'Категория', 'Комментарий']);
    tr.getRange('A:A').setNumberFormat('dd.MM.yyyy');
  }
  let arc = ss.getSheetByName('Архив');
  if (!arc) {
    arc = ss.insertSheet('Архив');
    arc.getRange('B:B').setNumberFormat('@'); // дата хранится текстом
    arc.appendRow(['Тип', 'Начало', 'Категория', 'Сумма']);
  }
}

// Папка «Траты: пользователи» на Диске владельца (ID запоминаем в свойствах)
function usersFolder_() {
  const saved = props_().getProperty('FOLDER_ID');
  if (saved) {
    try { return DriveApp.getFolderById(saved); } catch (err) { /* папку удалили: создадим заново */ }
  }
  const folder = DriveApp.createFolder('Траты: пользователи');
  props_().setProperty('FOLDER_ID', folder.getId());
  return folder;
}

// Создаёт новую таблицу пользователя с листами «Траты» и «Архив»
function createUserSheet_(name) {
  const ss = SpreadsheetApp.create('Траты - ' + String(name).slice(0, 60));
  const first = ss.getSheets()[0];
  first.setName('Траты');
  first.appendRow(['Дата', 'Сумма', 'Категория', 'Комментарий']);
  first.getRange('A:A').setNumberFormat('dd.MM.yyyy');
  ensureSheets_(ss); // добавит «Архив»
  DriveApp.getFileById(ss.getId()).moveTo(usersFolder_());
  return ss;
}

// ---------------------------------------------------------------------
// Запуск один раз вручную: создаёт реестр и записывает владельца
// первым пользователем (его таблица = текущая главная)
// ---------------------------------------------------------------------
function setupMulti() {
  const ownerId = props_().getProperty('OWNER_ID');
  usersSheet_();
  invitesSheet_();
  if (!getUser_(ownerId)) {
    usersSheet_().appendRow([ownerId, 'Владелец', registry_().getId(), 'active', new Date()]);
  }
}

// ---------------------------------------------------------------------
// Точка входа
// ---------------------------------------------------------------------
function doPost(e) {
  try {
    const body = JSON.parse(e.postData.contents);
    const tgUser = verify_(body.initData || '');
    if (!tgUser) return out_({ ok: false, error: 'denied' });

    const action = body.action || 'add';
    const user = getUser_(tgUser.id);

    // Эти действия доступны и незарегистрированным
    if (action === 'whoami') {
      // Если в реестре пустое имя, заполняем его данными из Telegram
      if (user && !visible_(user.name)) {
        try {
          user.name = displayName_(tgUser);
          usersSheet_().getRange(user.row, 2).setValue(user.name);
        } catch (err) { console.error('Не удалось обновить имя'); }
      }
      return out_({
        ok: true,
        registered: !!user && user.status === 'active',
        disabled: !!user && user.status !== 'active',
        isAdmin: isOwner_(tgUser.id),
        name: user ? user.name : '',
        categories: ALLOWED
      });
    }
    if (action === 'register') return out_(register_(tgUser, body));

    if (!user) return out_({ ok: false, error: 'not_registered' });
    if (user.status !== 'active') return out_({ ok: false, error: 'disabled' });

    // Админ-действия
    if (action === 'createInvite' || action === 'listUsers' || action === 'setUserStatus') {
      if (!isOwner_(tgUser.id)) return out_({ ok: false, error: 'denied' });
      if (action === 'createInvite') return out_(createInvite_());
      if (action === 'listUsers') return out_(listUsers_());
      return out_(setUserStatus_(body));
    }

    // Обычные действия работают с таблицей самого пользователя
    let ss;
    try {
      ss = userSs_(user);
    } catch (err) {
      return out_({ ok: false, error: 'no_access_to_sheet' });
    }
    if (action === 'summary') return out_(summary_(ss));
    if (action === 'archive') return out_(archiveList_(ss));
    return out_(add_(ss, body));
  } catch (err) {
    console.error(err && err.message ? err.message : 'server error');
    return out_({ ok: false, error: 'server error' });
  }
}

// ---------------------------------------------------------------------
// Регистрация по инвайт-коду
// ---------------------------------------------------------------------
function register_(tgUser, body) {
  const cache = CacheService.getScriptCache();
  const failKey = 'fails_' + tgUser.id;
  const fails = Number(cache.get(failKey) || 0);
  if (fails >= MAX_FAILS_PER_HOUR) return { ok: false, error: 'too_many_attempts' };

  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    if (getUser_(tgUser.id)) return { ok: false, error: 'already_registered' };

    const code = String(body.code || '').trim().toUpperCase();
    const invites = invitesSheet_();
    const rows = invites.getDataRange().getValues();
    let inviteRow = -1;
    for (let i = 1; i < rows.length; i++) {
      if (String(rows[i][0]).toUpperCase() === code && code !== '') { inviteRow = i + 1; break; }
    }
    const bad = inviteRow === -1 ||
      rows[inviteRow - 1][3] !== '' ||
      new Date(rows[inviteRow - 1][2]).getTime() < Date.now();
    if (bad) {
      cache.put(failKey, String(fails + 1), 3600);
      return { ok: false, error: 'bad_invite' };
    }

    // Необязательная почта для доступа на просмотр. Проверяем до создания таблицы,
    // чтобы опечатка не сжигала инвайт
    const email = String(body.email || '').trim();
    if (email && !/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email)) return { ok: false, error: 'bad_email' };

    const name = displayName_(tgUser);

    // Таблицу создаёт скрипт: она лежит на Диске владельца в отдельной папке
    let ss;
    try {
      ss = createUserSheet_(name);
    } catch (err) {
      console.error('Не удалось создать таблицу: ' + (err && err.message ? err.message : 'unknown'));
      return { ok: false, error: 'create_failed' };
    }

    usersSheet_().appendRow([String(tgUser.id), name, ss.getId(), 'active', new Date()]);
    invites.getRange(inviteRow, 4, 1, 2).setValues([[String(tgUser.id), new Date()]]);
    cache.remove(failKey);

    // Доступ только на просмотр. Если не получилось, регистрация всё равно успешна
    let shared = false;
    if (email) {
      try {
        DriveApp.getFileById(ss.getId()).addViewer(email);
        shared = true;
      } catch (err) {
        console.error('Не удалось выдать доступ на просмотр');
      }
    }
    return { ok: true, emailGiven: !!email, shared: shared };
  } finally {
    lock.releaseLock();
  }
}

// ---------------------------------------------------------------------
// Админ-функции
// ---------------------------------------------------------------------
function createInvite_() {
  // Без похожих символов (0/O, 1/I)
  const alphabet = 'ABCDEFGHJKLMNPQRSTUVWXYZ23456789';
  let code = '';
  for (let i = 0; i < 10; i++) code += alphabet.charAt(Math.floor(Math.random() * alphabet.length));
  const now = new Date();
  const exp = new Date(now.getTime() + INVITE_DAYS * 24 * 3600 * 1000);
  invitesSheet_().appendRow([code, now, exp, '', '']);
  return { ok: true, code: code, days: INVITE_DAYS };
}

function listUsers_() {
  return {
    ok: true,
    users: allUsers_().map(function (u) {
      // Если имя в листе _Users пустое, показываем хотя бы Telegram ID
      return { id: u.id, name: visible_(u.name) ? u.name : ('ID ' + u.id), status: u.status, owner: isOwner_(u.id) };
    })
  };
}

function setUserStatus_(body) {
  const id = String(body.userId || '');
  const status = body.status === 'disabled' ? 'disabled' : 'active';
  if (isOwner_(id)) return { ok: false, error: 'cannot_change_owner' };
  const u = getUser_(id);
  if (!u) return { ok: false, error: 'no_such_user' };
  usersSheet_().getRange(u.row, 4).setValue(status);
  return { ok: true };
}

// ---------------------------------------------------------------------
// Работа с тратами (всё принимает таблицу пользователя ss)
// ---------------------------------------------------------------------
function add_(ss, body) {
  const amount = Number(String(body.amount).replace(',', '.'));
  // Старое название (например, из закэшированного приложения) сохраняем под новым
  body.category = canon_(body.category);
  if (!(amount > 0) || ALLOWED.indexOf(body.category) === -1) {
    return { ok: false, error: 'bad data' };
  }
  let date = new Date();
  if (body.date) {
    const p = String(body.date).split('-');
    date = new Date(Number(p[0]), Number(p[1]) - 1, Number(p[2]));
  }
  const comment = String(body.comment || '').slice(0, 200).replace(/^[=+\-@]/, "'$&");
  ensureSheets_(ss);
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    ss.getSheetByName('Траты').appendRow([date, amount, body.category, comment]);
  } finally {
    lock.releaseLock();
  }
  return { ok: true };
}

// Сумма по категориям за период: от from (включительно) до to (не включая)
function sumByCategory_(ss, from, to) {
  const sh = ss.getSheetByName('Траты');
  const res = {};
  ALLOWED.forEach(function (c) { res[c] = 0; });
  if (!sh) return res;
  const rows = sh.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    const d = rows[i][0], a = Number(rows[i][1]), c = canon_(rows[i][2]);
    if (!(d instanceof Date) || !(a > 0) || !res.hasOwnProperty(c)) continue;
    if (d >= from && d < to) res[c] += a;
  }
  return res;
}

function weekStart_(now) {
  const d = new Date(now.getFullYear(), now.getMonth(), now.getDate());
  d.setDate(d.getDate() - (d.getDay() + 6) % 7); // понедельник
  return d;
}

function summary_(ss) {
  const now = new Date();
  const ws = weekStart_(now);
  const ms = new Date(now.getFullYear(), now.getMonth(), 1);
  return {
    ok: true,
    week: sumByCategory_(ss, ws, new Date(ws.getFullYear(), ws.getMonth(), ws.getDate() + 7)),
    month: sumByCategory_(ss, ms, new Date(now.getFullYear(), now.getMonth() + 1, 1))
  };
}

// Запись итогов периода в лист «Архив». Повторный запуск не создаёт дубли
function archive_(ss, type, from, to) {
  ensureSheets_(ss);
  const sh = ss.getSheetByName('Архив');
  const key = Utilities.formatDate(from, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  const exists = sh.getDataRange().getValues().some(function (r) {
    return r[0] === type && String(r[1]) === key;
  });
  if (exists) return;
  const sums = sumByCategory_(ss, from, to);
  ALLOWED.forEach(function (c) { sh.appendRow([type, key, c, sums[c]]); });
}

function archiveList_(ss) {
  const sh = ss.getSheetByName('Архив');
  if (!sh) return { ok: true, periods: [] };
  const rows = sh.getDataRange().getValues().slice(1);
  const map = {}, order = [];
  rows.forEach(function (r) {
    const k = r[0] + '|' + r[1];
    if (!map[k]) { map[k] = { type: r[0], start: String(r[1]), sums: {} }; order.push(k); }
    const c = canon_(r[2]);
    map[k].sums[c] = (map[k].sums[c] || 0) + Number(r[3]);
  });
  return { ok: true, periods: order.reverse().slice(0, 12).map(function (k) { return map[k]; }) };
}

// ---------------------------------------------------------------------
// Задачи по расписанию: перебирают всех активных пользователей.
// Ошибка у одного (например, отозван доступ) не останавливает остальных
// ---------------------------------------------------------------------
function forEachActiveUser_(fn) {
  allUsers_().forEach(function (u) {
    if (u.status !== 'active') return;
    try {
      fn(u, userSs_(u));
    } catch (err) {
      console.error('Ошибка у пользователя ' + u.id + ': ' + (err && err.message ? err.message : 'unknown'));
    }
  });
}

function archiveWeek() {
  const ws = weekStart_(new Date());
  const from = new Date(ws.getFullYear(), ws.getMonth(), ws.getDate() - 7);
  forEachActiveUser_(function (u, ss) { archive_(ss, 'Неделя', from, ws); });
}

function archiveMonth() {
  const n = new Date();
  const from = new Date(n.getFullYear(), n.getMonth() - 1, 1);
  const to = new Date(n.getFullYear(), n.getMonth(), 1);
  forEachActiveUser_(function (u, ss) { archive_(ss, 'Месяц', from, to); });
}

// Запусти один раз вручную: создаёт расписание архива
function setupTriggers() {
  ScriptApp.newTrigger('archiveWeek').timeBased().onWeekDay(ScriptApp.WeekDay.MONDAY).atHour(8).create();
  ScriptApp.newTrigger('archiveMonth').timeBased().onMonthDay(1).atHour(8).create();
}

// ---------------------------------------------------------------------
// Отчёты в Telegram
// ---------------------------------------------------------------------
function fmt_(n) { return (Math.round(n * 100) / 100).toString().replace('.', ','); }

function total_(o) {
  let s = 0;
  for (const k in o) s += o[k];
  return s;
}

function dayLabel_(d) { return Utilities.formatDate(d, Session.getScriptTimeZone(), 'dd.MM'); }

// Собирает текст отчёта: cur и prev это суммы по категориям
function report_(title, cur, prev, prevName) {
  const t = total_(cur), p = total_(prev);
  if (t === 0) return title + '\n\nТрат за период нет.';
  let text = title + '\n\nВсего: ' + fmt_(t) + '\n';
  let top = ALLOWED[0];
  ALLOWED.forEach(function (c) {
    if (cur[c] > 0) text += c + ': ' + fmt_(cur[c]) + '\n';
    if (cur[c] > cur[top]) top = c;
  });
  text += '\nБольше всего ушло на: ' + top + ' (' + fmt_(cur[top]) + ').';
  if (p > 0) {
    const diff = t - p;
    const pct = Math.round(Math.abs(diff) / p * 100);
    text += '\nПо сравнению с ' + prevName + ' (' + fmt_(p) + '): ' +
      (diff > 0 ? 'больше на ' : 'меньше на ') + fmt_(Math.abs(diff)) + ' (' + pct + '%).';
  } else {
    text += '\nЗа ' + prevName + ' трат нет, сравнить не с чем.';
  }
  return text;
}

// chatId - Telegram ID получателя. Если человек не нажимал Start у бота, Telegram вернёт 403
function sendMessage_(chatId, text) {
  try {
    const resp = UrlFetchApp.fetch(
      'https://api.telegram.org/bot' + props_().getProperty('BOT_TOKEN') + '/sendMessage', {
        method: 'post',
        contentType: 'application/json',
        payload: JSON.stringify({ chat_id: chatId, text: text }),
        muteHttpExceptions: true
      });
    if (resp.getResponseCode() !== 200) {
      console.error('Telegram вернул код ' + resp.getResponseCode() + ' для ' + chatId);
    }
  } catch (err) {
    console.error('Не удалось отправить сообщение'); // без деталей, чтобы токен не попал в журнал
  }
}

function sendWeeklyReport() {
  const ws = weekStart_(new Date());
  const from = new Date(ws.getFullYear(), ws.getMonth(), ws.getDate() - 7);
  const before = new Date(ws.getFullYear(), ws.getMonth(), ws.getDate() - 14);
  const last = new Date(ws.getFullYear(), ws.getMonth(), ws.getDate() - 1);
  const title = 'Итоги недели ' + dayLabel_(from) + ' - ' + dayLabel_(last);
  forEachActiveUser_(function (u, ss) {
    sendMessage_(u.id, report_(title,
      sumByCategory_(ss, from, ws), sumByCategory_(ss, before, from), 'позапрошлой неделей'));
  });
}

function sendMonthlyReport() {
  const n = new Date();
  const from = new Date(n.getFullYear(), n.getMonth() - 1, 1);
  const to = new Date(n.getFullYear(), n.getMonth(), 1);
  const before = new Date(n.getFullYear(), n.getMonth() - 2, 1);
  const title = 'Итоги месяца ' + Utilities.formatDate(from, Session.getScriptTimeZone(), 'MM.yyyy');
  forEachActiveUser_(function (u, ss) {
    sendMessage_(u.id, report_(title,
      sumByCategory_(ss, from, to), sumByCategory_(ss, before, from), 'предыдущим месяцем'));
  });
}

// Запусти один раз вручную
function setupReportTriggers() {
  ScriptApp.newTrigger('sendWeeklyReport').timeBased().onWeekDay(ScriptApp.WeekDay.MONDAY).atHour(9).create();
  ScriptApp.newTrigger('sendMonthlyReport').timeBased().onMonthDay(1).atHour(9).create();
}
