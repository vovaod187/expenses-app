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
    sh.appendRow(['код', 'создан', 'истекает', 'использован кем', 'когда', 'тип', 'таблица', 'пригласил']);
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
      since: rows[i][4],
      reminder: String(rows[i][5] || '') !== 'off'
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
    tr.appendRow(['Дата', 'Сумма', 'Категория', 'Комментарий', 'Кто']);
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
  first.appendRow(['Дата', 'Сумма', 'Категория', 'Комментарий', 'Кто']);
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
        categories: ALLOWED,
        // Версия сервера: приложение включает новые функции, только если сервер их знает
        // (старый сервер принял бы незнакомое действие за запись новой траты)
        version: 4,
        reminder: user ? user.reminder : true,
        partner: user && user.status === 'active' ? partnerNames_(user) : ''
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
    if (action === 'history') return out_(history_(ss));
    if (action === 'update') return out_(update_(ss, body, user));
    if (action === 'partnerInvite') return out_(createPartnerInvite_(user));
    if (action === 'setReminder') return out_(setReminder_(user, body));
    if (action === 'joinPartner') return out_(joinPartner_(tgUser, user, body));
    if (action === 'leavePartner') return out_(leavePartner_(user));
    if (action === 'delete') return out_(delete_(ss, body));
    if (action === 'setLimit') return out_(setLimit_(ss, body));
    if (action === 'setBudget') return out_(setBudget_(ss, body));
    return out_(add_(ss, body, user));
  } catch (err) {
    console.error(err && err.message ? err.message : 'server error');
    return out_({ ok: false, error: 'server error' });
  }
}

// ---------------------------------------------------------------------
// Регистрация по инвайт-коду
// ---------------------------------------------------------------------
// Ищет действующий код в _Invites. Неверный код увеличивает счётчик попыток (не больше 5 в час).
// Колонки: код | создан | истекает | использован кем | когда | тип | таблица | пригласил
function findInvite_(tgUser, codeRaw) {
  const cache = CacheService.getScriptCache();
  const failKey = 'fails_' + tgUser.id;
  const fails = Number(cache.get(failKey) || 0);
  if (fails >= MAX_FAILS_PER_HOUR) return { error: 'too_many_attempts' };
  const code = String(codeRaw || '').trim().toUpperCase();
  const invites = invitesSheet_();
  const rows = invites.getDataRange().getValues();
  for (let i = 1; i < rows.length && code !== ''; i++) {
    const r = rows[i];
    if (String(r[0]).toUpperCase() !== code) continue;
    if (r[3] !== '' || new Date(r[2]).getTime() < Date.now()) break;
    return {
      row: i + 1, sheet: invites, failKey: failKey,
      type: String(r[5] || ''), sheetId: String(r[6] || ''), inviterId: String(r[7] || '')
    };
  }
  cache.put(failKey, String(fails + 1), 3600);
  return { error: 'bad_invite' };
}

function useInvite_(inv, telegramId) {
  inv.sheet.getRange(inv.row, 4, 1, 2).setValues([[String(telegramId), new Date()]]);
  CacheService.getScriptCache().remove(inv.failKey);
}

function activeMembersOf_(sheetId) {
  return allUsers_().filter(function (u) { return u.status === 'active' && u.sheetId === sheetId; });
}

// Старые траты без отметки «Кто» записывал тот, кто вёл таблицу один, подписываем его именем
function backfillWho_(ss, name) {
  const sh = ss.getSheetByName('Траты');
  if (!sh || sh.getLastRow() < 2) return;
  if (sh.getRange(1, 5).getValue() === '') sh.getRange(1, 5).setValue('Кто');
  const n = sh.getLastRow() - 1;
  const range = sh.getRange(2, 5, n, 1);
  const vals = range.getValues();
  let changed = false;
  vals.forEach(function (v) { if (v[0] === '') { v[0] = name; changed = true; } });
  if (changed) range.setValues(vals);
}

function register_(tgUser, body) {
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    if (getUser_(tgUser.id)) return { ok: false, error: 'already_registered' };
    const inv = findInvite_(tgUser, body.code);
    if (inv.error) return { ok: false, error: inv.error };

    // Необязательная почта для доступа на просмотр. Проверяем до создания таблицы,
    // чтобы опечатка не сжигала инвайт
    const email = String(body.email || '').trim();
    if (email && !/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email)) return { ok: false, error: 'bad_email' };

    const name = displayName_(tgUser);
    let ss, partnerName = '';
    if (inv.type === 'partner') {
      // Код партнёра: подключаемся к таблице пригласившего, новую таблицу не создаём
      const members = activeMembersOf_(inv.sheetId);
      if (members.length >= 2) return { ok: false, error: 'partner_limit' };
      try { ss = SpreadsheetApp.openById(inv.sheetId); } catch (err) { return { ok: false, error: 'no_access_to_sheet' }; }
      const inviter = getUser_(inv.inviterId);
      partnerName = inviter ? inviter.name : '';
      if (partnerName) backfillWho_(ss, partnerName);
    } else {
      // Таблицу создаёт скрипт: она лежит на Диске владельца в отдельной папке
      try {
        ss = createUserSheet_(name);
      } catch (err) {
        console.error('Не удалось создать таблицу: ' + (err && err.message ? err.message : 'unknown'));
        return { ok: false, error: 'create_failed' };
      }
    }

    usersSheet_().appendRow([String(tgUser.id), name, ss.getId(), 'active', new Date()]);
    useInvite_(inv, tgUser.id);

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
    return { ok: true, emailGiven: !!email, shared: shared, partner: partnerName };
  } finally {
    lock.releaseLock();
  }
}

// Уже зарегистрированный пользователь вводит код партнёра и переходит в его таблицу.
// Его прежняя таблица остаётся на Диске, но в приложении больше не показывается
function joinPartner_(tgUser, user, body) {
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    const inv = findInvite_(tgUser, body.code);
    if (inv.error) return { ok: false, error: inv.error };
    if (inv.type !== 'partner' || inv.sheetId === user.sheetId) return { ok: false, error: 'bad_invite' };
    if (activeMembersOf_(user.sheetId).length > 1) return { ok: false, error: 'already_shared' };
    if (activeMembersOf_(inv.sheetId).length >= 2) return { ok: false, error: 'partner_limit' };
    let ss;
    try { ss = SpreadsheetApp.openById(inv.sheetId); } catch (err) { return { ok: false, error: 'no_access_to_sheet' }; }
    const inviter = getUser_(inv.inviterId);
    if (inviter) backfillWho_(ss, inviter.name);
    usersSheet_().getRange(user.row, 3).setValue(inv.sheetId);
    useInvite_(inv, tgUser.id);
    return { ok: true, partner: inviter ? inviter.name : '' };
  } finally {
    lock.releaseLock();
  }
}

// Код, по которому партнёр подключится к таблице пользователя (общий бюджет максимум на двоих)
function createPartnerInvite_(user) {
  if (activeMembersOf_(user.sheetId).length >= 2) return { ok: false, error: 'partner_limit' };
  const code = newCode_();
  const now = new Date();
  const exp = new Date(now.getTime() + INVITE_DAYS * 24 * 3600 * 1000);
  const sh = invitesSheet_();
  if (sh.getRange(1, 6).getValue() === '') sh.getRange(1, 6, 1, 3).setValues([['тип', 'таблица', 'пригласил']]);
  sh.appendRow([code, now, exp, '', '', 'partner', user.sheetId, user.id]);
  return { ok: true, code: code, days: INVITE_DAYS };
}

// Выход из общего бюджета: уходящий получает новую пустую таблицу, общие траты остаются у партнёра
function leavePartner_(user) {
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    if (activeMembersOf_(user.sheetId).length < 2) return { ok: false, error: 'not_shared' };
    let ss;
    try { ss = createUserSheet_(user.name); } catch (err) { return { ok: false, error: 'create_failed' }; }
    usersSheet_().getRange(user.row, 3).setValue(ss.getId());
    return { ok: true };
  } finally {
    lock.releaseLock();
  }
}

// Имена остальных участников общей таблицы
function partnerNames_(user) {
  return activeMembersOf_(user.sheetId)
    .filter(function (u) { return u.id !== user.id; })
    .map(function (u) { return visible_(u.name) ? u.name : 'ID ' + u.id; }).join(', ');
}

// ---------------------------------------------------------------------
// Админ-функции
// ---------------------------------------------------------------------
// Случайный код без похожих символов (0/O, 1/I)
function newCode_() {
  const alphabet = 'ABCDEFGHJKLMNPQRSTUVWXYZ23456789';
  let code = '';
  for (let i = 0; i < 10; i++) code += alphabet.charAt(Math.floor(Math.random() * alphabet.length));
  return code;
}

function createInvite_() {
  const code = newCode_();
  const now = new Date();
  const exp = new Date(now.getTime() + INVITE_DAYS * 24 * 3600 * 1000);
  invitesSheet_().appendRow([code, now, exp, '', '']);
  return { ok: true, code: code, days: INVITE_DAYS };
}

function listUsers_() {
  const users = allUsers_();
  return {
    ok: true,
    users: users.map(function (u) {
      // Если имя в листе _Users пустое, показываем хотя бы Telegram ID
      const shared = u.status === 'active' ? partnerNames_(u) : '';
      return { id: u.id, name: visible_(u.name) ? u.name : ('ID ' + u.id), status: u.status, owner: isOwner_(u.id), shared: shared };
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
// Разбор и проверка траты из запроса. Возвращает null, если данные неверные
function parseExpense_(body) {
  const amount = Number(String(body.amount).replace(/\s/g, '').replace(',', '.'));
  // Старое название (например, из закэшированного приложения) сохраняем под новым
  const category = canon_(body.category);
  if (!(amount > 0) || amount > 1e8 || ALLOWED.indexOf(category) === -1) return null;
  let date = new Date();
  if (body.date) {
    const p = String(body.date).split('-');
    date = new Date(Number(p[0]), Number(p[1]) - 1, Number(p[2]));
    if (isNaN(date.getTime())) return null;
  }
  const comment = String(body.comment || '').slice(0, 200).replace(/^[=+\-@]/, "'$&");
  return { date: date, amount: Math.round(amount * 100) / 100, category: category, comment: comment };
}

function add_(ss, body, user) {
  const e = parseExpense_(body);
  if (!e) return { ok: false, error: 'bad data' };
  ensureSheets_(ss);
  const sh = ss.getSheetByName('Траты');
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    sh.appendRow([e.date, e.amount, e.category, e.comment, user ? user.name : '']);
    if (sh.getRange(1, 5).getValue() === '') sh.getRange(1, 5).setValue('Кто'); // старые таблицы
  } finally {
    lock.releaseLock();
  }
  checkLimits_(ss, e.category, user);
  return { ok: true };
}

// Находит строку траты и сверяет, что это та же запись, что видел пользователь.
// Если таблицу меняли (например, удалили строку выше), возвращает 0
function findExpenseRow_(sh, body) {
  const row = Number(body.row);
  const o = body.orig || {};
  if (!(row >= 2) || row > sh.getLastRow()) return 0;
  const r = sh.getRange(row, 1, 1, 4).getValues()[0];
  const same = r[0] instanceof Date &&
    Utilities.formatDate(r[0], Session.getScriptTimeZone(), 'yyyy-MM-dd') === String(o.date) &&
    Math.abs(Number(r[1]) - Number(o.amount)) < 0.005 &&
    canon_(r[2]) === String(o.category) &&
    String(r[3] || '') === String(o.comment || '');
  return same ? row : 0;
}

function update_(ss, body, user) {
  const e = parseExpense_(body);
  if (!e) return { ok: false, error: 'bad data' };
  const sh = ss.getSheetByName('Траты');
  if (!sh) return { ok: false, error: 'stale' };
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    const row = findExpenseRow_(sh, body);
    if (!row) return { ok: false, error: 'stale' };
    sh.getRange(row, 1, 1, 4).setValues([[e.date, e.amount, e.category, e.comment]]);
  } finally {
    lock.releaseLock();
  }
  checkLimits_(ss, e.category, user);
  return { ok: true };
}

function delete_(ss, body) {
  const sh = ss.getSheetByName('Траты');
  if (!sh) return { ok: false, error: 'stale' };
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    const row = findExpenseRow_(sh, body);
    if (!row) return { ok: false, error: 'stale' };
    sh.deleteRow(row);
  } finally {
    lock.releaseLock();
  }
  return { ok: true };
}

// ---------------------------------------------------------------------
// Настройки пользователя: лист «Настройки» в его таблице (Ключ | Значение)
//   budget:month            - бюджет на месяц
//   limit:week:<категория>  - лимит на неделю
//   limit:month:<категория> - лимит на месяц
//   notified:<период>:<ключ периода>:<категория> - уведомление о 80% уже отправлено
// ---------------------------------------------------------------------
function settingsSheet_(ss) {
  let sh = ss.getSheetByName('Настройки');
  if (!sh) {
    sh = ss.insertSheet('Настройки');
    sh.getRange('A:A').setNumberFormat('@');
    sh.appendRow(['Ключ', 'Значение']);
  }
  return sh;
}

function getSettings_(ss) {
  const sh = ss.getSheetByName('Настройки');
  const res = {};
  if (!sh) return res;
  sh.getDataRange().getValues().slice(1).forEach(function (r) {
    if (r[0] !== '') res[String(r[0])] = r[1];
  });
  return res;
}

// value === null удаляет настройку
function setSetting_(ss, key, value) {
  const sh = settingsSheet_(ss);
  const rows = sh.getDataRange().getValues();
  for (let i = rows.length - 1; i >= 1; i--) {
    if (String(rows[i][0]) === key) {
      if (value === null) sh.deleteRow(i + 1);
      else sh.getRange(i + 1, 2).setValue(value);
      return;
    }
  }
  if (value !== null) sh.appendRow([key, value]);
}

function limitsFrom_(settings) {
  const res = { week: {}, month: {} };
  Object.keys(settings).forEach(function (k) {
    const m = k.match(/^limit:(week|month):(.+)$/);
    if (m && ALLOWED.indexOf(m[2]) !== -1 && Number(settings[k]) > 0) res[m[1]][m[2]] = Number(settings[k]);
  });
  return res;
}

function parseMoney_(v) {
  if (v === null || v === undefined || String(v).trim() === '') return 0;
  const n = Number(String(v).replace(/\s/g, '').replace(',', '.'));
  return n >= 0 && n <= 1e8 ? Math.round(n * 100) / 100 : NaN;
}

function setLimit_(ss, body) {
  const period = body.period === 'week' ? 'week' : body.period === 'month' ? 'month' : '';
  const category = canon_(body.category);
  const amount = parseMoney_(body.amount);
  if (!period || ALLOWED.indexOf(category) === -1 || isNaN(amount)) return { ok: false, error: 'bad data' };
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    setSetting_(ss, 'limit:' + period + ':' + category, amount > 0 ? amount : null);
  } finally {
    lock.releaseLock();
  }
  return { ok: true };
}

function setBudget_(ss, body) {
  const amount = parseMoney_(body.amount);
  if (isNaN(amount)) return { ok: false, error: 'bad data' };
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    setSetting_(ss, 'budget:month', amount > 0 ? amount : null);
  } finally {
    lock.releaseLock();
  }
  return { ok: true };
}

// Все активные пользователи, которые ведут траты в этой таблице (для общего бюджета их двое)
function sheetMembers_(ss) {
  const id = ss.getId();
  return allUsers_().filter(function (u) { return u.status === 'active' && u.sheetId === id; });
}

// После записи: если по категории потрачено 80% лимита, один раз за период пишем в Telegram
function checkLimits_(ss, category, user) {
  try {
    const settings = getSettings_(ss);
    const limits = limitsFrom_(settings);
    const now = new Date();
    const tz = Session.getScriptTimeZone();
    const ws = weekStart_(now);
    const periods = [
      { p: 'week', key: Utilities.formatDate(ws, tz, 'yyyy-MM-dd'), from: ws,
        to: new Date(ws.getFullYear(), ws.getMonth(), ws.getDate() + 7), name: 'неделю' },
      { p: 'month', key: Utilities.formatDate(now, tz, 'yyyy-MM'), from: new Date(now.getFullYear(), now.getMonth(), 1),
        to: new Date(now.getFullYear(), now.getMonth() + 1, 1), name: 'месяц' }
    ];
    periods.forEach(function (x) {
      const limit = limits[x.p][category];
      if (!limit) return;
      const flag = 'notified:' + x.p + ':' + x.key + ':' + category;
      if (settings[flag]) return;
      const spent = sumByCategory_(ss, x.from, x.to)[category];
      if (spent < limit * 0.8) return;
      setSetting_(ss, flag, new Date());
      const text = category + ': потрачено ' + fmt_(spent) + ' из ' + fmt_(limit) + ' BYN за ' + x.name +
        ' (' + Math.round(spent / limit * 100) + '% лимита).';
      const ids = user ? [user.id] : [];
      sheetMembers_(ss).forEach(function (u) { if (ids.indexOf(u.id) === -1) ids.push(u.id); });
      ids.forEach(function (id) { sendMessage_(id, text); });
    });
  } catch (err) {
    console.error('Проверка лимитов: ' + (err && err.message ? err.message : 'unknown'));
  }
}

// Сумма по категориям за период: от from (включительно) до to (не включая)
// Все строки листа «Траты» (с заголовком). Читаем один раз и передаём в расчёты
function expenseRows_(ss) {
  const sh = ss.getSheetByName('Траты');
  return sh ? sh.getDataRange().getValues() : [];
}

function sumByCategory_(ss, from, to, rowsCache) {
  const res = {};
  ALLOWED.forEach(function (c) { res[c] = 0; });
  const rows = rowsCache || expenseRows_(ss);
  for (let i = 1; i < rows.length; i++) {
    const d = rows[i][0], a = Number(rows[i][1]), c = canon_(rows[i][2]);
    if (!(d instanceof Date) || !(a > 0) || !res.hasOwnProperty(c)) continue;
    if (d >= from && d < to) res[c] += a;
  }
  return res;
}

// Сколько внёс каждый участник за период (колонка «Кто»)
function whoSums_(ss, from, to, rowsCache) {
  const res = {};
  const rows = rowsCache || expenseRows_(ss);
  for (let i = 1; i < rows.length; i++) {
    const d = rows[i][0], a = Number(rows[i][1]);
    if (!(d instanceof Date) || !(a > 0) || d < from || d >= to) continue;
    const who = String(rows[i][4] || '') || 'Без отметки';
    res[who] = Math.round(((res[who] || 0) + a) * 100) / 100;
  }
  return res;
}

// Серия: сколько дней подряд есть хотя бы одна трата. Если сегодня ещё ничего нет,
// серия не прерывается, а считается по вчерашний день
function streakInfo_(rows) {
  const tz = Session.getScriptTimeZone();
  const key = function (d) { return Utilities.formatDate(d, tz, 'yyyy-MM-dd'); };
  const days = {};
  for (let i = 1; i < rows.length; i++) {
    if (rows[i][0] instanceof Date && Number(rows[i][1]) > 0) days[key(rows[i][0])] = true;
  }
  const now = new Date();
  const d = new Date(now.getFullYear(), now.getMonth(), now.getDate());
  const today = !!days[key(d)];
  if (!today) d.setDate(d.getDate() - 1);
  let streak = 0;
  while (days[key(d)] && streak < 3660) { streak++; d.setDate(d.getDate() - 1); }
  return { streak: streak, today: today };
}

function plural_(n, forms) {
  const a = Math.abs(n) % 100, b = a % 10;
  if (a > 10 && a < 20) return forms[2];
  if (b > 1 && b < 5) return forms[1];
  if (b === 1) return forms[0];
  return forms[2];
}

// Включить или выключить вечернее напоминание (своё у каждого, даже в общем бюджете)
function setReminder_(user, body) {
  const sh = usersSheet_();
  if (sh.getRange(1, 6).getValue() === '') sh.getRange(1, 6).setValue('напоминание');
  sh.getRange(user.row, 6).setValue(body.on ? '' : 'off');
  return { ok: true, reminder: !!body.on };
}

// По расписанию в 21:00: напоминаем тем, у кого в таблице за сегодня нет ни одной траты.
// В общем бюджете достаточно, чтобы записал кто-то один
function sendReminders() {
  const cache = {};
  forEachActiveUser_(function (u, ss) {
    if (!u.reminder) return;
    const info = cache[u.sheetId] || (cache[u.sheetId] = streakInfo_(expenseRows_(ss)));
    if (info.today) return;
    let text = 'Сегодня в приложении ещё нет трат. Если траты были, запиши их, пока помнится.';
    if (info.streak >= 2) {
      text += '\nСерия: ' + info.streak + ' ' + plural_(info.streak, ['день', 'дня', 'дней']) + ' подряд, не прерывай её.';
    }
    text += '\n\nВыключить напоминание можно в приложении: Настройки.';
    sendMessage_(u.id, text);
  });
}

// Запусти один раз вручную: напоминание каждый день около 21:00
function setupReminderTrigger() {
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === 'sendReminders') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('sendReminders').timeBased().everyDays(1).atHour(21).nearMinute(0).create();
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
  const ds = new Date(now.getFullYear(), now.getMonth(), now.getDate());
  const settings = getSettings_(ss);
  const rows = expenseRows_(ss);
  const we = new Date(ws.getFullYear(), ws.getMonth(), ws.getDate() + 7);
  const me = new Date(now.getFullYear(), now.getMonth() + 1, 1);
  const st = streakInfo_(rows);
  return {
    ok: true,
    week: sumByCategory_(ss, ws, we, rows),
    month: sumByCategory_(ss, ms, me, rows),
    today: total_(sumByCategory_(ss, ds, new Date(ds.getFullYear(), ds.getMonth(), ds.getDate() + 1), rows)),
    byWho: { week: whoSums_(ss, ws, we, rows), month: whoSums_(ss, ms, me, rows) },
    streak: st.streak,
    todayHas: st.today,
    budget: Number(settings['budget:month']) > 0 ? Number(settings['budget:month']) : 0,
    limits: limitsFrom_(settings)
  };
}

// Дата начала периода в виде 'yyyy-MM-dd'. В старых листах «Архив» таблица
// могла сама превратить текст в дату, поэтому принимаем оба варианта
function dayKey_(v) {
  if (v instanceof Date) return Utilities.formatDate(v, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  return String(v);
}

// Траты текущего месяца, новые сверху (для главной и вкладки «История»)
function history_(ss) {
  const sh = ss.getSheetByName('Траты');
  if (!sh) return { ok: true, items: [] };
  const rows = sh.getDataRange().getValues();
  const now = new Date();
  const from = new Date(now.getFullYear(), now.getMonth(), 1);
  const to = new Date(now.getFullYear(), now.getMonth() + 1, 1);
  const tz = Session.getScriptTimeZone();
  const items = [];
  for (let i = 1; i < rows.length; i++) {
    const d = rows[i][0], a = Number(rows[i][1]);
    if (!(d instanceof Date) || !(a > 0) || d < from || d >= to) continue;
    items.push({
      row: i + 1,
      date: Utilities.formatDate(d, tz, 'yyyy-MM-dd'),
      amount: a,
      category: canon_(rows[i][2]),
      comment: String(rows[i][3] || ''),
      who: String(rows[i][4] || '')
    });
  }
  items.sort(function (x, y) {
    if (x.date !== y.date) return x.date < y.date ? 1 : -1;
    return y.row - x.row;
  });
  return { ok: true, items: items.slice(0, 300) };
}

// Запись итогов периода в лист «Архив». Повторный запуск не создаёт дубли
function archive_(ss, type, from, to) {
  ensureSheets_(ss);
  const sh = ss.getSheetByName('Архив');
  const key = Utilities.formatDate(from, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  const exists = sh.getDataRange().getValues().some(function (r) {
    return r[0] === type && dayKey_(r[1]) === key;
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
    const start = dayKey_(r[1]);
    const k = r[0] + '|' + start;
    if (!map[k]) { map[k] = { type: r[0], start: start, sums: {} }; order.push(k); }
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
