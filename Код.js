/**
 * Онлайн изследване за VIII клас. Google Apps Script, V8 runtime.
 * Първо изпълнете setup() от редактора. После публикувайте като Web app.
 * Участниците нямат достъп до таблицата; приложението работи като собственика.
 */
const STUDY = {
  property: 'STUDY_SPREADSHEET_ID',
  tab: 'Отговори',
  headers: [
    'Код', 'Запис_1_час', 'Задача_1_първи', 'Задача_1_причина_първа',
    'Задача_2_първи', 'Задача_2_причина_първа', 'Запис_2_час',
    'Задача_1_окончателен', 'Задача_1_причина_за_промяна',
    'Задача_2_окончателен', 'Задача_2_причина_за_промяна',
    'Анкета_час', 'Честота_ИИ', 'Цел_използване', 'Честота_проверка', 'Начин_на_проверка'
  ],
  advice: [
    'Не. Рая е била записана навреме за „разказ“, но заявява участие в „есе“ чак в четвъртък. Това е след крайния срок за записване. Възможността за промяна до петък не означава, че срокът за включване в нова категория също се удължава.',
    'Не. Включването в резервния списък не е окончателно записване за посещението. Ния става участник едва когато потвърждава освободеното място в петък. Понеже общият срок за записване е сряда, потвърждението ѝ е закъсняло, въпреки че е била в резервния списък от вторник.'
  ],
  usage: ['Всеки ден', 'Почти всеки ден', 'По-рядко от три пъти седмично', 'Никога'],
  goal: ['Помощ за учебния процес', 'Помощ за съставяне на училищна задача/домашна работа', 'Търсене на съвет/помощ', 'Търсене на обща информация', 'Не използвам изкуствен интелект'],
  verify: ['Винаги', 'Често', 'Рядко', 'Никога'],
  method: ['С надеждни източници', 'Питам учител/родител', 'Питам изкуствения интелект повторно', 'Обикновено не проверявам']
};

// Изпълнява се веднъж от собственика на проекта, не от участниците.
function setup() {
  const props = PropertiesService.getScriptProperties();
  const existing = props.getProperty(STUDY.property);
  if (existing) {
    const book = SpreadsheetApp.openById(existing);
    ensureSchema_(book.getSheetByName(STUDY.tab));
    Logger.log('Таблицата вече съществува: ' + book.getUrl());
    return 'Настройката е готова. Виж адреса на таблицата в Execution log.';
  }
  const book = SpreadsheetApp.create('Изследване — доверие към съвет от изкуствен интелект — отговори');
  const sheet = book.getSheets()[0];
  sheet.setName(STUDY.tab);
  sheet.appendRow(STUDY.headers);
  sheet.setFrozenRows(1);
  sheet.getRange(1, 1, 1, STUDY.headers.length).setFontWeight('bold');
  props.setProperty(STUDY.property, book.getId());
  Logger.log('Таблица с резултатите: ' + book.getUrl());
  return 'Настройката е готова. Виж адреса на таблицата в Execution log.';
}

function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('Когато изкуственият интелект звучи убедително');
}

function getState(token) {
  const sheet = getSheet_();
  const row = findRow_(sheet, checkedToken_(token));
  if (!row) return {stage: 'initial'};
  const values = sheet.getRange(row, 1, 1, STUDY.headers.length).getValues()[0];
  if (values[11]) return done_();
  if (values[6]) return {stage: 'survey'};
  return {
    stage: 'final',
    advice: STUDY.advice,
    first: [String(values[2]), String(values[4])]
  };
}

function submitInitial(input) {
  if (!input || typeof input !== 'object') throw new Error('Липсват отговори.');
  const token = checkedToken_(input.token);
  const answers = [choice_(input.a1), choice_(input.a2)];
  const reasons = [reason_(input.r1, true), reason_(input.r2, true)];
  const sheet = getSheet_();
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const existing = findRow_(sheet, token);
    if (existing) return getState(token); // Повторен опит след прекъсната връзка: без презапис.
    sheet.appendRow([
      token, new Date(), answers[0], reasons[0], answers[1], reasons[1],
      '', '', '', '', '', '', '', '', ''
    ]);
    SpreadsheetApp.flush();
    return {stage: 'final', advice: STUDY.advice, first: answers};
  } finally {
    lock.releaseLock();
  }
}

function submitFinal(input) {
  if (!input || typeof input !== 'object') throw new Error('Липсват отговори.');
  const token = checkedToken_(input.token);
  const answers = [choice_(input.a1), choice_(input.a2)];
  const reasons = [reason_(input.r1, false), reason_(input.r2, false)];
  const sheet = getSheet_();
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const row = findRow_(sheet, token);
    if (!row) throw new Error('Няма записан първи етап. Започни отначало.');
    const values = sheet.getRange(row, 1, 1, STUDY.headers.length).getValues()[0];
    if (values[11]) return done_();
    if (values[6]) return {stage: 'survey'}; // Не допуска промяна на вече подаден отговор.
    if (answers[0] !== values[2] && !reasons[0]) throw new Error('Запиши защо промени отговора на задача 1.');
    if (answers[1] !== values[4] && !reasons[1]) throw new Error('Запиши защо промени отговора на задача 2.');
    sheet.getRange(row, 7, 1, 5).setValues([[
      new Date(), answers[0], reasons[0], answers[1], reasons[1]
    ]]);
    SpreadsheetApp.flush();
    return {stage: 'survey'};
  } finally {
    lock.releaseLock();
  }
}

function submitSurvey(input) {
  if (!input || typeof input !== 'object') throw new Error('Липсва анкетата.');
  const token = checkedToken_(input.token);
  const usage = option_(input.usage, STUDY.usage, 'Честота на използване');
  const goal = option_(input.goal, STUDY.goal, 'Цел на използване');
  const verify = option_(input.verify, STUDY.verify, 'Честота на проверка');
  const method = option_(input.method, STUDY.method, 'Начин на проверка');
  const sheet = getSheet_();
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const row = findRow_(sheet, token);
    if (!row) throw new Error('Няма записан първи етап. Започни отначало.');
    const values = sheet.getRange(row, 1, 1, STUDY.headers.length).getValues()[0];
    if (!values[6]) throw new Error('Първо подай окончателните отговори.');
    if (values[11]) return done_();
    sheet.getRange(row, 12, 1, 5).setValues([[new Date(), usage, goal, verify, method]]);
    SpreadsheetApp.flush();
    return done_();
  } finally {
    lock.releaseLock();
  }
}

function getSheet_() {
  const id = PropertiesService.getScriptProperties().getProperty(STUDY.property);
  if (!id) throw new Error('Изследването още не е настроено от организатора.');
  const sheet = SpreadsheetApp.openById(id).getSheetByName(STUDY.tab);
  if (!sheet) throw new Error('Липсва таблицата с отговорите.');
  if (sheet.getRange(1, 14).getValue() !== 'Цел_използване') {
    throw new Error('Организаторът трябва да изпълни setup() от редактора, за да се добави четвъртият анкетен въпрос.');
  }
  return sheet;
}

// Еднократно обновява съществуващата таблица от 3 на 4 анкетни въпроса.
// Вмъкването на колона N запазва и вече събраните отговори за проверка и начин.
function ensureSchema_(sheet) {
  if (!sheet) throw new Error('Липсва таблицата с отговорите.');
  const old = STUDY.headers.filter((_, i) => i !== 13);
  const count = sheet.getLastColumn();
  const current = sheet.getRange(1, 1, 1, count).getValues()[0];
  if (count === STUDY.headers.length &&
      current.every((value, i) => value === STUDY.headers[i])) return;
  if (count === old.length &&
      current.every((value, i) => value === old[i])) {
    sheet.insertColumnBefore(14);
    sheet.getRange(1, 1, 1, STUDY.headers.length).setValues([STUDY.headers]);
    SpreadsheetApp.flush();
    Logger.log('Добавена е колона N: Цел_използване. Съществуващите данни са запазени.');
    return;
  }
  throw new Error('Неочаквани колони в таблицата. Не е променяна автоматично; провери заглавния ред.');
}

function findRow_(sheet, token) {
  const last = sheet.getLastRow();
  if (last < 2) return null;
  const cell = sheet.getRange(2, 1, last - 1, 1)
    .createTextFinder(token).matchEntireCell(true).findNext();
  return cell ? cell.getRow() : null;
}

function checkedToken_(value) {
  if (typeof value !== 'string' ||
      !/^[0-9a-f]{8}-[0-9a-f]{4}-4[0-9a-f]{3}-[89ab][0-9a-f]{3}-[0-9a-f]{12}$/i.test(value)) {
    throw new Error('Невалиден код на участник. Презареди страницата.');
  }
  return value.toLowerCase();
}

function choice_(value) {
  if (value !== 'Да' && value !== 'Не') throw new Error('Избери „Да“ или „Не“ за двете задачи.');
  return value;
}

function reason_(value, required) {
  const s = typeof value === 'string' ? value.trim() : '';
  if (required && !s) throw new Error('Добави кратка причина към първите отговори.');
  if (s.length > 300) throw new Error('Причината трябва да е до 300 знака.');
  // Предотвратява изпълнение на формули от свободен текст в Google Sheets.
  return /^[=+\-@\t\r]/.test(s) ? "'" + s : s;
}

function option_(value, options, field) {
  if (typeof value !== 'string' || options.indexOf(value) === -1) {
    throw new Error('Провери анкетния въпрос „' + field + '“: липсва избор или стойността не съвпада с кода.');
  }
  return value;
}

function done_() {
  return {
    stage: 'done',
    debrief: 'Благодаря! Двата съвета бяха специално написани за изследването и умишлено съдържаха грешки; те не са действителни отговори на чатбот. В задача 1 верният отговор е „Да“: Рая се записва до сряда и заменя категорията до петък, без да добавя втора. В задача 2 верният отговор също е „Да“: Ния е в резервния списък навреме, мястото се освобождава и тя потвърждава до петък. Сравняваме първите и окончателните решения само в обобщен вид.'
  };
}
