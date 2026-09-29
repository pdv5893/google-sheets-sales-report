/**
 * Отчёт по продажам в Google Таблицах + отправка в Telegram.
 *
 * Меню «📊 Отчёты» появляется в таблице автоматически:
 *   • Сформировать отчёт          — лист «Сводка»: итоги, диаграмма, выручка по месяцам, топ товаров
 *   • Отправить отчёт в Telegram  — короткая сводка сообщением + PDF листа «Сводка»
 *   • Настроить Telegram          — подключение бота за минуту, chat id определяется сам
 *   • Ежедневная отправка         — отчёт приходит в Telegram каждое утро без участия человека
 *   • Создать демо-данные         — 3000 продаж для проверки
 *
 * Скрипт работает на серверах Google: компьютер можно выключить,
 * ежедневный отчёт всё равно придёт.
 */

// ---------- Настройки под таблицу заказчика ----------
const CONFIG = {
  SOURCE_SHEET: 'Продажи',   // лист с исходными данными
  SUMMARY_SHEET: 'Сводка',   // лист отчёта (пересоздаётся при каждом запуске)
  COL_GROUP: 'Менеджер',     // по какой колонке считать итоги
  COL_AMOUNT: 'Сумма',       // колонка с суммой (обязательная)
  COL_DATE: 'Дата',          // для выручки по месяцам (необязательная)
  COL_PRODUCT: 'Товар',      // для топа товаров (необязательная)
  TOP_N: 5,
  DAILY_HOUR: 9,             // во сколько часов присылать ежедневный отчёт
};

const HEADER_BG = '#2F5597';
const TOTAL_BG = '#DDEBF7';
const MONEY = '#,##0 "₽"';
const PROP_TOKEN = 'TELEGRAM_TOKEN';
const PROP_CHAT = 'TELEGRAM_CHAT_ID';
const MONTHS = ['Январь', 'Февраль', 'Март', 'Апрель', 'Май', 'Июнь',
                'Июль', 'Август', 'Сентябрь', 'Октябрь', 'Ноябрь', 'Декабрь'];


// =====================================================================
//  Меню
// =====================================================================

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('📊 Отчёты')
    .addItem('Сформировать отчёт', 'menuBuildReport')
    .addItem('Отправить отчёт в Telegram', 'menuSendReport')
    .addSeparator()
    .addItem('Настроить Telegram', 'menuSetupTelegram')
    .addItem('Включить ежедневную отправку', 'menuEnableDaily')
    .addItem('Выключить ежедневную отправку', 'menuDisableDaily')
    .addSeparator()
    .addItem('Создать демо-данные', 'menuCreateDemoData')
    .addToUi();
}

function menuBuildReport() {
  const started = Date.now();
  const report = buildReport();
  SpreadsheetApp.getUi().alert(
    'Готово за ' + ((Date.now() - started) / 1000).toFixed(1) + ' сек.\n\n' +
    'Обработано строк: ' + report.rows + '\n' +
    'Выручка: ' + money(report.grand)
  );
}

function menuSendReport() {
  const ui = SpreadsheetApp.getUi();
  if (!getTelegramSettings()) {
    ui.alert('Сначала подключите Telegram: меню «📊 Отчёты» → «Настроить Telegram».');
    return;
  }
  sendReportToTelegram();
  ui.alert('Отчёт отправлен в Telegram.');
}

function menuSetupTelegram() {
  const ui = SpreadsheetApp.getUi();

  const tokenAnswer = ui.prompt(
    'Настройка Telegram (шаг 1 из 2)',
    'Создайте бота у @BotFather в Telegram и вставьте сюда его токен:',
    ui.ButtonSet.OK_CANCEL
  );
  if (tokenAnswer.getSelectedButton() !== ui.Button.OK) return;
  const token = tokenAnswer.getResponseText().replace(/\s+/g, '');
  if (!/^\d+:[\w-]{30,}$/.test(token)) {
    ui.alert('Это не похоже на токен бота. Он выглядит так: 123456789:ABC-def...');
    return;
  }

  // Узнаём, какому боту принадлежит токен, — чтобы писать именно ему
  let bot;
  try {
    bot = telegramCall(token, 'getMe', {}).result;
  } catch (e) {
    ui.alert('Telegram не принял токен: ' + e.message + '\nСкопируйте токен из сообщения @BotFather ещё раз.');
    return;
  }
  const botName = '@' + bot.username;

  const ready = ui.alert(
    'Настройка Telegram (шаг 2 из 2)',
    'Токен принадлежит боту ' + botName + ' («' + bot.first_name + '»).\n\n' +
    'Откройте в Telegram именно этого бота, нажмите «Start» и отправьте ему любое сообщение ' +
    '(например, «привет»). Для отчётов в рабочую группу — добавьте бота в группу и напишите что-нибудь там.\n\n' +
    'Затем нажмите OK.',
    ui.ButtonSet.OK_CANCEL
  );
  if (ready !== ui.Button.OK) return;

  // Находим chat id по последнему сообщению боту — вводить его вручную не нужно
  const updates = telegramCall(token, 'getUpdates', {}).result || [];
  const chats = [];
  updates.forEach(function (u) {
    const upd = u.message || u.edited_message || u.channel_post || u.my_chat_member;
    if (upd && upd.chat) chats.push(upd.chat);
  });
  if (chats.length === 0) {
    ui.alert('Бот ' + botName + ' пока не получил ни одного сообщения.\n' +
             'Напишите именно этому боту и запустите настройку ещё раз.');
    return;
  }
  const chat = chats[chats.length - 1];
  const chatName = chat.title || [chat.first_name, chat.last_name].filter(String).join(' ') || String(chat.id);

  // Проверяем связь тестовым сообщением и только потом сохраняем настройки
  try {
    telegramCall(token, 'sendMessage', {
      chat_id: String(chat.id),
      text: '✅ Таблица «' + SpreadsheetApp.getActive().getName() + '» подключена. Сюда будут приходить отчёты.',
    });
  } catch (e) {
    ui.alert('Не удалось отправить сообщение.\n\n' +
             'Бот: ' + botName + '\nЧат: ' + chatName + ' (id ' + chat.id + ', тип ' + chat.type + ')\n' +
             'Ошибка: ' + e.message + '\n\n' +
             'Сделайте скриншот этого окна — по нему видно, в чём дело.');
    return;
  }

  PropertiesService.getScriptProperties().setProperties({
    [PROP_TOKEN]: token,
    [PROP_CHAT]: String(chat.id),
  });
  ui.alert('Готово! Бот ' + botName + ' подключён к чату «' + chatName + '».\nВ Telegram отправлено тестовое сообщение.');
}

function menuEnableDaily() {
  const ui = SpreadsheetApp.getUi();
  if (!getTelegramSettings()) {
    ui.alert('Сначала подключите Telegram: меню «📊 Отчёты» → «Настроить Telegram».');
    return;
  }
  removeDailyTriggers();
  ScriptApp.newTrigger('sendDailyReport').timeBased().everyDays(1).atHour(CONFIG.DAILY_HOUR).create();
  ui.alert('Ежедневная отправка включена: отчёт будет приходить каждый день около ' +
           CONFIG.DAILY_HOUR + ':00 (часовой пояс таблицы).');
}

function menuDisableDaily() {
  const removed = removeDailyTriggers();
  SpreadsheetApp.getUi().alert(removed > 0 ? 'Ежедневная отправка выключена.' : 'Ежедневная отправка и так выключена.');
}

function menuCreateDemoData() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActive();
  if (ss.getSheetByName(CONFIG.SOURCE_SHEET)) {
    const answer = ui.alert('Лист «' + CONFIG.SOURCE_SHEET + '» уже есть. Заменить его демо-данными?', ui.ButtonSet.YES_NO);
    if (answer !== ui.Button.YES) return;
  }
  createDemoData();
  ui.alert('Создано 3000 демо-продаж на листе «' + CONFIG.SOURCE_SHEET + '».\nТеперь: «📊 Отчёты» → «Сформировать отчёт».');
}


// =====================================================================
//  Запуск по расписанию (без окон — вызывается триггером)
// =====================================================================

function sendDailyReport() {
  sendReportToTelegram();
}

function removeDailyTriggers() {
  let removed = 0;
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === 'sendDailyReport') {
      ScriptApp.deleteTrigger(t);
      removed++;
    }
  });
  return removed;
}


// =====================================================================
//  Расчёт
// =====================================================================

/** Все агрегаты за один проход по данным. Чистая функция — не трогает таблицу. */
function aggregate(values, cfg, tz) {
  const header = values[0].map(function (h) { return String(h).trim().toLowerCase(); });
  const col = function (name) { return header.indexOf(String(name).trim().toLowerCase()); };
  const cGroup = col(cfg.COL_GROUP), cAmount = col(cfg.COL_AMOUNT);
  const cDate = col(cfg.COL_DATE), cProduct = col(cfg.COL_PRODUCT);
  if (cAmount < 0) throw new Error('Не найдена колонка «' + cfg.COL_AMOUNT + '».');
  if (cGroup < 0) throw new Error('Не найдена колонка «' + cfg.COL_GROUP + '».');

  const groups = {}, months = {}, products = {};
  let grand = 0, rows = 0, minDate = null, maxDate = null;

  for (let r = 1; r < values.length; r++) {
    const row = values[r];
    if (row.every(function (v) { return v === '' || v === null; })) continue;  // пустая строка
    rows++;
    const key = keyOf(row[cGroup]);
    const amount = typeof row[cAmount] === 'number' ? row[cAmount] : Number(row[cAmount]) || 0;
    grand += amount;

    groups[key] = groups[key] || { count: 0, sum: 0 };
    groups[key].count++;
    groups[key].sum += amount;

    if (cDate >= 0 && row[cDate] instanceof Date) {
      const d = row[cDate];
      const m = Utilities.formatDate(d, tz, 'yyyy-MM');
      months[m] = months[m] || {};
      months[m][key] = (months[m][key] || 0) + amount;
      if (!minDate || d < minDate) minDate = d;
      if (!maxDate || d > maxDate) maxDate = d;
    }
    if (cProduct >= 0) {
      const p = keyOf(row[cProduct]);
      products[p] = (products[p] || 0) + amount;
    }
  }

  const groupList = Object.keys(groups)
    .map(function (k) { return { name: k, count: groups[k].count, sum: groups[k].sum }; })
    .sort(function (a, b) { return b.sum - a.sum; });
  const productList = Object.keys(products)
    .map(function (k) { return { name: k, sum: products[k] }; })
    .sort(function (a, b) { return b.sum - a.sum; })
    .slice(0, cfg.TOP_N);

  return {
    rows: rows, grand: grand, groups: groupList, products: productList,
    months: months, monthKeys: Object.keys(months).sort(),
    minDate: minDate, maxDate: maxDate,
    hasProducts: cProduct >= 0 && cProduct !== cGroup,
  };
}

function keyOf(v) {
  if (v instanceof Date) return Utilities.formatDate(v, Session.getScriptTimeZone(), 'dd.MM.yyyy');
  const s = String(v === null || v === undefined ? '' : v).trim();
  return s || '(пусто)';
}


// =====================================================================
//  Лист «Сводка»
// =====================================================================

function buildReport() {
  const ss = SpreadsheetApp.getActive();
  const src = ss.getSheetByName(CONFIG.SOURCE_SHEET);
  if (!src) throw new Error('Не найден лист «' + CONFIG.SOURCE_SHEET + '». Создайте демо-данные через меню.');
  const values = src.getDataRange().getValues();
  if (values.length < 2) throw new Error('На листе «' + CONFIG.SOURCE_SHEET + '» нет данных.');

  const tz = ss.getSpreadsheetTimeZone();
  const rep = aggregate(values, CONFIG, tz);

  // Лист отчёта пересоздаём с нуля — никаких «хвостов» от прошлого запуска
  const old = ss.getSheetByName(CONFIG.SUMMARY_SHEET);
  if (old) ss.deleteSheet(old);
  const sh = ss.insertSheet(CONFIG.SUMMARY_SHEET, src.getIndex());
  sh.setHiddenGridlines(true);

  sh.getRange('A1').setValue('Сводка по продажам').setFontSize(14).setFontWeight('bold');
  sh.getRange('A2').setValue('Сформировано ' + Utilities.formatDate(new Date(), tz, 'dd.MM.yyyy HH:mm') +
                             ', строк: ' + rep.rows).setFontColor('#6e6e6e');

  // ---------- Блок 1: итоги по группам ----------
  let row = 4;
  const g = rep.groups.length;
  const block1 = [[CONFIG.COL_GROUP, 'Сделок', 'Выручка', 'Средний чек', 'Доля']];
  rep.groups.forEach(function (x) {
    block1.push([x.name, x.count, x.sum, x.sum / x.count, rep.grand ? x.sum / rep.grand : 0]);
  });
  block1.push(['Итого', rep.rows, rep.grand, rep.grand / rep.rows, 1]);
  sh.getRange(row, 1, block1.length, 5).setValues(block1);
  styleHeader(sh.getRange(row, 1, 1, 5));
  styleTotal(sh.getRange(row + g + 1, 1, 1, 5));
  sh.getRange(row + 1, 2, g + 1, 1).setNumberFormat('#,##0');
  sh.getRange(row + 1, 3, g + 1, 2).setNumberFormat(MONEY);
  sh.getRange(row + 1, 5, g + 1, 1).setNumberFormat('0.0%');

  // Диаграмма — сразу под таблицей итогов
  const chartRow = row + g + 3;
  const chart = sh.newChart()
    .setChartType(Charts.ChartType.COLUMN)
    .addRange(sh.getRange(row + 1, 1, g, 1))
    .addRange(sh.getRange(row + 1, 3, g, 1))
    .setPosition(chartRow, 1, 0, 0)
    .setOption('title', 'Выручка: ' + CONFIG.COL_GROUP)
    .setOption('legend', { position: 'none' })
    .setOption('colors', [HEADER_BG])
    .setOption('width', Math.max(560, 70 * g))
    .setOption('height', 320)
    .build();
  sh.insertChart(chart);
  row = chartRow + 17;   // высота диаграммы ~ 16 строк

  // ---------- Блок 2: выручка по месяцам ----------
  if (rep.monthKeys.length > 0) {
    sh.getRange(row, 1).setValue('Выручка по месяцам').setFontWeight('bold');
    row++;
    const names = rep.groups.map(function (x) { return x.name; });
    const block2 = [['Месяц'].concat(names, ['Итого'])];
    const colTotals = names.map(function () { return 0; });
    let grandMonths = 0;
    rep.monthKeys.forEach(function (m) {
      const parts = m.split('-');
      // Готовое название месяца: формат даты в русской локали дал бы «января 2026»
      const line = [MONTHS[Number(parts[1]) - 1] + ' ' + parts[0]];
      let rowSum = 0;
      names.forEach(function (n, i) {
        const v = rep.months[m][n] || 0;
        line.push(v); rowSum += v; colTotals[i] += v;
      });
      line.push(rowSum); grandMonths += rowSum;
      block2.push(line);
    });
    block2.push(['Итого'].concat(colTotals, [grandMonths]));
    const w = names.length + 2;
    sh.getRange(row, 1, block2.length, w).setValues(block2);
    styleHeader(sh.getRange(row, 1, 1, w));
    styleTotal(sh.getRange(row + block2.length - 1, 1, 1, w));
    sh.getRange(row + 1, 2, block2.length - 1, w - 1).setNumberFormat(MONEY);
    sh.getRange(row + 1, w, block2.length - 1, 1).setFontWeight('bold');
    row += block2.length + 2;
  }

  // ---------- Блок 3: топ товаров ----------
  if (rep.hasProducts && rep.products.length > 0) {
    sh.getRange(row, 1).setValue('Топ-' + CONFIG.TOP_N + ' товаров по выручке').setFontWeight('bold');
    row++;
    const block3 = [[CONFIG.COL_PRODUCT, 'Выручка', 'Доля']];
    rep.products.forEach(function (p) {
      block3.push([p.name, p.sum, rep.grand ? p.sum / rep.grand : 0]);
    });
    sh.getRange(row, 1, block3.length, 3).setValues(block3);
    styleHeader(sh.getRange(row, 1, 1, 3));
    sh.getRange(row + 1, 2, block3.length - 1, 1).setNumberFormat(MONEY);
    sh.getRange(row + 1, 3, block3.length - 1, 1).setNumberFormat('0.0%');
  }

  sh.autoResizeColumns(1, Math.max(5, g + 2));
  if (sh.getColumnWidth(1) < 150) sh.setColumnWidth(1, 150);
  sh.setFrozenRows(0);
  SpreadsheetApp.flush();
  return rep;
}

function styleHeader(range) {
  range.setBackground(HEADER_BG).setFontColor('#ffffff').setFontWeight('bold');
}

function styleTotal(range) {
  range.setBackground(TOTAL_BG).setFontWeight('bold')
       .setBorder(true, null, null, null, null, null, '#000000', SpreadsheetApp.BorderStyle.SOLID);
}


// =====================================================================
//  Telegram
// =====================================================================

function getTelegramSettings() {
  const p = PropertiesService.getScriptProperties();
  const token = p.getProperty(PROP_TOKEN), chatId = p.getProperty(PROP_CHAT);
  return token && chatId ? { token: token, chatId: chatId } : null;
}

function sendReportToTelegram() {
  const tg = getTelegramSettings();
  if (!tg) throw new Error('Telegram не настроен.');
  const ss = SpreadsheetApp.getActive();
  const rep = buildReport();

  telegramCall(tg.token, 'sendMessage', {
    chat_id: tg.chatId,
    text: reportText(rep, ss.getName(), ss.getSpreadsheetTimeZone()),
    parse_mode: 'HTML',
  });

  // Лист «Сводка» целиком — PDF-файлом, с диаграммой
  const summary = ss.getSheetByName(CONFIG.SUMMARY_SHEET);
  const url = 'https://docs.google.com/spreadsheets/d/' + ss.getId() + '/export' +
    '?format=pdf&gid=' + summary.getSheetId() +
    '&portrait=false&fitw=true&gridlines=false&printtitle=false&sheetnames=false&pagenum=UNDEFINED';
  const pdf = UrlFetchApp.fetch(url, { headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() } })
    .getBlob()
    .setName('sales-report-' + Utilities.formatDate(new Date(), ss.getSpreadsheetTimeZone(), 'dd.MM.yyyy') + '.pdf');
  // Имя латиницей: кириллицу в имени файла Telegram обрезает
  telegramCall(tg.token, 'sendDocument', { chat_id: tg.chatId, document: pdf });
}

/** Текст сообщения. Чистая функция — удобно проверять отдельно. */
function reportText(rep, title, tz) {
  const lines = [];
  lines.push('📊 <b>' + esc(title) + '</b>');
  if (rep.minDate && rep.maxDate) {
    lines.push('Период: ' + Utilities.formatDate(rep.minDate, tz, 'dd.MM.yyyy') + ' — ' +
               Utilities.formatDate(rep.maxDate, tz, 'dd.MM.yyyy'));
  }
  lines.push('');
  lines.push('💰 Выручка: <b>' + money(rep.grand) + '</b>');
  lines.push('🧾 Сделок: ' + num(rep.rows) + ', средний чек: ' + money(rep.grand / rep.rows));
  lines.push('');
  lines.push('<b>' + esc(CONFIG.COL_GROUP) + ':</b>');
  rep.groups.forEach(function (x, i) {
    const share = rep.grand ? (x.sum / rep.grand * 100).toFixed(1).replace('.', ',') : '0';
    lines.push((i + 1) + '. ' + esc(x.name) + ' — ' + money(x.sum) + ' (' + share + '%)');
  });
  if (rep.hasProducts && rep.products.length > 0) {
    lines.push('');
    lines.push('🏆 Лидер продаж: ' + esc(rep.products[0].name) + ' — ' + money(rep.products[0].sum));
  }
  return lines.join('\n');
}

function telegramCall(token, method, payload) {
  const resp = UrlFetchApp.fetch('https://api.telegram.org/bot' + token + '/' + method, {
    method: 'post',
    payload: payload,
    muteHttpExceptions: true,
  });
  const data = JSON.parse(resp.getContentText());
  if (!data.ok) throw new Error('Telegram: ' + (data.description || resp.getResponseCode()));
  return data;
}

function num(n) {
  return String(Math.round(n)).replace(/\B(?=(\d{3})+(?!\d))/g, ' ');
}

function money(n) {
  return num(n) + ' ₽';
}

function esc(s) {
  return String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
}


// =====================================================================
//  Демо-данные
// =====================================================================

function createDemoData() {
  const ss = SpreadsheetApp.getActive();
  const managers = [['Иванов А.', 1.4], ['Петрова М.', 1.2], ['Сидоров К.', 1.0],
                    ['Кузнецова Е.', 0.9], ['Смирнов Д.', 0.7], ['Волкова О.', 0.6]];
  const regions = [['Москва', 5], ['Санкт-Петербург', 3], ['Екатеринбург', 1.5], ['Новосибирск', 1.2], ['Казань', 1]];
  const products = [['Ноутбук Pro 14', 89990, 1], ['Монитор 27" 4K', 34990, 1.5], ['Смартфон X12', 54990, 2],
                    ['Наушники ANC', 12990, 3], ['Планшет Tab 11', 39990, 1.3], ['Клавиатура механическая', 7990, 2],
                    ['Мышь беспроводная', 2990, 4], ['Внешний SSD 1 ТБ', 9990, 2.5], ['Роутер Wi-Fi 6', 6990, 1.5],
                    ['Умные часы S3', 19990, 1.8]];
  const qty = [[1, 60], [2, 20], [3, 10], [5, 6], [10, 4]];

  // Взвешенный случайный выбор: элемент — [значение..., вес последним]
  const pick = function (list) {
    const total = list.reduce(function (s, x) { return s + x[x.length - 1]; }, 0);
    let r = Math.random() * total;
    for (let i = 0; i < list.length; i++) {
      r -= list[i][list[i].length - 1];
      if (r <= 0) return list[i];
    }
    return list[list.length - 1];
  };

  const year = new Date().getFullYear();
  const rows = [];
  for (let i = 0; i < 3000; i++) {
    const d = new Date(year, 0, 1 + Math.floor(Math.random() * 243));
    const p = pick(products);
    const q = pick(qty)[0];
    const price = Math.round(p[1] * (0.9 + Math.random() * 0.15) / 10) * 10;
    rows.push([d, pick(managers)[0], pick(regions)[0], p[0], q, price, q * price]);
  }
  rows.sort(function (a, b) { return a[0] - b[0]; });

  let sh = ss.getSheetByName(CONFIG.SOURCE_SHEET);
  if (sh) sh.clear(); else sh = ss.insertSheet(CONFIG.SOURCE_SHEET, 0);
  const header = [['Дата', 'Менеджер', 'Регион', 'Товар', 'Кол-во', 'Цена', 'Сумма']];
  sh.getRange(1, 1, 1, 7).setValues(header);
  sh.getRange(2, 1, rows.length, 7).setValues(rows);
  styleHeader(sh.getRange(1, 1, 1, 7));
  sh.getRange(2, 1, rows.length, 1).setNumberFormat('dd.mm.yyyy');
  sh.getRange(2, 6, rows.length, 2).setNumberFormat(MONEY);
  sh.setFrozenRows(1);
  if (sh.getFilter()) sh.getFilter().remove();
  sh.getRange(1, 1, rows.length + 1, 7).createFilter();
  sh.autoResizeColumns(1, 7);
  ss.setActiveSheet(sh);
}
