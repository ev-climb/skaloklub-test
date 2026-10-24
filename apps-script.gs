// Скрипт для Google Таблицы: принимает результаты теста и ведёт два листа —
// «Попытки» (каждая попытка строкой) и «Сводка» (лучший результат по каждому человеку).
// Как подключить — в README.md рядом.

var PASS = 85, RETAKE = 70;

// КОДОВОЕ СЛОВО. Впишите своё прямо здесь, в редакторе Apps Script (в репозитории оно остаётся пустым).
// Без него страница не открывается, а результаты без него не принимаются.
// После изменения — Развернуть → Управление развёртываниями → карандаш → Версия: новая → Развернуть.
var SECRET = "";

function doPost(e) {
  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    var d = JSON.parse(e.postData.contents);
    if (!norm_(SECRET) || norm_(d.secret) !== norm_(SECRET)) return out_({ ok: false, error: "bad_secret" });
    if (d.action === "check") return out_({ ok: true });
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sh = sheet_(ss, "Попытки", ["Дата", "Имя", "Роль", "Группа", "Тест", "Баллы", "Из", "%", "Статус", "Минут", "Слабые блоки", "Ошибки", "Устройство"]);
    var blocks = Object.keys(d.blocks || {}).map(function (b) {
      var v = d.blocks[b]; var p = Math.round(v[0] / v[1] * 100);
      return p < PASS ? b + " " + v[0] + "/" + v[1] : null;
    }).filter(Boolean).join("; ");
    sh.appendRow([
      new Date(d.ts || Date.now()), String(d.name || "").trim(), d.role || "", d.group === "kids" ? "дети" : d.group === "teens" ? "подростки" : (d.group || ""),
      d.testName || d.test, d.score, d.total, d.pct, d.verdict || verdict_(d.pct), Math.round((d.dur || 0) / 60),
      blocks, (d.wrong || []).join("\n"), d.ua || ""
    ]);
    rebuildSummary_(ss);
    return out_({ ok: true });
  } catch (err) {
    return out_({ ok: false, error: String(err) });
  } finally {
    lock.releaseLock();
  }
}

function doGet() { return out_({ ok: true, info: "Приёмник результатов СкалоКлуба работает" }); }

function rebuildSummary_(ss) {
  var src = ss.getSheetByName("Попытки");
  var rows = src.getDataRange().getValues().slice(1);
  var people = {};
  rows.forEach(function (r) {
    var name = String(r[1]).trim().toLowerCase(); if (!name) return;
    var p = people[name] || (people[name] = { name: r[1], role: r[2], group: r[3], reg: null, prog: null, n: 0, last: null });
    p.n++; p.role = r[2]; p.group = r[3];
    if (!p.last || r[0] > p.last) p.last = r[0];
    var isReg = String(r[4]).indexOf("Критерии") === 0;
    var key = isReg ? "reg" : "prog";
    if (!p[key] || r[7] > p[key].pct) p[key] = { pct: r[7], date: r[0], blocks: r[10] };
  });
  var out = Object.keys(people).map(function (k) { return people[k]; }).sort(function (a, b) { return b.last - a.last; });
  var sh = sheet_(ss, "Сводка", ["Имя", "Роль", "Группа", "Критерии, %", "Программа, %", "Статус", "Попыток", "Последняя", "Что подтянуть"]);
  if (sh.getLastRow() > 1) sh.getRange(2, 1, sh.getLastRow() - 1, sh.getLastColumn()).clearContent().setBackground(null);
  if (!out.length) return;
  var data = out.map(function (p) {
    var minp = (p.reg && p.prog) ? Math.min(p.reg.pct, p.prog.pct) : null;
    var weak = [p.reg && p.reg.blocks, p.prog && p.prog.blocks].filter(Boolean).join("; ");
    return [p.name, p.role, p.group, p.reg ? p.reg.pct : "", p.prog ? p.prog.pct : "", minp === null ? "Не все тесты" : verdict_(minp), p.n, p.last, weak];
  });
  var rng = sh.getRange(2, 1, data.length, data[0].length);
  rng.setValues(data);
  // подсветка статуса
  var colors = { "Готов": "#e5f4e4", "Доучить и пересдать": "#f9eeda", "Нужен разбор с управляющим": "#fbebeb", "Не все тесты": "#f0efec" };
  data.forEach(function (r, i) { sh.getRange(i + 2, 6).setBackground(colors[r[5]] || null); });
  sh.getRange(2, 8, data.length, 1).setNumberFormat("dd.mm.yyyy hh:mm");
}

function sheet_(ss, name, header) {
  var sh = ss.getSheetByName(name);
  if (!sh) { sh = ss.insertSheet(name); }
  if (sh.getLastRow() === 0) { sh.appendRow(header); sh.setFrozenRows(1); sh.getRange(1, 1, 1, header.length).setFontWeight("bold"); }
  return sh;
}
function norm_(s) { return String(s || "").trim().toLowerCase().replace(/ё/g, "е").replace(/\s+/g, " "); }
function verdict_(p) { return p >= PASS ? "Готов" : p >= RETAKE ? "Доучить и пересдать" : "Нужен разбор с управляющим"; }
function out_(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }

// Для проверки из редактора: Выполнить → testPost. Добавит тестовую строку.
function testPost() {
  doPost({ postData: { contents: JSON.stringify({ secret: SECRET, name: "Проверка Скрипта", role: "сопровождающий", group: "kids", test: "reg", testName: "Критерии · сопровождающий", score: 27, total: 30, pct: 90, blocks: { "Безопасность": [5, 6], "Чаты": [4, 4] }, wrong: ["Пример вопроса"], dur: 600, ts: Date.now() }) } });
}
