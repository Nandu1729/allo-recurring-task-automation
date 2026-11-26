import fs from 'node:fs/promises';
import vm from 'node:vm';
import assert from 'node:assert/strict';

const source = await fs.readFile(new URL('./Allo Ops.gs', import.meta.url), 'utf8');
const results = [];
const emails = [];
const alerts = [];
const props = new Map();
let book, selected, actor = 'reviewer@example.test', time, handlers = [], failSend = false, advanceOnQuota = '';
let locked = false;
const format = t => new Date(+t + 330 * 60000).toISOString().slice(0, 16).replace('T', ' ');
const date = s => new Date(s.replace(' ', 'T') + ':00+05:30');
class Clock extends Date { static now() { return +date(time); } }

class Range {
  constructor(sheet, row, col, height = 1, width = 1) {
    Object.assign(this, { sheet, row, col, height, width });
  }
  getSheet() { return this.sheet; }
  getRow() { return this.row; }
  getNumRows() { return this.height; }
  getCell(r, c) { return new Range(this.sheet, this.row + r - 1, this.col + c - 1); }
  getValues() {
    return Array.from({ length: this.height }, (_, r) => Array.from({ length: this.width }, (_, c) =>
      this.sheet.cells.get(`${this.row + r},${this.col + c}`) ?? ''));
  }
  getDisplayValues() {
    return this.getValues().map(row => row.map(value => value instanceof Date ? format(value) : String(value)));
  }
  getFormula() { return ''; }
  getDisplayValue() { return this.getDisplayValues()[0][0]; }
  setValues(values) {
    assert.equal(values.length, this.height);
    for (const [r, row] of values.entries()) {
      assert.equal(row.length, this.width);
      for (const [c, value] of row.entries()) this.sheet.cells.set(`${this.row + r},${this.col + c}`, value);
    }
    return this;
  }
  setValue(value) { return this.setValues([[value]]); }
  setBackground() { return this; }
  setNumberFormat() { return this; }
}

class Sheet {
  constructor(name, id) { this.name = name; this.id = id; this.cells = new Map(); this.hidden = new Set(); }
  getSheetId() { return this.id; }
  getName() { return this.name; }
  getLastRow() {
    return Math.max(5, ...[...this.cells].filter(([, value]) => value !== '').map(([key]) => Number(key.split(',')[0])));
  }
  getMaxRows() { return 1000; }
  setFrozenRows() {}
  setHiddenGridlines() {}
  hideColumns() {}
  hideRows(start, count) { for (let i = 0; i < count; i++) this.hidden.add(start + i); }
  getRange(row, col, height, width) {
    if (typeof row === 'number') return new Range(this, row, col, height, width);
    const parts = row.replaceAll('$', '').split(':').map(address => /^([A-Z])(\d+)$/.exec(address));
    const [, column, start] = parts[0];
    return new Range(this, +start, column.charCodeAt(0) - 64, parts[1] ? +parts[1][2] - +start + 1 : 1);
  }
}

const context = vm.createContext({ Date: Clock, console: { error: () => {} },
  Session: { getActiveUser: () => ({ getEmail: () => actor }) },
  SpreadsheetApp: {
    getActiveSpreadsheet: () => book, openById: () => book, flush: () => {},
    getUi: () => ({ alert: (...args) => { alerts.push(args[0]); return 'OK'; }, ButtonSet: { OK_CANCEL: 1 }, Button: { OK: 'OK' } }),
  },
  PropertiesService: { getScriptProperties: () => ({ getProperty: key => props.get(key), setProperty: (k, v) => props.set(k, v) }) },
  LockService: { getScriptLock: () => ({ waitLock: () => { assert(!locked); locked = true; }, releaseLock: () => { locked = false; } }) },
  MailApp: { getRemainingDailyQuota: () => {
    if (advanceOnQuota) {
      time = advanceOnQuota;
      advanceOnQuota = '';
    }
    return 50;
  }, sendEmail: mail => {
    assert(locked);
    const latest = records('Alert Log', 9).at(-1);
    assert.equal(latest[3], 'CLAIMED');
    emails.push(mail);
    if (failSend) throw new Error('Mock transport error');
  } },
  ScriptApp: { getProjectTriggers: () => handlers.slice(), newTrigger: name => ({ timeBased() { return this; },
    everyMinutes(n) { assert.equal(n, 5); return this; }, create() { handlers.push({ getHandlerFunction: () => name }); } }),
  deleteTrigger: trigger => { handlers = handlers.filter(item => item !== trigger); } },
});
vm.runInContext(source + '\nglobalThis.logic = { OPS, initialize, runNow, submitProof, acknowledge, installClock, clockCheck };', context);
const { OPS, initialize, runNow, submitProof, acknowledge, installClock, clockCheck } = context.logic;
const people = [...OPS.people];
const master = [
  ['DAILY', 'Daily Operations Update', people[0], 'employee.a@example.test', 'Workday', '2026-11-26', '17:00', 'WhatsApp', 'Type POSTED after sharing the update'],
  ['PORTAL-A', 'Portal Task A', people[1], 'employee.b@example.test', 7, '2026-11-26', '17:00', 'Ticket', 'Portal Ticket ID'],
  ['PORTAL-B', 'Portal Task B', people[2], 'employee.c@example.test', 15, '2026-11-26', '17:00', 'Ticket', 'Portal Ticket ID'],
  ['PORTAL-C', 'Portal Task C', people[3], 'employee.d@example.test', 30, '2026-11-26', '17:00', 'Ticket', 'Portal Ticket ID'],
];
const defaults = ['DEMO', 'manager@example.test', 'leadership@example.test', OPS.demoClock, 60, 180];

function reset(mode = 'DEMO') {
  const tabs = new Map([...people, 'Manager Queue', 'Task Master', 'Settings', 'Alert Log'].map((name, i) => [name, new Sheet(name, i)]));
  book = { getSheetByName: name => tabs.get(name), getId: () => 'test-book', getUrl: () => 'https://docs.google.com/spreadsheets/d/test-book/edit',
    getActiveRange: () => selected, toast() {}, setSpreadsheetTimeZone() {} };
  props.clear(); emails.length = 0; handlers = []; failSend = false; advanceOnQuota = '';
  actor = 'reviewer@example.test';
  time = OPS.demoClock;
  book.getSheetByName('Settings').getRange('B6:B11').setValues([mode, ...defaults.slice(1)].map(value => [value]));
  book.getSheetByName('Task Master').getRange(6, 1, 4, 9).setValues(master);
  initialize();
}

function records(name, width = OPS.width) {
  const sheet = book.getSheetByName(name);
  return sheet.getRange(6, 1, sheet.getLastRow() - 5, width).getValues();
}
function move(clock, scan = runNow) {
  time = clock;
  book.getSheetByName('Settings').getRange('B9').setValue(clock);
  scan();
}
function enter(person, proof, row = 6) {
  selected = book.getSheetByName(person).getRange(row, 5);
  selected.setValue(proof);
  submitProof();
}
function ack(id) {
  const index = records('Manager Queue', 12).findIndex(row => row[11] === id);
  assert(index >= 0);
  selected = book.getSheetByName('Manager Queue').getRange(index + 6, 1);
  acknowledge();
}
function log(id, stage) { return records('Alert Log', 9).find(row => row[0] === id + '|' + stage); }
function test(label, action, evidence) {
  action();
  results.push([label, 'PASS', evidence]);
}

reset();
test('1 · Recurring cycle', () => {
  enter(people[1], 'TKT-7001');
  const [closed, next] = records(people[1]);
  assert.equal(closed[2], 'Completed');
  assert.equal(next[2], 'Scheduled');
  assert.equal(format(next[1]), '2026-12-03 17:00');
  move('2026-12-03 09:00');
  assert.equal(records(people[1])[1][2], 'Pending');
}, 'Portal A closed on 26 Nov. Its next cycle was scheduled for 3 Dec, 17:00 IST, then became pending on 3 Dec.');

reset();
test('2 · Single-point completion', () => {
  enter(people[1], 'TKT-7001');
  const row = records(people[1])[0];
  assert.equal(row[4], ''); assert.equal(row[12], 'TKT-7001');
  assert.equal(row[7], actor); assert.equal(format(row[6]), OPS.demoClock);
  assert(book.getSheetByName(people[1]).hidden.has(6));
  assert.equal(records('Manager Queue', 12).find(item => item[11] === row[5])[10], row[12]);
}, 'One proof entry on Employee B was archived with actor and time. The manager view updated automatically.');

test('3 · Daily proof', () => {
  assert.throws(() => enter(people[0], 'yes'), /enter POSTED/);
  enter(people[0], 'POSTED');
  assert.equal(records(people[0])[0][12], 'POSTED');
  move('2026-11-30 09:00');
  assert.deepEqual(records(people[0]).map(row => row[5]), ['DAILY/2026-11-26','DAILY/2026-11-27','DAILY/2026-11-30']);
}, 'Validated POSTED, recorded the account and timestamp, opened Friday and Monday cycles, and skipped the weekend.');

reset();
enter(people[0], 'POSTED');
enter(people[1], 'TKT-7001');
const portalB = 'PORTAL-B/2026-11-26', portalC = 'PORTAL-C/2026-11-26';
test('4 · Same-day manager alert', () => {
  assert.equal(records('Alert Log', 9).length, 0);
  move('2026-11-26 17:00');
  const row = log(portalB, 'MANAGER');
  assert.equal(row[3], 'SIMULATED'); assert.equal(format(row[4]), '2026-11-26 17:00');
  assert.equal(row[5], 'manager@example.test'); assert.match(row[7], /Employee C.*\nTask: Portal Task B/);
}, 'At 17:00 IST the scan logged manager emails for Portal B and C. The payload included employee, task, due time and a queue link.');

test('5 · Manager acknowledgement', () => {
  move('2026-11-26 17:15'); ack(portalB);
  const reviewed = records('Manager Queue', 12).find(row => row[11] === portalB);
  assert.equal(reviewed[3], 'Overdue'); assert.equal(reviewed[6], actor);
  assert.equal(format(reviewed[5]), '2026-11-26 17:15');
  move('2026-11-26 17:30'); enter(people[2], 'TKT-15001');
  const resolved = records('Manager Queue', 12).find(row => row[11] === portalB);
  assert.equal(resolved[3], 'Completed'); assert.equal(format(resolved[8]), '2026-11-26 17:30');
  assert.equal(format(records(people[2])[1][1]), '2026-12-11 17:00');
}, 'Portal B stayed overdue after the 17:15 review and changed to Completed when ticket proof arrived at 17:30.');

test('6 · Missing-review escalation', () => {
  move('2026-11-26 17:59'); assert(!log(portalC, 'LEADERSHIP'));
  move('2026-11-26 18:00');
  assert.match(log(portalC, 'LEADERSHIP')[7], /acknowledgement overdue/);
  assert(!log(portalB, 'LEADERSHIP'));
}, 'At 18:00 IST the unresolved and unreviewed Portal C escalated. Resolved Portal B generated no leadership email.');

test('7 · Reviewed but unresolved escalation', () => {
  reset(); enter(people[0], 'POSTED'); enter(people[1], 'TKT-7001'); enter(people[2], 'TKT-15001');
  move('2026-11-26 17:00'); move('2026-11-26 17:15'); ack(portalC);
  move('2026-11-26 18:00'); assert(!log(portalC, 'LEADERSHIP'));
  move('2026-11-26 19:59'); assert(!log(portalC, 'LEADERSHIP'));
  move('2026-11-26 20:00'); assert.match(log(portalC, 'LEADERSHIP')[7], /resolution SLA/);
  enter(people[3], 'TKT-30001'); assert.equal(format(records(people[3])[1][1]), '2026-12-26 17:00');
}, 'With a timely review, Portal C skipped 18:00 escalation and escalated at 20:00 IST because it was still open.');

test('8 · Repeated runs and submissions', () => {
  const counts = [...people, 'Alert Log'].map(name => records(name, name === 'Alert Log' ? 9 : OPS.width).length);
  runNow(); runNow();
  assert.throws(() => enter(people[1], 'TKT-DUPLICATE'), /already completed/);
  assert.equal(records(people[1])[0][12], 'TKT-7001');
  assert.throws(() => ack(portalC), /already acknowledged/);
  assert.deepEqual([...people, 'Alert Log'].map(name => records(name, name === 'Alert Log' ? 9 : OPS.width).length), counts);
  assert.equal(emails.length, 0);
}, 'Two repeat scans created zero new cycle rows or alerts. Repeated completion and acknowledgement were rejected.');

test('Additional · Live transport and access', () => {
  reset('LIVE');
  assert.throws(() => enter(people[1], 'TKT-7001'), /Only the assigned/);
  actor = 'employee.b@example.test'; enter(people[1], 'TKT-7001');
  actor = 'manager@example.test'; installClock(); installClock(); assert.equal(handlers.length, 1);
  time = '2026-11-26 17:00'; advanceOnQuota = '2026-11-26 17:03'; clockCheck(); clockCheck();
  assert.equal(emails.length, 3);
  assert.equal(log(portalC, 'MANAGER')[3], 'SENT');
  assert.equal(format(log(portalC, 'MANAGER')[4]), '2026-11-26 17:03');
  actor = 'employee.d@example.test'; assert.throws(() => ack(portalC), /Only the configured manager/);
  time = '2026-11-26 18:02'; clockCheck(); assert.equal(emails.length, 3);
  failSend = true; time = '2026-11-26 18:03'; clockCheck(); clockCheck();
  assert.equal(emails.length, 6);
  assert.equal(log(portalC, 'LEADERSHIP')[3], 'CHECK DELIVERY');
}, 'Mock clock installed once. Live mode enforced account roles, claimed each email before sending and flagged ambiguous delivery.');

test('Additional · Recipient limit', () => {
  reset();
  const recipients = Array.from({ length: 50 }, (_, i) => `leadership${i}@example.test`);
  const cell = book.getSheetByName('Settings').getRange('B8');
  cell.setValue([...recipients, recipients[0]].join(','));
  runNow();
  cell.setValue([...recipients, 'overflow@example.test'].join(','));
  assert.throws(() => runNow(), /no more than 50 distinct leadership/);
  assert.equal(emails.length, 0);
}, 'The setting accepts 50 distinct recipients and rejects 51 before any alert can be claimed.');

console.table(results.map(([scenario, outcome]) => ({ scenario, outcome })));
