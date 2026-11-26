const OPS = {
  zone: 'Asia/Kolkata', offset: 330 * 60000, day: 86400000,
  first: 6, width: 14, managerWidth: 12, demoClock: '2026-11-26 16:45',
  people: ['Employee A', 'Employee B', 'Employee C', 'Employee D'],
  cycleHeaders: ['Task', 'Due (IST)', 'Status', 'Required proof', 'Enter proof here', 'Cycle ID',
    'Completed at', 'Completed by', 'Reviewed at', 'Reviewed by', 'Opens (IST)', 'Proof type', 'Recorded proof', 'Assignee email'],
  managerHeaders: ['Employee', 'Task', 'Due (IST)', 'Status', 'Manager alert', 'Reviewed at',
    'Reviewed by', 'Leadership', 'Completed at', 'Completed by', 'Proof', 'Cycle ID'],
  logHeaders: ['Alert ID', 'Cycle ID', 'Stage', 'Outcome', 'Recorded at (IST)', 'To', 'Subject', 'Message', 'Detail'],
};

function onOpen() {
  SpreadsheetApp.getUi().createMenu('Allo Ops')
    .addItem('Submit proof for selected row', 'submitProof')
    .addItem('Acknowledge selected manager row', 'acknowledge')
    .addSeparator().addItem('Refresh queues and check alerts', 'runNow')
    .addItem('Initialize / verify workbook', 'initialize')
    .addItem('Install 5-minute live check', 'installClock').addToUi();
}

function initialize() {
  interactive_(() => locked_(() => {
    const book = SpreadsheetApp.getActiveSpreadsheet();
    readState_(book);
    book.setSpreadsheetTimeZone(OPS.zone);
    for (const sheet of [...OPS.people, 'Manager Queue', 'Alert Log']) {
      const tab = book.getSheetByName(sheet);
      tab.setFrozenRows(5);
      tab.setHiddenGridlines(true);
      if (OPS.people.includes(sheet)) {
        tab.getRange(OPS.first, 5, tab.getMaxRows() - 5, 1).setNumberFormat('@');
        tab.hideColumns(11, 4);
      }
    }
    PropertiesService.getScriptProperties().setProperty('BOOK_ID', book.getId());
    evaluateBook_(book);
    return 'Workbook initialized. The Settings tab shows the active mode.';
  }));
}

function installClock() {
  interactive_(() => locked_(() => {
    const book = SpreadsheetApp.getActiveSpreadsheet();
    const state = readState_(book);
    require_(state.config.mode === 'LIVE', 'Set Settings mode to LIVE before installing the clock.');
    const owner = actor_();
    require_(owner === state.config.manager, 'Install the clock from the configured manager account.');
    const props = PropertiesService.getScriptProperties();
    require_(!props.getProperty('CLOCK_OWNER') || props.getProperty('CLOCK_OWNER') === owner,
      'A different account installed the clock. See the Start Here tab for ownership transfer.');
    props.setProperty('BOOK_ID', book.getId());
    props.setProperty('CLOCK_OWNER', owner);
    const triggers = ScriptApp.getProjectTriggers().filter(trigger => trigger.getHandlerFunction() === 'clockCheck');
    if (!triggers.length) ScriptApp.newTrigger('clockCheck').timeBased().everyMinutes(5).create();
    triggers.slice(1).forEach(trigger => ScriptApp.deleteTrigger(trigger));
    return 'Five-minute trigger installed under ' + owner + '.';
  }));
}

function clockCheck() {
  clockCheckBook_(PropertiesService.getScriptProperties().getProperty('BOOK_ID'));
}

function clockCheckBook_(id) {
  locked_(() => {
    require_(id, 'Initialize the workbook first.');
    const book = SpreadsheetApp.openById(id);
    const state = readState_(book);
    if (state.config.mode === 'LIVE') scan_(book, state, true);
  });
}

function runNow() {
  interactive_(() => locked_(() => {
    const book = SpreadsheetApp.getActiveSpreadsheet();
    const state = readState_(book);
    const live = state.config.mode === 'LIVE';
    require_(!live || actor_() === state.config.manager, 'Only the configured manager can check live alerts manually.');
    scan_(book, state, live);
    return 'Queues refreshed and alert rules checked.';
  }));
}

function submitProof() {
  interactive_(() => {
    const book = SpreadsheetApp.getActiveSpreadsheet();
    const selection = book.getActiveRange();
    const person = selection.getSheet().getName();
    require_(OPS.people.includes(person) && selection.getNumRows() === 1 && selection.getRow() >= OPS.first,
      'Select a single task row on your employee tab.');
    const row = selection.getRow();
    // Capture the cycle, input and actor from this selection before taking the lock.
    const selected = selection.getSheet().getRange(row, 5, 1, 2);
    const [proof, id] = selected.getDisplayValues()[0];
    require_(!selected.getCell(1, 1).getFormula(), 'Enter the proof as plain text.');
    const actor = actor_();
    return locked_(() => {
      const state = readState_(book);
      const cycle = state.cycles.find(item => item.id === id && item.person === person);
      require_(cycle, 'Select a task cycle.');
      const proofCell = cycle.sheet.getRange(cycle.row, 5);
      require_(!proofCell.getFormula() && proofCell.getDisplayValue() === proof,
        'The proof changed. Review the current input and submit again.');
      commitCompletion_(book, state, cycle, proof, actor);
      return 'Recorded ' + id + '. The next cycle is tracked automatically.';
    });
  });
}

function acknowledge() {
  interactive_(() => {
    const book = SpreadsheetApp.getActiveSpreadsheet();
    const selection = book.getActiveRange();
    require_(selection.getSheet().getName() === 'Manager Queue' && selection.getNumRows() === 1 &&
      selection.getRow() >= OPS.first, 'Select one row in Manager Queue.');
    const selected = selection.getSheet().getRange(selection.getRow(), 1, 1, OPS.managerWidth).getDisplayValues()[0];
    const id = selected[11];
    require_(id, 'Select a task cycle.');
    const ui = SpreadsheetApp.getUi();
    const answer = ui.alert('Acknowledge review', selected[0] + ' | ' + selected[1] + ' | ' + id, ui.ButtonSet.OK_CANCEL);
    if (answer !== ui.Button.OK) return '';
    const actor = actor_();
    return locked_(() => {
      const state = readState_(book);
      const cycle = state.cycles.find(item => item.id === id);
      require_(cycle, 'This cycle is no longer available.');
      commitReview_(book, state, cycle, actor);
      return 'Review recorded for ' + id + '.';
    });
  });
}

function commitCompletion_(book, state, cycle, proof, actor) {
  complete_(state, cycle, proof, actor);
  // Proof, completion date and actor commit together on the authoritative cycle row.
  cycle.sheet.getRange(cycle.row, 3, 1, 11).setValues([['Completed', cycle.required, '', cycle.id,
    new Date(cycle.completedAt), cycle.completedBy, cycle.reviewedAt === null ? '' : new Date(cycle.reviewedAt),
    cycle.reviewedBy, new Date(cycle.opens), cycle.type, literal_(cycle.proof)]]);
  scan_(book, state);
}

function commitReview_(book, state, cycle, actor) {
  review_(state, cycle, actor);
  cycle.sheet.getRange(cycle.row, 9, 1, 2).setValues([[new Date(cycle.reviewedAt), cycle.reviewedBy]]);
  scan_(book, state);
}

function evaluateBook_(book) {
  scan_(book, readState_(book));
}

function scan_(book, state, sendLive = false) {
  const added = generate_(state);
  for (const name of OPS.people) {
    const additions = added.filter(cycle => cycle.person === name);
    if (!additions.length) continue;
    const sheet = book.getSheetByName(name);
    const first = Math.max(OPS.first, sheet.getLastRow() + 1);
    const range = rows_(sheet, first, additions.length, OPS.width);
    range.setValues(additions.map(cycle => cycleRow_(cycle, state.now)));
    additions.forEach((cycle, index) => { cycle.sheet = sheet; cycle.row = first + index; });
  }
  for (const name of OPS.people) refreshEmployee_(book.getSheetByName(name), state);
  if (state.config.mode === 'DEMO' || sendLive) {
    for (const cycle of state.cycles) {
      const alert = nextAlert_(state, cycle);
      if (alert) deliver_(book, state, cycle, alert);
    }
  }
  writeManager_(book, state);
}

function readState_(book) {
  const requiredSheets = [...OPS.people, 'Manager Queue', 'Task Master', 'Settings', 'Alert Log'];
  requiredSheets.forEach(name => require_(book.getSheetByName(name), 'Missing sheet: ' + name));
  const settings = book.getSheetByName('Settings').getRange('B6:B11').getDisplayValues().flat();
  const config = {
    mode: settings[0].trim().toUpperCase(), manager: settings[1].trim().toLowerCase(),
    leaders: settings[2].trim().toLowerCase().split(',').map(email => email.trim()).filter(Boolean),
    clock: settings[3].trim(), reviewMinutes: Number(settings[4]), resolveMinutes: Number(settings[5]),
  };
  require_(['DEMO', 'LIVE'].includes(config.mode), 'Settings B6 must be DEMO or LIVE.');
  require_(validEmail_(config.manager) && config.leaders.length && config.leaders.every(validEmail_),
    'Enter a manager email and at least one leadership email in Settings.');
  config.leaders = [...new Set(config.leaders)];
  require_(Number.isInteger(config.reviewMinutes) && config.reviewMinutes > 0 &&
    Number.isInteger(config.resolveMinutes) && config.resolveMinutes >= config.reviewMinutes,
    'Enter positive whole SLA minutes, with resolution at least as long as review.');
  const now = config.mode === 'DEMO' ? parseIst_(config.clock) : Date.now();
  const ids = new Set();
  const tasks = data_(book.getSheetByName('Task Master'), 9, true).filter(row => row[0]).map(row => {
    const [id, name, person, email, cadence, firstDay, cutoff, type, required] = row.map(value => String(value).trim());
    const days = cadence === 'Workday' ? 1 : Number(cadence);
    require_(/^[A-Z0-9-]+$/.test(id) && !ids.has(id) && name && OPS.people.includes(person) &&
      validEmail_(email) && Number.isInteger(days) && days >= 1 && days <= 365 &&
      ['Ticket', 'WhatsApp'].includes(type) && required, 'Invalid or duplicate task master entry: ' + id);
    ids.add(id);
    const first = parseIst_(firstDay + ' 00:00');
    parseIst_(firstDay + ' ' + cutoff);
    return { id, name, person, email: email.toLowerCase(), cadence, days, first, cutoff, type, required };
  });
  const seen = new Set();
  const cycles = OPS.people.flatMap(person => {
    const sheet = book.getSheetByName(person);
    return data_(sheet, OPS.width).flatMap((row, index) => {
      if (!row[5]) return [];
      const [name, due, , required, , id, completed, completedBy, reviewed, reviewedBy, opens, type, proof, email] = row;
      require_(!seen.has(id) && due instanceof Date && opens instanceof Date,
        'Invalid or duplicate cycle row: ' + id);
      seen.add(id);
      return [{ name, due: due.getTime(), required, proof: String(proof), id, taskId: id.split('/')[0],
        completedAt: completed instanceof Date ? completed.getTime() : null, completedBy,
        reviewedAt: reviewed instanceof Date ? reviewed.getTime() : null, reviewedBy, opens: opens.getTime(),
        type, person, email, sheet, row: OPS.first + index }];
    });
  });
  const notices = new Map(data_(book.getSheetByName('Alert Log'), 9).filter(row => row[0]).map(row => [row[0], {
    id: row[0], status: row[3], at: row[4] instanceof Date ? row[4].getTime() : null,
  }]));
  require_(config.mode !== 'LIVE' || ![...notices.values()].some(notice => notice.status === 'SIMULATED'),
    'This workbook has simulated alerts. Use a fresh copy for LIVE mode.');
  return { config, now, tasks, cycles, notices };
}

function generate_(state) {
  const added = [];
  const today = dayStart_(state.now);
  for (const task of state.tasks) {
    const existing = state.cycles.filter(cycle => cycle.taskId === task.id);
    const latest = existing.reduce((last, cycle) => !last || cycle.opens > last.opens ? cycle : last, null);
    if (task.cadence === 'Workday') {
      let day = latest ? latest.opens + OPS.day : task.first;
      while (day <= today) {
        const weekday = new Date(day + OPS.offset).getUTCDay();
        if (weekday >= 1 && weekday <= 5) added.push(newCycle_(task, day));
        day += OPS.day;
      }
    } else if (!latest || latest.completedAt !== null) {
      const day = latest ? dayStart_(latest.completedAt) + task.days * OPS.day : task.first;
      added.push(newCycle_(task, day));
    }
  }
  state.cycles.push(...added);
  return added;
}

function newCycle_(task, opens) {
  const day = ist_(opens).slice(0, 10);
  return { id: task.id + '/' + day, taskId: task.id, name: task.name, person: task.person, email: task.email, opens,
    due: parseIst_(day + ' ' + task.cutoff), required: task.required, type: task.type, proof: '',
    completedAt: null, completedBy: '', reviewedAt: null, reviewedBy: '' };
}

function complete_(state, cycle, input, actor) {
  require_(actor, 'Your Google account email is required.');
  require_(state.config.mode === 'DEMO' || cycle.email === actor,
    'Only the assigned employee can submit this cycle.');
  require_(cycle.completedAt === null, 'This cycle is already completed.');
  require_(state.now >= cycle.opens, 'This cycle has not opened yet.');
  const text = input.trim();
  require_(cycle.type === 'WhatsApp' ? /^POSTED$/i.test(text) : /^[A-Za-z0-9][A-Za-z0-9._/-]{0,79}$/.test(text),
    cycle.type === 'WhatsApp' ? 'After posting to WhatsApp, enter POSTED.' : 'Enter a valid portal Ticket ID (up to 80 characters).');
  cycle.proof = cycle.type === 'WhatsApp' ? 'POSTED' : text;
  cycle.completedAt = state.now;
  cycle.completedBy = actor;
}

function review_(state, cycle, actor) {
  require_(actor && (state.config.mode === 'DEMO' || state.config.manager === actor),
    'Only the configured manager can acknowledge in LIVE mode.');
  require_(state.notices.has(cycle.id + '|MANAGER'), 'This cycle has no manager alert to acknowledge.');
  require_(cycle.reviewedAt === null, 'The manager already acknowledged this cycle.');
  cycle.reviewedAt = state.now;
  cycle.reviewedBy = actor;
}

function status_(cycle, now) {
  if (cycle.completedAt !== null) return 'Completed';
  if (now < cycle.opens) return 'Scheduled';
  return now >= cycle.due ? 'Overdue' : 'Pending';
}

function nextAlert_(state, cycle) {
  if (status_(cycle, state.now) !== 'Overdue') return null;
  const manager = state.notices.get(cycle.id + '|MANAGER');
  if (!manager) return { stage: 'MANAGER', reason: 'Task overdue' };
  if (state.notices.has(cycle.id + '|LEADERSHIP') || manager.at === null) return null;
  const elapsed = (state.now - manager.at) / 60000;
  const reviewDeadline = manager.at + state.config.reviewMinutes * 60000;
  if (elapsed >= state.config.reviewMinutes && (cycle.reviewedAt === null || cycle.reviewedAt > reviewDeadline)) {
    return { stage: 'LEADERSHIP', reason: 'Manager acknowledgement overdue and task still open' };
  }
  if (elapsed >= state.config.resolveMinutes) return { stage: 'LEADERSHIP', reason: 'Task still open beyond resolution SLA' };
  return null;
}

function deliver_(book, state, cycle, alert) {
  const stage = alert.stage;
  const id = cycle.id + '|' + stage;
  const demo = state.config.mode === 'DEMO';
  const recipients = stage === 'MANAGER' ? [state.config.manager] : state.config.leaders;
  if (!demo && MailApp.getRemainingDailyQuota() < recipients.length) {
    console.error('Daily mail quota unavailable for ' + id);
    return;
  }
  const managerSheet = book.getSheetByName('Manager Queue');
  const subject = '[Allo Ops] ' + stage + ' | ' + cycle.person + ' | ' + cycle.name;
  const body = [alert.reason, 'Employee: ' + cycle.person, 'Task: ' + cycle.name, 'Cycle: ' + cycle.id,
    'Due (IST): ' + ist_(cycle.due), 'Required proof: ' + cycle.required,
    'Manager review: ' + (cycle.reviewedAt === null ? 'Awaiting acknowledgement' : ist_(cycle.reviewedAt)),
    stage === 'MANAGER' ? 'Contact the employee, then select the cycle in Manager Queue and acknowledge review from Allo Ops.' :
      'Follow up with the manager and employee on the open cycle.',
    book.getUrl() + '#gid=' + managerSheet.getSheetId()].join('\n');
  const sheet = book.getSheetByName('Alert Log');
  const row = Math.max(OPS.first, sheet.getLastRow() + 1);
  const record = { id, status: demo ? 'SIMULATED' : 'CLAIMED', at: state.now };
  rows_(sheet, row, 1, 9).setValues([[id, cycle.id, stage, record.status, new Date(record.at),
    recipients.join(', '), subject, body, '']]);
  sheet.getRange(row, 5).setNumberFormat('dd mmm yyyy HH:mm');
  state.notices.set(id, record);
  SpreadsheetApp.flush();
  if (demo) return;
  try {
    MailApp.sendEmail({ to: recipients.join(','), subject, body });
  } catch (error) {
    record.status = 'CHECK DELIVERY';
    sheet.getRange(row, 4).setValue(record.status);
    sheet.getRange(row, 9).setValue(literal_(String(error)));
    console.error('Delivery requires review for ' + id + ': ' + error);
    return;
  }
  record.status = 'SENT';
  sheet.getRange(row, 4).setValue(record.status);
}

function cycleRow_(cycle, now) {
  return [cycle.name, new Date(cycle.due), status_(cycle, now), cycle.required, '', cycle.id,
    '', '', '', '', new Date(cycle.opens), cycle.type, '', cycle.email];
}

function refreshEmployee_(sheet, state) {
  const count = sheet.getLastRow() - OPS.first + 1;
  const cycles = state.cycles.filter(cycle => cycle.person === sheet.getName());
  const counts = { Pending: 0, Overdue: 0, Scheduled: 0 };
  if (count > 0) {
    const statuses = Array.from({ length: count }, () => ['']);
    const statusesToHide = [];
    for (const cycle of cycles) {
      const status = status_(cycle, state.now);
      statuses[cycle.row - OPS.first] = [status];
      if (status === 'Completed') statusesToHide.push(cycle.row);
      else counts[status]++;
    }
    sheet.getRange(OPS.first, 3, count, 1).setValues(statuses);
    for (let i = 0; i < statusesToHide.length; i++) {
      const first = statusesToHide[i];
      while (statusesToHide[i + 1] === statusesToHide[i] + 1) i++;
      sheet.hideRows(first, statusesToHide[i] - first + 1);
    }
    for (const col of [2, 7, 9, 11]) sheet.getRange(OPS.first, col, count, 1).setNumberFormat('dd mmm yyyy HH:mm');
    sheet.getRange(OPS.first, 5, count, 1).setBackground('#FFF1D5').setNumberFormat('@');
  }
  sheet.getRange('A4').setValue(counts.Pending + ' pending  ·  ' + counts.Overdue + ' overdue  ·  ' + counts.Scheduled + ' scheduled');
  sheet.getRange('C4').setValue(state.config.mode + ' · ' + ist_(state.now) + ' IST');
}

function writeManager_(book, state) {
  const sheet = book.getSheetByName('Manager Queue');
  const rank = { Overdue: 0, Pending: 1, Scheduled: 2, Completed: 3 };
  const cycleMap = new Map(state.cycles.map(cycle => [cycle.id, cycle]));
  // Keep existing IDs on stable rows so concurrent refreshes cannot move a selected cycle.
  const visible = data_(sheet, OPS.managerWidth).map(row => cycleMap.get(row[11])).filter(Boolean);
  const visibleIds = new Set(visible.map(cycle => cycle.id));
  visible.push(...state.cycles.filter(cycle => !visibleIds.has(cycle.id)).sort((a, b) =>
    rank[status_(a, state.now)] - rank[status_(b, state.now)] || a.due - b.due));
  const pending = visible.filter(cycle => status_(cycle, state.now) === 'Pending').length;
  const overdue = visible.filter(cycle => status_(cycle, state.now) === 'Overdue').length;
  const review = visible.filter(cycle => status_(cycle, state.now) === 'Overdue' && cycle.reviewedAt === null &&
    state.notices.has(cycle.id + '|MANAGER')).length;
  sheet.getRange('A4').setValue(pending + ' pending  ·  ' + overdue + ' overdue  ·  ' + review + ' awaiting review');
  sheet.getRange('D4').setValue(state.config.mode + ' · ' + ist_(state.now) + ' IST');
  if (!visible.length) return;
  const time = value => value === null ? '' : new Date(value);
  const notice = (cycle, stage) => {
    const record = state.notices.get(cycle.id + '|' + stage);
    return record ? record.status + ' · ' + ist_(record.at) : '';
  };
  rows_(sheet, OPS.first, visible.length, OPS.managerWidth).setValues(visible.map(cycle => [cycle.person,
    cycle.name, new Date(cycle.due), status_(cycle, state.now), notice(cycle, 'MANAGER'), time(cycle.reviewedAt),
    cycle.reviewedBy, notice(cycle, 'LEADERSHIP'), time(cycle.completedAt), cycle.completedBy,
    cycle.completedAt === null ? '' : literal_(cycle.proof), cycle.id]));
  for (const column of [3, 6, 9]) sheet.getRange(OPS.first, column, visible.length, 1).setNumberFormat('dd mmm yyyy HH:mm');
}

function data_(sheet, width, display) {
  const count = sheet.getLastRow() - OPS.first + 1;
  if (count < 1) return [];
  const range = sheet.getRange(OPS.first, 1, count, width);
  return display ? range.getDisplayValues() : range.getValues();
}

function rows_(sheet, start, count, width) {
  const needed = start + count - 1 - sheet.getMaxRows();
  if (needed > 0) sheet.insertRowsAfter(sheet.getMaxRows(), Math.max(needed, 100));
  return sheet.getRange(start, 1, count, width);
}

function actor_() {
  const email = Session.getActiveUser().getEmail().trim().toLowerCase();
  require_(email, 'Google did not provide your account email. Use an authorized Google Workspace account.');
  return email;
}

function interactive_(action) {
  const book = SpreadsheetApp.getActiveSpreadsheet();
  try {
    const message = action();
    if (message) book.toast(message, 'Allo Ops', 8);
  } catch (error) {
    SpreadsheetApp.getUi().alert(error.message || String(error));
    throw error;
  }
}

function locked_(action) {
  const lock = LockService.getScriptLock();
  lock.waitLock(25000);
  try {
    return action();
  } finally {
    try { SpreadsheetApp.flush(); } finally { lock.releaseLock(); }
  }
}

function require_(condition, message) { if (!condition) throw new Error(message); }
function validEmail_(value) { return /^[^\s@,]+@[^\s@,]+\.[^\s@,]+$/.test(value); }
function literal_(value) { return /^[=+\-@']/.test(value) ? "'" + value : value; }
function dayStart_(time) { return Math.floor((time + OPS.offset) / OPS.day) * OPS.day - OPS.offset; }
function ist_(time) { return new Date(time + OPS.offset).toISOString().replace('T', ' ').slice(0, 16); }

function parseIst_(text) {
  require_(/^\d{4}-\d{2}-\d{2} \d{2}:\d{2}$/.test(text), 'Use YYYY-MM-DD HH:MM (IST): ' + text);
  const time = Date.parse(text.replace(' ', 'T') + ':00+05:30');
  require_(Number.isFinite(time) && ist_(time) === text, 'Invalid IST date or time: ' + text);
  return time;
}
