const CONFIG = {
  spreadsheetId: '1xt-n4ovYRPuSR60PbRVZ6zTdw9jnzGMCv2M6eEpNEqw',
  trialsSheet: 'Prove',
  configSheet: 'Config',
  calendarId: 'riccardo.calli@gmail.com',
  ownerEmail: 'riccardo.calli@gmail.com',
  timezone: 'Europe/Rome',
  startHour: 19,
  endHour: 20,
  endMinute: 30
};

function doGet(e) {
  try {
    const action = (e && e.parameter && e.parameter.action) || 'availability';
    if (action !== 'availability') return json_({ ok: false, error: 'Azione non valida.' });
    return json_({ ok: true, dates: getAvailability_() });
  } catch (err) {
    return json_({ ok: false, error: 'Disponibilità non disponibile.' });
  }
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    const payload = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    const result = createBooking_(payload);
    return json_(result);
  } catch (err) {
    return json_({ ok: false, error: 'Prenotazione non completata.' });
  } finally {
    try { lock.releaseLock(); } catch (_) {}
  }
}

function getAvailability_() {
  const sheet = sheet_();
  const rows = sheet.getDataRange().getValues();
  const counts = {};
  rows.slice(1).forEach(row => {
    const date = normalizeDate_(row[1]);
    const status = String(row[6] || '').toLowerCase();
    if (date && status !== 'annullata' && status !== 'cancellata') counts[date] = (counts[date] || 0) + 1;
  });

  const max = Number(configValue_('Max nuove prove per lezione')) || 10;
  const horizon = Number(configValue_('Durata finestra prenotabile (settimane)')) || 3;
  const today = new Date();
  const end = new Date(today);
  end.setDate(end.getDate() + horizon * 7);
  const dates = [];

  for (let d = new Date(today); d <= end; d.setDate(d.getDate() + 1)) {
    const day = d.getDay();
    if (day !== 2 && day !== 4) continue;
    const key = Utilities.formatDate(d, CONFIG.timezone, 'yyyy-MM-dd');
    const used = counts[key] || 0;
    const remaining = Math.max(0, max - used);
    dates.push({
      date: key,
      label: Utilities.formatDate(d, CONFIG.timezone, 'EEEE d MMMM'),
      remaining: remaining,
      full: remaining === 0,
      status: remaining === 0 ? 'red' : remaining <= 3 ? 'orange' : 'green'
    });
  }
  return dates;
}

function createBooking_(p) {
  const name = clean_(p.name);
  const age = Number(p.age);
  const email = clean_(p.email).toLowerCase();
  const phone = clean_(p.phone);
  const trialDate = clean_(p.date);

  if (!name || !email || !phone || !trialDate || !Number.isInteger(age)) return { ok: false, error: 'Compila tutti i campi richiesti.' };
  if (age < 18) return { ok: false, error: 'Il corso è riservato ai maggiorenni.' };
  if (!/^\d{4}-\d{2}-\d{2}$/.test(trialDate)) return { ok: false, error: 'Data non valida.' };
  if (!/^\S+@\S+\.\S+$/.test(email)) return { ok: false, error: 'Email non valida.' };
  if (!isBookableDay_(trialDate)) return { ok: false, error: 'Questa data non è prenotabile.' };
  if (!isBeforeSameDayCutoff_(trialDate)) return { ok: false, error: 'Le prenotazioni per oggi sono chiuse alle 12:00.' };

  const duplicate = sheet_().getDataRange().getValues().slice(1).some(row =>
    normalizeDate_(row[1]) === trialDate &&
    String(row[2]).trim().toLowerCase() === name.toLowerCase() &&
    String(row[5]).trim().toLowerCase() === email &&
    !['annullata', 'cancellata'].includes(String(row[6] || '').trim().toLowerCase())
  );
  if (duplicate) return { ok: false, error: 'Esiste già una prenotazione per questa persona e questa data.' };

  const availability = getAvailability_().find(x => x.date === trialDate);
  if (!availability || availability.full) return { ok: false, error: 'La lezione è completa o non più disponibile.' };

  const attribution = p.attribution || {};
  const now = new Date();
  sheet_().appendRow([
    now,
    parseDate_(trialDate),
    name,
    age,
    phone,
    email,
    'Prenotata',
    '',
    '',
    '',
    '',
    clean_(attribution.utm_campaign || attribution.campaign || ''),
    clean_(attribution.utm_content || attribution.creative || ''),
    attributionNote_(attribution)
  ]);

  sendEmails_(name, age, phone, email, trialDate);
  updateCalendar_(trialDate);
  return { ok: true, message: 'Prenotazione registrata. Controlla la tua email.' };
}

function sendEmails_(name, age, phone, email, trialDate) {
  const dateLabel = formatDate_(trialDate);
  const spot = getSpot_();
  const spotText = spot.name && spot.name !== 'DA DEFINIRE'
    ? spot.name + ' — ' + spot.address + (spot.maps ? '\nGoogle Maps: ' + spot.maps : '')
    : 'Ti comunicherò il punto di ritrovo preciso prima della prova.';

  const adminSubject = 'NUOVA PROVA – ' + name + ' – ' + dateLabel;
  const adminBody = 'Nome e cognome: ' + name + '\nEtà: ' + age + '\nTelefono: ' + phone + '\nEmail: ' + email + '\nData della prova: ' + dateLabel;
  MailApp.sendEmail(CONFIG.ownerEmail, adminSubject, adminBody);

  const subject = 'Conferma prova Parkour – ' + dateLabel;
  const body = 'Ciao!\n\nGrazie per avermi contattato e per l’interesse verso il corso di Parkour!\n\nLa tua richiesta per partecipare a una prova è stata registrata. Ti aspetto il ' + dateLabel + ' dalle 19:00 alle 20:30.\n\nIl corso è rivolto a persone maggiorenni e non serve avere già esperienza. Gli allenamenti vengono adattati al livello di ciascuno.\n\nIl corso si svolge il martedì e il giovedì dalle 19:00 alle 20:30, con allenamenti itineranti a Padova in zone vicine al centro.\n\nPer la tua prova non devi fare altro: ti aspetto alle 19:00.\n\n' + spotText + '\n\nSe dovessi avere un imprevisto, avvisami rispondendo a questa email.\n\nA presto!\n\nSaiu';
  MailApp.sendEmail({ to: email, subject: subject, body: body, name: 'Riccardo Calli' });
}

function updateCalendar_(trialDate) {
  const cal = CalendarApp.getCalendarById(CONFIG.calendarId);
  if (!cal) return;
  const start = parseDate_(trialDate);
  start.setHours(CONFIG.startHour, 0, 0, 0);
  const end = new Date(start);
  end.setHours(CONFIG.endHour, CONFIG.endMinute, 0, 0);
  const events = cal.getEventsForDay(start).filter(e => /^PROV[AE] – /.test(e.getTitle()));
  const people = sheet_().getDataRange().getValues().slice(1)
    .filter(row => normalizeDate_(row[1]) === trialDate && String(row[6]).toLowerCase() !== 'annullata')
    .map(row => String(row[2]) + ' ' + String(row[3]));
  const title = (people.length === 1 ? 'PROVA – ' : 'PROVE – ') + people.join(' | ');
  const spot = getSpot_();
  const location = spot.name && spot.name !== 'DA DEFINIRE' ? [spot.name, spot.address].filter(Boolean).join(' — ') : '';
  const description = 'Persone in prova: ' + people.join(', ');
  if (events.length) {
    events[0].setTitle(title);
    events[0].setDescription(description);
    if (location) events[0].setLocation(location);
  } else {
    const event = cal.createEvent(title, start, end, { description: description, location: location });
    event.addPopupReminder(45);
  }
}

function sheet_() { return SpreadsheetApp.openById(CONFIG.spreadsheetId).getSheetByName(CONFIG.trialsSheet); }
function configValue_(key) {
  const rows = SpreadsheetApp.openById(CONFIG.spreadsheetId).getSheetByName(CONFIG.configSheet).getDataRange().getValues();
  const row = rows.find(r => String(r[0]).trim() === key);
  return row ? row[1] : '';
}
function getSpot_() {
  return { name: String(configValue_('Spot nome') || 'DA DEFINIRE'), address: String(configValue_('Spot indirizzo') || ''), maps: String(configValue_('Spot Google Maps') || '') };
}
function normalizeDate_(v) { return v instanceof Date && !isNaN(v) ? Utilities.formatDate(v, CONFIG.timezone, 'yyyy-MM-dd') : ''; }
function parseDate_(s) { const p = s.split('-').map(Number); return new Date(p[0], p[1] - 1, p[2]); }
function formatDate_(s) { return Utilities.formatDate(parseDate_(s), CONFIG.timezone, 'EEEE d MMMM yyyy'); }
function isBookableDay_(s) { const d = parseDate_(s); const day = d.getDay(); return day === 2 || day === 4; }
function isBeforeSameDayCutoff_(s) {
  const now = new Date();
  const today = Utilities.formatDate(now, CONFIG.timezone, 'yyyy-MM-dd');
  if (s !== today) return true;
  const cutoff = new Date(now);
  cutoff.setHours(12, 0, 0, 0);
  return now < cutoff;
}
function clean_(v) { return String(v == null ? '' : v).trim().slice(0, 300); }
function attributionNote_(a) { return Object.keys(a).filter(k => k.indexOf('utm_') === 0 || /fbclid|gclid|campaign|creative|term|source|medium/.test(k)).map(k => k + '=' + clean_(a[k])).join(' | '); }
function json_(obj) { return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON); }
