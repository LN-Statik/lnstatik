/**
 * SB Kapela – prenos akcií do Google kalendára (účet simpleb4nd@gmail.com)
 *
 * Číta kalendár "SB Kapela" z Outlooku (zverejnený ICS odkaz) a každých 15 minút
 * ho prepíše do Google kalendára "SB Kapela", ktorý je zdieľaný s členmi kapely.
 * Beží na serveroch Google, nezávisí od počítača ani od appky.
 *
 * Nastavenie (raz):
 *  1. Do ICS_URL vlož odkaz ICS zo zverejneného Outlook kalendára "SB Kapela".
 *  2. Hore vyber funkciu "nastavit" a klikni Spustiť (povoľ prístup ku kalendáru).
 *  Hotovo – odteraz sa to synchronizuje samo.
 */
const ICS_URL = 'SEM_VLOZ_ICS_ODKAZ';
const NAZOV_KALENDARA = 'SB Kapela';
const ZONA = 'Europe/Bratislava';

function nastavit() {
  ScriptApp.getProjectTriggers().forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger('synchronizovat').timeBased().everyMinutes(15).create();
  synchronizovat();
}

function kalendar_() {
  const c = CalendarApp.getCalendarsByName(NAZOV_KALENDARA);
  if (c.length) return c[0];
  const n = CalendarApp.createCalendar(NAZOV_KALENDARA, { timeZone: ZONA, color: CalendarApp.Color.ORANGE });
  return n;
}

function synchronizovat() {
  if (ICS_URL.indexOf('http') !== 0) throw new Error('Najprv vlož ICS_URL.');
  const txt = UrlFetchApp.fetch(ICS_URL, { muteHttpExceptions: false }).getContentText();
  const udalosti = parseIcs_(txt);
  const cal = kalendar_();
  const od = new Date(); od.setFullYear(od.getFullYear() - 2);
  const po = new Date(); po.setFullYear(po.getFullYear() + 3);
  const existujuce = {};
  cal.getEvents(od, po).forEach(e => { const u = e.getTag('uid'); if (u) existujuce[u] = e; else if (/^\[SB\]/.test(e.getDescription())) e.deleteEvent(); });
  const videne = {};
  udalosti.forEach(u => {
    if (!u.start || u.start < od || u.start > po) return;
    videne[u.uid] = true;
    const podpis = [u.title, u.start.getTime(), u.end.getTime(), u.allDay, u.location, u.desc].join('|');
    let e = existujuce[u.uid];
    if (e && e.getTag('sig') === podpis) return;
    if (e) e.deleteEvent();
    e = u.allDay ? cal.createAllDayEvent(u.title, u.start, u.end, { location: u.location, description: u.desc })
                 : cal.createEvent(u.title, u.start, u.end, { location: u.location, description: u.desc });
    e.setTag('uid', u.uid); e.setTag('sig', podpis);
    e.removeAllReminders();
    if (!/^ZRUŠENÉ/.test(u.title)) e.addPopupReminder(24 * 60);
  });
  Object.keys(existujuce).forEach(uid => { if (!videne[uid]) existujuce[uid].deleteEvent(); });
}

function parseIcs_(txt) {
  const lines = txt.replace(/\r\n[ \t]/g, '').replace(/\n[ \t]/g, '').split(/\r?\n/);
  const out = []; let cur = null;
  const un = s => (s || '').replace(/\\n/gi, '\n').replace(/\\,/g, ',').replace(/\\;/g, ';').replace(/\\\\/g, '\\');
  lines.forEach(l => {
    if (l === 'BEGIN:VEVENT') { cur = {}; return; }
    if (l === 'END:VEVENT') { if (cur && cur.uid) out.push(cur); cur = null; return; }
    if (!cur) return;
    const i = l.indexOf(':'); if (i < 0) return;
    const head = l.slice(0, i), val = l.slice(i + 1);
    const name = head.split(';')[0].toUpperCase();
    if (name === 'UID') cur.uid = val;
    else if (name === 'SUMMARY') cur.title = un(val);
    else if (name === 'LOCATION') cur.location = un(val);
    else if (name === 'DESCRIPTION') cur.desc = un(val);
    else if (name === 'DTSTART' || name === 'DTEND') {
      const allDay = /VALUE=DATE(?!-)/i.test(head) || /^\d{8}$/.test(val);
      const d = datum_(val, allDay);
      if (name === 'DTSTART') { cur.start = d; cur.allDay = allDay; } else cur.end = d;
    }
  });
  out.forEach(u => {
    u.title = u.title || 'Akcia'; u.location = u.location || ''; u.desc = '[SB] ' + (u.desc || '').trim();
    if (!u.end) { u.end = new Date(u.start.getTime() + (u.allDay ? 864e5 : 6 * 36e5)); }
  });
  return out;
}

function datum_(v, allDay) {
  const y = +v.slice(0, 4), m = +v.slice(4, 6) - 1, d = +v.slice(6, 8);
  if (allDay) return new Date(y, m, d);
  const hh = +v.slice(9, 11), mi = +v.slice(11, 13), ss = +v.slice(13, 15) || 0;
  if (/Z$/.test(v)) return new Date(Date.UTC(y, m, d, hh, mi, ss));
  // miestny čas (Outlook používa stredoeurópsky čas)
  const s = Utilities.formatString('%04d-%02d-%02dT%02d:%02d:%02d', y, m + 1, d, hh, mi, ss);
  const off = Utilities.formatDate(new Date(Date.UTC(y, m, d, 12)), ZONA, 'XXX');
  return new Date(s + off);
}
