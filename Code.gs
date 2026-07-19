/**
 * Lulu & Muhammad — Wedding invitation backend
 * Tabs: wishes | songs | rsvp | photos   (بتتعمل لوحدها أول مرة)
 * Photos folder: "Lulu-Muhammad-Photos" في الـ Drive بتاعك
 */

const PHOTOS_FOLDER = 'Lulu-Muhammad-Photos';

/* ---------- helpers ---------- */
function ensureSheet(name, headers) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    sh.appendRow(headers);
  }
  return sh;
}
function jsonOut(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
function photosFolder() {
  const it = DriveApp.getFoldersByName(PHOTOS_FOLDER);
  return it.hasNext() ? it.next() : DriveApp.createFolder(PHOTOS_FOLDER);
}

/* ---------- POST ---------- */
function doPost(e) {
  try {
    const body = JSON.parse(e.postData.contents);
    const now = new Date();

    if (body.type === 'wish') {
      const name = String(body.name || '').slice(0, 60).trim();
      const message = String(body.message || '').slice(0, 280).trim();
      if (!name || !message) return jsonOut({ ok: false, error: 'missing fields' });
      ensureSheet('wishes', ['timestamp', 'name', 'message']).appendRow([now, name, message]);
      return jsonOut({ ok: true });
    }

    if (body.type === 'song') {
      const song = String(body.song || '').slice(0, 80).trim();
      if (!song) return jsonOut({ ok: false, error: 'missing song' });
      ensureSheet('songs', ['timestamp', 'song']).appendRow([now, song]);
      return jsonOut({ ok: true });
    }

    if (body.type === 'rsvp') {
      const name = String(body.name || '').slice(0, 60).trim();
      const attending = body.attending === 'yes' ? 'yes' : 'no';
      const count = Math.max(0, Math.min(10, Number(body.count) || 0));
      if (!name) return jsonOut({ ok: false, error: 'missing name' });
      ensureSheet('rsvp', ['timestamp', 'name', 'attending', 'count'])
        .appendRow([now, name, attending, attending === 'yes' ? count : 0]);
      return jsonOut({ ok: true });
    }

    if (body.type === 'photo') {
      const data = String(body.data || '');
      if (!data) return jsonOut({ ok: false, error: 'no data' });
      const filename = ('guest-' + now.getTime() + '-' +
        String(body.filename || 'photo.jpg').replace(/[^\w.\-]/g, '_')).slice(0, 90);
      const blob = Utilities.newBlob(Utilities.base64Decode(data), 'image/jpeg', filename);
      photosFolder().createFile(blob);
      ensureSheet('photos', ['timestamp', 'filename']).appendRow([now, filename]);
      return jsonOut({ ok: true });
    }

    return jsonOut({ ok: false, error: 'unknown type' });
  } catch (err) {
    return jsonOut({ ok: false, error: String(err) });
  }
}

/* ---------- GET ---------- */
function doGet(e) {
  try {
    const type = (e.parameter && e.parameter.type) || '';

    if (type === 'wishes') {
      const sh = ensureSheet('wishes', ['timestamp', 'name', 'message']);
      const rows = sh.getDataRange().getValues().slice(1);   // skip header
      const wishes = rows.reverse().slice(0, 30)             // newest first
        .map(r => ({ name: String(r[1]), message: String(r[2]) }));
      return jsonOut({ ok: true, wishes: wishes });
    }

    if (type === 'songs') {
      const sh = ensureSheet('songs', ['timestamp', 'song']);
      const rows = sh.getDataRange().getValues().slice(1);
      const songs = rows.reverse().slice(0, 20).map(r => String(r[1]));
      return jsonOut({ ok: true, songs: songs });
    }

    if (type === 'rsvp_total') {
      const sh = ensureSheet('rsvp', ['timestamp', 'name', 'attending', 'count']);
      const rows = sh.getDataRange().getValues().slice(1);
      let total = 0;
      rows.forEach(r => { if (r[2] === 'yes') total += Number(r[3]) || 0; });
      return jsonOut({ ok: true, total: total });
    }

    return jsonOut({ ok: true, hello: 'Lulu & Muhammad backend 🤍' });
  } catch (err) {
    return jsonOut({ ok: false, error: String(err) });
  }
}
