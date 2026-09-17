/**
 * VISITKORT – Google Apps Script API  (v3)
 * Kører usynligt bag GitHub Pages. Ingen HTML her – kun JSON og vCard.
 * Opsætning: se README.md
 */
const CFG = {
  PUBLIC_BASE: 'https://bernt84.github.io/visitkort/', // adressen på GitHub-siden (med / til sidst)
  FIRST_ADMIN: 'martin',
  FIRST_ADMIN_PW: 'SkiftMig123!',                      // skal skiftes ved første login
  SESSION_SEC: 21600,                                  // 6 timer (max i CacheService)
  MAX_FAIL: 5,
  LOCK_SEC: 900,
  PHOTO_MAX: 48000,                                    // tegn (en celle i Sheets kan max 50.000)
  MAX_LINKS: 4,
  MAX_CARDS: 10,                                       // profiler pr. bruger
  FOLDER_NAME: 'Visitkort',                            // mappe i Google Drive (oprettes automatisk)
  SHEET_NAME: 'Visitkort data',                        // regnearket (oprettes automatisk i mappen)
  BACKUP_KEEP: 14                                      // antal daglige backups der gemmes
};

const SH_USERS = 'Brugere';
const SH_CARDS = 'Kort';
const U_HEAD = ['username','name','hash','salt','role','active','mustChange','created','lastLogin'];
const C_HEAD = ['slug','firstName','lastName','title','company','phone','email','web','linkedin',
                'address','note','color','updated','updatedBy','theme','photo',
                'links','hidden','views','saves','lastView','owner','label','sort'];
const TEXT_FIELDS = ['firstName','lastName','title','company','phone','email','web','linkedin','address','note'];
const PUBLIC_FIELDS = ['slug','firstName','lastName','title','company','phone','email','web','linkedin',
                       'address','note','color','theme','photo','links'];
const FONTS = ['Archivo','Space Grotesk','IBM Plex Sans','Source Serif 4','Barlow Condensed'];

/* ================= Kør én gang ================= */
// Opretter mappen "Visitkort" i Google Drive, regnearket i mappen,
// arkene, din admin-bruger og den daglige backup. Kan køres igen uden skade.
function setup() {
  const folder = getFolder_();
  const props = PropertiesService.getScriptProperties();
  let id = props.getProperty('SHEET_ID');
  let file = null;
  if (id) { try { file = DriveApp.getFileById(id); if (file.isTrashed()) file = null; } catch (e) { file = null; } }

  let isNew = false;
  if (!file) {
    const ss = SpreadsheetApp.create(CFG.SHEET_NAME);
    id = ss.getId();
    props.setProperty('SHEET_ID', id);
    file = DriveApp.getFileById(id);
    isNew = true;
  }
  // sørg for at regnearket ligger i mappen
  let inFolder = false;
  const parents = file.getParents();
  while (parents.hasNext()) if (parents.next().getId() === folder.getId()) inFolder = true;
  if (!inFolder) file.moveTo(folder);

  sheet_(SH_USERS, U_HEAD);
  sheet_(SH_CARDS, C_HEAD);
  if (isNew) {
    const ss = ss_();
    ss.getSheets().forEach(sh => {
      if ([SH_USERS, SH_CARDS].indexOf(sh.getName()) < 0 && ss.getSheets().length > 2) ss.deleteSheet(sh);
    });
  }

  if (!findUser_(CFG.FIRST_ADMIN)) {
    createUserRow_(CFG.FIRST_ADMIN, 'Martin Bernt Christoffersen', 'admin', CFG.FIRST_ADMIN_PW);
    const c = findCard_(CFG.FIRST_ADMIN);
    c.title = 'QHSE Manager'; c.company = 'Nicon Industries'; c.address = 'Esbjerg, Denmark'; c.label = 'Arbejde';
    writeRow_(SH_CARDS, C_HEAD, c);
  }
  setupBackup_();
  Logger.log('Mappe: %s', folder.getUrl());
  Logger.log('Regneark: %s', ss_().getUrl());
  Logger.log('Klar. Log ind som "%s" med "%s"', CFG.FIRST_ADMIN, CFG.FIRST_ADMIN_PW);
}

/* ================= Mappe og backup ================= */
function getFolder_() {
  const props = PropertiesService.getScriptProperties();
  const fid = props.getProperty('FOLDER_ID');
  if (fid) { try { const f = DriveApp.getFolderById(fid); if (!f.isTrashed()) return f; } catch (e) {} }
  const found = DriveApp.getRootFolder().getFoldersByName(CFG.FOLDER_NAME);
  const folder = found.hasNext() ? found.next() : DriveApp.getRootFolder().createFolder(CFG.FOLDER_NAME);
  props.setProperty('FOLDER_ID', folder.getId());
  return folder;
}

function backupFolder_() {
  const main = getFolder_();
  const it = main.getFoldersByName('Backup');
  return it.hasNext() ? it.next() : main.createFolder('Backup');
}

// Daglig kopi af regnearket i Visitkort/Backup. Ældre kopier slettes.
function backup() {
  const folder = backupFolder_();
  const stamp = Utilities.formatDate(new Date(), 'Europe/Copenhagen', 'yyyy-MM-dd');
  const name = 'Visitkort backup ' + stamp;
  const same = folder.getFilesByName(name);
  while (same.hasNext()) same.next().setTrashed(true);
  DriveApp.getFileById(ss_().getId()).makeCopy(name, folder);

  const files = [];
  const all = folder.getFiles();
  while (all.hasNext()) {
    const f = all.next();
    if (f.getName().indexOf('Visitkort backup ') === 0) files.push(f);
  }
  files.sort((a, b) => b.getName().localeCompare(a.getName()));
  files.slice(CFG.BACKUP_KEEP).forEach(f => f.setTrashed(true));
}

function setupBackup_() {
  const exists = ScriptApp.getProjectTriggers().some(t => t.getHandlerFunction() === 'backup');
  if (!exists) ScriptApp.newTrigger('backup').timeBased().everyDays(1).atHour(3).create();
}

/* ================= HTTP ================= */
// GET  ?c=slug    -> kortets data som JSON (offentligt)
// GET  ?vcf=slug  -> kontaktfil (offentligt)
function doGet(e) {
  const p = (e && e.parameter) || {};
  try {
    if (p.vcf) {
      const c = publicCard_(p.vcf);
      if (!c) return ContentService.createTextOutput('Card not found.');
      bump_(c, 'saves');
      return ContentService.createTextOutput(vcard_(c)).setMimeType(ContentService.MimeType.VCARD);
    }
    if (p.c) {
      const c = publicCard_(p.c);
      if (!c) return json_({ ok: false, error: 'Kortet findes ikke.' });
      if (p.count !== '0') bump_(c, 'views');
      return json_({ ok: true, data: cardOut_(c, true) });
    }
    return json_({ ok: true, data: 'visitkort-api' });
  } catch (err) {
    return json_({ ok: false, error: String(err.message || err) });
  }
}

// POST body: {"action":"...","token":"...","p":{...}}   (Content-Type: text/plain)
function doPost(e) {
  try {
    const body = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    return json_({ ok: true, data: api_(String(body.action || ''), body.token, body.p || {}) });
  } catch (err) {
    return json_({ ok: false, error: String(err.message || err) });
  }
}

function json_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}

/* ================= API ================= */
function api_(action, token, p) {
  if (action === 'login') return login_(p.username, p.password);

  const me = session_(token);
  const isAdmin = me.role === 'admin';
  const needAdmin = () => { if (!isAdmin) throw new Error('Kun administrator har adgang.'); };
  if (me.mustChange === 'TRUE' && ['changePassword','logout','me'].indexOf(action) < 0)
    throw new Error('Skift din adgangskode først.');

  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    switch (action) {
      case 'me': return pubUser_(me);

      case 'logout':
        CacheService.getScriptCache().remove('s_' + token);
        return true;

      case 'changePassword': {
        if (hash_(p.oldPassword || '', me.salt) !== me.hash) throw new Error('Nuværende adgangskode er forkert.');
        checkPw_(p.newPassword);
        me.salt = Utilities.getUuid();
        me.hash = hash_(p.newPassword, me.salt);
        me.mustChange = 'FALSE';
        writeRow_(SH_USERS, U_HEAD, me);
        return pubUser_(me);
      }

      case 'listCards': {
        const owner = clean_(p.owner) || me.username;
        if (owner !== me.username) needAdmin();
        if (!findUser_(owner)) throw new Error('Brugeren findes ikke.');
        return cardsOf_(owner).map(c => cardOut_(c, false));
      }

      case 'getCard': {
        const c = p.slug ? findCard_(clean_(p.slug)) : cardsOf_(me.username)[0];
        if (!c) throw new Error('Profilen findes ikke.');
        if (!canEdit_(me, c)) throw new Error('Du har ikke adgang til denne profil.');
        return cardOut_(c, false);
      }

      case 'saveCard': {
        const c = findCard_(clean_(p.slug));
        if (!c) throw new Error('Profilen findes ikke.');
        if (!canEdit_(me, c)) throw new Error('Du har ikke adgang til denne profil.');
        const d = p.data || {};
        TEXT_FIELDS.forEach(k => { c[k] = String(d[k] == null ? '' : d[k]).trim().slice(0, k === 'note' ? 500 : 200); });
        if (!c.firstName && !c.lastName) throw new Error('Skriv mindst et navn.');
        c.web = fixUrl_(c.web);
        c.linkedin = fixLinkedIn_(c.linkedin);
        if (c.email && !/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(c.email)) throw new Error('Mailadressen ser forkert ud.');
        c.label = cleanLabel_(d.label) || c.label || 'Profil';
        c.owner = ownerOf_(c);
        c.theme = cleanTheme_(d.theme);
        c.links = cleanLinks_(d.links);
        c.hidden = d.hidden ? 'TRUE' : 'FALSE';
        if ('photo' in d) {
          const ph = String(d.photo || '');
          if (ph && (ph.indexOf('data:image/jpeg;base64,') !== 0 || ph.length > CFG.PHOTO_MAX))
            throw new Error('Billedet kunne ikke gemmes. Prøv et andet billede.');
          c.photo = ph;
        }
        c.updated = new Date().toISOString();
        c.updatedBy = me.username;
        writeRow_(SH_CARDS, C_HEAD, c);
        return cardOut_(c, false);
      }

      case 'createCard': {
        const owner = clean_(p.owner) || me.username;
        if (owner !== me.username) needAdmin();
        const u = findUser_(owner);
        if (!u) throw new Error('Brugeren findes ikke.');
        const list = cardsOf_(owner);
        if (list.length >= CFG.MAX_CARDS) throw new Error('Du kan højst have ' + CFG.MAX_CARDS + ' profiler.');
        const label = cleanLabel_(p.label);
        if (!label) throw new Error('Giv profilen et navn.');

        let slug = clean_(p.slug) || slugify_(owner + '-' + label);
        if (!/^[a-z0-9-]{2,40}$/.test(slug)) throw new Error('Link-navn: 2-40 tegn, kun a-z, 0-9 og bindestreg.');
        const taken = x => !!(findCard_(x) || findUser_(x));
        if (taken(slug)) {
          if (p.slug) throw new Error('Link-navnet er allerede brugt. Vælg et andet.');
          let i = 2;
          while (taken(slug + '-' + i)) i++;
          slug = slug + '-' + i;
        }

        const c = {};
        const src = p.copyFrom ? findCard_(clean_(p.copyFrom)) : null;
        if (src) {
          if (!canEdit_(me, src)) throw new Error('Du har ikke adgang til profilen, der skal kopieres.');
          C_HEAD.forEach(k => c[k] = src[k] || '');
        } else {
          const parts = String(u.name || '').split(/\s+/);
          c.lastName = parts.length > 1 ? parts.pop() : '';
          c.firstName = parts.join(' ');
          c.theme = ''; c.photo = ''; c.links = '[]';
        }
        Object.assign(c, {
          slug: slug, owner: owner, label: label, sort: String(list.length),
          hidden: 'FALSE', views: '0', saves: '0', lastView: '',
          updated: new Date().toISOString(), updatedBy: me.username
        });
        writeRow_(SH_CARDS, C_HEAD, c);
        return cardOut_(c, false);
      }

      case 'deleteCard': {
        const c = findCard_(clean_(p.slug));
        if (!c) throw new Error('Profilen findes ikke.');
        if (!canEdit_(me, c)) throw new Error('Du har ikke adgang til denne profil.');
        if (cardsOf_(ownerOf_(c)).length <= 1) throw new Error('Der skal være mindst én profil.');
        sheet_(SH_CARDS, C_HEAD).deleteRow(c._row);
        return true;
      }

      case 'listUsers': {
        needAdmin();
        const byOwner = {};
        rows_(SH_CARDS, C_HEAD).forEach(c => {
          const o = ownerOf_(c);
          (byOwner[o] = byOwner[o] || []).push(c);
        });
        return rows_(SH_USERS, U_HEAD).map(u => {
          const o = pubUser_(u);
          const list = byOwner[u.username] || [];
          o.cards = list.length;
          o.labels = list.map(c => c.label || 'Profil');
          o.title = list[0] ? list[0].title : '';
          o.hasPhoto = list.some(c => c.photo);
          o.views = list.reduce((n, c) => n + Number(c.views || 0), 0);
          o.saves = list.reduce((n, c) => n + Number(c.saves || 0), 0);
          o.publicUrl = list[0] ? publicUrl_(list[0].slug) : '';
          return o;
        });
      }

      case 'createUser': {
        needAdmin();
        const username = clean_(p.username);
        if (!/^[a-z0-9-]{2,30}$/.test(username)) throw new Error('Brugernavn: 2-30 tegn, kun a-z, 0-9 og bindestreg.');
        if (findUser_(username) || findCard_(username)) throw new Error('Brugernavnet er allerede brugt.');
        const name = String(p.name || '').trim().slice(0, 100);
        if (!name) throw new Error('Skriv brugerens navn.');
        const pw = tempPw_();
        createUserRow_(username, name, p.role === 'admin' ? 'admin' : 'user', pw);
        return { username: username, tempPassword: pw, loginUrl: CFG.PUBLIC_BASE };
      }

      case 'updateUser': {
        needAdmin();
        const u = findUser_(clean_(p.username));
        if (!u) throw new Error('Brugeren findes ikke.');
        const self = u.username === me.username;
        if ('role' in p) {
          if (self && p.role !== 'admin') throw new Error('Du kan ikke fjerne din egen admin-rolle.');
          u.role = p.role === 'admin' ? 'admin' : 'user';
        }
        if ('active' in p) {
          if (self && !p.active) throw new Error('Du kan ikke deaktivere dig selv.');
          u.active = p.active ? 'TRUE' : 'FALSE';
          if (!p.active) killSessions_(u.username);
        }
        if ('name' in p && String(p.name).trim()) u.name = String(p.name).trim().slice(0, 100);
        writeRow_(SH_USERS, U_HEAD, u);
        return pubUser_(u);
      }

      case 'resetPassword': {
        needAdmin();
        const u = findUser_(clean_(p.username));
        if (!u) throw new Error('Brugeren findes ikke.');
        const pw = tempPw_();
        u.salt = Utilities.getUuid();
        u.hash = hash_(pw, u.salt);
        u.mustChange = 'TRUE';
        writeRow_(SH_USERS, U_HEAD, u);
        killSessions_(u.username);
        CacheService.getScriptCache().remove('f_' + u.username);
        return { username: u.username, tempPassword: pw, loginUrl: CFG.PUBLIC_BASE };
      }

      case 'resetStats': {
        needAdmin();
        cardsOf_(clean_(p.username)).forEach(c => {
          setCell_(SH_CARDS, C_HEAD, c._row, 'views', '0');
          setCell_(SH_CARDS, C_HEAD, c._row, 'saves', '0');
        });
        return true;
      }

      case 'deleteUser': {
        needAdmin();
        const name = clean_(p.username);
        if (name === me.username) throw new Error('Du kan ikke slette dig selv.');
        const u = findUser_(name);
        if (!u) throw new Error('Brugeren findes ikke.');
        const sh = sheet_(SH_CARDS, C_HEAD);
        cardsOf_(name).map(c => c._row).sort((a, b) => b - a).forEach(r => sh.deleteRow(r));
        sheet_(SH_USERS, U_HEAD).deleteRow(u._row);
        killSessions_(name);
        return true;
      }
    }
    throw new Error('Ukendt handling: ' + action);
  } finally {
    lock.releaseLock();
  }
}

/* ================= Login & sessioner ================= */
function login_(username, pw) {
  username = clean_(username);
  const cache = CacheService.getScriptCache();
  const fk = 'f_' + username;
  const fails = Number(cache.get(fk) || 0);
  if (fails >= CFG.MAX_FAIL) throw new Error('For mange forkerte forsøg. Prøv igen om 15 minutter.');

  const u = findUser_(username);
  const ok = u && u.active === 'TRUE' && hash_(pw || '', u.salt) === u.hash;
  if (!ok) {
    if (!u) hash_(pw || '', 'x'); // samme svartid uanset om brugeren findes
    cache.put(fk, String(fails + 1), CFG.LOCK_SEC);
    throw new Error('Forkert brugernavn eller adgangskode.');
  }
  cache.remove(fk);
  const token = Utilities.getUuid() + Utilities.getUuid();
  cache.put('s_' + token, username, CFG.SESSION_SEC);
  const props = PropertiesService.getScriptProperties();
  const list = JSON.parse(props.getProperty('sess_' + username) || '[]').slice(-9);
  list.push(token);
  props.setProperty('sess_' + username, JSON.stringify(list));
  setCell_(SH_USERS, U_HEAD, u._row, 'lastLogin', new Date().toISOString());
  return { token: token, user: pubUser_(u) };
}

function session_(token) {
  if (!token) throw new Error('SESSION: Log ind igen.');
  const cache = CacheService.getScriptCache();
  const name = cache.get('s_' + token);
  if (!name) throw new Error('SESSION: Du er logget ud. Log ind igen.');
  const u = findUser_(name);
  if (!u || u.active !== 'TRUE') throw new Error('SESSION: Brugeren er deaktiveret.');
  cache.put('s_' + token, name, CFG.SESSION_SEC);
  return u;
}

function killSessions_(username) {
  const props = PropertiesService.getScriptProperties();
  const list = JSON.parse(props.getProperty('sess_' + username) || '[]');
  const cache = CacheService.getScriptCache();
  list.forEach(t => cache.remove('s_' + t));
  props.deleteProperty('sess_' + username);
}

function hash_(pw, salt) {
  let h = salt + '|' + pw;
  for (let i = 0; i < 300; i++) {
    h = Utilities.base64Encode(
      Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, h, Utilities.Charset.UTF_8));
  }
  return h;
}
function checkPw_(pw) {
  if (!pw || String(pw).length < 8) throw new Error('Adgangskoden skal være mindst 8 tegn.');
}
function tempPw_() { return Utilities.getUuid().replace(/-/g, '').slice(0, 10); }

/* ================= Data ================= */
const _checked = {};
let _ss = null;
function ss_() {
  if (_ss) return _ss;
  const id = PropertiesService.getScriptProperties().getProperty('SHEET_ID');
  if (!id) throw new Error('Systemet er ikke sat op. Kør setup i Apps Script.');
  _ss = SpreadsheetApp.openById(id);
  return _ss;
}

function sheet_(name, head) {
  const ss = ss_();
  let sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    if (sh.getMaxColumns() < head.length) sh.insertColumnsAfter(sh.getMaxColumns(), head.length - sh.getMaxColumns());
    sh.getRange(1, 1, sh.getMaxRows(), head.length).setNumberFormat('@'); // alt som tekst (+45 osv.)
    sh.getRange(1, 1, 1, head.length).setValues([head]).setFontWeight('bold');
    sh.setFrozenRows(1);
    _checked[name] = true;
  } else if (!_checked[name]) {
    // Tilføj nye kolonner til ældre ark (opgradering)
    const cur = sh.getRange(1, 1, 1, Math.max(sh.getLastColumn(), 1)).getValues()[0].map(String).filter(String);
    const missing = head.filter(k => cur.indexOf(k) < 0);
    if (missing.length) {
      const start = cur.length + 1;
      const needCols = start + missing.length - 1 - sh.getMaxColumns();
      if (needCols > 0) sh.insertColumnsAfter(sh.getMaxColumns(), needCols);
      sh.getRange(1, start, sh.getMaxRows(), missing.length).setNumberFormat('@');
      sh.getRange(1, start, 1, missing.length).setValues([missing]).setFontWeight('bold');
    }
    _checked[name] = true;
  }
  return sh;
}

function header_(sh) { return sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(String); }

function rows_(name, head) {
  const v = sheet_(name, head).getDataRange().getValues();
  const h = v.shift().map(String);
  return v.map((r, i) => {
    const o = { _row: i + 2 };
    h.forEach((k, j) => { if (k) o[k] = String(r[j]); });
    return o;
  }).filter(o => o[h[0]] !== '');
}

function writeRow_(name, head, obj) {
  const sh = sheet_(name, head);
  const h = header_(sh);
  const row = h.map(k => (!k || obj[k] == null) ? '' : String(obj[k]));
  const r = obj._row || sh.getLastRow() + 1;
  sh.getRange(r, 1, 1, h.length).setNumberFormat('@').setValues([row]);
  obj._row = r;
}

function setCell_(name, head, row, key, value) {
  const sh = sheet_(name, head);
  const col = header_(sh).indexOf(key) + 1;
  if (col > 0) sh.getRange(row, col).setNumberFormat('@').setValue(String(value));
}

function bump_(c, key) {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(3000)) return; // tællinger må gerne springes over ved travlhed
  try {
    const sh = sheet_(SH_CARDS, C_HEAD);
    const h = header_(sh);
    const col = h.indexOf(key) + 1;
    if (col < 1) return;
    const cell = sh.getRange(c._row, col);
    cell.setNumberFormat('@').setValue(String(Number(cell.getValue() || 0) + 1));
    if (key === 'views') setCell_(SH_CARDS, C_HEAD, c._row, 'lastView', new Date().toISOString());
  } finally { lock.releaseLock(); }
}

function findUser_(username) { return rows_(SH_USERS, U_HEAD).filter(u => u.username === username)[0] || null; }
function findCard_(slug) { return rows_(SH_CARDS, C_HEAD).filter(c => c.slug === slug)[0] || null; }
function publicCard_(slug) {
  const c = findCard_(clean_(slug));
  if (!c || c.hidden === 'TRUE') return null;
  const u = findUser_(ownerOf_(c));
  return u && u.active === 'TRUE' ? c : null;
}

// Ældre kort uden "owner" tilhører brugeren med samme navn som linket
function ownerOf_(c) { return c.owner || c.slug; }
function cardsOf_(owner) {
  return rows_(SH_CARDS, C_HEAD)
    .filter(c => ownerOf_(c) === owner)
    .sort((a, b) => (Number(a.sort || 0) - Number(b.sort || 0)) || (a._row - b._row));
}
function canEdit_(me, c) { return me.role === 'admin' || ownerOf_(c) === me.username; }
function cleanLabel_(s) { return String(s || '').replace(/[<>]/g, '').trim().slice(0, 30); }
function slugify_(s) {
  return String(s || '').toLowerCase()
    .replace(/æ/g, 'ae').replace(/ø/g, 'oe').replace(/å/g, 'aa')
    .normalize('NFD').replace(/[\u0300-\u036f]/g, '')
    .replace(/[^a-z0-9]+/g, '-').replace(/^-+|-+$/g, '').slice(0, 40) || 'profil';
}

function createUserRow_(username, name, role, pw) {
  const salt = Utilities.getUuid();
  writeRow_(SH_USERS, U_HEAD, {
    username: username, name: name, hash: hash_(pw, salt), salt: salt, role: role,
    active: 'TRUE', mustChange: 'TRUE', created: new Date().toISOString(), lastLogin: ''
  });
  const parts = name.split(/\s+/);
  const last = parts.length > 1 ? parts.pop() : '';
  writeRow_(SH_CARDS, C_HEAD, {
    slug: username, owner: username, label: 'Arbejde', sort: '0',
    firstName: parts.join(' '), lastName: last, theme: '', photo: '', links: '[]',
    hidden: 'FALSE', views: '0', saves: '0', updated: new Date().toISOString(), updatedBy: username
  });
}

function pubUser_(u) {
  return { username: u.username, name: u.name, role: u.role,
           active: u.active === 'TRUE', mustChange: u.mustChange === 'TRUE', lastLogin: u.lastLogin };
}

function cardOut_(c, isPublic) {
  const o = {};
  (isPublic ? PUBLIC_FIELDS : C_HEAD).forEach(k => o[k] = c[k] || '');
  o.hidden = c.hidden === 'TRUE';
  if (!isPublic) {
    o.views = Number(c.views || 0); o.saves = Number(c.saves || 0);
    o.owner = ownerOf_(c); o.label = c.label || 'Profil';
    delete o.updatedBy;
  }
  o.publicUrl = publicUrl_(c.slug);
  o.vcfUrl = ScriptApp.getService().getUrl() + '?vcf=' + encodeURIComponent(c.slug);
  return o;
}

function publicUrl_(slug) { return CFG.PUBLIC_BASE + 'card.html?c=' + encodeURIComponent(slug); }

function cleanTheme_(t) {
  try { t = typeof t === 'string' ? JSON.parse(t || '{}') : (t || {}); } catch (e) { t = {}; }
  const hex = v => /^#[0-9a-f]{6}$/i.test(v || '') ? v : '';
  return JSON.stringify({
    preset: String(t.preset || '').replace(/[^a-z]/g, '').slice(0, 20),
    bg: hex(t.bg), card: hex(t.card), text: hex(t.text), accent: hex(t.accent),
    font: FONTS.indexOf(t.font) >= 0 ? t.font : '',
    shape: t.shape === 'square' ? 'square' : 'round',
    layout: t.layout === 'top' ? 'top' : 'side',
    lang: t.lang === 'da' ? 'da' : 'en'
  });
}

function cleanLinks_(list) {
  try { list = typeof list === 'string' ? JSON.parse(list || '[]') : (list || []); } catch (e) { list = []; }
  if (!Array.isArray(list)) list = [];
  return JSON.stringify(list
    .map(l => ({ label: String((l && l.label) || '').trim().slice(0, 40), url: fixUrl_((l && l.url) || '') }))
    .filter(l => l.url && /^https?:\/\/[^\s]+\.[^\s]+/i.test(l.url))
    .map(l => ({ label: l.label || l.url.replace(/^https?:\/\/(www\.)?/i, '').split('/')[0], url: l.url.slice(0, 300) }))
    .slice(0, CFG.MAX_LINKS));
}

function clean_(s) { return String(s || '').trim().toLowerCase(); }

function fixUrl_(u) {
  u = String(u || '').trim();
  if (!u) return '';
  if (/^(javascript|data|vbscript):/i.test(u)) return '';
  if (!/^https?:\/\//i.test(u)) u = 'https://' + u;
  return u;
}
function fixLinkedIn_(u) {
  u = String(u || '').trim();
  if (!u) return '';
  if (!/linkedin\.com/i.test(u)) u = 'https://www.linkedin.com/in/' + u.replace(/^@/, '').replace(/^\/+/, '');
  return fixUrl_(u);
}

/* ================= vCard ================= */
function vcard_(c) {
  const e = s => String(s || '').replace(/\\/g, '\\\\').replace(/\n/g, '\\n').replace(/([,;])/g, '\\$1');
  const L = ['BEGIN:VCARD', 'VERSION:3.0',
    'N:' + e(c.lastName) + ';' + e(c.firstName) + ';;;',
    'FN:' + e([c.firstName, c.lastName].join(' ').trim())];
  if (c.company) L.push('ORG:' + e(c.company));
  if (c.title) L.push('TITLE:' + e(c.title));
  if (c.phone) L.push('TEL;TYPE=CELL,VOICE:' + c.phone.replace(/[^\d+]/g, ''));
  if (c.email) L.push('EMAIL;TYPE=INTERNET,WORK:' + c.email);
  if (c.web) L.push('URL;TYPE=WORK:' + c.web);
  let n = 1;
  if (c.linkedin) {
    L.push('item' + n + '.URL:' + c.linkedin, 'item' + n + '.X-ABLabel:LinkedIn');
    L.push('X-SOCIALPROFILE;TYPE=linkedin:' + c.linkedin);
    n++;
  }
  JSON.parse(c.links || '[]').forEach(l => {
    L.push('item' + n + '.URL:' + l.url, 'item' + n + '.X-ABLabel:' + e(l.label));
    n++;
  });
  let lang = 'en';
  try { lang = JSON.parse(c.theme || '{}').lang === 'da' ? 'da' : 'en'; } catch (err) {}
  L.push('item' + n + '.URL:' + publicUrl_(c.slug), 'item' + n + '.X-ABLabel:' + (lang === 'da' ? 'Visitkort' : 'Business card'));
  if (c.address) L.push('ADR;TYPE=WORK:;;' + e(c.address) + ';;;;');
  if (c.note) L.push('NOTE:' + e(c.note));
  if (c.photo && c.photo.indexOf('data:image/jpeg;base64,') === 0) {
    L.push(fold_('PHOTO;ENCODING=b;TYPE=JPEG:' + c.photo.split(',')[1]));
  }
  L.push('REV:' + new Date().toISOString());
  L.push('END:VCARD');
  return L.join('\r\n') + '\r\n';
}

// vCard-linjer må max være 75 tegn; fortsættelseslinjer starter med mellemrum
function fold_(line) {
  const out = [line.slice(0, 75)];
  for (let i = 75; i < line.length; i += 74) out.push(' ' + line.slice(i, i + 74));
  return out.join('\r\n');
}
