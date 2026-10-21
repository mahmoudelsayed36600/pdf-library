/**
 * المكتبة المحمية — الخادم (Google Apps Script)  الإصدار 3
 * - الملفات تُخزَّن مشفّرة على Drive (تُشفَّر عند الرفع) فيرسلها الخادم كما هي بسرعة.
 * - السيريال يعمل على عدد أجهزة يحدده الأدمن.
 */

var SHEETS = {
  serials: ['serial', 'name', 'status', 'expiry', 'device', 'folders', 'created', 'lastSeen', 'note', 'maxDevices', 'present', 'presentMin'],
  folders: ['id', 'name', 'order', 'created'],
  files:   ['id', 'folderId', 'name', 'visible', 'size', 'added', 'order', 'parts', 'fkey', 'psize', 'pages', 'cover']
};
var OLD_CHUNK = 6 * 1024 * 1024;       // حجم الأجزاء القديمة
var MAX_FILE = 400 * 1024 * 1024;      // أقصى حجم للملف

/* ===================== الإعداد ===================== */
function setup() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  Object.keys(SHEETS).forEach(function (n) {
    var sh = ss.getSheetByName(n) || ss.insertSheet(n);
    sh.getRange(1, 1, 1, SHEETS[n].length).setValues([SHEETS[n]]).setFontWeight('bold');
    sh.setFrozenRows(1);
  });
  var p = PropertiesService.getScriptProperties();
  if (!p.getProperty('ROOT')) {
    var root = DriveApp.createFolder('PDF_Library_' + ss.getId().slice(0, 6));
    p.setProperty('ROOT', root.getId());
    p.setProperty('ANN', root.createFolder('annotations').getId());
  }
  if (!p.getProperty('ADMIN_PASS')) p.setProperty('ADMIN_PASS', hash_('admin123'));
  return 'تم الإعداد. كلمة مرور الأدمن المؤقتة: admin123';
}

/* ===================== أدوات ===================== */
function hash_(s) {
  var b = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, String(s), Utilities.Charset.UTF_8);
  return b.map(function (x) { return ('0' + (x & 255).toString(16)).slice(-2); }).join('');
}
function prop_(k) { return PropertiesService.getScriptProperties().getProperty(k); }
function sheet_(n) {
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(n);
  if (!sh) throw new Error('شغّل الدالة setup أولاً');
  if (sh.getLastColumn() < SHEETS[n].length) sh.getRange(1, 1, 1, SHEETS[n].length).setValues([SHEETS[n]]).setFontWeight('bold');
  return sh;
}
var MEMO_ = {};
function rows_(n) {
  if (MEMO_[n]) return MEMO_[n];
  // نسخة مخبّأة للفولدرات والملفات (تُمسح عند أي تعديل)
  var cache = CacheService.getScriptCache(), ck = 'rows_' + n + '_' + (prop_('VER_' + n) || 0);
  if (n !== 'serials') {
    var c = cache.get(ck);
    if (c) { try { return (MEMO_[n] = JSON.parse(c)); } catch (e) {} }
  }
  var sh = sheet_(n), v = sh.getDataRange().getValues(), h = v.shift();
  var out = v.map(function (r, i) {
    var o = { _row: i + 2 };
    h.forEach(function (k, j) { o[k] = r[j] instanceof Date ? r[j].toISOString() : r[j]; });
    return o;
  }).filter(function (o) { return String(o[h[0]]) !== ''; });
  if (n !== 'serials') { try { cache.put(ck, JSON.stringify(out), 21600); } catch (e) {} }
  return (MEMO_[n] = out);
}
function bump_(n) {
  delete MEMO_[n];
  var p = PropertiesService.getScriptProperties();
  p.setProperty('VER_' + n, String((parseInt(p.getProperty('VER_' + n), 10) || 0) + 1));
}
function append_(n, obj) {
  sheet_(n).appendRow(SHEETS[n].map(function (k) { return obj[k] === undefined ? '' : obj[k]; }));
  bump_(n);
}
function update_(n, row, obj) {
  var sh = sheet_(n);
  Object.keys(obj).forEach(function (k) {
    var c = SHEETS[n].indexOf(k);
    if (c >= 0) sh.getRange(row, c + 1).setValue(obj[k]);
  });
  bump_(n);
}
function delRow_(n, row) { sheet_(n).deleteRow(row); bump_(n); }
function json_(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }
function text_(s) { return ContentService.createTextOutput(s).setMimeType(ContentService.MimeType.TEXT); }
function genSerial_() {
  var c = 'ABCDEFGHJKLMNPQRSTUVWXYZ23456789', s = '';
  for (var i = 0; i < 16; i++) { s += c.charAt(Math.floor(Math.random() * c.length)); if (i % 4 === 3 && i < 15) s += '-'; }
  return s;
}
function iso_(d) { return d ? new Date(d).toISOString() : ''; }
function isTrue_(v) { return v === true || v === 'TRUE' || v === 'true'; }
function xorHex_(a, b) {   // يغلّف مفتاح الملف بمفتاح الجهاز (hex ^ hex)
  var out = '';
  for (var i = 0; i < a.length; i += 2) {
    var x = parseInt(a.substr(i, 2), 16) ^ parseInt(b.substr(i % b.length, 2), 16);
    out += ('0' + (x & 255).toString(16)).slice(-2);
  }
  return out;
}

/* ===================== نقطة الدخول ===================== */
function doGet() { return json_({ ok: true, app: 'pdf-library', v: 3 }); }

function doPost(e) {
  var req;
  try { req = JSON.parse(e.postData.contents || '{}'); } catch (x) { return json_({ ok: false, error: 'طلب غير صالح' }); }
  if (req.action === 'part') {               // جزء من الملف: نص مباشر بدون JSON (أسرع)
    try { return text_(partData_(req)); } catch (err) { return text_('ERR:' + (err && err.message || err)); }
  }
  var lock = null;
  try {
    var a = req.action, fn = API[a];
    if (!fn) return json_({ ok: false, error: 'إجراء غير معروف' });
    if (/^admin(Add|Rename|Delete|File|Create|Serial|Password|UploadFinish|MigrateFinish|SetCover)|^activate$/.test(a)) {
      lock = LockService.getScriptLock(); lock.waitLock(20000);
    }
    return json_(fn(req));
  } catch (err) {
    return json_({ ok: false, error: String(err && err.message || err) });
  } finally {
    if (lock) try { lock.releaseLock(); } catch (x) {}
  }
}

/* ===================== السيريال والأجهزة ===================== */
function findSerial_(serial) {
  serial = String(serial || '').trim().toUpperCase();
  var list = rows_('serials');
  for (var i = 0; i < list.length; i++) if (String(list[i].serial).toUpperCase() === serial) return list[i];
  return null;
}
function devices_(s) {
  var d = String(s.device || '').trim();
  if (!d) return [];
  if (d.charAt(0) === '[') { try { return JSON.parse(d); } catch (e) { return []; } }
  return [{ id: d, label: 'جهاز', at: '' }];          // صيغة قديمة: جهاز واحد
}
function maxDev_(s) { return Math.max(1, parseInt(s.maxDevices, 10) || 1); }
function presMin_(s) { return Math.min(Math.max(parseInt(s.presentMin, 10) || 60, 5), 480); }
function checkUser_(req, bind) {
  var s = findSerial_(req.serial);
  if (!s) throw new Error('INVALID:السيريال غير صحيح أو تم حذفه');
  if (s.status !== 'active') throw new Error('INVALID:هذا السيريال موقوف. تواصل مع الإدارة');
  if (s.expiry && new Date(s.expiry) < new Date()) throw new Error('INVALID:انتهت صلاحية السيريال. تواصل مع الإدارة للتجديد');
  if (!req.device) throw new Error('معرّف الجهاز مفقود');
  var minV = parseInt(prop_('MIN_VER'), 10) || 0;
  if (minV && (parseInt(req.ver, 10) || 0) < minV)
    throw new Error('INVALID:يوجد إصدار أحدث من البرنامج، ويجب التحديث للمتابعة. حمّل الإصدار الجديد من: ' + (prop_('DL_LINK') || ''));
  var list = devices_(s), mine = list.filter(function (d) { return d.id === req.device; })[0];
  if (!mine) {
    if (!bind) throw new Error('INVALID:هذا الجهاز غير مفعّل على السيريال');
    var mx = maxDev_(s);
    if (list.length >= mx) throw new Error('INVALID:تم الوصول للحد الأقصى من الأجهزة (' + mx + ') لهذا السيريال. تواصل مع الإدارة');
    list.push({ id: req.device, label: String(req.label || 'جهاز').slice(0, 40), at: new Date().toISOString() });
    update_('serials', s._row, { device: JSON.stringify(list), lastSeen: new Date().toISOString() });
    s.device = JSON.stringify(list);
  } else if (!s.lastSeen || Date.now() - new Date(s.lastSeen).getTime() > 3600000) {
    update_('serials', s._row, { lastSeen: new Date().toISOString() });   // تحديث كل ساعة فقط
  }
  return s;
}
function allowedFolders_(s) {
  var f = String(s.folders || '').trim();
  if (f === '*' || f === '') return null;
  return f.split(',').map(function (x) { return x.trim(); });
}
function checkAdmin_(req) {
  if (!req.pass || hash_(req.pass) !== prop_('ADMIN_PASS')) throw new Error('كلمة مرور الأدمن غير صحيحة');
}
function userFile_(req) {
  var s = checkUser_(req, false), allow = allowedFolders_(s);
  var f = rows_('files').filter(function (x) { return x.id === req.fileId; })[0];
  if (!f || !isTrue_(f.visible)) throw new Error('INVALID_FILE:الملف غير متاح');
  if (allow && allow.indexOf(String(f.folderId)) < 0) throw new Error('INVALID_FILE:غير مسموح بفتح هذا الملف');
  return f;
}
function fileParts_(f) { return String(f.parts || '').split(',').filter(String); }

/* ===================== تقديم أجزاء الملف ===================== */
function partData_(req) {
  var f = userFile_(req), part = parseInt(req.part, 10) || 0, parts = fileParts_(f);
  if (f.fkey) {                                   // مشفّر مسبقاً: يُرسل كما هو
    if (part >= parts.length) throw new Error('جزء غير موجود');
    return DriveApp.getFileById(parts[part]).getBlob().getDataAsString();
  }
  // ملف قديم: يُشفَّر الآن بمفتاح الجهاز (أبطأ)
  var bytes;
  if (parts.length) bytes = DriveApp.getFileById(parts[part]).getBlob().getBytes();
  else bytes = DriveApp.getFileById(f.id).getBlob().getBytes().slice(part * OLD_CHUNK, (part + 1) * OLD_CHUNK);
  var key = hash_(req.serial + req.device + req.fileId), kb = [];
  for (var i = 0; i < key.length; i++) kb.push(key.charCodeAt(i));
  for (var j = 0; j < bytes.length; j++) bytes[j] = bytes[j] ^ kb[j % 64];
  return Utilities.base64Encode(bytes);
}

/* ===================== واجهة البرمجة ===================== */
var API = {

  /* ---------- المستخدم ---------- */
  activate: function (req) {
    var s = checkUser_(req, true);
    return { ok: true, name: s.name, expiry: iso_(s.expiry) };
  },

  library: function (req) {
    var s = checkUser_(req, false);
    var allow = allowedFolders_(s);
    var folders = rows_('folders').filter(function (f) { return !allow || allow.indexOf(String(f.id)) >= 0; })
      .sort(function (a, b) { return (a.order || 0) - (b.order || 0); })
      .map(function (f) { return { id: String(f.id), name: f.name }; });
    var ids = folders.map(function (f) { return f.id; });
    var files = rows_('files').filter(function (x) { return isTrue_(x.visible) && ids.indexOf(String(x.folderId)) >= 0; })
      .sort(function (a, b) { return (a.order || 0) - (b.order || 0); })
      .map(function (x) { return { id: x.id, folderId: String(x.folderId), name: x.name, size: Number(x.size) || 0, added: iso_(x.added), pages: Number(x.pages) || 0, cover: String(x.cover || '') }; });
    return { ok: true, name: s.name, expiry: iso_(s.expiry), devices: devices_(s).length, maxDevices: maxDev_(s), present: isTrue_(s.present), presentMin: presMin_(s), folders: folders, files: files };
  },

  fileInfo: function (req) {
    var f = userFile_(req), parts = fileParts_(f), size = Number(f.size) || 0;
    var enc = !!f.fkey, psize = enc ? (Number(f.psize) || OLD_CHUNK) : OLD_CHUNK;
    var n = (enc || parts.length) ? parts.length : Math.max(1, Math.ceil(size / OLD_CHUNK));
    var dk = hash_(req.serial + '|' + req.device + '|' + req.fileId);
    return { ok: true, mode: enc ? 'enc' : 'legacy', parts: n, psize: psize, size: size, k: enc ? xorHex_(String(f.fkey), dk) : '' };
  },

  covers: function (req) {          // صور الأغلفة المطلوبة فقط
    checkUser_(req, false);
    var want = (req.ids || []).slice(0, 60), out = {};
    var valid = {}; rows_('files').forEach(function (x) { if (x.cover) valid[String(x.cover)] = 1; });
    want.forEach(function (id) { if (valid[id]) try { out[id] = DriveApp.getFileById(id).getBlob().getDataAsString(); } catch (e) {} });
    return { ok: true, covers: out };
  },

  getAnn: function (req) {
    checkUser_(req, false);
    var it = DriveApp.getFolderById(prop_('ANN')).getFilesByName(annName_(req));
    return { ok: true, data: it.hasNext() ? it.next().getBlob().getDataAsString() : '' };
  },
  saveAnn: function (req) {
    checkUser_(req, false);
    var folder = DriveApp.getFolderById(prop_('ANN'));
    var name = annName_(req), it = folder.getFilesByName(name), data = String(req.data || '');
    if (data.length > 8000000) throw new Error('حجم الملاحظات كبير جداً');
    if (it.hasNext()) it.next().setContent(data); else folder.createFile(name, data, 'application/json');
    return { ok: true };
  },

  /* ---------- الأدمن ---------- */
  adminLogin: function (req) { checkAdmin_(req); return { ok: true }; },

  adminData: function (req) {
    checkAdmin_(req);
    return {
      ok: true, minVer: parseInt(prop_('MIN_VER'), 10) || 0,
      serials: rows_('serials').map(function (s) {
        return { serial: s.serial, name: s.name, status: s.status, expiry: iso_(s.expiry),
                 devices: devices_(s).map(function (d) { return { id: String(d.id).slice(0, 8), label: d.label, at: d.at }; }),
                 maxDevices: maxDev_(s), present: isTrue_(s.present), presentMin: presMin_(s), folders: String(s.folders || '*'), created: iso_(s.created), lastSeen: iso_(s.lastSeen), note: s.note };
      }),
      folders: rows_('folders').slice().sort(function (a, b) { return (a.order || 0) - (b.order || 0); }).map(function (f) { return { id: String(f.id), name: f.name }; }),
      files: rows_('files').slice().sort(function (a, b) { return (a.order || 0) - (b.order || 0); })
        .map(function (x) { return { id: x.id, folderId: String(x.folderId), name: x.name, visible: isTrue_(x.visible), size: Number(x.size) || 0, added: iso_(x.added),
                                     enc: !!x.fkey, cover: !!x.cover, parts: fileParts_(x).length || Math.max(1, Math.ceil((Number(x.size) || 0) / OLD_CHUNK)) }; })
    };
  },

  adminAddFolder: function (req) {
    checkAdmin_(req);
    var id = 'F' + Date.now().toString(36);
    append_('folders', { id: id, name: req.name, order: rows_('folders').length + 1, created: new Date() });
    return { ok: true, id: id };
  },
  adminRenameFolder: function (req) {
    checkAdmin_(req);
    var f = rows_('folders').filter(function (x) { return String(x.id) === req.id; })[0];
    if (!f) throw new Error('الفولدر غير موجود');
    update_('folders', f._row, { name: req.name });
    return { ok: true };
  },
  adminDeleteFolder: function (req) {
    checkAdmin_(req);
    var files = rows_('files').filter(function (x) { return String(x.folderId) === req.id; });
    files.sort(function (a, b) { return b._row - a._row; }).forEach(function (x) { trashFile_(x); sheet_('files').deleteRow(x._row); });
    bump_('files');
    var f = rows_('folders').filter(function (x) { return String(x.id) === req.id; })[0];
    if (f) delRow_('folders', f._row);
    return { ok: true };
  },

  // رفع جزء (مشفّر مسبقاً من المتصفح، يُخزَّن كنص base64)
  adminUploadPart: function (req) {
    checkAdmin_(req);
    var data = String(req.data || '');
    if (data.length > 8 * 1024 * 1024) throw new Error('الجزء أكبر من المسموح');
    var file = DriveApp.getFolderById(prop_('ROOT')).createFile(req.uploadId + '.p' + req.index, data, 'text/plain');
    return { ok: true, index: req.index, id: file.getId() };
  },
  adminUploadFinish: function (req) {
    checkAdmin_(req);
    if (!rows_('folders').some(function (x) { return String(x.id) === req.folderId; })) {
      trashIds_(req.parts); throw new Error('اختر فولدر صحيح');
    }
    if (!req.parts || !req.parts.length || req.parts.some(function (x) { return !x; })) throw new Error('لم تكتمل أجزاء الملف');
    if (Number(req.size) > MAX_FILE) { trashIds_(req.parts); throw new Error('الملف أكبر من 400 ميجا'); }
    append_('files', { id: req.uploadId, folderId: req.folderId, name: String(req.name).replace(/\.pdf$/i, ''), visible: true,
                       size: Number(req.size) || 0, added: new Date(), order: rows_('files').length + 1,
                       parts: req.parts.join(','), fkey: req.key, psize: req.psize, pages: req.pages || '' });
    return { ok: true, id: req.uploadId };
  },
  adminUploadCancel: function (req) { checkAdmin_(req); trashIds_(req.parts); return { ok: true }; },

  // تحويل ملف قديم إلى الصيغة السريعة (جزء جزء)
  adminMigratePart: function (req) {
    checkAdmin_(req);
    var f = rows_('files').filter(function (x) { return x.id === req.id; })[0];
    if (!f) throw new Error('الملف غير موجود');
    if (f.fkey) throw new Error('الملف محوّل بالفعل');
    var parts = fileParts_(f), part = parseInt(req.part, 10) || 0, bytes;
    if (parts.length) bytes = DriveApp.getFileById(parts[part]).getBlob().getBytes();
    else bytes = DriveApp.getFileById(f.id).getBlob().getBytes().slice(part * OLD_CHUNK, (part + 1) * OLD_CHUNK);
    var kb = [];
    for (var i = 0; i < 64; i += 2) kb.push(parseInt(req.key.substr(i, 2), 16));
    for (var j = 0; j < bytes.length; j++) bytes[j] = bytes[j] ^ kb[j % 32];
    var nf = DriveApp.getFolderById(prop_('ROOT')).createFile(f.id + '.m' + part, Utilities.base64Encode(bytes), 'text/plain');
    return { ok: true, id: nf.getId() };
  },
  adminMigrateFinish: function (req) {
    checkAdmin_(req);
    var f = rows_('files').filter(function (x) { return x.id === req.id; })[0];
    if (!f) { trashIds_(req.parts); throw new Error('الملف غير موجود'); }
    if (f.fkey) { trashIds_(req.parts); return { ok: true }; }
    var old = fileParts_(f); if (!old.length) old = [f.id];
    update_('files', f._row, { parts: req.parts.join(','), fkey: req.key, psize: OLD_CHUNK });
    trashIds_(old);
    return { ok: true };
  },

  // قراءة الملف للأدمن (لإنشاء الغلاف من الصفحة الأولى)
  adminPart: function (req) {
    checkAdmin_(req);
    var f = rows_('files').filter(function (x) { return x.id === req.id; })[0];
    if (!f) throw new Error('الملف غير موجود');
    var parts = fileParts_(f), part = parseInt(req.part, 10) || 0, size = Number(f.size) || 0, data;
    var n = (f.fkey || parts.length) ? parts.length : Math.max(1, Math.ceil(size / OLD_CHUNK));
    if (f.fkey) data = DriveApp.getFileById(parts[part]).getBlob().getDataAsString();
    else if (parts.length) data = Utilities.base64Encode(DriveApp.getFileById(parts[part]).getBlob().getBytes());
    else data = Utilities.base64Encode(DriveApp.getFileById(f.id).getBlob().getBytes().slice(part * OLD_CHUNK, (part + 1) * OLD_CHUNK));
    return { ok: true, parts: n, key: f.fkey ? String(f.fkey) : '', data: data };
  },
  adminSetCover: function (req) {
    checkAdmin_(req);
    var f = rows_('files').filter(function (x) { return x.id === req.id; })[0];
    if (!f) throw new Error('الملف غير موجود');
    var data = String(req.data || '');
    if (data.indexOf('data:image/') !== 0 || data.length > 400000) throw new Error('صورة الغلاف غير صالحة');
    if (f.cover) trashIds_([f.cover]);
    var cf = DriveApp.getFolderById(prop_('ROOT')).createFile(f.id + '.cover', data, 'text/plain');
    var u = { cover: cf.getId() }; if (req.pages) u.pages = req.pages;
    update_('files', f._row, u);
    return { ok: true, cover: cf.getId() };
  },

  adminFile: function (req) {   // إظهار/إخفاء/إعادة تسمية/نقل
    checkAdmin_(req);
    var f = rows_('files').filter(function (x) { return x.id === req.id; })[0];
    if (!f) throw new Error('الملف غير موجود');
    var u = {};
    if (req.visible !== undefined) u.visible = !!req.visible;
    if (req.name) u.name = req.name;
    if (req.folderId) u.folderId = req.folderId;
    update_('files', f._row, u);
    return { ok: true };
  },
  adminDeleteFile: function (req) {
    checkAdmin_(req);
    var f = rows_('files').filter(function (x) { return x.id === req.id; })[0];
    if (!f) throw new Error('الملف غير موجود');
    trashFile_(f); delRow_('files', f._row);
    return { ok: true };
  },

  adminCreateSerials: function (req) {
    checkAdmin_(req);
    var n = Math.min(Math.max(parseInt(req.count, 10) || 1, 1), 200), out = [], rows = [];
    var exp = req.days ? new Date(Date.now() + req.days * 86400000) : '';
    var existing = rows_('serials').map(function (s) { return s.serial; });
    for (var i = 0; i < n; i++) {
      var s; do { s = genSerial_(); } while (existing.indexOf(s) >= 0);
      existing.push(s);
      var name = n > 1 ? (req.name || 'مستخدم') + ' ' + (i + 1) : (req.name || 'مستخدم');
      var o = { serial: s, name: name, status: 'active', expiry: exp, device: '', folders: req.folders || '*', created: new Date(), note: req.note || '',
                maxDevices: Math.min(Math.max(parseInt(req.maxDevices, 10) || 1, 1), 50),
                present: !!req.present, presentMin: Math.min(Math.max(parseInt(req.presentMin, 10) || 60, 5), 480) };
      rows.push(SHEETS.serials.map(function (k) { return o[k] === undefined ? '' : o[k]; }));
      out.push(s);
    }
    var sh = sheet_('serials');
    sh.getRange(sh.getLastRow() + 1, 1, rows.length, SHEETS.serials.length).setValues(rows);
    delete MEMO_.serials;
    return { ok: true, serials: out };
  },
  adminSerial: function (req) {   // تعديل: الحالة، التجديد، الصلاحيات، الأجهزة، الاسم
    checkAdmin_(req);
    var s = findSerial_(req.serial);
    if (!s) throw new Error('السيريال غير موجود');
    var u = {};
    if (req.status) u.status = req.status;
    if (req.name !== undefined) u.name = req.name;
    if (req.folders !== undefined) u.folders = req.folders;
    if (req.present !== undefined) u.present = !!req.present;
    if (req.presentMin !== undefined) u.presentMin = Math.min(Math.max(parseInt(req.presentMin, 10) || 60, 5), 480);
    if (req.maxDevices !== undefined) u.maxDevices = Math.min(Math.max(parseInt(req.maxDevices, 10) || 1, 1), 50);
    if (req.resetDevice) u.device = '';
    if (req.removeDevice) u.device = JSON.stringify(devices_(s).filter(function (d) { return String(d.id).slice(0, 8) !== req.removeDevice; }));
    if (req.renewDays) {
      var base = s.expiry && new Date(s.expiry) > new Date() ? new Date(s.expiry) : new Date();
      u.expiry = new Date(base.getTime() + req.renewDays * 86400000);
      u.status = 'active';
    }
    if (req.noExpiry) u.expiry = '';
    update_('serials', s._row, u);
    return { ok: true };
  },
  adminDeleteSerial: function (req) {
    checkAdmin_(req);
    var s = findSerial_(req.serial);
    if (!s) throw new Error('السيريال غير موجود');
    delRow_('serials', s._row);
    return { ok: true };
  },
  adminUpdatePolicy: function (req) {
    checkAdmin_(req);
    var p = PropertiesService.getScriptProperties();
    p.setProperty('MIN_VER', String(parseInt(req.min, 10) || 0));
    if (req.link) p.setProperty('DL_LINK', String(req.link).slice(0, 300));
    return { ok: true };
  },
  adminPassword: function (req) {
    checkAdmin_(req);
    if (!req.newPass || String(req.newPass).length < 6) throw new Error('كلمة المرور يجب ألا تقل عن 6 أحرف');
    PropertiesService.getScriptProperties().setProperty('ADMIN_PASS', hash_(req.newPass));
    return { ok: true };
  }
};

function trashIds_(ids) { (ids || []).forEach(function (id) { if (id) try { DriveApp.getFileById(id).setTrashed(true); } catch (e) {} }); }
function trashFile_(f) { var ids = fileParts_(f); trashIds_(ids.length ? ids : [f.id]); if (f.cover) trashIds_([f.cover]); }
function annName_(req) { return hash_(String(req.serial).toUpperCase() + '|' + req.fileId) + '.json'; }
