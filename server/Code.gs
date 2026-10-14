/**
 * مكتبة PDF المحمية — الخادم (Google Apps Script)
 * يُلصق هذا الملف في: Google Sheet ← الإضافات ← Apps Script
 * ثم تشغيل الدالة setup مرة واحدة، ثم النشر كتطبيق ويب (Anyone).
 */

var SHEETS = {
  serials: ['serial', 'name', 'status', 'expiry', 'device', 'folders', 'created', 'lastSeen', 'note'],
  folders: ['id', 'name', 'order', 'created'],
  files:   ['id', 'folderId', 'name', 'visible', 'size', 'added', 'order', 'parts']
};
var CHUNK = 6 * 1024 * 1024;          // حجم الجزء الواحد (6 ميجا)
var MAX_FILE = 300 * 1024 * 1024;     // أقصى حجم للملف (300 ميجا)

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
  if (!p.getProperty('SECRET')) p.setProperty('SECRET', Utilities.getUuid());
  var d = ss.getSheetByName('Sheet1') || ss.getSheetByName('ورقة1');
  if (d && ss.getSheets().length > 3) ss.deleteSheet(d);
  return 'تم الإعداد. كلمة مرور الأدمن المؤقتة: admin123';
}

/* ===================== أدوات ===================== */
function hash_(s) {
  var b = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, String(s), Utilities.Charset.UTF_8);
  return b.map(function (x) { return ('0' + (x & 255).toString(16)).slice(-2); }).join('');
}
function sheet_(n) { return SpreadsheetApp.getActiveSpreadsheet().getSheetByName(n); }
function heal_(n) {
  var sh = sheet_(n), h = SHEETS[n];
  if (sh.getLastColumn() < h.length) sh.getRange(1, 1, 1, h.length).setValues([h]).setFontWeight('bold');
  return sh;
}
function rows_(n) {
  heal_(n);
  var sh = sheet_(n), v = sh.getDataRange().getValues(), h = v.shift();
  return v.map(function (r, i) {
    var o = { _row: i + 2 };
    h.forEach(function (k, j) { o[k] = r[j]; });
    return o;
  }).filter(function (o) { return String(o[h[0]]) !== ''; });
}
function append_(n, obj) {
  heal_(n).appendRow(SHEETS[n].map(function (k) { return obj[k] === undefined ? '' : obj[k]; }));
}
function update_(n, row, obj) {
  var sh = sheet_(n);
  Object.keys(obj).forEach(function (k) {
    var c = SHEETS[n].indexOf(k);
    if (c >= 0) sh.getRange(row, c + 1).setValue(obj[k]);
  });
}
function delRow_(n, row) { sheet_(n).deleteRow(row); }
function out_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}
function genSerial_() {
  var c = 'ABCDEFGHJKLMNPQRSTUVWXYZ23456789', s = '';
  for (var i = 0; i < 16; i++) { s += c.charAt(Math.floor(Math.random() * c.length)); if (i % 4 === 3 && i < 15) s += '-'; }
  return s;
}
function iso_(d) { return d ? new Date(d).toISOString() : ''; }

/* ===================== نقطة الدخول ===================== */
function doGet() { return out_({ ok: true, app: 'pdf-library', v: 1 }); }

function doPost(e) {
  var lock = LockService.getScriptLock();
  try {
    var req = JSON.parse(e.postData.contents || '{}');
    var a = req.action;
    var isWrite = /^(admin|activate|saveAnn)/.test(a) && !/^(adminLogin|adminData|adminUploadPart|adminUploadCancel)$/.test(a);
    if (isWrite) lock.waitLock(20000);
    var fn = API[a];
    if (!fn) return out_({ ok: false, error: 'إجراء غير معروف' });
    return out_(fn(req));
  } catch (err) {
    return out_({ ok: false, error: String(err && err.message || err) });
  } finally {
    try { lock.releaseLock(); } catch (x) {}
  }
}

/* ===================== التحقق ===================== */
function findSerial_(serial) {
  serial = String(serial || '').trim().toUpperCase();
  var list = rows_('serials');
  for (var i = 0; i < list.length; i++) if (String(list[i].serial).toUpperCase() === serial) return list[i];
  return null;
}
function checkUser_(req, bind) {
  var s = findSerial_(req.serial);
  if (!s) throw new Error('السيريال غير صحيح أو تم حذفه');
  if (s.status !== 'active') throw new Error('هذا السيريال موقوف. تواصل مع الإدارة');
  if (s.expiry && new Date(s.expiry) < new Date()) throw new Error('انتهت صلاحية السيريال. تواصل مع الإدارة للتجديد');
  if (!req.device) throw new Error('معرّف الجهاز مفقود');
  if (s.device && s.device !== req.device) throw new Error('هذا السيريال مفعّل على جهاز آخر');
  if (!s.device) {
    if (!bind) throw new Error('السيريال غير مفعّل على هذا الجهاز');
    update_('serials', s._row, { device: req.device });
    s.device = req.device;
  }
  update_('serials', s._row, { lastSeen: new Date().toISOString() });
  return s;
}
function allowedFolders_(s) {
  var f = String(s.folders || '').trim();
  if (f === '*' ) return null; // الكل
  return f ? f.split(',').map(function (x) { return x.trim(); }) : [];
}
function checkAdmin_(req) {
  var p = PropertiesService.getScriptProperties();
  if (!req.pass || hash_(req.pass) !== p.getProperty('ADMIN_PASS')) throw new Error('كلمة مرور الأدمن غير صحيحة');
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
    var files = rows_('files').filter(function (x) { return x.visible === true || x.visible === 'TRUE'; })
      .filter(function (x) { return ids.indexOf(String(x.folderId)) >= 0; })
      .sort(function (a, b) { return (a.order || 0) - (b.order || 0); })
      .map(function (x) { return { id: x.id, folderId: String(x.folderId), name: x.name, size: x.size, added: iso_(x.added) }; });
    return { ok: true, name: s.name, expiry: iso_(s.expiry), folders: folders, files: files };
  },

  getFile: function (req) {
    var s = checkUser_(req, false);
    var allow = allowedFolders_(s);
    var f = rows_('files').filter(function (x) { return x.id === req.fileId; })[0];
    if (!f || !(f.visible === true || f.visible === 'TRUE')) throw new Error('الملف غير متاح');
    if (allow && allow.indexOf(String(f.folderId)) < 0) throw new Error('غير مسموح بفتح هذا الملف');
    var part = parseInt(req.part, 10) || 0, parts = String(f.parts || '').split(',').filter(String), bytes, total;
    if (parts.length) {                       // ملف مقسّم لأجزاء
      if (part >= parts.length) throw new Error('جزء غير موجود');
      bytes = DriveApp.getFileById(parts[part]).getBlob().getBytes();
      total = parts.length;
    } else {                                  // ملف قديم بملف واحد
      var all = DriveApp.getFileById(f.id).getBlob().getBytes();
      total = Math.max(1, Math.ceil(all.length / CHUNK));
      bytes = all.slice(part * CHUNK, (part + 1) * CHUNK);
    }
    var key = hash_(req.serial + req.device + req.fileId), kb = [];
    for (var i = 0; i < key.length; i++) kb.push(key.charCodeAt(i));
    for (var j = 0; j < bytes.length; j++) bytes[j] = bytes[j] ^ kb[j % kb.length];
    return { ok: true, part: part, parts: total, size: Number(f.size) || 0, data: Utilities.base64Encode(bytes) };
  },

  getAnn: function (req) {
    checkUser_(req, false);
    var name = annName_(req);
    var it = DriveApp.getFolderById(PropertiesService.getScriptProperties().getProperty('ANN')).getFilesByName(name);
    return { ok: true, data: it.hasNext() ? it.next().getBlob().getDataAsString() : '' };
  },

  saveAnn: function (req) {
    checkUser_(req, false);
    var folder = DriveApp.getFolderById(PropertiesService.getScriptProperties().getProperty('ANN'));
    var name = annName_(req), it = folder.getFilesByName(name);
    var data = String(req.data || '');
    if (data.length > 8000000) throw new Error('حجم الملاحظات كبير جداً');
    if (it.hasNext()) it.next().setContent(data); else folder.createFile(name, data, 'application/json');
    return { ok: true };
  },

  /* ---------- الأدمن ---------- */
  adminLogin: function (req) { checkAdmin_(req); return { ok: true }; },

  adminData: function (req) {
    checkAdmin_(req);
    return {
      ok: true,
      serials: rows_('serials').map(function (s) {
        return { serial: s.serial, name: s.name, status: s.status, expiry: iso_(s.expiry), device: s.device ? 'مرتبط' : '',
                 folders: String(s.folders || ''), created: iso_(s.created), lastSeen: iso_(s.lastSeen), note: s.note };
      }),
      folders: rows_('folders').sort(function (a, b) { return (a.order || 0) - (b.order || 0); })
        .map(function (f) { return { id: String(f.id), name: f.name }; }),
      files: rows_('files').sort(function (a, b) { return (a.order || 0) - (b.order || 0); })
        .map(function (x) { return { id: x.id, folderId: String(x.folderId), name: x.name, visible: x.visible === true || x.visible === 'TRUE', size: x.size, added: iso_(x.added) }; })
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
    files.reverse().forEach(function (x) { trashFile_(x); delRow_('files', x._row); });
    var f = rows_('folders').filter(function (x) { return String(x.id) === req.id; })[0];
    if (f) delRow_('folders', f._row);
    return { ok: true };
  },

  adminUpload: function (req) {   // رفع ملف صغير مرة واحدة (للتوافق)
    checkAdmin_(req);
    var r = API.adminUploadPart({ pass: req.pass, uploadId: 'U' + Date.now().toString(36), index: 0, data: req.data });
    return API.adminUploadFinish({ pass: req.pass, uploadId: r.uploadId, parts: [r.id], name: req.name, folderId: req.folderId, size: r.size });
  },
  adminUploadPart: function (req) {
    checkAdmin_(req);
    var bytes = Utilities.base64Decode(req.data);
    if (bytes.length > CHUNK + 1024) throw new Error('الجزء أكبر من المسموح');
    var root = DriveApp.getFolderById(PropertiesService.getScriptProperties().getProperty('ROOT'));
    var file = root.createFile(Utilities.newBlob(bytes, 'application/octet-stream', req.uploadId + '.part' + req.index));
    return { ok: true, uploadId: req.uploadId, index: req.index, id: file.getId(), size: bytes.length };
  },
  adminUploadFinish: function (req) {
    checkAdmin_(req);
    if (!rows_('folders').some(function (x) { return String(x.id) === req.folderId; })) {
      (req.parts || []).forEach(function (id) { try { DriveApp.getFileById(id).setTrashed(true); } catch (e) {} });
      throw new Error('اختر فولدر صحيح');
    }
    if (!req.parts || !req.parts.length) throw new Error('لم يتم رفع أي جزء');
    if (Number(req.size) > MAX_FILE) throw new Error('الملف أكبر من 300 ميجا');
    append_('files', { id: req.uploadId, folderId: req.folderId, name: String(req.name).replace(/\.pdf$/i, ''), visible: true,
                       size: Number(req.size) || 0, added: new Date(), order: rows_('files').length + 1, parts: req.parts.join(',') });
    return { ok: true, id: req.uploadId };
  },
  adminUploadCancel: function (req) {
    checkAdmin_(req);
    (req.parts || []).forEach(function (id) { try { DriveApp.getFileById(id).setTrashed(true); } catch (e) {} });
    return { ok: true };
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
    trashFile_(f);
    delRow_('files', f._row);
    return { ok: true };
  },

  adminCreateSerials: function (req) {
    checkAdmin_(req);
    var n = Math.min(Math.max(parseInt(req.count, 10) || 1, 1), 200), out = [];
    var exp = req.days ? new Date(Date.now() + req.days * 86400000) : '';
    var existing = rows_('serials').map(function (s) { return s.serial; });
    for (var i = 0; i < n; i++) {
      var s; do { s = genSerial_(); } while (existing.indexOf(s) >= 0);
      existing.push(s);
      var name = n > 1 ? (req.name || 'مستخدم') + ' ' + (i + 1) : (req.name || 'مستخدم');
      append_('serials', { serial: s, name: name, status: 'active', expiry: exp, device: '', folders: req.folders || '*', created: new Date(), note: req.note || '' });
      out.push(s);
    }
    return { ok: true, serials: out };
  },
  adminSerial: function (req) {   // تعديل: الحالة، التجديد، الصلاحيات، فك الجهاز، الاسم
    checkAdmin_(req);
    var s = findSerial_(req.serial);
    if (!s) throw new Error('السيريال غير موجود');
    var u = {};
    if (req.status) u.status = req.status;
    if (req.name !== undefined) u.name = req.name;
    if (req.folders !== undefined) u.folders = req.folders;
    if (req.resetDevice) u.device = '';
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
  adminPassword: function (req) {
    checkAdmin_(req);
    if (!req.newPass || String(req.newPass).length < 6) throw new Error('كلمة المرور يجب ألا تقل عن 6 أحرف');
    PropertiesService.getScriptProperties().setProperty('ADMIN_PASS', hash_(req.newPass));
    return { ok: true };
  }
};

function trashFile_(f) {
  var ids = String(f.parts || '').split(',').filter(String);
  if (!ids.length) ids = [f.id];
  ids.forEach(function (id) { try { DriveApp.getFileById(id).setTrashed(true); } catch (e) {} });
}
function annName_(req) { return hash_(String(req.serial).toUpperCase() + '|' + req.fileId) + '.json'; }
