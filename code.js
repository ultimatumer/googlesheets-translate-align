/*
 * googlesheets-translate-align
 * Copyright (c) 2025 Aleksandr Zhdanov
 * Licensed under the MIT
 * See LICENSE and NOTICE files for details.
 */
/***** UI *****/
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Translate & Align')
    .addItem('Open panel', 'showSidebar')
    .addToUi();
}
function showSidebar() {
  var html = HtmlService.createHtmlOutputFromFile('Sidebar').setTitle('Translate & Align');
  SpreadsheetApp.getUi().showSidebar(html);
}

/***** Headers & detection *****/
function listHeaders(sheetName) {
  var ss = SpreadsheetApp.getActive();
  var sheet = ss.getSheetByName(sheetName);
  if (!sheet) throw new Error('Sheet not found: ' + sheetName);
  var values = sheet.getDataRange().getValues();
  if (!values.length) return { headers: [] };
  return { headers: values[0].map(function (h) { return (h || '').toString(); }) };
}

/** Детект источника:
 * 1) EN/English с данными — приоритетный источник
 * 2) иначе "От редакторов" > "От технарей"
 */
function detectSourceServer(sheetName) {
  var ss = SpreadsheetApp.getActive();
  var sheet = ss.getSheetByName(sheetName);
  if (!sheet) throw new Error('Sheet not found: ' + sheetName);

  var values = sheet.getDataRange().getValues();
  var headers = (values[0] || []).map(function (h) { return (h || '').toString(); });
  var map = {};
  headers.forEach(function (h) { map[String(h).toLowerCase()] = String(h); });

  var found = '';
  var enHeader = map['en'] || map['english'] || 'EN';
  var enHasData = false;

  if (enHeader && headers.indexOf(enHeader) !== -1) {
    var enIdx = headers.indexOf(enHeader);
    for (var r = 1; r < values.length; r++) {
      var v = normalize_(values[r][enIdx]);
      if (v) { enHasData = true; break; }
    }
    if (enHasData) found = enHeader;
  }
  if (!found) {
    var a = map['от технарей'];
    var b = map['от редакторов'];
    if (a && !b) found = a;
    else if (a && b) found = b;
  }
  return { found: found, headers: headers, enHeader: enHeader || '', enHasData: enHasData };
}

/***** Props & utils *****/
function _dp() { return PropertiesService.getDocumentProperties(); }
function _sp() { return PropertiesService.getScriptProperties(); }
function _newId_() { return String(Date.now()) + '_' + Math.floor(Math.random() * 1e6); }
function _maskKey_(k) { if (!k) return ''; return k.slice(0, 4) + '…' + k.slice(-4); }
function _recordDeepLError_(obj) {
  try { _sp().setProperty('DEEPL_LAST', JSON.stringify(obj)); } catch (e) { }
}
function _getDeepLLast_() {
  var raw = _sp().getProperty('DEEPL_LAST') || '';
  if (!raw) return null;
  try { return JSON.parse(raw); } catch (e) { return null; }
}

/***** DeepL keys presence for UI *****/
function checkDeepLKeys() { return { hasKeys: getDeepLKeys_().length > 0 }; }

/***** Glossary (Script Properties) *****/
// { srcLang?: "EN", tgtLang: "DE", from: "Drawing Dog", to: "Zeichenhund" }
function glossaryList() {
  var raw = _sp().getProperty('GLOSSARY_JSON') || '[]';
  try { return JSON.parse(raw); } catch (e) { return []; }
}
function glossarySave(entries) {
  if (!Array.isArray(entries)) throw new Error('Bad glossary payload');
  entries.forEach(function (x) {
    if (!x || !x.tgtLang || !x.from || !x.to) {
      throw new Error('Each entry must have tgtLang, from, to');
    }
  });
  _sp().setProperty('GLOSSARY_JSON', JSON.stringify(entries));
  return { ok: true, count: entries.length };
}

/***** Start job *****/
function startJob(cfg) {
  if (!cfg || !cfg.sheetName || !cfg.targets || !cfg.targets.length) {
    throw new Error('Please fill Sheet and at least one target.');
  }
  var ss = SpreadsheetApp.getActive();
  var sheet = ss.getSheetByName(cfg.sheetName);
  if (!sheet) throw new Error('Sheet not found: ' + cfg.sheetName);

  var values = sheet.getDataRange().getValues();
  if (values.length < 2) throw new Error('Sheet has no data.');
  var headers = values[0].map(function (h) { return (h || '').toString(); });
  var totalRows = values.length - 1;

  var sourceHeader = (cfg.sourceHeader || '').toString().trim();
  var enInfo = detectSourceServer(cfg.sheetName);
  if (!sourceHeader) {
    if (enInfo.enHasData && enInfo.enHeader) {
      sourceHeader = enInfo.enHeader;
    } else if (cfg.autoDetect) {
      var map = {};
      headers.forEach(function (h) { map[String(h).toLowerCase()] = String(h); });
      var a = map['от технарей'];
      var b = map['от редакторов'];
      if (a && !b) sourceHeader = a;
      else if (a && b) sourceHeader = b;
      else throw new Error('No source column detected. Please specify Source header manually.');
    }
  }
  if (!sourceHeader) throw new Error('Source header not specified.');
  var srcIdx = headers.indexOf(sourceHeader);
  if (srcIdx === -1) throw new Error('Source header not found: ' + sourceHeader);

  var strict = (cfg.strict !== false);
  var langMap = cfg.langMap || defaultLangMap_(cfg.targets);

  // EN column exists but empty -> add EN as target EN->EN (на самом деле EN-US)
  var enWasAutoAdded = false;
  var hasEnHeader = headers.indexOf(enInfo.enHeader || 'EN') !== -1
    || headers.indexOf('EN') !== -1
    || headers.indexOf('English') !== -1;
  if (hasEnHeader && !enInfo.enHasData && cfg.targets.indexOf('EN') === -1) {
    cfg.targets.push('EN');
    enWasAutoAdded = true;
    if (!langMap['EN']) langMap['EN'] = 'EN-US';
  }

  var effectiveTargets = [];
  var missing = [];
  cfg.targets.forEach(function (t) {
    if (headers.indexOf(t) === -1) {
      if (enWasAutoAdded && t === 'EN') { headers.push('EN'); effectiveTargets.push('EN'); }
      else if (strict) { missing.push(t); }
      else { headers.push(t); effectiveTargets.push(t); }
    } else {
      effectiveTargets.push(t);
    }
  });
  if (strict && missing.length) {
    throw new Error('Strict mode: missing target columns: ' + missing.join(', ') + '. Add them to header row or disable strict.');
  }

  // Гарантируем, что DeepL всегда получит EN-US
  if (effectiveTargets.indexOf('EN') !== -1 && !langMap['EN']) {
    langMap['EN'] = 'EN-US';
  }

  var outName = cfg.writeBack || 'Aligned';
  var out = ss.getSheetByName(outName);
  if (out) ss.deleteSheet(out);
  out = ss.insertSheet(outName);
  out.getRange(1, 1, 1, headers.length).setValues([headers]);

  var jobId = _newId_();
  var state = {
    id: jobId,
    sheetName: cfg.sheetName,
    outName: outName,
    headers: headers,
    sourceHeader: sourceHeader,
    srcIdx: srcIdx,
    targets: effectiveTargets,
    langMap: langMap || {},
    autoTranslate: !!cfg.autoTranslate,
    fillMissing: !!cfg.fillMissing,
    strict: strict,
    totalRows: totalRows,
    cursor: 0,
    batchRows: Math.max(20, Math.min(1000, Number(cfg.batchRows) || 150))
  };
  _dp().setProperty('JOB_' + jobId, JSON.stringify(state));
  return { jobId: jobId, totalRows: totalRows, outName: outName, sourceHeader: sourceHeader };
}

/***** Process one batch *****/
function processJob(jobId) {
  var p = _dp().getProperty('JOB_' + jobId);
  if (!p) throw new Error('Job not found or finished.');
  var st = JSON.parse(p);

  var ss = SpreadsheetApp.getActive();
  var inSheet = ss.getSheetByName(st.sheetName);
  var outSheet = ss.getSheetByName(st.outName);
  if (!inSheet || !outSheet) throw new Error('Sheet missing.');

  if (st.cursor >= st.totalRows) {
    _dp().deleteProperty('JOB_' + jobId);
    return { done: true, cursor: st.cursor, total: st.totalRows, outName: st.outName };
  }

  var take = Math.min(st.batchRows, st.totalRows - st.cursor);
  var startRow = 2 + st.cursor;
  var values = inSheet.getRange(startRow, 1, take, st.headers.length).getValues();

  var src = values.map(function (r) { return normalize_(r[st.srcIdx]); });

  var tgtRawByCol = {};
  st.targets.forEach(function (col) {
    var idx = st.headers.indexOf(col);
    tgtRawByCol[col] = (idx >= 0) ? values.map(function (r) { return normalize_(r[idx]); }) : new Array(take).fill('');
  });

  var keys = getDeepLKeys_();
  var glossary = glossaryList();

  if (st.autoTranslate) {
    if (!keys.length) throw new Error('DeepL keys not configured. Set Script property DEEPL_KEYS.');
    st.targets.forEach(function (col) {
      var code = st.langMap[col] ? st.langMap[col] : col;
      var translated = deeplBatchTranslate_(src, code, keys);
      translated = applyGlossaryBatch_(translated, code, glossary);
      var idx = st.headers.indexOf(col);
      for (var r = 0; r < take; r++) values[r][idx] = translated[r] || '';
    });
  } else {
    st.targets.forEach(function (col) {
      var aligned = alignSegment_(src, tgtRawByCol[col], 0).out;
      if (st.fillMissing && keys.length) {
        var code = st.langMap[col] ? st.langMap[col] : col;
        for (var i = 0; i < aligned.length; i++) {
          if (!aligned[i] || !aligned[i].trim()) {
            var t = deeplTranslateWithFailover_(src[i], code, keys) || '';
            aligned[i] = applyGlossarySingle_(t, code, glossary);
          } else {
            aligned[i] = applyGlossarySingle_(aligned[i], code, glossary);
          }
        }
      } else {
        for (var j = 0; j < aligned.length; j++) {
          aligned[j] = applyGlossarySingle_(aligned[j], st.langMap[col] || col, glossary);
        }
      }
      var idx2 = st.headers.indexOf(col);
      for (var r2 = 0; r2 < take; r2++) values[r2][idx2] = aligned[r2] || '';
    });
  }

  var writeStart = 2 + st.cursor;
  outSheet.getRange(writeStart, 1, take, st.headers.length).setValues(values);

  st.cursor += take;
  _dp().setProperty('JOB_' + jobId, JSON.stringify(st));
  var done = st.cursor >= st.totalRows;
  if (done) _dp().deleteProperty('JOB_' + jobId);

  return { done: done, cursor: st.cursor, total: st.totalRows, outName: st.outName };
}

/***** Alignment helpers *****/
function alignSegment_(sourceArr, targetArr, jStart) {
  var i = 0, j = Math.max(0, Number(jStart) || 0);
  var N = sourceArr.length, M = targetArr.length;
  var out = [];
  while (i < N) {
    var s = normalize_(sourceArr[i]);
    var t = (j < M) ? normalize_(targetArr[j]) : '';
    var sim = similarity_(s, t);
    if (j < M && (sim >= 0.35 || (!s && !t))) { out.push(t); i++; j++; continue; }
    var merged = false;
    if (j < M) {
      for (var k = 2; k <= 3; k++) {
        if (i + k > N) break;
        var sMerge = normalize_(sourceArr.slice(i, i + k).join(' '));
        var simM = similarity_(sMerge, t);
        if (simM >= 0.35) {
          var split = trySplitToK_(t, k);
          if (split && split.length >= k) {
            for (var dx = 0; dx < k; dx++) out.push(split[dx]);
            i += k; j += 1; merged = true; break;
          }
        }
      }
      if (merged) continue;
    }
    out.push(''); i++;
  }
  return { out: out, jEnd: j };
}
function trySplitToK_(text, k) {
  var sents = sentSplit_(text);
  if (sents.length >= k) return sents.slice(0, k);
  var parts = text.split(/(?<=[\.!?…,:;])\s+/);
  parts = parts.filter(function (x) { return x && x.trim(); });
  if (parts.length >= k) return parts.slice(0, k);
  return null;
}
function sentSplit_(text) {
  text = normalize_(text);
  if (!text) return [];
  var parts = text.split(/(?<=[\.!?…])\s+|\n+/);
  var out = [];
  for (var i = 0; i < parts.length; i++) {
    var t = (parts[i] || '').trim();
    if (t) out.push(t);
  }
  return out;
}
function normalize_(s) {
  if (s === null || s === undefined) return '';
  s = String(s).trim().replace(/\s+/g, ' ');
  return s;
}
function similarity_(a, b) {
  if (!a && !b) return 1.0;
  if (!a || !b) return 0.0;
  var lenScore = (Math.min(a.length, b.length) / Math.max(a.length, b.length));
  var ja = jaccard_(a, b);
  return 0.5 * lenScore + 0.5 * ja;
}
function jaccard_(a, b) {
  var A = (a.toLowerCase().match(/\w+/g) || []).reduce(function (s, x) { s[x] = 1; return s; }, {});
  var B = (b.toLowerCase().match(/\w+/g) || []).reduce(function (s, x) { s[x] = 1; return s; }, {});
  var inter = 0, union = 0, seen = {};
  for (var k in A) { union++; if (B[k]) inter++; seen[k] = 1; }
  for (var k2 in B) { if (!seen[k2]) union++; }
  if (union === 0) return 1;
  return inter / union;
}

/***** DeepL (failover + instrumentation) *****/
function getDeepLKeys_() {
  var s = _sp().getProperty('DEEPL_KEYS') || '';
  return s.split(',').map(function (x) { return x.trim(); }).filter(String);
}
function deeplTranslateWithFailover_(text, targetLang, keys) {
  if (!text || !text.trim()) return '';
  for (var i = 0; i < keys.length; i++) {
    var key = keys[i];
    var endpoint = key.endsWith(':fx') ? 'https://api-free.deepl.com/v2/translate' : 'https://api.deepl.com/v2/translate';
    var options = {
      method: 'post',
      muteHttpExceptions: true,
      contentType: 'application/x-www-form-urlencoded',
      headers: { 'Authorization': 'DeepL-Auth-Key ' + key },
      payload: 'text=' + encodeURIComponent(text) + '&target_lang=' + encodeURIComponent(targetLang)
    };
    try {
      var resp = UrlFetchApp.fetch(endpoint, options);
      var code = resp.getResponseCode();
      if (code === 200) {
        return JSON.parse(resp.getContentText()).translations[0].text || '';
      }
      _recordDeepLError_({ ts: new Date().toISOString(), where: 'single', key: _maskKey_(key), endpoint: endpoint, code: code, body: String(resp.getContentText()).slice(0, 500) });
      if (code === 456 || code === 429 || code === 403) continue;
    } catch (e) {
      _recordDeepLError_({ ts: new Date().toISOString(), where: 'single', key: _maskKey_(key), endpoint: endpoint, code: 'EXC', body: String(e) });
      continue;
    }
  }
  return '';
}
function deeplBatchTranslate_(texts, targetLang, keys) {
  var out = new Array(texts.length).fill('');
  var batchSize = 50;
  for (var i = 0; i < texts.length; i += batchSize) {
    var slice = texts.slice(i, i + batchSize);
    var res = deeplBatch_(slice, targetLang, keys);
    for (var j = 0; j < slice.length; j++) out[i + j] = res[j] || '';
    Utilities.sleep(150);
  }
  return out;
}
function deeplBatch_(texts, targetLang, keys) {
  for (var k = 0; k < keys.length; k++) {
    var key = keys[k];
    var endpoint = key.endsWith(':fx') ? 'https://api-free.deepl.com/v2/translate' : 'https://api.deepl.com/v2/translate';
    var parts = ['target_lang=' + encodeURIComponent(targetLang)];
    for (var i = 0; i < texts.length; i++) parts.push('text=' + encodeURIComponent(texts[i] || ''));
    var options = {
      method: 'post',
      muteHttpExceptions: true,
      contentType: 'application/x-www-form-urlencoded',
      headers: { 'Authorization': 'DeepL-Auth-Key ' + key },
      payload: parts.join('&')
    };
    try {
      var resp = UrlFetchApp.fetch(endpoint, options);
      var code = resp.getResponseCode();
      if (code === 200) {
        var data = JSON.parse(resp.getContentText());
        var arr = (data.translations || []).map(function (t) { return t.text || ''; });
        var out = new Array(texts.length).fill('');
        for (var j = 0; j < Math.min(arr.length, texts.length); j++) out[j] = arr[j] || '';
        return out;
      }
      _recordDeepLError_({ ts: new Date().toISOString(), where: 'batch', key: _maskKey_(key), endpoint: endpoint, code: code, body: String(resp.getContentText()).slice(0, 500) });
      if (code === 456 || code === 429 || code === 403) continue;
      throw new Error('DeepL error ' + code + ': ' + resp.getContentText());
    } catch (e) {
      _recordDeepLError_({ ts: new Date().toISOString(), where: 'batch', key: _maskKey_(key), endpoint: endpoint, code: 'EXC', body: String(e) });
      if (k === keys.length - 1) throw e;
      continue;
    }
  }
  return new Array(texts.length).fill('');
}

/***** Quick DeepL key test (manual) *****/
function testDeepL() {
  var keys = getDeepLKeys_();
  if (!keys.length) throw new Error('DEEPL_KEYS not set');
  var key = keys[0];
  var endpoint = key.endsWith(':fx') ? 'https://api-free.deepl.com/v2/translate' : 'https://api.deepl.com/v2/translate';
  var resp = UrlFetchApp.fetch(endpoint, {
    method: 'post',
    contentType: 'application/x-www-form-urlencoded',
    headers: { 'Authorization': 'DeepL-Auth-Key ' + key },
    payload: 'text=Hello world&target_lang=DE'
  });
  Logger.log(resp.getResponseCode() + ' ' + resp.getContentText());
}

/***** Diagnostics exposed to UI *****/
function deepLHealth() {
  var keys = getDeepLKeys_();
  if (!keys.length) return { ok: false, items: [], note: 'No DEEPL_KEYS' };
  var items = [];
  for (var i = 0; i < keys.length; i++) {
    var key = keys[i];
    var base = key.endsWith(':fx') ? 'https://api-free.deepl.com/v2' : 'https://api.deepl.com/v2';
    var headers = { 'Authorization': 'DeepL-Auth-Key ' + key };
    var usageCode = null, usageBody = null, trCode = null, trBody = null;
    try {
      var u = UrlFetchApp.fetch(base + '/usage', { method: 'get', headers: headers, muteHttpExceptions: true });
      usageCode = u.getResponseCode(); usageBody = u.getContentText();
    } catch (e) { usageCode = 'EXC'; usageBody = String(e); }
    try {
      var t = UrlFetchApp.fetch(base + '/translate', {
        method: 'post', headers: headers, muteHttpExceptions: true,
        contentType: 'application/x-www-form-urlencoded', payload: 'text=ping&target_lang=DE'
      });
      trCode = t.getResponseCode(); trBody = t.getContentText();
    } catch (e) { trCode = 'EXC'; trBody = String(e); }
    items.push({ key: _maskKey_(key), endpoint: base, usage_code: usageCode, translate_code: trCode, usage_body: usageBody, translate_body: trBody });
  }
  var ok = items.some(function (x) { return x.usage_code === 200 && x.translate_code === 200; });
  return { ok: ok, items: items };
}
function getDeepLDashboard() {
  return {
    hasKeys: getDeepLKeys_().length > 0,
    health: deepLHealth(),
    last: _getDeepLLast_()
  };
}

/***** Defaults *****/
function defaultLangMap_(targets) {
  var map = {};
  (targets || []).forEach(function (t) {
    if (t === 'CH') map['CH'] = 'ZH-HANT';
    else if (t === 'GR') map['GR'] = 'EL';
    else if (t === 'EN') map['EN'] = map['EN'] || 'EN-US';
  });
  return map;
}

/***** Glossary application *****/
function applyGlossaryBatch_(arr, tgtCode, glossary) {
  var out = new Array(arr.length);
  for (var i = 0; i < arr.length; i++) {
    out[i] = applyGlossarySingle_(arr[i], tgtCode, glossary);
  }
  return out;
}
function applyGlossarySingle_(text, tgtCode, glossary) {
  var t = normalize_(text);
  if (!t) return t;
  var tgt = String(tgtCode || '').toUpperCase();
  var rules = (glossary || []).filter(function (g) {
    return g && g.tgtLang && g.from && g.to && String(g.tgtLang).toUpperCase() === tgt;
  });
  rules.forEach(function (g) {
    try {
      var from = String(g.from).trim();
      if (!from) return;
      var pattern = new RegExp(escapeRegExp_(from), 'gi');
      t = t.replace(pattern, g.to);
    } catch (e) { }
  });
  return t;
}
function escapeRegExp_(s) {
  return String(s).replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
}

/***** Panel *****/
function onHomepage(e) {
  const card = CardService.newCardBuilder()
    .setHeader(CardService.newCardHeader().setTitle('Translate & Align'))
    .addSection(CardService.newCardSection().addWidget(
      CardService.newTextButton()
        .setText('Открыть панель')
        .setOnClickAction(CardService.newAction().setFunctionName('openSidebar'))
    ))
    .build();
  return card;
}
function openSidebar(e) {
  showSidebar();
  return CardService.newActionResponseBuilder().build();
}

