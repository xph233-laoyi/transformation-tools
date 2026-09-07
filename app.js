/*!
 * app.js —— 对白翻译工作台·在线版（纯浏览器）
 * 状态/交互：文件解析 → 逐集翻译(双候选) → 点选手改跳过 → 导出 docx+纯英txt。
 * 关键：key 与剧本都只在本浏览器里；翻译按集、按批直连 api.deepseek.com。
 */
(function () {
  'use strict';
  var $ = function (s) { return document.querySelector(s); };
  var S = {
    key: '', model: 'deepseek-v4-flash',
    file: null,          // {name, base, buf, rawXml, docId, parsed}
    episodes: [],        // [{ep, lines:[{pi, zh, rec, alt, text, skip, src}]}]
    choices: {},         // String(pi) -> 文本 或 {skip:true}
    candidates: {},      // ep -> [rec,alt,...]（当前模型）
    ac: null,            // AbortController
    timer: null
  };

  /* ---------- localStorage ---------- */
  function lsGet(k) { try { return localStorage.getItem(k); } catch (e) { return null; } }
  function lsSet(k, v) { try { localStorage.setItem(k, v); } catch (e) { /* 超限则忽略 */ } }
  function lsDel(k) { try { localStorage.removeItem(k); } catch (e) { } }
  function choiceKey() { return 'dzb.ch.' + S.file.docId; }
  function candKey(ep) { return 'dzb.c.' + S.file.docId + '|' + encodeURIComponent(S.model) + '.e' + ep; }

  function loadChoices() {
    S.choices = {};
    try { S.choices = JSON.parse(lsGet(choiceKey()) || '{}'); } catch (e) { S.choices = {}; }
  }
  function saveChoices() { lsSet(choiceKey(), JSON.stringify(S.choices)); }
  function loadCandidates() {
    S.candidates = {};
    if (!S.file) return;
    S.file.parsed.eps.forEach(function (ep) {
      try {
        var c = JSON.parse(lsGet(candKey(ep)) || 'null');
        var need = 2 * (S.file.parsed.dialByEp[ep] || []).length;
        if (Array.isArray(c) && c.length === need) S.candidates[ep] = c;
      } catch (e) { /* ignore */ }
    });
  }
  function saveCandidates(ep) { lsSet(candKey(ep), JSON.stringify(S.candidates[ep])); }

  /* ---------- 常用 ---------- */
  function saveStat(m) { $('#savestat').textContent = m || ''; }
  function showErr(m) {
    var e = $('#err');
    if (!m) { e.style.display = 'none'; e.textContent = ''; return; }
    e.style.display = 'block';
    e.textContent = '⚠ ' + m;
  }
  function keyNow() { return $('#key').value.trim(); }
  function modelNow() { return $('#model').value.trim() || 'deepseek-v4-flash'; }

  /* ---------- 文件载入 ---------- */
  async function loadFile(file) {
    showErr(null); saveStat('');
    var buf = await file.arrayBuffer();
    var xml;
    try {
      var zip = await JSZip.loadAsync(buf);
      var entry = zip.file('word/document.xml');
      if (!entry) throw new Error('不是有效的 .docx（缺少 word/document.xml）');
      xml = await entry.async('string');
    } catch (e) {
      showErr('读取 docx 失败：' + (e && e.message) + '。请确认文件没损坏、确实是 Word 的 .docx。');
      return;
    }
    var parsed;
    try { parsed = DocxTools.parseDocxXml(xml); }
    catch (e) { showErr('解析剧本失败：' + (e && e.message)); return; }
    if (!parsed.dial.length) {
      showErr('没在这份 docx 里识别到台词行。请确认格式是「角色名：台词」分行、可用「第X集」分集。');
      return;
    }
    S.file = {
      name: file.name,
      base: file.name.replace(/\.docx$/i, ''),
      buf: buf, rawXml: xml, parsed: parsed,
      docId: encodeURIComponent(file.name) + ':' + (file.size || Date.now())
    };
    loadChoices(); loadCandidates();
    buildEpisodes();
    showWorkbench();
  }

  /* ---------- 状态合成（同 python Store.state） ---------- */
  function lineState(pi, zh, rec, alt) {
    var c = S.choices[String(pi)];
    if (c && typeof c === 'object' && c.skip) return { skip: true, text: '', src: 'skip' };
    if (typeof c === 'string' && c) {
      return { skip: false, text: c, src: (c === rec ? 'A' : (c === alt ? 'B' : 'custom')) };
    }
    return { skip: false, text: rec || '', src: (rec ? 'A' : 'none') };
  }

  function buildEpisodes() {
    S.episodes = [];
    var P = S.file.parsed;
    P.eps.forEach(function (ep) {
      var dl = P.dialByEp[ep] || [];
      var cand = S.candidates[ep] || null;
      var lines = dl.map(function (d, k) {
        var rec = cand ? (cand[2 * k] || '') : '';
        var alt = cand ? (cand[2 * k + 1] || '') : '';
        var st = lineState(d.pi, d.zh, rec, alt);
        return { pi: d.pi, zh: d.zh, rec: rec, alt: alt, text: st.text, skip: st.skip, src: st.src };
      });
      S.episodes.push({ ep: ep, lines: lines });
    });
  }

  function chosenText(ln) { return (!ln.skip && ln.text) ? ln.text.trim() : ''; }

  /* ---------- 视图切换 ---------- */
  function showWorkbench() {
    $('#docxname').textContent = '剧本：' + S.file.name;
    $('#drop').style.display = 'none';
    $('#nav').style.display = 'flex';
    $('#main').style.display = 'block';
    $('#importBtn').style.display = '';
    $('#translateAllBtn').style.display = '';
    $('#exportBtn').style.display = '';
    renderAll();
  }

  /* ---------- 渲染 ---------- */
  function renderAll() {
    var main = $('#main'), nav = $('#nav');
    main.innerHTML = ''; nav.innerHTML = '';
    S.episodes.forEach(function (ep) {
      var done = ep.lines.filter(function (l) { return !l.skip && chosenText(l); }).length;
      var h = document.createElement('div'); h.className = 'ep';
      h.innerHTML = '第' + ep.ep + '集 <span class="st" id="st' + ep.ep + '">' + done + '/' + ep.lines.length + '</span>';
      var tr = document.createElement('button');
      tr.className = 'btn sm'; tr.dataset.ep = ep.ep;
      if (S.candidates[ep.ep]) { tr.classList.add('grey'); tr.textContent = '重新翻译本集'; }
      else tr.textContent = '翻译本集(双候选)';
      tr.onclick = function () { translateEp(ep.ep); };
      h.appendChild(tr);
      h.id = 'anch' + ep.ep;
      main.appendChild(h);
      var a = document.createElement('a'); a.href = '#ep' + ep.ep; a.textContent = '第' + ep.ep + '集'; a.id = 'ep' + ep.ep;
      nav.appendChild(a);
      ep.lines.forEach(function (ln) { h.appendChild(lineEl(ep, ln)); });
    });
    S.episodes.forEach(function (ep) { var t = document.getElementById('anch' + ep.ep); if (t) t.id = ''; });
  }

  function lineEl(ep, ln) {
    var div = document.createElement('div');
    div.className = 'line' + (ln.skip ? ' skip' : '');
    div.dataset.pi = ln.pi; div.dataset.ep = ep.ep;
    var spk = ln.zh.indexOf('：') >= 0 ? ln.zh.slice(0, ln.zh.indexOf('：') + 1) : '';
    var txt = spk ? ln.zh.slice(spk.length) : ln.zh;
    div.innerHTML =
      '<div class="zh"><span class="spk"></span><span class="body"></span></div>' +
      '<div class="opts">' +
      '  <button class="opt A"><span class="tag">A · 推荐</span><span class="txt"></span></button>' +
      '  <button class="opt B" style="display:none"><span class="tag">B · 备选</span><span class="txt"></span></button>' +
      '</div>' +
      '<textarea spellcheck="false"></textarea>' +
      '<div class="aux"><label><input type="checkbox" class="skipcb"> 跳过这句（保留中文不译）</label></div>';
    div.querySelector('.spk').textContent = spk;
    div.querySelector('.body').textContent = txt;
    var oa = div.querySelector('.opt.A'), ob = div.querySelector('.opt.B');
    var ta = div.querySelector('textarea');
    if (ln.rec) { oa.querySelector('.txt').textContent = ln.rec; oa.style.display = ''; } else oa.style.display = 'none';
    if (ln.alt) { ob.querySelector('.txt').textContent = ln.alt; ob.style.display = ''; } else ob.style.display = 'none';
    if (ln.rec) oa.onclick = function () { pick(ln, oa, ta, ln.rec); };
    if (ln.alt) ob.onclick = function () { pick(ln, ob, ta, ln.alt); };
    if (!ln.skip) {
      ta.value = ln.text || '';
      if (ln.text && ln.src === 'custom') ta.classList.add('edited');
      if (ln.src === 'A' || ln.src === 'B') (ln.src === 'A' ? oa : ob).classList.add('active');
    }
    ta.addEventListener('input', function () {
      ta.classList.add('edited');
      clearTimeout(S.timer);
      var pi = ln.pi;
      S.timer = setTimeout(function () { saveChoice(pi, ta.value); }, 600);
    });
    var cb = div.querySelector('.skipcb'); cb.checked = ln.skip;
    cb.onchange = function () { saveSkip(ln.pi, cb.checked, div); };
    return div;
  }

  function pick(ln, btn, ta, val) {
    ta.value = val; ta.classList.add('edited');
    saveChoice(ln.pi, val);
    var line = document.querySelectorAll('.line[data-pi="' + ln.pi + '"] .opt');
    line.forEach(function (x) { x.classList.remove('active'); });
    btn.classList.add('active');
    ln.text = val; ln.skip = false; ln.src = 'A';
  }

  function saveChoice(pi, text) {
    if (text && text.trim()) S.choices[String(pi)] = text;
    else delete S.choices[String(pi)];
    saveChoices(); saveStat('已保存');
    refreshEpCount(epOf(pi));
  }
  function saveSkip(pi, skip, div) {
    if (skip) S.choices[String(pi)] = { skip: true };
    else delete S.choices[String(pi)];
    saveChoices(); saveStat('已保存');
    div.classList.toggle('skip', skip);
    // 同步内存里该句的 skip 状态，避免导出/计数错
    findLine(pi).skip = skip;
    refreshEpCount(epOf(pi));
  }
  function epOf(pi) {
    for (var i = 0; i < S.episodes.length; i++)
      for (var j = 0; j < S.episodes[i].lines.length; j++)
        if (S.episodes[i].lines[j].pi === pi) return S.episodes[i].ep;
    return null;
  }
  function findLine(pi) {
    for (var i = 0; i < S.episodes.length; i++)
      for (var j = 0; j < S.episodes[i].lines.length; j++)
        if (S.episodes[i].lines[j].pi === pi) return S.episodes[i].lines[j];
    return null;
  }
  function refreshEpCount(ep) {
    if (ep == null) return;
    var all = [], target = null;
    for (var i = 0; i < S.episodes.length; i++) if (S.episodes[i].ep === ep) target = S.episodes[i];
    if (!target) return;
    var done = target.lines.filter(function (l) { return !l.skip && chosenText(l); }).length;
    var el = document.getElementById('st' + ep); if (el) el.textContent = done + '/' + target.lines.length;
  }

  /* ---------- 翻译 ---------- */
  function setBusy(id, on) {
    ['translateAllBtn', 'stopBtn', 'exportBtn', 'importBtn'].forEach(function (x) {
      var b = document.getElementById(x); if (b) b.disabled = on && x !== 'stopBtn';
    });
    $('#stopBtn').style.display = on ? '' : 'none';
    $('#translateAllBtn').textContent = on ? '翻译中…' : '全部翻译(逐集)';
    document.querySelectorAll('button.btn.sm[data-ep]').forEach(function (b) { b.disabled = on; });
  }
  function epButton(ep) {
    return document.querySelector('button.btn.sm[data-ep="' + ep + '"]');
  }

  async function translateEp(ep) {
    var key = keyNow();
    if (!key) { showErr('请先在上方填入你的 DeepSeek API key（只存你浏览器，用于翻译）。'); return; }
    var dl = S.file.parsed.dialByEp[ep] || [];
    if (!dl.length) return;
    var model = modelNow();
    if (model !== S.model) { S.model = model; loadCandidates(); rebuildCandidates(); }
    S.ac = new AbortController();
    var btn = epButton(ep);
    var old = btn ? btn.textContent : '';
    if (btn) { btn.disabled = true; btn.textContent = '翻译中…'; }
    saveStat('第' + ep + '集翻译中（分批直连 DeepSeek，约几秒/批）…');
    showErr(null);
    try {
      var zhLines = dl.map(function (d) { return d.zh; });
      var cand = await Translate.translateTwoCandidates({
        apiKey: key, model: model, zhLines: zhLines, per: 24, signal: S.ac.signal,
        onChunk: function (c, t) { saveStat('第' + ep + '集 翻译批 ' + c + '/' + t + ' …'); }
      });
      S.candidates[ep] = cand;
      saveCandidates(ep);
      buildEpisodes();
      renderAll();
      saveStat('第' + ep + '集翻译完成，已缓存（自动保存，刷新不丢）。');
    } catch (e) {
      if (e && e.name === 'AbortError') { saveStat('已停止。'); }
      else showErr((e && e.friendly) || ('翻译失败：' + (e && e.message)));
    } finally {
      S.ac = null;
      if (btn && btn.isConnected) { btn.disabled = false; btn.textContent = old; }
    }
  }

  async function translateAll() {
    var key = keyNow();
    if (!key) { showErr('请先在上方填入你的 DeepSeek API key。'); return; }
    var model = modelNow();
    if (model !== S.model) { S.model = model; loadCandidates(); }
    S.ac = new AbortController();
    setBusy(true);
    showErr(null);
    try {
      for (var i = 0; i < S.episodes.length; i++) {
        var ep = S.episodes[i].ep;
        if (S.ac.signal.aborted) break;
        if (S.candidates[ep]) continue;         // 已有缓存跳过
        await translateEp(ep);
        if (S.ac && S.ac.signal.aborted) break;
      }
      if (!(S.ac && S.ac.signal.aborted)) saveStat('全部未翻译的集已完成。');
      else saveStat('已停止翻译。');
    } finally {
      setBusy(false);
    }
  }

  function rebuildCandidates() {
    buildEpisodes(); renderAll();
  }

  /* ---------- 导出 ---------- */
  function buildTxt() {
    var lines = [], lastEp = null;
    S.episodes.forEach(function (ep) {
      ep.lines.forEach(function (ln) {
        var t = chosenText(ln);
        if (!t) return;
        if (ep.ep !== lastEp) { if (lastEp != null) lines.push(''); lines.push('[第' + ep.ep + '集]'); lastEp = ep.ep; }
        lines.push(t);
      });
    });
    return lines.join('\n');
  }

  async function doExport() {
    var b = $('#exportBtn'); b.disabled = true; b.textContent = '导出中…';
    saveStat(''); showErr(null);
    try {
      var chosen = {};
      var n = 0;
      S.episodes.forEach(function (ep) {
        ep.lines.forEach(function (ln) {
          var t = chosenText(ln);
          if (t) { chosen[ln.pi] = t; n++; }
        });
      });
      if (!n) { showErr('还没有任何可导出的台词：先翻译或手填英文，且别勾「跳过」。'); b.disabled = false; b.textContent = '确认并导出成品'; return; }
      var xml = DocxTools.exportDocxXml(S.file.rawXml, chosen);
      var zip = await JSZip.loadAsync(S.file.buf);
      zip.file('word/document.xml', xml);
      var blob = await zip.generateAsync({
        type: 'blob',
        mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
        compression: 'DEFLATE',
        compressionOptions: { level: 6 }
      });
      download(S.file.base + '_成品_中上英下.docx', blob);
      download(S.file.base + '_成品_纯英.txt', new Blob([buildTxt()], { type: 'text/plain;charset=utf-8' }));
      saveStat('已导出 ' + n + ' 句，正在下载两个成品文件。');
    } catch (e) {
      showErr('导出失败：' + (e && e.message));
    } finally {
      b.disabled = false; b.textContent = '确认并导出成品';
    }
  }

  function download(name, blob) {
    var url = URL.createObjectURL(blob);
    var a = document.createElement('a');
    a.href = url; a.download = name; a.style.display = 'none';
    document.body.appendChild(a); a.click();
    setTimeout(function () { URL.revokeObjectURL(url); a.remove(); }, 4000);
  }

  /* ---------- 验证 key ---------- */
  async function checkKey() {
    var key = keyNow();
    if (!key) { showErr('请先填入 key 再验证。'); return; }
    var model = modelNow();
    showErr(null); saveStat('验证中…');
    var btn = $('#checkKey'); btn.disabled = true;
    try {
      var r = await fetch('https://api.deepseek.com/chat/completions', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json', Authorization: 'Bearer ' + key },
        body: JSON.stringify({ model: model, messages: [{ role: 'user', content: 'ping' }], max_tokens: 1 })
      });
      if (r.ok) { saveStat('✓ key 有效，模型可访问。'); }
      else {
        var msg = ''; try { var j = await r.json(); msg = (j.error && j.error.message) || j.message || ''; } catch (e) {}
        showErr('验证失败（HTTP ' + r.status + '）' + (msg ? '：' + msg : ''));
      }
    } catch (e) {
      showErr('网络请求失败：' + (e && e.message));
    } finally { btn.disabled = false; }
  }

  /* ---------- 绑定事件 ---------- */
  function bind() {
    $('#pickBtn').onclick = function () { $('#file').click(); };
    $('#file').addEventListener('change', function (e) {
      var f = e.target.files[0]; if (f) loadFile(f); e.target.value = '';
    });
    $('#demoBtn').onclick = async function () {
      try {
        var r = await fetch('sample/demo_script.docx');
        if (!r.ok) throw new Error('HTTP ' + r.status);
        var blob = await r.blob();
        var f = new File([blob], 'demo_script.docx', { type: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document' });
        loadFile(f);
      } catch (e) { showErr('示例剧本加载失败：' + (e && e.message)); }
    };
    var dz = $('#drop');
    ['dragover', 'dragenter'].forEach(function (ev) {
      dz.addEventListener(ev, function (e) { e.preventDefault(); dz.classList.add('on'); });
    });
    ['dragleave', 'drop'].forEach(function (ev) {
      dz.addEventListener(ev, function (e) { e.preventDefault(); dz.classList.remove('on'); });
    });
    dz.addEventListener('drop', function (e) {
      var f = e.dataTransfer && e.dataTransfer.files && e.dataTransfer.files[0];
      if (f && /\.docx$/i.test(f.name)) loadFile(f);
      else if (f) showErr('只支持 .docx 文件。');
    });
    $('#importBtn').onclick = function () { $('#file').click(); };
    $('#translateAllBtn').onclick = translateAll;
    $('#stopBtn').onclick = function () { if (S.ac) S.ac.abort(); };
    $('#exportBtn').onclick = doExport;
    $('#checkKey').onclick = checkKey;
    $('#keyVis').onclick = function () {
      var k = $('#key'); k.type = (k.type === 'password' ? 'text' : 'password');
      this.textContent = (k.type === 'password' ? '显示' : '隐藏');
    };
    $('#key').addEventListener('input', function () { S.key = this.value; lsSet('dzb.key', this.value); });
    $('#model').addEventListener('change', function () { lsSet('dzb.model', this.value.trim()); });
    // 模型变了 → 已缓存候选可能不属于新模型，重读缓存并重渲染
    $('#model').addEventListener('input', function () {
      var m = this.value.trim() || 'deepseek-v4-flash';
      if (m !== S.model && S.file) { S.model = m; loadCandidates(); buildEpisodes(); renderAll(); }
    });
    // 恢复偏好
    $('#key').value = lsGet('dzb.key') || '';
    $('#model').value = lsGet('dzb.model') || 'deepseek-v4-flash';
    S.model = $('#model').value.trim() || 'deepseek-v4-flash';
    S.key = $('#key').value;
  }

  bind();
})();
