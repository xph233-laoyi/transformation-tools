/*!
 * docx.js —— 对白翻译工作台·在线版 docx 读写核心（纯浏览器，环境无关）
 * 解析/导出规则与原版 translate_docx.py / server.py 保持一致：
 *   - 只认 body 直属段落（同 python-docx Document.paragraphs）
 *   - 集标题: 第X集/第X回；台词 = 含“：”且冒号前是短角色名、不以 △/【/人物 开头的行
 *   - 导出 = 把“选定英文”作为新段落插到对应中文台词段之后，段落格式克隆原段(pPr)
 *
 * 浏览器用 window.DocxTools，Node 用 module.exports（便于离线自测）。
 */
(function (global) {
  'use strict';
  const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
  const XMLNS_XML = 'http://www.w3.org/XML/1998/namespace';
  const EP_RE = /第\s*[0-9一二三四五六七八九十百千〇零]+\s*[集回]/;

  const CN = { 一: 1, 二: 2, 三: 3, 四: 4, 五: 5, 六: 6, 七: 7, 八: 8, 九: 9, 十: 10 };

  function to_num(t) {
    t = (t || '').trim();
    if (/^\d+$/.test(t)) return parseInt(t, 10);
    if (t in CN) return CN[t];
    const xi = t.indexOf('十');
    if (xi >= 0) {
      const a = xi === 0 ? '' : t.slice(0, xi);   // 十位
      const b = t.slice(xi + 1);                  // 个位
      const hi = a ? (CN[a] !== undefined ? CN[a] : (a === '' ? 1 : null)) : 1;
      if (hi === null) return null;
      return hi * 10 + (b ? (CN[b] !== undefined ? CN[b] : 0) : 0);
    }
    return null;
  }

  function is_dialogue(s) {
    s = (s || '').trim();
    if (!s || s.indexOf('：') < 0) return false;
    const head = s.split('：', 1)[0];
    if (!head || ['人物', '时间', '地点', '场景', '字幕', '标题'].indexOf(head) >= 0) return false;
    if ('△【(（［#'.indexOf(s[0]) >= 0) return false;
    if (head.indexOf('人物') === 0 || head.length > 24) return false;
    return !!(s.split('：', 2)[1] || '').trim();
  }

  /** 取某段落的可见文本：按文档顺序拼所有 w:t（近似 python-docx paragraph.text）。 */
  function paraText(el) {
    let out = '';
    const walk = function (node) {
      for (let n = node.firstChild; n; n = n.nextSibling) {
        if (n.nodeType === 3) { out += n.nodeValue; continue; }
        if (n.nodeType !== 1) continue;
        const ln = n.localName || n.nodeName;
        if (ln === 't') out += n.textContent || '';
        else if (ln === 'tab') out += '\t';
        else walk(n);
      }
    };
    walk(el);
    return out;
  }

  function newEl(doc, name) { return doc.createElementNS(W, 'w:' + name); }
  function isName(node, name) {
    return node && node.nodeType === 1 &&
      (node.localName === name || (node.nodeName || '').split(':').pop() === name);
  }

  /**
   * 解析 document.xml 文本 → 剧本结构。
   * @returns {Object} {
   *   paras: [{pi, text, ep}],          // body 直属段落（pi=在其中的序号，作稳定 key）
   *   dial:  [{pi, ep, zh}],            // 识别的台词（zh=冒号后的正文）
   *   epCount, dialByEp
   * }
   */
  function parseDocxXml(xmlString) {
    const DOMP = global.DOMParser;
    const doc = new DOMP().parseFromString(xmlString, 'application/xml');
    if (doc.getElementsByTagName('parsererror').length) {
      throw new Error('docx 里的 word/document.xml 解析失败');
    }
    const bodies = doc.getElementsByTagNameNS(W, 'body');
    const body = bodies.length ? bodies[0] : null;
    if (!body) throw new Error('未找到 word 正文(body)');
    // body 直属段落
    const pels = [];
    for (let n = body.firstChild; n; n = n.nextSibling) {
      if (n.nodeType === 1 && isName(n, 'p')) pels.push(n);
    }
    const paras = [], dial = [];
    let cur = 1;
    for (let pi = 0; pi < pels.length; pi++) {
      const text = paraText(pels[pi]).trim();
      const m = text.match(EP_RE);
      if (m) {
        const nn = to_num(m[0].replace(/第\s*|\s*[集回]/g, ''));
        if (nn) cur = nn;
      }
      paras.push({ pi, text, ep: cur });
      if (is_dialogue(text)) {
        const zh = text.split('：', 2)[1].trim();
        dial.push({ pi, ep: cur, zh });
      }
    }
    const dialByEp = {};
    dial.forEach(function (d) { (dialByEp[d.ep] = dialByEp[d.ep] || []).push(d); });
    const eps = Object.keys(dialByEp).map(Number).sort(function (a, b) { return a - b; });
    return { paras: paras, dial: dial, dialByEp: dialByEp, eps: eps };
  }

  /**
   * 生成新 docx：把 chosen 里 (pi→英文) 插到对应中文段之后。
   * @returns 新的 document.xml 字符串
   */
  function exportDocxXml(xmlString, chosen) {
    const DOMP = global.DOMParser, XS = global.XMLSerializer;
    const decl = (xmlString.match(/^\s*<\?xml[^>]*\?>\s*/) || [''])[0];
    const doc = new DOMP().parseFromString(xmlString, 'application/xml');
    const bodies = doc.getElementsByTagNameNS(W, 'body');
    const body = bodies.length ? bodies[0] : null;
    if (!body) throw new Error('未找到 word 正文(body)');
    const pels = [];
    for (let n = body.firstChild; n; n = n.nextSibling) {
      if (n.nodeType === 1 && isName(n, 'p')) pels.push(n);
    }
    // 已选中的台词段 pi（升序），插英文在各自原段后
    const pis = Object.keys(chosen).map(Number).sort(function (a, b) { return a - b; });
    for (let k = 0; k < pis.length; k++) {
      const pi = pis[k];
      const text = (chosen[pi] || '').trim();
      if (!text) continue;
      const orig = pels[pi];
      if (!orig) continue;
      const newP = orig.cloneNode(true);
      // 只保留 pPr，去掉其它内容
      for (let i = newP.childNodes.length - 1; i >= 0; i--) {
        const c = newP.childNodes[i];
        if (!(c.nodeType === 1 && isName(c, 'pPr'))) newP.removeChild(c);
      }
      // 原段落首个 run 的 rPr，复制到新 run
      let rpr = null;
      const runs = orig.getElementsByTagNameNS(W, 'r');
      if (runs.length) {
        const rps = runs[0].getElementsByTagNameNS(W, 'rPr');
        if (rps.length) rpr = rps[0].cloneNode(true);
      }
      const r = newEl(doc, 'r');
      if (rpr) r.appendChild(rpr);
      const t = newEl(doc, 't');
      t.setAttributeNS(XMLNS_XML, 'xml:space', 'preserve');
      t.textContent = text;
      r.appendChild(t);
      newP.appendChild(r);
      orig.parentNode.insertBefore(newP, orig.nextSibling);
    }
    let out = new XS().serializeToString(doc);
    // 若序列化本身不带 XML 声明再补一个，避免重复
    if (decl && !/^\s*<\?xml/i.test(out)) out = decl + out;
    return out;
  }

  const api = { to_num, is_dialogue, paraText, parseDocxXml, exportDocxXml, W };
  if (typeof module !== 'undefined' && module.exports) module.exports = api;
  global.DocxTools = api;
})(typeof window !== 'undefined' ? window : globalThis);
