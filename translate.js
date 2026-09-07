/*!
 * translate.js —— 对白翻译工作台·在线版 DeepSeek 调用（纯浏览器，用户自带 key）
 * 逻辑复刻 translate_docx.py 的 chat()：双候选、JSON 输出、按集分批、429/5xx 重试。
 * key 只存在浏览器、只随请求发往 api.deepseek.com，不经任何中转。
 */
(function (global) {
  'use strict';
  const API = 'https://api.deepseek.com/chat/completions';

  const SYSTEM_TWO =
    '你是欧美竖屏短剧对白翻译。把每条中文台词给【2个】不同写法的美式英语字幕候选，' +
    '推荐的那个排在前面。规则：口语自然、字幕短句；保留人名/专有名词；忽略括号舞台提示；' +
    '两个候选在风格/用词上要有明显区别但都自然、都忠于原意。' +
    '输出 JSON 对象 {"lines": ["句1推荐", "句1备选", "句2推荐", "句2备选", ...]}，' +
    '顺序为每句【推荐、备选】交替，共 2×N 条，不要任何多余文字。';

  const SYSTEM_SINGLE =
    '你是欧美竖屏短剧对白翻译。把中文台词逐句译成自然、口语化的美式英语字幕文本。' +
    '规则：1) 口语自然像真人说话，字幕短句优先；2) 保留人名/地名/专有名词；' +
    '3) 忽略并删除台词中间的括号舞台提示（如（指着弗兰克）），只译括号外的台词；' +
    '4) 严格按输入条目顺序翻译，输出 JSON 对象 {"lines": ["第1条译文", ...]}，' +
    '条数与输入一致，不要任何多余文字。';

  function sleep(ms) { return new Promise(function (res) { setTimeout(res, ms); }); }

  function userFriendly(r) {
    // r = {status, bodyMessage}
    if (r.status === 401 || r.status === 403) {
      return 'DeepSeek key 无效或无权限（HTTP ' + r.status + '）。请检查你填的 key。';
    }
    if (r.status === 402) return 'DeepSeek 账户余额不足（HTTP 402），请充值后再试。';
    if (r.status === 429) return 'DeepSeek 限流/并发过高（HTTP 429），稍后会自动重试。';
    if (r.status >= 400 && r.status < 500) {
      return 'DeepSeek 请求被拒（HTTP ' + r.status + '）' + (r.bodyMessage ? '：' + r.bodyMessage : '');
    }
    return 'DeepSeek 服务异常（HTTP ' + r.status + '）' + (r.bodyMessage ? '：' + r.bodyMessage : '');
  }

  /**
   * 一次调用（一批 zhLines）。
   * @returns Promise<string[]>  长度 = two ? 2n : n
   */
  async function chatDeepSeek(opts) {
    const apiKey = opts.apiKey, model = opts.model, zhLines = opts.zhLines;
    const two = opts.two !== false;
    const signal = opts.signal;
    const n = zhLines.length;
    if (!apiKey) { const e = new Error('请先在上方填入你的 DeepSeek API key'); e.friendly = e.message; throw e; }
    if (!n) return [];
    const sysp = two ? SYSTEM_TWO : SYSTEM_SINGLE;
    const target = two ? 2 * n : n;
    const numbered = zhLines.map(function (x, i) { return (i + 1) + '. ' + x; }).join('\n');
    const need = two ? '每条给两个候选，' : '';
    const payload = {
      model: model || 'deepseek-v4-flash',
      messages: [
        { role: 'system', content: sysp },
        { role: 'user', content: '共 ' + n + ' 条台词，' + need + '请翻译：\n' + numbered }
      ],
      temperature: two ? 0.7 : 0.6,
      response_format: { type: 'json_object' },
      max_tokens: 12000,
    };
    const retries = 3;
    for (let attempt = 0; attempt < retries; attempt++) {
      let resp;
      try {
        resp = await fetch(API, {
          method: 'POST',
          headers: { 'Content-Type': 'application/json', Authorization: 'Bearer ' + apiKey },
          body: JSON.stringify(payload),
          signal: signal,
        });
      } catch (e) {
        if (e && e.name === 'AbortError') throw e;
        if (attempt < retries - 1) { await sleep(2000 * (attempt + 1)); continue; }
        const fe = new Error('网络/跨域请求失败，请检查网络后重试（' + e.message + '）。若你用了广告拦截或代理，请放行 api.deepseek.com。');
        fe.friendly = fe.message; throw fe;
      }
      if (resp.status === 200) {
        try {
          const j = await resp.json();
          const content = (j.choices && j.choices[0] && j.choices[0].message && j.choices[0].message.content) || '';
          let lines = JSON.parse(content).lines;
          if (!Array.isArray(lines) || lines.length !== target) {
            throw new Error('返回条数不符 (' + (Array.isArray(lines) ? lines.length : '?') + '/' + target + ')');
          }
          const out = lines.map(function (x) { return String(x).trim(); });
          if (two) for (let i = 0; i < out.length; i++) if (!out[i]) out[i] = out[Math.floor(i / 2) * 2];
          return out;
        } catch (e) {
          // JSON/条数错误 → 让模型重来一次
          if (attempt < retries - 1) { await sleep(2000 * (attempt + 1)); continue; }
          const fe = new Error('DeepSeek 返回内容无法解析（' + (e && e.message) + '），请重试或换模型。');
          fe.friendly = fe.message; throw fe;
        }
      } else {
        let msg = '';
        try { const j = await resp.json(); msg = (j.error && (j.error.message || '')) || j.message || ''; } catch (e) { /* ignore */ }
        if ((resp.status === 429 || resp.status >= 500) && attempt < retries - 1) {
          await sleep(2000 * (attempt + 1)); continue;
        }
        const fe = new Error(userFriendly({ status: resp.status, bodyMessage: msg }));
        fe.friendly = fe.message; throw fe;
      }
    }
    const fe = new Error('DeepSeek 多次重试仍失败，请稍后再试。');
    fe.friendly = fe.message; throw fe;
  }

  /**
   * 整集（可能拆多批）翻译成双候选。
   * @returns Promise<string[]> 长度 = 2 * zhLines.length（推荐、备选交替）
   */
  async function translateTwoCandidates(opts) {
    const zh = opts.zhLines, per = opts.per || 24;
    const apiKey = opts.apiKey, model = opts.model, signal = opts.signal;
    const out = [];
    for (let start = 0; start < zh.length; start += per) {
      const chunk = zh.slice(start, start + per);
      const part = await chatDeepSeek({ apiKey, model, zhLines: chunk, two: true, signal });
      out.push.apply(out, part);
      if (opts.onChunk) opts.onChunk(start / per + 1, Math.ceil(zh.length / per));
    }
    return out;
  }

  const api = { chatDeepSeek, translateTwoCandidates, SYSTEM_TWO, SYSTEM_SINGLE };
  if (typeof module !== 'undefined' && module.exports) module.exports = api;
  global.Translate = api;
})(typeof window !== 'undefined' ? window : globalThis);
