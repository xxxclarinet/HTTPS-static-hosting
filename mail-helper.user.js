// ==UserScript==
// @name         邮件助手 (Outlook Web)
// @namespace    https://xxxclarinet.github.io/HTTPS-static-hosting/
// @version      1.0.0
// @description  Outlook 网页版右下角悬浮“邮件助手”：快捷发送（深链接预填）、邮件模板、收件人分组、文本片段。零权限、不越权。
// @match        https://outlook.office.com/*
// @match        https://outlook.office365.com/*
// @match        https://outlook.cloud.microsoft/*
// @run-at       document-idle
// @grant        none
// ==/UserScript==

/*
 * 邮件助手（用户脚本版）
 * - 快捷发送：用 Outlook Web 深链接打开预填好的撰写窗口（to/cc/subject/body）。
 * - 模板 / 收件人分组 / 文本片段：localStorage 持久化，面板内管理。
 * - 占位符：{date} {time} {weekday} 自动替换，其余 {token} 使用时弹窗询问。
 * - 全程 DOM API 构建界面（不使用 innerHTML），兼容 OWA 的 Trusted Types。
 */

(function () {
  'use strict';

  // ---------------- 存储 ----------------

  const STORE_KEY = 'mailHelper.state.v1';

  const DEFAULT_STATE = {
    settings: {
      composeBase: 'https://outlook.office.com/mail/deeplink/compose',
      signature: ''
    },
    templates: [
      {
        id: 'tpl-demo',
        name: '示例模板',
        to: '',
        cc: '',
        subject: '关于 {topic} 的通知 {date}',
        body: 'Hi,\n\n{topic} 的相关信息如下：\n\n\n谢谢！'
      }
    ],
    groups: [],
    snippets: [
      { id: 'sni-demo', name: '示例片段', text: '如有疑问，请直接回复本邮件。' }
    ]
  };

  function loadState() {
    try {
      const raw = localStorage.getItem(STORE_KEY);
      if (!raw) return JSON.parse(JSON.stringify(DEFAULT_STATE));
      const parsed = JSON.parse(raw);
      return {
        settings: { ...DEFAULT_STATE.settings, ...(parsed.settings || {}) },
        templates: Array.isArray(parsed.templates) ? parsed.templates : [],
        groups: Array.isArray(parsed.groups) ? parsed.groups : [],
        snippets: Array.isArray(parsed.snippets) ? parsed.snippets : []
      };
    } catch (e) {
      return JSON.parse(JSON.stringify(DEFAULT_STATE));
    }
  }

  function saveState() {
    localStorage.setItem(STORE_KEY, JSON.stringify(state));
  }

  let state = loadState();
  const uid = (p) => p + '-' + Date.now().toString(36) + Math.random().toString(36).slice(2, 6);

  // ---------------- 占位符 ----------------

  function applyPlaceholders(text) {
    if (!text) return '';
    const now = new Date();
    const pad = (n) => String(n).padStart(2, '0');
    const builtins = {
      date: `${now.getFullYear()}-${pad(now.getMonth() + 1)}-${pad(now.getDate())}`,
      time: `${pad(now.getHours())}:${pad(now.getMinutes())}`,
      weekday: ['周日', '周一', '周二', '周三', '周四', '周五', '周六'][now.getDay()]
    };
    const asked = {};
    return text.replace(/\{([a-zA-Z0-9_\u4e00-\u9fa5]+)\}/g, (m, key) => {
      const k = key.toLowerCase();
      if (k in builtins) return builtins[k];
      if (key in asked) return asked[key];
      const v = window.prompt(`请填写占位符 {${key}} 的值：`, '');
      asked[key] = v == null ? m : v;
      return asked[key];
    });
  }

  // ---------------- 深链接与撰写 ----------------

  function withSignature(body) {
    const sig = (state.settings.signature || '').trim();
    if (!sig) return body || '';
    return (body || '') + '\n\n' + sig;
  }

  function buildComposeUrl(fields) {
    const base = state.settings.composeBase || DEFAULT_STATE.settings.composeBase;
    const p = new URLSearchParams();
    if (fields.to) p.set('to', fields.to);
    if (fields.cc) p.set('cc', fields.cc);
    if (fields.bcc) p.set('bcc', fields.bcc);
    if (fields.subject) p.set('subject', fields.subject);
    if (fields.body) p.set('body', fields.body);
    return base + '?' + p.toString();
  }

  function openCompose(fields) {
    window.open(buildComposeUrl(fields), '_blank', 'noopener');
  }

  // 邮件地址合并（追加、去重、用逗号分隔）
  function mergeRecipients(existing, addition) {
    const split = (s) => (s || '').split(/[;,]/).map((x) => x.trim()).filter(Boolean);
    const set = new Set(split(existing));
    for (const a of split(addition)) set.add(a);
    return Array.from(set).join('; ');
  }

  // ---------------- 样式 ----------------

  function injectStyles() {
    const css = `
    #mh-fab{position:fixed;right:16px;bottom:16px;z-index:2147483646;background:#0f6cbd;color:#fff;
      border:none;border-radius:20px;padding:9px 14px;font:13px "Segoe UI",sans-serif;cursor:pointer;
      box-shadow:0 2px 8px rgba(0,0,0,.3);}
    #mh-panel{position:fixed;right:16px;bottom:60px;z-index:2147483646;width:360px;max-height:80vh;overflow:auto;
      background:#fff;color:#201f1e;border:1px solid #c8c6c4;border-radius:8px;box-shadow:0 6px 24px rgba(0,0,0,.25);
      font:13px "Segoe UI",sans-serif;padding:12px;}
    #mh-panel h2{font-size:15px;margin:0 0 8px;}
    #mh-panel details{border:1px solid #e1dfdd;border-radius:6px;padding:6px 8px;margin:8px 0;}
    #mh-panel summary{cursor:pointer;font-weight:600;}
    #mh-panel label{display:block;margin:6px 0;color:#605e5c;font-size:12px;}
    #mh-panel input[type=text],#mh-panel textarea,#mh-panel select{width:100%;font-size:12px;padding:5px;
      border:1px solid #c8c6c4;border-radius:3px;box-sizing:border-box;}
    #mh-panel button{font-size:12px;padding:6px 10px;border-radius:4px;cursor:pointer;border:1px solid #c8c6c4;background:#fff;margin:3px 4px 0 0;}
    #mh-panel button.mh-primary{background:#0f6cbd;color:#fff;border-color:#0f6cbd;}
    #mh-row{display:flex;gap:6px;flex-wrap:wrap;}
    #mh-status{font-size:12px;margin:6px 0;min-height:16px;}
    #mh-status.err{color:#a4262c;} #mh-status.ok{color:#107c10;}
    .mh-sub{color:#605e5c;font-size:11px;margin:2px 0 6px;}
    `;
    const style = document.createElement('style');
    style.textContent = css;
    document.head.appendChild(style);
  }

  // ---------------- DOM 辅助（不用 innerHTML） ----------------

  function el(tag, attrs, children) {
    const e = document.createElement(tag);
    if (attrs) {
      for (const k in attrs) {
        const v = attrs[k];
        if (k === 'class') e.className = v;
        else if (k === 'text') e.textContent = v;
        else if (k === 'value') e.value = v;
        else if (k === 'onclick') e.addEventListener('click', v);
        else if (k === 'onchange') e.addEventListener('change', v);
        else e.setAttribute(k, v);
      }
    }
    if (children) for (const c of children) {
      if (c == null) continue;
      e.appendChild(typeof c === 'string' ? document.createTextNode(c) : c);
    }
    return e;
  }

  // ---------------- 界面 ----------------

  let panel, statusEl;
  let qTo, qCc, qSubject, qBody;
  let tplSel, grpSel, sniSel;
  let setBase, setSig;

  function setStatus(msg, kind) {
    statusEl.textContent = msg || '';
    statusEl.className = kind || '';
  }

  function fillSelect(sel, items) {
    sel.replaceChildren(el('option', { value: '', text: '— 选择 —' }));
    for (const it of items) sel.appendChild(el('option', { value: it.id, text: it.name }));
  }

  function refreshSelects() {
    fillSelect(tplSel, state.templates);
    fillSelect(grpSel, state.groups);
    fillSelect(sniSel, state.snippets);
  }

  function currentFields() {
    return {
      to: qTo.value.trim(),
      cc: qCc.value.trim(),
      subject: qSubject.value,
      body: qBody.value
    };
  }

  function buildUI() {
    const fab = el('button', { id: 'mh-fab', text: '邮件助手',
      onclick: () => { panel.style.display = panel.style.display === 'none' ? 'block' : 'none'; } });
    document.body.appendChild(fab);

    // —— 快捷发送 ——
    qTo = el('input', { type: 'text', value: '' });
    qCc = el('input', { type: 'text', value: '' });
    qSubject = el('input', { type: 'text', value: '' });
    qBody = el('textarea', { rows: '5', value: '' });

    const quick = el('div', null, [
      el('label', null, ['收件人', qTo]),
      el('label', null, ['抄送', qCc]),
      el('label', null, ['主题', qSubject]),
      el('label', null, ['正文', qBody]),
      el('div', { id: 'mh-row' }, [
        el('button', { class: 'mh-primary', text: '打开预填邮件', onclick: onOpenCompose }),
        el('button', { text: '复制正文', onclick: onCopyBody }),
        el('button', { text: '清空', onclick: onClear })
      ])
    ]);

    // —— 模板 ——
    tplSel = el('select');
    const tplSection = el('details', null, [
      el('summary', { text: '邮件模板' }),
      el('div', { class: 'mh-sub', text: '把上方“快捷发送”内容存为模板，或载入已有模板。占位符 {date}{time}{weekday} 自动替换，其它 {x} 使用时询问。' }),
      tplSel,
      el('div', { id: 'mh-row' }, [
        el('button', { text: '载入到上方', onclick: onLoadTemplate }),
        el('button', { text: '存为模板', onclick: onSaveTemplate }),
        el('button', { text: '删除', onclick: onDeleteTemplate })
      ])
    ]);

    // —— 收件人分组 ——
    grpSel = el('select');
    const grpSection = el('details', null, [
      el('summary', { text: '收件人分组' }),
      el('div', { class: 'mh-sub', text: '把上方“收件人”存为分组，或把分组成员加到收件人/抄送。' }),
      grpSel,
      el('div', { id: 'mh-row' }, [
        el('button', { text: '加到收件人', onclick: () => onUseGroup('to') }),
        el('button', { text: '加到抄送', onclick: () => onUseGroup('cc') }),
        el('button', { text: '存当前收件人为分组', onclick: onSaveGroup }),
        el('button', { text: '删除', onclick: onDeleteGroup })
      ])
    ]);

    // —— 文本片段 ——
    sniSel = el('select');
    const sniSection = el('details', null, [
      el('summary', { text: '文本片段' }),
      el('div', { class: 'mh-sub', text: '常用文本，复制后到撰写窗口粘贴。' }),
      sniSel,
      el('div', { id: 'mh-row' }, [
        el('button', { text: '复制片段', onclick: onCopySnippet }),
        el('button', { text: '新建片段', onclick: onNewSnippet }),
        el('button', { text: '删除', onclick: onDeleteSnippet })
      ])
    ]);

    // —— 设置 ——
    setBase = el('input', { type: 'text', value: state.settings.composeBase });
    setSig = el('textarea', { rows: '3', value: state.settings.signature });
    const setSection = el('details', null, [
      el('summary', { text: '设置' }),
      el('label', null, ['Compose 基础地址（不同 Outlook Web 域名可能不同）', setBase]),
      el('label', null, ['默认签名（追加到正文末尾）', setSig]),
      el('div', { id: 'mh-row' }, [el('button', { text: '保存设置', onclick: onSaveSettings })])
    ]);

    statusEl = el('div', { id: 'mh-status' });

    panel = el('div', { id: 'mh-panel' }, [
      el('h2', { text: '邮件助手' }),
      quick,
      statusEl,
      tplSection,
      grpSection,
      sniSection,
      setSection
    ]);
    panel.style.display = 'none';
    document.body.appendChild(panel);

    refreshSelects();
  }

  // ---------------- 快捷发送动作 ----------------

  function onOpenCompose() {
    const f = currentFields();
    if (!f.to && !f.subject && !f.body) { setStatus('请至少填写收件人/主题/正文之一。', 'err'); return; }
    openCompose({
      to: f.to,
      cc: f.cc,
      subject: applyPlaceholders(f.subject),
      body: withSignature(applyPlaceholders(f.body))
    });
    setStatus('已打开预填撰写窗口，请核对后发送/存草稿。', 'ok');
  }

  async function onCopyBody() {
    try {
      await navigator.clipboard.writeText(withSignature(applyPlaceholders(qBody.value)));
      setStatus('正文已复制到剪贴板。', 'ok');
    } catch (e) {
      setStatus('复制失败：' + e.message, 'err');
    }
  }

  function onClear() {
    qTo.value = ''; qCc.value = ''; qSubject.value = ''; qBody.value = '';
    setStatus('已清空。', 'ok');
  }

  // ---------------- 模板 ----------------

  function onLoadTemplate() {
    const t = state.templates.find((x) => x.id === tplSel.value);
    if (!t) { setStatus('请先选择一个模板。', 'err'); return; }
    qTo.value = t.to || ''; qCc.value = t.cc || '';
    qSubject.value = t.subject || ''; qBody.value = t.body || '';
    setStatus(`已载入模板：${t.name}`, 'ok');
  }

  function onSaveTemplate() {
    const name = window.prompt('模板名称：', '');
    if (!name) return;
    const f = currentFields();
    state.templates.push({ id: uid('tpl'), name, to: f.to, cc: f.cc, subject: f.subject, body: f.body });
    saveState(); refreshSelects();
    setStatus(`已保存模板：${name}`, 'ok');
  }

  function onDeleteTemplate() {
    const t = state.templates.find((x) => x.id === tplSel.value);
    if (!t) { setStatus('请先选择一个模板。', 'err'); return; }
    if (!window.confirm(`删除模板「${t.name}」？`)) return;
    state.templates = state.templates.filter((x) => x.id !== t.id);
    saveState(); refreshSelects();
    setStatus('已删除模板。', 'ok');
  }

  // ---------------- 收件人分组 ----------------

  function onUseGroup(field) {
    const g = state.groups.find((x) => x.id === grpSel.value);
    if (!g) { setStatus('请先选择一个分组。', 'err'); return; }
    const target = field === 'cc' ? qCc : qTo;
    target.value = mergeRecipients(target.value, g.members);
    setStatus(`已把「${g.name}」加到${field === 'cc' ? '抄送' : '收件人'}。`, 'ok');
  }

  function onSaveGroup() {
    const members = qTo.value.trim();
    if (!members) { setStatus('当前“收件人”为空。', 'err'); return; }
    const name = window.prompt('分组名称：', '');
    if (!name) return;
    state.groups.push({ id: uid('grp'), name, members });
    saveState(); refreshSelects();
    setStatus(`已保存分组：${name}`, 'ok');
  }

  function onDeleteGroup() {
    const g = state.groups.find((x) => x.id === grpSel.value);
    if (!g) { setStatus('请先选择一个分组。', 'err'); return; }
    if (!window.confirm(`删除分组「${g.name}」？`)) return;
    state.groups = state.groups.filter((x) => x.id !== g.id);
    saveState(); refreshSelects();
    setStatus('已删除分组。', 'ok');
  }

  // ---------------- 文本片段 ----------------

  async function onCopySnippet() {
    const s = state.snippets.find((x) => x.id === sniSel.value);
    if (!s) { setStatus('请先选择一个片段。', 'err'); return; }
    try {
      await navigator.clipboard.writeText(applyPlaceholders(s.text));
      setStatus(`已复制片段：${s.name}`, 'ok');
    } catch (e) {
      setStatus('复制失败：' + e.message, 'err');
    }
  }

  function onNewSnippet() {
    const name = window.prompt('片段名称：', '');
    if (!name) return;
    const text = window.prompt('片段内容：', '');
    if (text == null) return;
    state.snippets.push({ id: uid('sni'), name, text });
    saveState(); refreshSelects();
    setStatus(`已新建片段：${name}`, 'ok');
  }

  function onDeleteSnippet() {
    const s = state.snippets.find((x) => x.id === sniSel.value);
    if (!s) { setStatus('请先选择一个片段。', 'err'); return; }
    if (!window.confirm(`删除片段「${s.name}」？`)) return;
    state.snippets = state.snippets.filter((x) => x.id !== s.id);
    saveState(); refreshSelects();
    setStatus('已删除片段。', 'ok');
  }

  // ---------------- 设置 ----------------

  function onSaveSettings() {
    const base = setBase.value.trim();
    if (!/^https:\/\/.+/.test(base)) { setStatus('Compose 基础地址需为 https:// 开头。', 'err'); return; }
    state.settings.composeBase = base;
    state.settings.signature = setSig.value;
    saveState();
    setStatus('设置已保存。', 'ok');
  }

  // ---------------- 启动 ----------------

  function init() {
    if (document.getElementById('mh-fab')) return;
    injectStyles();
    buildUI();
  }

  let tries = 0;
  const timer = setInterval(() => {
    tries++;
    if (document.body) { init(); clearInterval(timer); }
    else if (tries > 50) { clearInterval(timer); }
  }, 300);
})();
