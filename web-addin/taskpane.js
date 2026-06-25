// 入口与编排：检测模式（读 / 撰写），驱动解析→清单→逐封填充。
import { inferCycle, formatCycle, cycleId } from './domain/cycle.js';
import { isValidLogin } from './domain/login.js';
import { renderSubject } from './domain/subject.js';
import { buildEmail } from './domain/recipient.js';
import { buildBodyHtml } from './domain/template.js';
import { mapImagesToLogins } from './domain/images.js';
import { loadConfig, saveConfig, DEFAULT_CONFIG } from './config.js';
import { savePlan, loadPlan, clearPlan } from './storage/planStore.js';
import {
  isComposeMode,
  getSourceBodyHtml,
  getSourceDate,
  getSourceAttachments,
  getAttachmentBase64,
  fillCompose
} from './outlook/mailbox.js';

const $ = (id) => document.getElementById(id);

let cfg = { ...DEFAULT_CONFIG };
let parsedRows = []; // 读模式：{login, imageBase64, imageName, contentType, thumbUrl}
let cycleLabel = '';
let cycleKey = '';

function setStatus(msg, kind = '') {
  const el = $('status');
  el.textContent = msg || '';
  el.className = 'status' + (kind ? ' ' + kind : '');
}

function extFromNameOrType(name, contentType) {
  const m = (name || '').match(/\.(png|jpe?g|gif|bmp|webp|tiff?)$/i);
  if (m) return m[0].toLowerCase();
  if (contentType && contentType.toLowerCase().includes('jpeg')) return '.jpg';
  if (contentType && contentType.toLowerCase().startsWith('image/')) {
    return '.' + contentType.toLowerCase().split('/')[1];
  }
  return '.png';
}

// ---------------- 初始化 ----------------

Office.onReady((info) => {
  if (info.host !== Office.HostType.Outlook) {
    setStatus('该加载项需在 Outlook 中运行。', 'error');
    return;
  }
  cfg = loadConfig();
  fillSettingsInputs();
  wireEvents();

  if (isComposeMode()) {
    $('mode-hint').textContent = '撰写模式：把分发清单逐封填充进新邮件。';
    $('compose-panel').classList.remove('hidden');
    refreshComposePanel();
  } else {
    $('mode-hint').textContent = '阅读模式：在源班次邮件上生成分发清单。';
    $('read-panel').classList.remove('hidden');
  }
});

function fillSettingsInputs() {
  $('cfg-domain').value = cfg.emailDomain;
  $('cfg-subject').value = cfg.subjectTemplate;
  $('cfg-body').value = cfg.bodyContentHtml;
  $('cfg-sign').value = cfg.signatureHtml;
  $('cfg-year').checked = !!cfg.cycleWithYear;
}

function wireEvents() {
  $('btn-parse').addEventListener('click', onParse);
  $('btn-save').addEventListener('click', onSavePlan);
  $('btn-save-config').addEventListener('click', onSaveConfig);
  $('btn-fill-next').addEventListener('click', onFillNext);
  $('btn-clear-plan').addEventListener('click', onClearPlan);
}

// ---------------- 配置 ----------------

async function onSaveConfig() {
  cfg = {
    emailDomain: $('cfg-domain').value.trim(),
    subjectTemplate: $('cfg-subject').value,
    bodyContentHtml: $('cfg-body').value,
    signatureHtml: $('cfg-sign').value,
    cycleWithYear: $('cfg-year').checked
  };
  try {
    await saveConfig(cfg);
    setStatus('配置已保存。', 'ok');
  } catch (e) {
    setStatus('保存配置失败：' + e.message, 'error');
  }
}

// ---------------- 读模式：解析 ----------------

async function onParse() {
  setStatus('正在解析源邮件…');
  $('btn-parse').disabled = true;
  try {
    const cycle = inferCycle(getSourceDate());
    cycleLabel = formatCycle(cycle, cfg.cycleWithYear);
    cycleKey = cycleId(cycle);
    $('cycle-info').classList.remove('hidden');
    $('cycle-info').innerHTML = `推断班次周期：<b>${cycleLabel}</b>`;

    const atts = getSourceAttachments();
    const { mapped, anomalies } = mapImagesToLogins(atts);

    parsedRows = [];
    for (const m of mapped) {
      const content = await getAttachmentBase64(m.attachmentId);
      if (content.format !== Office.MailboxEnums.AttachmentContentFormat.Base64) {
        anomalies.push({ type: 'UnreadableImage', name: m.name, candidates: [] });
        continue;
      }
      const att = atts.find((a) => a.id === m.attachmentId) || {};
      const ext = extFromNameOrType(m.name, att.contentType);
      parsedRows.push({
        login: m.login,
        imageBase64: content.content,
        imageName: `${m.login}${ext}`,
        contentType: att.contentType || 'image/png',
        thumbUrl: `data:${att.contentType || 'image/png'};base64,${content.content}`
      });
    }

    renderRows();
    renderAnomalies(anomalies);
    $('btn-save').classList.toggle('hidden', parsedRows.length === 0);
    setStatus(`解析完成：${parsedRows.length} 位组员，${anomalies.length} 条待处理。`, 'ok');
  } catch (e) {
    setStatus('解析失败：' + e.message, 'error');
  } finally {
    $('btn-parse').disabled = false;
  }
}

function previewFor(login) {
  let email = '';
  let subject = '';
  try {
    email = buildEmail(login, $('cfg-domain').value.trim());
  } catch (e) {
    email = '（' + e.message + '）';
  }
  try {
    subject = renderSubject($('cfg-subject').value, cycleLabel, login);
  } catch (e) {
    subject = '（' + e.message + '）';
  }
  return { email, subject };
}

function renderRows() {
  const wrap = $('rows');
  wrap.innerHTML = '';
  parsedRows.forEach((r, idx) => {
    const { email, subject } = previewFor(r.login);
    const div = document.createElement('div');
    div.className = 'row';
    div.innerHTML = `
      <img class="thumb" src="${r.thumbUrl}" alt="" />
      <div class="meta">
        <div><input type="text" value="${r.login}" data-idx="${idx}" class="login-input" /></div>
        <div class="email">${email}</div>
        <div class="subject">${subject}</div>
      </div>`;
    wrap.appendChild(div);
  });
  wrap.querySelectorAll('.login-input').forEach((inp) => {
    inp.addEventListener('input', (e) => {
      const i = Number(e.target.dataset.idx);
      parsedRows[i].login = e.target.value.trim();
      // 更新该行的 email/subject 显示
      const { email, subject } = previewFor(parsedRows[i].login);
      const meta = e.target.closest('.meta');
      meta.querySelector('.email').textContent = email;
      meta.querySelector('.subject').textContent = subject;
    });
  });
}

function renderAnomalies(anomalies) {
  const wrap = $('anomalies');
  wrap.innerHTML = '';
  for (const a of anomalies) {
    const div = document.createElement('div');
    div.className = 'anomaly';
    const cand = a.candidates && a.candidates.length ? `候选：${a.candidates.join(', ')}` : '';
    div.textContent = `⚠ ${a.name || ''} — ${anomalyText(a.type)} ${cand}`;
    wrap.appendChild(div);
  }
}

function anomalyText(type) {
  switch (type) {
    case 'NoLoginFound': return '文件名未识别出 Login';
    case 'AmbiguousLogin': return '文件名含多个 Login 候选，需人工确认';
    case 'DuplicateLogin': return '该 Login 已对应另一张图';
    case 'UnreadableImage': return '图片内容无法读取';
    default: return '待处理';
  }
}

// ---------------- 保存计划 ----------------

async function onSavePlan() {
  const domain = $('cfg-domain').value.trim();
  const subjectTpl = $('cfg-subject').value;
  const bodyContentHtml = $('cfg-body').value;
  const signatureHtml = $('cfg-sign').value;

  const members = [];
  const problems = [];
  for (const r of parsedRows) {
    if (!isValidLogin(r.login)) {
      problems.push(`${r.login || '(空)'}: Login 格式非法`);
      continue;
    }
    try {
      const email = buildEmail(r.login, domain);
      const subject = renderSubject(subjectTpl, cycleLabel, r.login);
      const bodyHtml = buildBodyHtml({
        login: r.login,
        bodyContentHtml,
        signatureHtml,
        imageCid: r.imageName
      });
      members.push({
        login: r.login,
        email,
        subject,
        bodyHtml,
        imageBase64: r.imageBase64,
        imageName: r.imageName,
        status: 'pending'
      });
    } catch (e) {
      problems.push(`${r.login}: ${e.message}`);
    }
  }

  if (members.length === 0) {
    setStatus('没有可保存的有效组员。' + (problems.length ? ' ' + problems.join('；') : ''), 'error');
    return;
  }

  const plan = {
    createdAt: new Date().toISOString(),
    cycleLabel,
    cycleId: cycleKey,
    emailDomain: domain,
    members
  };
  try {
    await savePlan(plan);
    let msg = `已保存分发清单：${members.length} 位组员。请新建空白邮件、在撰写窗口打开本加载项逐封填充。`;
    if (problems.length) msg += ` 另有 ${problems.length} 条被跳过：${problems.join('；')}`;
    setStatus(msg, 'ok');
  } catch (e) {
    setStatus('保存清单失败：' + e.message, 'error');
  }
}

// ---------------- 撰写模式：逐封填充 ----------------

async function refreshComposePanel() {
  const plan = await loadPlan();
  const wrap = $('pending-list');
  wrap.innerHTML = '';
  if (!plan || !plan.members || plan.members.length === 0) {
    $('compose-status').textContent = '暂无分发清单。请先在源邮件（阅读模式）里解析并保存清单。';
    $('btn-fill-next').disabled = true;
    return;
  }
  const pending = plan.members.filter((m) => m.status === 'pending');
  $('compose-status').innerHTML =
    `周期 <b>${plan.cycleLabel}</b>：共 ${plan.members.length} 位，待处理 <b>${pending.length}</b> 位。`;
  $('btn-fill-next').disabled = pending.length === 0;

  plan.members.forEach((m) => {
    const div = document.createElement('div');
    div.className = 'row' + (m.status === 'done' ? ' done' : '');
    div.innerHTML = `
      <div class="meta">
        <div><b>${m.login}</b> ${m.status === 'done' ? '✓ 已填充' : ''}</div>
        <div class="email">${m.email}</div>
        <div class="subject">${m.subject}</div>
      </div>`;
    wrap.appendChild(div);
  });
}

async function onFillNext() {
  $('btn-fill-next').disabled = true;
  setStatus('正在填充当前邮件…');
  try {
    const plan = await loadPlan();
    const next = plan.members.find((m) => m.status === 'pending');
    if (!next) {
      setStatus('清单已全部填充完毕。', 'ok');
      await refreshComposePanel();
      return;
    }
    await fillCompose(next);
    next.status = 'done';
    await savePlan(plan);
    setStatus(`已填充 ${next.login}。请按 Ctrl+S 存为草稿，再新建下一封。`, 'ok');
    await refreshComposePanel();
  } catch (e) {
    setStatus('填充失败：' + e.message, 'error');
    $('btn-fill-next').disabled = false;
  }
}

async function onClearPlan() {
  await clearPlan();
  setStatus('已清空分发清单。', 'ok');
  await refreshComposePanel();
}
