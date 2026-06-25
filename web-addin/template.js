// 邮件正文 HTML 拼接（对应需求 R6）。
//
// 顺序：问候语（Dear {login}）→ 主体内容 → 班次图（内嵌，用 cid 引用）→ 落款。
// 段与段之间留一行空行。

/**
 * HTML 转义，避免 login 等文本破坏结构。
 * @param {string} s
 * @returns {string}
 */
export function escapeHtml(s) {
  return String(s)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;');
}

/**
 * 生成问候语（R6.2，固定规则，不可编辑）。
 * @param {string} login
 */
export function buildGreetingHtml(login) {
  return `<p>Dear ${escapeHtml(login)},</p>`;
}

/**
 * 拼接整封正文。
 * @param {object} args
 * @param {string} args.login
 * @param {string} args.bodyContentHtml 用户维护的主体内容（HTML 片段）
 * @param {string} args.signatureHtml 用户维护的落款（HTML 片段）
 * @param {string|null} [args.imageCid] 班次图的 cid（即内嵌附件名）；无图则不插入
 * @returns {string} 完整 HTML 正文
 */
export function buildBodyHtml({ login, bodyContentHtml, signatureHtml, imageCid }) {
  const greeting = buildGreetingHtml(login);
  const img = imageCid
    ? `<p><img src="cid:${escapeHtml(imageCid)}" alt="${escapeHtml(login)} shift" /></p>`
    : '';
  const parts = [greeting, bodyContentHtml || '', img, signatureHtml || ''].filter(
    (p) => p && p.trim().length > 0
  );
  // 段间空行
  return parts.join('\n<p></p>\n');
}
