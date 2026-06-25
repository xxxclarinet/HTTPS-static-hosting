// Office.js 封装层：把回调式 API 包成 Promise，并区分读模式 / 撰写模式。
// 这是本加载项里唯一直接依赖 Office.js 的地方。

/**
 * 当前是否撰写模式（有可写的收件人对象）。
 * @returns {boolean}
 */
export function isComposeMode() {
  const item = Office.context.mailbox.item;
  return !!(item && item.to && typeof item.to.setAsync === 'function');
}

function promisify(fn) {
  return new Promise((resolve, reject) => {
    fn((res) => {
      if (res.status === Office.AsyncResultStatus.Succeeded) resolve(res.value);
      else reject(new Error(res.error ? res.error.message : 'Office API 调用失败'));
    });
  });
}

// ---------- 读模式：源邮件 ----------

/**
 * 读取当前（源）邮件的 HTML 正文。
 * @returns {Promise<string>}
 */
export function getSourceBodyHtml() {
  const item = Office.context.mailbox.item;
  return promisify((cb) => item.body.getAsync(Office.CoercionType.Html, cb));
}

/**
 * 源邮件的接收/创建日期（用于推断周期）。
 * @returns {Date}
 */
export function getSourceDate() {
  const item = Office.context.mailbox.item;
  return item.dateTimeCreated ? new Date(item.dateTimeCreated) : new Date();
}

/**
 * 源邮件附件元数据列表。
 * @returns {Array<{id: string, name: string, contentType?: string, isInline?: boolean}>}
 */
export function getSourceAttachments() {
  const item = Office.context.mailbox.item;
  return (item.attachments || []).map((a) => ({
    id: a.id,
    name: a.name,
    contentType: a.contentType,
    isInline: a.isInline
  }));
}

/**
 * 读取某个附件的内容（期望返回 base64）。
 * @param {string} attachmentId
 * @returns {Promise<{format: string, content: string}>}
 */
export function getAttachmentBase64(attachmentId) {
  const item = Office.context.mailbox.item;
  return promisify((cb) => item.getAttachmentContentAsync(attachmentId, cb)).then((c) => ({
    format: c.format,
    content: c.content
  }));
}

// ---------- 撰写模式：新邮件 ----------

/**
 * 把一位组员的内容填充到当前撰写中的邮件，并内嵌班次图。
 * @param {object} member
 * @param {string} member.email
 * @param {string} member.subject
 * @param {string} member.bodyHtml 已包含 <img src="cid:imageName"> 的 HTML
 * @param {string|null} [member.imageBase64]
 * @param {string|null} [member.imageName] 作为 cid 引用名
 * @returns {Promise<void>}
 */
export async function fillCompose(member) {
  const item = Office.context.mailbox.item;

  await promisify((cb) => item.to.setAsync([{ emailAddress: member.email }], cb));
  await promisify((cb) => item.subject.setAsync(member.subject, cb));
  await promisify((cb) =>
    item.body.setAsync(member.bodyHtml, { coercionType: Office.CoercionType.Html }, cb)
  );

  if (member.imageBase64 && member.imageName) {
    await promisify((cb) =>
      item.addFileAttachmentFromBase64Async(
        member.imageBase64,
        member.imageName,
        { isInline: true },
        cb
      )
    );
  }
}
