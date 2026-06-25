// 班次图片识别与 Login 映射（独立图模式，对应需求 R7 的核心子集）。
//
// 本版只实现"独立图模式"：每位组员一张独立图片，按文件名（必要时配合相邻文本）配人。
// 模板切分模式（整张大图按网格裁剪）留待后续迭代。

import { deriveLoginFromFileName, isValidLogin } from './login.js';

const IMAGE_EXT_RE = /\.(png|jpe?g|gif|bmp|webp|tiff?)$/i;

/**
 * 判断一个附件是否是图片。
 * @param {{name?: string, contentType?: string}} att
 * @returns {boolean}
 */
export function isImageAttachment(att) {
  if (!att) return false;
  if (att.contentType && att.contentType.toLowerCase().startsWith('image/')) return true;
  if (att.name && IMAGE_EXT_RE.test(att.name)) return true;
  return false;
}

/**
 * 把一组图片附件映射到 Login。
 * 优先用文件名推断；文件名得到唯一候选则直接采用，否则标记为待人工处理。
 *
 * @param {Array<{id: string, name: string, contentType?: string}>} attachments
 * @returns {{
 *   mapped: Array<{login: string, attachmentId: string, name: string}>,
 *   anomalies: Array<{type: string, attachmentId: string, name: string, candidates: string[]}>
 * }}
 */
export function mapImagesToLogins(attachments) {
  const images = (attachments || []).filter(isImageAttachment);
  const mapped = [];
  const anomalies = [];
  const usedLogin = new Map(); // login -> attachmentId（检测重复）

  for (const att of images) {
    const candidates = deriveLoginFromFileName(att.name);
    if (candidates.length === 1 && isValidLogin(candidates[0])) {
      const login = candidates[0];
      if (usedLogin.has(login)) {
        anomalies.push({
          type: 'DuplicateLogin',
          attachmentId: att.id,
          name: att.name,
          candidates
        });
      } else {
        usedLogin.set(login, att.id);
        mapped.push({ login, attachmentId: att.id, name: att.name });
      }
    } else if (candidates.length === 0) {
      anomalies.push({
        type: 'NoLoginFound',
        attachmentId: att.id,
        name: att.name,
        candidates
      });
    } else {
      anomalies.push({
        type: 'AmbiguousLogin',
        attachmentId: att.id,
        name: att.name,
        candidates
      });
    }
  }

  return { mapped, anomalies };
}
