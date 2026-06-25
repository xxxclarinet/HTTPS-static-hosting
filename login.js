// Member_Login 合法性与解析（对应需求 R3.2）。
//
// 合法 Login：仅小写字母、数字与可选连字符，长度 2..32，且不以连字符开头/结尾。

const LOGIN_RE = /^[a-z0-9](?:[a-z0-9-]{0,30}[a-z0-9])?$/;

/**
 * @param {string} s
 * @returns {boolean}
 */
export function isValidLogin(s) {
  return typeof s === 'string' && s.length >= 2 && s.length <= 32 && LOGIN_RE.test(s);
}

/**
 * 从一段文本中抽取所有"看起来像 Login"的候选 token（去重，保持出现顺序）。
 * 用于图片相邻文本（caption）解析。
 * @param {string} text
 * @returns {string[]}
 */
export function extractLoginCandidates(text) {
  if (!text || typeof text !== 'string') return [];
  const tokens = text.toLowerCase().split(/[^a-z0-9-]+/);
  const seen = new Set();
  const out = [];
  for (let t of tokens) {
    t = t.replace(/^-+/, '').replace(/-+$/, '');
    if (isValidLogin(t) && !seen.has(t)) {
      seen.add(t);
      out.push(t);
    }
  }
  return out;
}

/**
 * 从文件名中推断 Login 候选（先去扩展名，再按非 Login 字符切分）。
 * @param {string} fileName 例如 "ytshunyu.png" 或 "shift_ytshunyu_0620.png"
 * @returns {string[]} 候选 Login 列表（可能为 0/1/多个）
 */
export function deriveLoginFromFileName(fileName) {
  if (!fileName || typeof fileName !== 'string') return [];
  const base = fileName.replace(/\.[^.]+$/, '');
  return extractLoginCandidates(base);
}
