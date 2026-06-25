// 收件人地址生成与校验（对应需求 R4.1）。
//
// 规则：Member_Email = {login}@{domain}
// Email_Domain：长度 4..255，仅小写字母/数字/连字符/英文句点，且至少含一个句点。

const DOMAIN_RE = /^[a-z0-9.-]+$/;

/**
 * @param {string} domain
 * @returns {{ok: boolean, reason?: string}}
 */
export function validateDomain(domain) {
  if (typeof domain !== 'string') return { ok: false, reason: '域名必须是字符串' };
  if (domain.length < 4 || domain.length > 255) {
    return { ok: false, reason: '域名长度需在 4..255 之间' };
  }
  if (!DOMAIN_RE.test(domain)) {
    return { ok: false, reason: '域名只能含小写字母、数字、连字符与英文句点' };
  }
  if (!domain.includes('.')) {
    return { ok: false, reason: '域名必须至少包含一个英文句点' };
  }
  return { ok: true };
}

/**
 * @param {string} login
 * @param {string} domain
 * @returns {string} login@domain
 * @throws 当 login 为空或域名非法时
 */
export function buildEmail(login, domain) {
  if (!login) throw new Error('Login 为空，无法生成收件人');
  const v = validateDomain(domain);
  if (!v.ok) throw new Error(v.reason);
  return `${login}@${domain}`;
}

/**
 * 基础邮件地址格式校验（用于分发前自检）。
 * @param {string} email
 * @returns {boolean}
 */
export function isValidEmail(email) {
  if (typeof email !== 'string') return false;
  const at = email.indexOf('@');
  if (at <= 0 || at !== email.lastIndexOf('@')) return false;
  const local = email.slice(0, at);
  const domain = email.slice(at + 1);
  if (!local) return false;
  return validateDomain(domain).ok;
}
