// 主题模板渲染（对应需求 R4.2、R4.3）。
//
// 支持占位符 {cycle} 与 {login}，区分大小写、严格花括号匹配；
// 模板总长上限 200；渲染后若仍残留花括号视为非法。

const MAX_TEMPLATE_LEN = 200;
const ALLOWED_PLACEHOLDERS = ['{cycle}', '{login}'];

/**
 * 校验模板：长度、是否只含允许的占位符（不允许孤立/嵌套花括号）。
 * @param {string} template
 * @returns {{ok: boolean, reason?: string}}
 */
export function validateSubjectTemplate(template) {
  if (typeof template !== 'string' || template.length === 0) {
    return { ok: false, reason: '主题模板不能为空' };
  }
  if (template.length > MAX_TEMPLATE_LEN) {
    return { ok: false, reason: `主题模板长度超过 ${MAX_TEMPLATE_LEN}` };
  }
  // 把允许的占位符抠掉后，不应再出现任何花括号
  let stripped = template;
  for (const p of ALLOWED_PLACEHOLDERS) {
    stripped = stripped.split(p).join('');
  }
  if (stripped.includes('{') || stripped.includes('}')) {
    return { ok: false, reason: '主题模板含有非法或不完整的花括号占位符' };
  }
  return { ok: true };
}

/**
 * 渲染主题。
 * @param {string} template
 * @param {string} cycleDisplay 已格式化的周期文本
 * @param {string} login
 * @returns {string}
 * @throws 当模板非法或 login 为空时
 */
export function renderSubject(template, cycleDisplay, login) {
  const v = validateSubjectTemplate(template);
  if (!v.ok) throw new Error(v.reason);
  if (!login) throw new Error('Login 为空，无法渲染主题');

  const result = template
    .split('{cycle}').join(cycleDisplay ?? '')
    .split('{login}').join(login);

  if (result.includes('{') || result.includes('}')) {
    throw new Error('主题渲染后仍存在未替换的占位符');
  }
  return result;
}
