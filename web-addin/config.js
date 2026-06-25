// 用户配置：存于 Outlook 的 roamingSettings（随邮箱漫游，体积小，仅文本）。
// 班次图等大数据走 IndexedDB（见 storage/planStore.js）。

const KEY = 'shiftAddin.config.v1';

export const DEFAULT_CONFIG = {
  emailDomain: 'example.com', // 首次使用请在"配置"面板改为你的真实邮件域名（仅存于本地邮箱，不进仓库）
  subjectTemplate: '班次 {cycle} {login}',
  bodyContentHtml: '<p>请查收你下一周期的班次安排，如有疑问请直接回复本邮件。</p>',
  signatureHtml: '<p>Best regards,</p>',
  cycleWithYear: false
};

/**
 * 从 roamingSettings 读取配置（与默认值合并）。
 * @returns {object}
 */
export function loadConfig() {
  try {
    const raw = Office.context.roamingSettings.get(KEY);
    if (!raw) return { ...DEFAULT_CONFIG };
    const parsed = typeof raw === 'string' ? JSON.parse(raw) : raw;
    return { ...DEFAULT_CONFIG, ...parsed };
  } catch (e) {
    return { ...DEFAULT_CONFIG };
  }
}

/**
 * 保存配置到 roamingSettings。
 * @param {object} cfg
 * @returns {Promise<void>}
 */
export function saveConfig(cfg) {
  return new Promise((resolve, reject) => {
    try {
      Office.context.roamingSettings.set(KEY, JSON.stringify(cfg));
      Office.context.roamingSettings.saveAsync((res) => {
        if (res.status === Office.AsyncResultStatus.Succeeded) resolve();
        else reject(new Error(res.error ? res.error.message : '保存配置失败'));
      });
    } catch (e) {
      reject(e);
    }
  });
}
