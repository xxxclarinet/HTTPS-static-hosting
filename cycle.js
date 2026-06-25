// 班次周期推断（对应需求 R2）。纯函数，便于单测与复用。
//
// 规则：每月 20 日开启新周期，区间为 [当月20日, 次月20日)（左闭右开）。
// - 源邮件日期 day 在 1..19  → 归属 [本月20日, 次月20日)
// - 源邮件日期 day 在 20..31 → 归属 [次月20日, 再次月20日)（20 日当日归入下一周期）

/**
 * @param {Date} received 源邮件的本地日历日期（用本地时区的年月日即可）
 * @returns {{start: Date, end: Date}} 起止边界（本地日期，end 为开区间）
 */
export function inferCycle(received) {
  const y = received.getFullYear();
  const m = received.getMonth(); // 0-11
  const d = received.getDate();

  let startYear = y;
  let startMonth = m;
  if (d >= 20) {
    // 20 日及以后归入下一周期
    startMonth = m + 1;
    if (startMonth > 11) {
      startMonth = 0;
      startYear = y + 1;
    }
  }

  const start = new Date(startYear, startMonth, 20);
  let endYear = startYear;
  let endMonth = startMonth + 1;
  if (endMonth > 11) {
    endMonth = 0;
    endYear = startYear + 1;
  }
  const end = new Date(endYear, endMonth, 20);
  return { start, end };
}

/**
 * 外部展示格式：默认 "M.D-M.D"（不带年份）；withYear=true 时 "YYYY.M.D-YYYY.M.D"。
 * @param {{start: Date, end: Date}} cycle
 * @param {boolean} [withYear]
 * @returns {string}
 */
export function formatCycle(cycle, withYear = false) {
  const s = cycle.start;
  const e = cycle.end;
  if (withYear) {
    return `${s.getFullYear()}.${s.getMonth() + 1}.${s.getDate()}-${e.getFullYear()}.${e.getMonth() + 1}.${e.getDate()}`;
  }
  return `${s.getMonth() + 1}.${s.getDate()}-${e.getMonth() + 1}.${e.getDate()}`;
}

/**
 * 内部稳定标识，形如 "cyc-2026-06-20"。
 * @param {{start: Date}} cycle
 */
export function cycleId(cycle) {
  const s = cycle.start;
  const mm = String(s.getMonth() + 1).padStart(2, '0');
  const dd = String(s.getDate()).padStart(2, '0');
  return `cyc-${s.getFullYear()}-${mm}-${dd}`;
}

/**
 * 校验手动调整后的区间是否为合法的 "20 日 → 次月 20 日" 连续一个月（对应 R2.5）。
 * @param {Date} start
 * @param {Date} end
 * @returns {boolean}
 */
export function isValidCycle(start, end) {
  if (!(start instanceof Date) || !(end instanceof Date)) return false;
  if (start.getDate() !== 20 || end.getDate() !== 20) return false;
  let em = start.getMonth() + 1;
  let ey = start.getFullYear();
  if (em > 11) {
    em = 0;
    ey += 1;
  }
  return end.getMonth() === em && end.getFullYear() === ey;
}
