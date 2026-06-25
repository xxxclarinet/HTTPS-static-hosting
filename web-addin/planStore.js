// 分发计划存储（IndexedDB）。
//
// 为什么用 IndexedDB：读模式（在源邮件上解析）与撰写模式（在新邮件里填充）是两个
// 独立的运行上下文，但同源（同一 Pages 地址），可共享 IndexedDB。班次图的 base64
// 体积较大，不适合放进 roamingSettings，故用 IndexedDB 承载整份计划。

const DB_NAME = 'shiftAddinDB';
const STORE = 'plan';
const PLAN_KEY = 'current';

function openDb() {
  return new Promise((resolve, reject) => {
    const req = indexedDB.open(DB_NAME, 1);
    req.onupgradeneeded = () => {
      const db = req.result;
      if (!db.objectStoreNames.contains(STORE)) {
        db.createObjectStore(STORE);
      }
    };
    req.onsuccess = () => resolve(req.result);
    req.onerror = () => reject(req.error);
  });
}

/**
 * 保存整份分发计划。
 * @param {object} plan
 * @returns {Promise<void>}
 */
export async function savePlan(plan) {
  const db = await openDb();
  return new Promise((resolve, reject) => {
    const tx = db.transaction(STORE, 'readwrite');
    tx.objectStore(STORE).put(plan, PLAN_KEY);
    tx.oncomplete = () => resolve();
    tx.onerror = () => reject(tx.error);
  });
}

/**
 * 读取当前分发计划。
 * @returns {Promise<object|null>}
 */
export async function loadPlan() {
  const db = await openDb();
  return new Promise((resolve, reject) => {
    const tx = db.transaction(STORE, 'readonly');
    const req = tx.objectStore(STORE).get(PLAN_KEY);
    req.onsuccess = () => resolve(req.result || null);
    req.onerror = () => reject(req.error);
  });
}

/**
 * 清空当前计划。
 * @returns {Promise<void>}
 */
export async function clearPlan() {
  const db = await openDb();
  return new Promise((resolve, reject) => {
    const tx = db.transaction(STORE, 'readwrite');
    tx.objectStore(STORE).delete(PLAN_KEY);
    tx.oncomplete = () => resolve();
    tx.onerror = () => reject(tx.error);
  });
}
