/**
 * Пример интеграции с CRM (JavaScript/fetch)
 * Замените BASE_URL на адрес вашего VPS.
 */

const BASE_URL = "http://your-vps-ip:8000";

/**
 * Запустить парсинг самовывоза и дождаться результата.
 * @param {string} category  - "rims" | "wheels" | "auto%20parts"
 * @param {string} city      - "astana" | "pavlodar"
 * @returns {Promise<Array>} - массив товаров с самовывозом
 */
async function parsePickup(category, city) {
  // 1. Запустить задачу
  const { task_id } = await fetch(`${BASE_URL}/api/parse/pickup`, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ category, city }),
  }).then((r) => r.json());

  // 2. Polling каждые 3 секунды
  while (true) {
    await new Promise((r) => setTimeout(r, 3000));

    const task = await fetch(`${BASE_URL}/api/tasks/${task_id}`).then((r) =>
      r.json()
    );

    console.log(`[${task.status}] ${task.message} (${task.progress}/${task.total})`);

    if (task.status === "done") {
      // 3a. Получить JSON
      const { results } = await fetch(
        `${BASE_URL}/api/tasks/${task_id}/json`
      ).then((r) => r.json());
      return results;
    }

    if (task.status === "error") {
      throw new Error(task.message);
    }
  }
}

/**
 * Запустить парсинг нашего магазина и скачать Excel.
 */
async function parseOurShopExcel(category) {
  const { task_id } = await fetch(`${BASE_URL}/api/parse/our-shop`, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ category, merchant_id: "Kama" }),
  }).then((r) => r.json());

  while (true) {
    await new Promise((r) => setTimeout(r, 3000));

    const task = await fetch(`${BASE_URL}/api/tasks/${task_id}`).then((r) =>
      r.json()
    );

    if (task.status === "done") {
      // Скачать Excel как Blob
      const blob = await fetch(
        `${BASE_URL}/api/tasks/${task_id}/excel`
      ).then((r) => r.blob());
      return blob; // сохранить или отобразить в CRM
    }

    if (task.status === "error") throw new Error(task.message);
  }
}

// ─── Пример использования ────────────────────────────────────────────────────

// parsePickup("rims", "pavlodar").then(console.log);
// parseOurShopExcel("rims").then((blob) => { /* save file */ });
