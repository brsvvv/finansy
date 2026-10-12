/**
 * Фінанси API — нова функція update_ (v2).
 * Замінює стару update_ повністю: виділіть стару функцію від «function update_(req) {»
 * до її закриваючої «}» і вставте замість неї цей код.
 *
 * Що вміє:
 *  - змінювати тип, категорію, підкатегорію (як і раніше — для «На перевірку»);
 *  - змінювати суму, дату й назву (нове — вікно «Редагування»);
 *  - видаляти рядок, якщо прийшло remove: true (нове).
 * Відповідь містить v: 2, щоб застосунок знав, що скрипт оновлено.
 */
function update_(req) {
  const sh = opsSheet_();
  const row = Number(req.row);
  if (!(row >= 2)) return { ok: false, error: 'Не вказано рядок' };
  const id = String(sh.getRange(row, 1).getValue() || '') || 'ROW' + row;
  if (id !== String(req.id)) return { ok: false, error: 'Рядок змінився в таблиці. Оновіть застосунок і спробуйте ще раз' };

  // Видалення
  if (req.remove === true) {
    sh.deleteRow(row);
    log_(req.request_id || id, 'app:delete', { row: row, id: id }, 'OK: видалено рядок ' + row);
    return { ok: true, v: 2, removed: true };
  }

  // Перевірка всіх полів до запису, щоб не записати половину
  const cat = CATS.indexOf(req.cat) > -1 ? req.cat : null;
  if (!cat) return { ok: false, error: 'Невідома категорія' };
  const type = TYPES.indexOf(req.type) > -1 ? req.type : 'Расход';

  let amount = null;
  if (req.amount !== undefined) {
    amount = Math.round(Math.abs(Number(req.amount)) * 100) / 100;
    if (!(amount > 0) || amount > 1e7) return { ok: false, error: 'Неправильна сума' };
  }
  let date = null;
  if (req.date !== undefined) {
    if (!/^\d{4}-\d{2}-\d{2}$/.test(String(req.date))) return { ok: false, error: 'Неправильна дата' };
    const p = String(req.date).split('-').map(Number);
    date = new Date(p[0], p[1] - 1, p[2], 12);
  }

  // Запис
  sh.getRange(row, 3, 1, 3).setValues([[type, cat, String(req.sub || '').slice(0, 60)]]);
  if (date) sh.getRange(row, 2).setValue(date).setNumberFormat('dd.MM.yyyy');
  if (req.desc !== undefined) sh.getRange(row, 6).setValue(String(req.desc).replace(/\s+/g, ' ').trim().slice(0, 120));
  if (amount !== null) sh.getRange(row, 7).setValue(amount).setNumberFormat('#,##0.00 "zł"');
  sh.getRange(row, 9).setValue(false);

  log_(req.request_id || id, 'app:update', { row: row, id: id, type: type, cat: cat, amount: amount, date: req.date, desc: req.desc }, 'OK: рядок ' + row);
  return { ok: true, v: 2 };
}
