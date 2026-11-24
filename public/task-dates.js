// 把业务日期当作 UTC 日期运算，避免时区和夏令时影响列差。
const DAY_MS = 24 * 60 * 60 * 1000;

function dateKey(date) {
  return date.toISOString().slice(0, 10);
}

export function addDays(day, count) {
  const date = new Date(`${day}T00:00:00Z`);
  date.setUTCDate(date.getUTCDate() + count);
  return dateKey(date);
}

export function addMonths(month, count) {
  const date = new Date(`${month}-01T00:00:00Z`);
  date.setUTCMonth(date.getUTCMonth() + count);
  return dateKey(date).slice(0, 7);
}

export function daysBetween(from, to) {
  return (Date.parse(`${to}T00:00:00Z`) - Date.parse(`${from}T00:00:00Z`)) / DAY_MS;
}

export function monthGrid(month) {
  const first = `${month}-01`;
  const offset = (new Date(`${first}T00:00:00Z`).getUTCDay() + 6) % 7;
  const monday = addDays(first, -offset);
  return Array.from({ length: 42 }, (_, index) => addDays(monday, index));
}
