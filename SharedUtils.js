// SharedUtils.gs
// Centralized date utilities for payroll system - Timezone: Asia/Kuwait

function parseFlexibleDate_(dateInput) {
  if (!dateInput) return null;
  if (dateInput instanceof Date && !isNaN(dateInput)) return new Date(dateInput);
  if (typeof dateInput === 'string') {
    let match = dateInput.match(/^(\d{1,2})[-\/](\d{1,2})[-\/](\d{4})$/);
    if (match) {
      const [, day, month, year] = match;
      return new Date(year, month - 1, day);
    }
    match = dateInput.match(/^(\d{4})[-\/](\d{1,2})[-\/](\d{1,2})$/);
    if (match) {
      const [, year, month, day] = match;
      return new Date(year, month - 1, day);
    }
    const parsed = new Date(dateInput);
    if (!isNaN(parsed)) return parsed;
  }
  return null;
}

function formatDdMmYyyy_(date) {
  const d = date instanceof Date ? date : parseFlexibleDate_(date);
  if (!d || isNaN(d)) return '';
  return `${String(d.getDate()).padStart(2, '0')}-${String(d.getMonth() + 1).padStart(2, '0')}-${d.getFullYear()}`;
}

function formatDdMmYyyyHm_(date) {
  const d = date instanceof Date ? date : parseFlexibleDate_(date);
  if (!d || isNaN(d)) return '';
  return `${formatDdMmYyyy_(d)} ${String(d.getHours()).padStart(2, '0')}:${String(d.getMinutes()).padStart(2, '0')}`;
}

function startOfDay_(date) {
  const d = date instanceof Date ? date : parseFlexibleDate_(date);
  if (!d || isNaN(d)) return null;
  const result = new Date(d);
  result.setHours(0, 0, 0, 0);
  return result;
}

function endOfDay_(date) {
  const d = date instanceof Date ? date : parseFlexibleDate_(date);
  if (!d || isNaN(d)) return null;
  const result = new Date(d);
  result.setHours(23, 59, 59, 999);
  return result;
}

function addDays_(date, days) {
  const d = date instanceof Date ? date : parseFlexibleDate_(date);
  if (!d || isNaN(d)) return null;
  const result = new Date(d);
  result.setDate(result.getDate() + days);
  return result;
}

function sameCalendarDate_(date1, date2) {
  const d1 = date1 instanceof Date ? date1 : parseFlexibleDate_(date1);
  const d2 = date2 instanceof Date ? date2 : parseFlexibleDate_(date2);
  if (!d1 || !d2 || isNaN(d1) || isNaN(d2)) return false;
  return d1.getFullYear() === d2.getFullYear() && 
         d1.getMonth() === d2.getMonth() && 
         d1.getDate() === d2.getDate();
}

function tz_() {
  return Session.getScriptTimeZone() || 'Asia/Kuwait';
}