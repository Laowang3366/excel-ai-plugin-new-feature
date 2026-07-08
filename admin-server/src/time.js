const BEIJING_OFFSET_MS = 8 * 60 * 60 * 1000;
const TIMESTAMP_FIELDS = ["created_at", "updated_at", "activated_at", "last_heartbeat", "expires_at"];

export function formatBeijingDateTime(value) {
  if (!value) return value;

  const match = /^(\d{4})-(\d{2})-(\d{2})[ T](\d{2}):(\d{2}):(\d{2})/.exec(String(value));
  if (!match) return value;

  const [, year, month, day, hour, minute, second] = match;
  const utcMs = Date.UTC(
    Number(year),
    Number(month) - 1,
    Number(day),
    Number(hour),
    Number(minute),
    Number(second)
  );

  return new Date(utcMs + BEIJING_OFFSET_MS).toISOString().slice(0, 19).replace("T", " ");
}

export function withBeijingDateTimes(record) {
  if (!record) return record;

  const converted = { ...record };
  for (const field of TIMESTAMP_FIELDS) {
    if (Object.prototype.hasOwnProperty.call(converted, field)) {
      converted[field] = formatBeijingDateTime(converted[field]);
    }
  }
  return converted;
}
