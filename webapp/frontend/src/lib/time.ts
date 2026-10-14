export function formatMinutes(minutes: number): string {
  const h = Math.floor(minutes / 60);
  const m = minutes % 60;
  const period = h < 12 ? "AM" : "PM";
  let h12 = h % 12;
  if (h12 === 0) h12 = 12;
  return `${h12}:${String(m).padStart(2, "0")} ${period}`;
}

export function minutesToInputValue(minutes: number): string {
  const h = Math.floor(minutes / 60);
  const m = minutes % 60;
  return `${String(h).padStart(2, "0")}:${String(m).padStart(2, "0")}`;
}

export function inputValueToMinutes(value: string): number {
  const [h, m] = value.split(":").map(Number);
  return h * 60 + m;
}

export function formatUtcTimestampLocally(value: string): string {
  const hasTimezone = /(?:Z|[+-]\d{2}:\d{2})$/i.test(value);
  const timestamp = new Date(hasTimezone ? value : `${value}Z`);
  return Number.isNaN(timestamp.getTime()) ? value : timestamp.toLocaleString();
}

export function parseTwelveHourTime(value: string): number | null {
  const match = /^\s*(\d{1,2})(?::(\d{2}))?\s*([ap])\.?m?\.?\s*$/i.exec(value);
  if (!match) return null;
  const hour = Number(match[1]);
  const minute = match[2] === undefined ? 0 : Number(match[2]);
  if (hour < 1 || hour > 12 || minute > 59) return null;
  const pm = match[3].toLowerCase() === "p";
  return ((hour % 12) + (pm ? 12 : 0)) * 60 + minute;
}
