/** Formats weekly minutes as hours (e.g. "12 hrs", "7.5 hrs", "1 hr"). */
export function formatHours(minutes: number): string {
  const hours = Math.round((minutes / 60) * 100) / 100;
  return `${hours} ${hours === 1 ? "hr" : "hrs"}`;
}
