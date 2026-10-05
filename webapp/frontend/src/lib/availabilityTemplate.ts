const TEMPLATE_HEADER = ["Pianist Name", "Day", "Start", "End", "Status"];
const TEMPLATE_ROWS = [
  ["Avery Example", "Monday", "08:00 AM", "11:30 AM", "Available"],
  ["Avery Example", "Monday", "02:00 PM", "05:00 PM", "Available"],
  ["Jordan Sample", "Tuesday", "09:00", "12:00", "Tentative"],
];

function escapeCsv(value: string): string {
  return /[",\r\n]/.test(value) ? `"${value.replaceAll('"', '""')}"` : value;
}

export function generateAvailabilityTemplateCsv(): string {
  return [TEMPLATE_HEADER, ...TEMPLATE_ROWS]
    .map((row) => row.map(escapeCsv).join(","))
    .join("\r\n");
}