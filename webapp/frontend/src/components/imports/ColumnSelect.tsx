type Props = {
  label: string;
  value: string | null;
  columns: string[];
  /** Text of the empty option. */
  emptyLabel?: string;
  /** "plain" renders a bare label; "field" renders the availability-map-field layout. */
  variant?: "plain" | "field";
  /** Receives the chosen column, or an empty string when the empty option is chosen. */
  onChange: (value: string) => void;
};

/** Source-column dropdown shared by the spreadsheet import dialogs. */
export function ColumnSelect({ label, value, columns, emptyLabel = "Choose column", variant = "plain", onChange }: Props) {
  const select = (
    <select aria-label={label} value={value ?? ""} onChange={(event) => onChange(event.target.value)}>
      <option value="">{emptyLabel}</option>
      {columns.map((column) => (
        <option key={column} value={column}>
          {column}
        </option>
      ))}
    </select>
  );
  if (variant === "field") {
    return (
      <label className="availability-map-field">
        <span>{label}</span>
        {select}
      </label>
    );
  }
  return (
    <label>
      {label}
      {select}
    </label>
  );
}
