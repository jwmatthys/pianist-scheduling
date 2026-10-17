import { formatUtcTimestampLocally } from "../lib/time";

interface Props {
  stale: boolean;
  generatedAt: string;
}

export function StaleReportBanner({ stale, generatedAt }: Props) {
  return (
    <>
      {stale && (
        <p className="stale-report-banner" role="status">
          This report is out of date. The underlying data has changed since it was generated.
          Select Generate report to refresh it.
        </p>
      )}
      <p className="muted report-generated-at">Generated {formatUtcTimestampLocally(generatedAt)}</p>
    </>
  );
}
