import { useCallback, useEffect, useState } from "react";
import { api } from "./api";
import type { ReportPayload } from "./reportExport";
import type { ReportTable } from "./reportTable";

/** Wire shape of a report persisted in the session database. */
export interface SavedReportInput {
  markdown: string;
  page_break_headings: string[];
  table: ReportTable;
  source_stamp: string;
}

export interface SavedReport extends SavedReportInput {
  generated_at: string;
}

/** Module-owned description of the report's current source data. */
export interface ReportSourceState {
  /** Opaque value that changes whenever the report's source data changes. */
  stamp: string;
  /** True when the module already knows its source result is stale. */
  sourceStale: boolean;
}

export interface GeneratedReport {
  payload: ReportPayload;
}

export interface SavedReportView {
  payload: ReportPayload;
  generatedAt: string;
  stale: boolean;
}

/**
 * Shared persistence for generated reports. Modules supply how to read their source
 * state and how to build the report; this hook stores the result in the session
 * (so it survives restarts) and reports whether it has become stale.
 */
export function useSavedReport(moduleId: string, fetchSourceState: () => Promise<ReportSourceState>) {
  const [report, setReport] = useState<SavedReportView | null>(null);
  const [loading, setLoading] = useState(true);

  useEffect(() => {
    let cancelled = false;
    async function load() {
      try {
        const saved = await api.getSavedReport(moduleId);
        if (!saved || cancelled) return;
        let stale = false;
        try {
          const current = await fetchSourceState();
          stale = current.sourceStale || current.stamp !== saved.source_stamp;
        } catch {
          // Source state is unavailable (for example no schedule yet); keep the saved report as is.
        }
        if (cancelled) return;
        setReport({
          payload: {
            markdown: saved.markdown,
            pageBreakBeforeHeadings: saved.page_break_headings,
            table: saved.table,
          },
          generatedAt: saved.generated_at,
          stale,
        });
      } catch {
        // A saved report is a convenience; failing to load it must not block generating a new one.
      } finally {
        if (!cancelled) setLoading(false);
      }
    }
    void load();
    return () => {
      cancelled = true;
    };
    // fetchSourceState is a stable module-level function.
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [moduleId]);

  /** Reads the source state first so a concurrent edit can only make the report look stale, never fresh. */
  const generate = useCallback(
    async (build: () => Promise<ReportPayload>) => {
      const source = await fetchSourceState();
      const payload = await build();
      const generatedAt = new Date().toISOString();
      setReport({ payload, generatedAt, stale: source.sourceStale });
      const saved = await api.saveReport(moduleId, {
        markdown: payload.markdown,
        page_break_headings: payload.pageBreakBeforeHeadings ?? [],
        table: payload.table,
        source_stamp: source.stamp,
      });
      setReport({ payload, generatedAt: saved.generated_at, stale: source.sourceStale });
    },
    // eslint-disable-next-line react-hooks/exhaustive-deps
    [moduleId],
  );

  return { report, loading, generate };
}
