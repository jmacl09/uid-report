import React, { useMemo, useState, useEffect } from "react";
import { saveToStorage } from "../api/saveToStorage";
import { Calendar, dateFnsLocalizer, Views } from "react-big-calendar";
import { format, parse, startOfWeek, getDay } from "date-fns";
import { enUS } from "date-fns/locale";
import "react-big-calendar/lib/css/react-big-calendar.css";

export type VsoStatus = "Draft" | "Approved" | "Rejected";

export interface VsoCalendarEvent {
  id: string;
  title: string;
  start: Date; // inclusive day
  end: Date;   // exclusive end (RBC convention)
  status: VsoStatus;
  summary?: string; // short preview of maintenanceReason
  dcCode?: string;
  spans?: string[];
  startTimeUtc?: string; // e.g., 12:00
  endTimeUtc?: string;   // e.g., 14:00
  subject?: string;
  notificationType?: string;
  location?: string;
  isp?: string;
  ispTicket?: string;
  impactExpected?: boolean;
  maintenanceReason?: string; // full text
}

interface Props {
  events: VsoCalendarEvent[];
  onEventClick?: (e: VsoCalendarEvent) => void;
  // Optional controlled date for the visible month; when provided, component becomes controlled
  date?: Date;
  onNavigate?: (newDate: Date) => void;
}

const locales = { "en-US": enUS } as const;
const localizer = dateFnsLocalizer({
  format,
  parse: (str: string, fmt: string, refDate: Date) => parse(str, fmt, refDate, { locale: enUS }),
  startOfWeek: () => startOfWeek(new Date(), { weekStartsOn: 0 }),
  getDay,
  locales,
});

const statusColors: Record<VsoStatus, { bg: string; border: string; text: string }> = {
  Draft: { bg: "rgba(255,193,7,0.18)", border: "#cc9900", text: "#ffdd57" },
  Approved: { bg: "rgba(0,200,83,0.18)", border: "#00cc55", text: "#9cf3c2" },
  Rejected: { bg: "rgba(214, 55, 55, 0.18)", border: "#cc3333", text: "#ffb3b3" },
};

const EventItem: React.FC<any> = ({ event }) => {
  const status: VsoStatus = (event.status as VsoStatus) || 'Draft';
  const spansPreview = useMemo(() => {
    const s = event.spans || [];
    if (!s.length) return "";
    if (s.length <= 3) return s.join(", ");
    return `${s.slice(0, 3).join(", ")}, (+${s.length - 3} more)`;
  }, [event.spans]);

  const tooltip = [
    `${event.title}`,
    event.dcCode ? `DC: ${event.dcCode}` : "",
    spansPreview ? `Spans: ${spansPreview}` : "",
  event.startTimeUtc && event.endTimeUtc ? `${event.startTimeUtc} - ${event.endTimeUtc}` : "",
    event.summary ? `Notes: ${event.summary}` : "",
  ]
    .filter(Boolean)
    .join("\n");

  const colors = statusColors[status];

  return (
    <div
      className="vso-cal-event"
      title={tooltip}
      style={{
        display: "flex",
        gap: 6,
        alignItems: "center",
        overflow: "hidden",
        textOverflow: "ellipsis",
        whiteSpace: "nowrap",
        color: colors.text,
      }}
    >
  <span className={`vso-cal-pill status-${status.toLowerCase()}`}>{status}</span>
      <span className="vso-cal-title" style={{ overflow: "hidden", textOverflow: "ellipsis" }}>
        {event.title}
      </span>
    </div>
  );
};

const VSOCalendar: React.FC<Props> = ({ events, onEventClick, date, onNavigate }) => {
  const eventPropGetter = (event: VsoCalendarEvent) => {
    const status: VsoStatus = (event.status as VsoStatus) || 'Draft';
    const colors = statusColors[status];
    return {
      style: {
        backgroundColor: colors.bg,
        border: `1px solid ${colors.border}`,
        color: colors.text,
        borderRadius: 8,
        padding: 4,
      },
    };
  };

  const handleSelectEvent = (e: VsoCalendarEvent) => {
    onEventClick?.(e);
  };

  // default to current month when not controlled
  const defaultDate = useMemo(() => new Date(), []);
  const [testStatus, setTestStatus] = useState<string>('');

  // When the calendar receives events, attempt to persist any new events to server-side storage.
  // Behavior:
  // - Determine a UID to save under by checking localStorage key 'vsoCalendarUid'. If not present,
  //   prompt the user once and persist the choice.
  // - Maintain a per-UID map of saved event ids in localStorage under 'vsoSaved_<uid>' to avoid
  //   re-saving the same event multiple times.
  useEffect(() => {
    let cancelled = false;

    const ensureUid = async (): Promise<string | null> => {
      try {
        const cached = localStorage.getItem('vsoCalendarUid');
        if (cached && cached.trim()) return cached;
      } catch (e) {}
      // Prompt the user for a UID to associate calendar saves with. Store for future runs.
      try {
        const u = window.prompt('Enter UID to save calendar entries under (e.g. 12345678901)');
        if (!u) return null;
        localStorage.setItem('vsoCalendarUid', u);
        return u;
      } catch (e) { return null; }
    };

    const loadSavedMap = (uid: string): Record<string, string> => {
      try {
        const raw = localStorage.getItem(`vsoSaved_${uid}`);
        if (!raw) return {};
        return JSON.parse(raw || '{}') as Record<string,string>;
      } catch (e) { return {}; }
    };

    const saveSavedMap = (uid: string, map: Record<string,string>) => {
      try { localStorage.setItem(`vsoSaved_${uid}`, JSON.stringify(map)); } catch (e) {}
    };

    const doSave = async () => {
      if (!events || !events.length) return;
      const uid = await ensureUid();
      if (!uid) return;

      const map = loadSavedMap(uid);

      // Find events that are not recorded in the saved map
      const toSave = events.filter((ev) => ev && ev.id && !map[ev.id]);
      if (!toSave.length) return;

      for (const ev of toSave) {
        if (cancelled) return;
        try {
          // Build a friendly description containing useful fields
          const description = [
            ev.summary || ev.maintenanceReason || '',
            ev.subject || '',
            ev.notificationType || '',
            ev.location || '',
            ev.spans && ev.spans.length ? `Spans: ${ev.spans.join(', ')}` : '',
            ev.start ? `Start: ${new Date(ev.start).toISOString()}` : '',
            ev.end ? `End: ${new Date(ev.end).toISOString()}` : '',
          ].filter(Boolean).join('\n');

          const owner = (() => { try { return localStorage.getItem('loggedInEmail') || 'VSO Calendar'; } catch { return 'VSO Calendar'; } })();

          const resText = await saveToStorage({
            category: 'Calendar',
            uid,
            title: ev.title || `VSO Event ${ev.id}`,
            description,
            owner,
            timestamp: ev.start || new Date(),
          });

          // backend returns JSON text with entity.rowKey; try to parse and record it
          try {
            const parsed = JSON.parse(resText || '{}');
            const entity = parsed?.entity || parsed?.Entity || null;
            const rowKey = entity?.rowKey || entity?.RowKey || parsed?.rowKey || null;
            // Record mapping of event id -> rowKey (or timestamp) so we don't re-save
            map[ev.id] = rowKey || (new Date()).toISOString();
            saveSavedMap(uid, map);
          } catch (e) {
            // If parse fails, still mark as saved to avoid repeated tries
            map[ev.id] = new Date().toISOString();
            saveSavedMap(uid, map);
          }
        } catch (e) {
          // Fail silently: we don't want to break the UI if server save fails.
          // Optionally log to console for debugging.
          // eslint-disable-next-line no-console
          console.warn('[VSOCalendar] Failed to save event to server', e);
        }
      }
    };

    doSave();

    return () => { cancelled = true; };
  }, [events]);

  return (
    <div className="calendar-panel table-container" style={{ width: "80%", maxWidth: 1200, margin: "22px auto" }}>
      <div className="section-title" style={{ marginBottom: 10 }}>Fiber Maintenance Calendar</div>
      <Calendar
        localizer={localizer}
        events={events}
        startAccessor="start"
        endAccessor="end"
        defaultView={Views.MONTH}
        views={[Views.MONTH]}
        popup
        toolbar
        defaultDate={defaultDate}
        date={date}
        eventPropGetter={eventPropGetter}
        components={{ event: EventItem }}
        onSelectEvent={handleSelectEvent}
  onNavigate={(d: Date) => onNavigate?.(d)}
        style={{ height: 650 }}
      />
      <div className="calendar-legend" style={{ display: "flex", gap: 10, marginTop: 8, color: "#9ab" }}>
        <span className="legend-item"><span className="legend-dot draft" /> Draft</span>
        <span className="legend-item"><span className="legend-dot approved" /> Approved</span>
        <span className="legend-item"><span className="legend-dot rejected" /> Rejected</span>
      </div>
          {/* Test save button for debugging server-side save behavior. Prompts for a UID to save under. */}
          <div style={{ marginTop: 12, display: 'flex', gap: 10, alignItems: 'center' }}>
            <button
              type="button"
              className="btn btn-secondary"
              onClick={async () => {
                // Prompt for UID so the tester can use a real UID from the app
                const uid = window.prompt('Enter UID to test save (e.g. CUST-12345)');
                if (!uid) return;
                try {
                  setTestStatus('Saving...');
                  const res = await saveToStorage({
                    category: 'Calendar',
                    uid,
                    title: `Calendar test save ${new Date().toISOString()}`,
                    description: 'Test entry saved from VSOCalendar UI',
                    owner: 'VSO Test Button',
                  });
                  setTestStatus(`Saved: ${String(res).slice(0, 200)}`);
                } catch (err: any) {
                  setTestStatus(`Error: ${err?.message || String(err)}`);
                }
              }}
            >
              Test save to server
            </button>
            <span style={{ color: '#678', fontSize: 12 }} aria-live="polite">{testStatus}</span>
          </div>
    </div>
  );
};

export default VSOCalendar;

// Example usage: save a calendar-related note to storage.
// This doesn't alter the UI; call from parent code when an event is created/approved, etc.
export async function saveCalendarEntryExample(uid: string, opts: { title: string; description?: string; owner?: string }) {
  try {
    const result = await saveToStorage({
      category: "Calendar",
      uid,
      title: opts.title,
      description: opts.description || "Scheduled maintenance window",
      owner: opts.owner || "Calendar Bot",
    });
    console.log(`[save] Calendar entry saved for UID ${uid}:`, result);
  } catch (e: any) {
    if (e?.status && e.status >= 500) {
      console.error("Server error while saving calendar entry:", e?.body || e?.message);
    } else {
      console.error("Failed to save calendar entry:", e?.body || e?.message || e);
    }
  }
}
