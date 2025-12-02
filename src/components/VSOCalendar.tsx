// VSOCalendar.tsx
import React, { useMemo, useEffect } from "react";
import { saveToStorage } from "../api/saveToStorage";
import { Calendar, dateFnsLocalizer, Views } from "react-big-calendar";
import { format, parse, startOfWeek, getDay } from "date-fns";
import { enUS } from "date-fns/locale";
import "react-big-calendar/lib/css/react-big-calendar.css";

export type VsoStatus = "Draft" | "Approved" | "Rejected";

export interface VsoCalendarEvent {
  id: string;
  title: string;
  start: Date; // inclusive
  end: Date;   // exclusive
  status: VsoStatus;
  summary?: string;
  dcCode?: string;
  spans?: string[];
  startTimeUtc?: string;
  endTimeUtc?: string;
  subject?: string;
  notificationType?: string;
  location?: string;
  isp?: string;
  ispTicket?: string;
  impactExpected?: boolean;
  maintenanceReason?: string;
}

interface Props {
  events: VsoCalendarEvent[];
  onEventClick?: (e: VsoCalendarEvent) => void;
  date?: Date;
  onNavigate?: (newDate: Date) => void;
}

const locales = { "en-US": enUS } as const;

const localizer = dateFnsLocalizer({
  format,
  parse: (str: string, fmt: string, refDate: Date) =>
    parse(str, fmt, refDate, { locale: enUS }),
  startOfWeek: () => startOfWeek(new Date(), { weekStartsOn: 0 }),
  getDay,
  locales,
});

const statusColors: Record<
  VsoStatus,
  { bg: string; border: string; text: string }
> = {
  Draft: { bg: "rgba(255,193,7,0.18)", border: "#cc9900", text: "#ffdd57" },
  Approved: { bg: "rgba(0,200,83,0.18)", border: "#00cc55", text: "#9cf3c2" },
  Rejected: { bg: "rgba(214,55,55,0.18)", border: "#cc3333", text: "#ffb3b3" },
};

const EventItem: React.FC<any> = ({ event }) => {
  const status = (event.status as VsoStatus) || "Draft";
  const colors = statusColors[status];

  const spansPreview = useMemo(() => {
    const s = event.spans || [];
    if (!s.length) return "";
    if (s.length <= 3) return s.join(", ");
    return `${s.slice(0, 3).join(", ")}, (+${s.length - 3} more)`;
  }, [event.spans]);

  const tooltip = [
    event.title,
    event.dcCode ? `DC: ${event.dcCode}` : "",
    spansPreview ? `Spans: ${spansPreview}` : "",
    event.startTimeUtc && event.endTimeUtc
      ? `${event.startTimeUtc} - ${event.endTimeUtc}`
      : "",
    event.summary ? `Notes: ${event.summary}` : "",
  ]
    .filter(Boolean)
    .join("\n");

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
      <span className={`vso-cal-pill status-${status.toLowerCase()}`}>
        {status}
      </span>
      <span style={{ overflow: "hidden", textOverflow: "ellipsis" }}>
        {event.title}
      </span>
    </div>
  );
};

const VSOCalendar: React.FC<Props> = ({
  events,
  onEventClick,
  date,
  onNavigate,
}) => {
  const eventPropGetter = (event: VsoCalendarEvent) => {
    const status = (event.status as VsoStatus) || "Draft";
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

  const defaultDate = useMemo(() => new Date(), []);

  // --- Auto-persist calendar events to Table Storage ---
  useEffect(() => {
    let cancelled = false;

    const loadSavedMap = (): Record<string, string> => {
      try {
        return JSON.parse(localStorage.getItem("vsoSaved") || "{}");
      } catch {
        return {};
      }
    };

    const saveSavedMap = (map: Record<string, string>) => {
      try {
        localStorage.setItem("vsoSaved", JSON.stringify(map));
      } catch {}
    };

    const persistEvents = async () => {
      if (!events || !events.length) return;

      const uid = "VsoCalendar"; // logical partition for all calendar entries
      const map = loadSavedMap();
      const newEvents = events.filter((ev) => ev.id && !map[ev.id]);

      if (!newEvents.length) return;

      for (const ev of newEvents) {
        if (cancelled) return;

        try {
          const description = [
            ev.summary || ev.maintenanceReason || "",
            ev.subject || "",
            ev.notificationType || "",
            ev.location || "",
            ev.spans?.length ? `Spans: ${ev.spans.join(", ")}` : "",
            ev.start ? `Start: ${ev.start.toISOString()}` : "",
            ev.end ? `End: ${ev.end.toISOString()}` : "",
          ]
            .filter(Boolean)
            .join("\n");

          const owner =
            localStorage.getItem("loggedInEmail") || "VSO Calendar";

          const res = await saveToStorage({
            category: "Calendar",
            uid,
            title: ev.title || `VSO Event ${ev.id}`,
            description,
            owner,
            timestamp: ev.start || new Date(),
          });

          try {
            const parsed = JSON.parse(res || "{}");
            const entity =
              parsed?.entity || parsed?.Entity || parsed || {};
            const rk =
              entity.RowKey ||
              entity.rowKey ||
              new Date().toISOString();
            map[ev.id] = rk;
            saveSavedMap(map);
          } catch {
            map[ev.id] = new Date().toISOString();
            saveSavedMap(map);
          }
        } catch (err) {
          console.warn("[VSOCalendar] Save failed:", err);
        }
      }
    };

    void persistEvents();
    return () => {
      cancelled = true;
    };
  }, [events]);

  return (
    <div
      className="calendar-panel table-container"
      style={{
        width: "80%",
        maxWidth: 1200,
        margin: "22px auto",
      }}
    >
      <div className="section-title" style={{ marginBottom: 10 }}>
        Fiber Maintenance Calendar
      </div>

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

      <div
        className="calendar-legend"
        style={{
          display: "flex",
          gap: 10,
          marginTop: 8,
          color: "#9ab",
        }}
      >
        <span className="legend-item">
          <span className="legend-dot draft" /> Draft
        </span>
        <span className="legend-item">
          <span className="legend-dot approved" /> Approved
        </span>
        <span className="legend-item">
          <span className="legend-dot rejected" /> Rejected
        </span>
      </div>
    </div>
  );
};

export default VSOCalendar;

// Example manual save helper
export async function saveCalendarEntryExample(
  uid: string,
  opts: { title: string; description?: string; owner?: string }
) {
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
      console.error("Server error while saving calendar entry:", e?.body);
    } else {
      console.error("Failed to save calendar entry:", e?.body || e);
    }
  }
}
