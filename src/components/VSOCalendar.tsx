import React, { useMemo } from "react";
import { Calendar, dateFnsLocalizer, Views, EventProps } from "react-big-calendar";
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

const EventItem: React.FC<EventProps<VsoCalendarEvent>> = ({ event }) => {
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
    event.startTimeUtc && event.endTimeUtc ? `UTC: ${event.startTimeUtc} - ${event.endTimeUtc}` : "",
    event.summary ? `Notes: ${event.summary}` : "",
  ]
    .filter(Boolean)
    .join("\n");

  const colors = statusColors[event.status];

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
      <span className={`vso-cal-pill status-${event.status.toLowerCase()}`}>{event.status}</span>
      <span className="vso-cal-title" style={{ overflow: "hidden", textOverflow: "ellipsis" }}>
        {event.title}
      </span>
    </div>
  );
};

const VSOCalendar: React.FC<Props> = ({ events, onEventClick, date, onNavigate }) => {
  const eventPropGetter = (event: VsoCalendarEvent) => {
    const colors = statusColors[event.status];
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
        onNavigate={(d) => onNavigate?.(d)}
        style={{ height: 650 }}
      />
      <div className="calendar-legend" style={{ display: "flex", gap: 10, marginTop: 8, color: "#9ab" }}>
        <span className="legend-item"><span className="legend-dot draft" /> Draft</span>
        <span className="legend-item"><span className="legend-dot approved" /> Approved</span>
        <span className="legend-item"><span className="legend-dot rejected" /> Rejected</span>
      </div>
    </div>
  );
};

export default VSOCalendar;
