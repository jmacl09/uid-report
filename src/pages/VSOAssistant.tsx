import React, { useState, useMemo, useEffect } from "react";
import {
  ComboBox,
  IComboBox,
  IComboBoxOption,
  Dropdown,
  IDropdownOption,
  PrimaryButton,
  IconButton,
  Checkbox,
  Spinner,
  SpinnerSize,
  MessageBar,
  MessageBarType,
  Dialog,
  DialogType,
  DialogFooter,
  DefaultButton,
  TextField,
  Text,
  DatePicker,
} from "@fluentui/react";
import "../Theme.css";
import datacenterOptions from "../data/datacenterOptions";
import { getRackElevationUrl } from "../data/MappedREs";

interface SpanData {
  SpanID: string;
  Diversity: string;
  IDF_A: string;
  SpliceRackA: string;
  WiringScope: string;
  Status: string;
  Color: string;
  OpticalLink?: string;
}

interface LogicAppResponse {
  Spans: SpanData[];
  RackElevationUrl?: string;
  DataCenter?: string;
}
interface MaintenanceWindow {
  startDate: Date | null;
  startTime: string;
  endDate: Date | null;
  endTime: string;
}

const VSOAssistant: React.FC = () => {
  const [facilityCodeA, setFacilityCodeA] = useState<string>("");
  const [diversity, setDiversity] = useState<string>();
  const [spliceRackA, setSpliceRackA] = useState<string>();
  const [loading, setLoading] = useState<boolean>(false);
  const [result, setResult] = useState<SpanData[]>([]);
  const [selectedSpans, setSelectedSpans] = useState<string[]>([]);
  const [error, setError] = useState<string | null>(null);
  const [showAll, setShowAll] = useState<boolean>(false);
  const [rackUrl, setRackUrl] = useState<string>();
  const [rackDC, setRackDC] = useState<string>();
  const [dcSearch, setDcSearch] = useState<string>("");
  const dcComboRef = React.useRef<IComboBox | null>(null);

  // Track whether a search was completed to show no-results banner
  const [searchDone, setSearchDone] = useState<boolean>(false);

  // Sorting state for results table
  const [sortBy, setSortBy] = useState<string>("");
  const [sortDir, setSortDir] = useState<"asc" | "desc">("asc");

  // === Stage 2: Compose Email state ===
  const [composeOpen, setComposeOpen] = useState<boolean>(false);
  const EMAIL_TO = "opticaldri@microsoft.com"; // fixed
  const EMAIL_LOGIC_APP_URL = "https://fibertools-dsavavdcfdgnh2cm.westeurope-01.azurewebsites.net:443/api/VSO/triggers/When_an_HTTP_request_is_received/invoke?api-version=2022-05-01&sp=%2Ftriggers%2FWhen_an_HTTP_request_is_received%2Frun&sv=1.0&sig=6ViXNM-TmW5F7Qd9_e4fz3IhRNqmNzKwovWvcmuNJto";
  const [subject, setSubject] = useState<string>("");
  const [notificationType, setNotificationType] = useState<string>("New Maintenance Scheduled");
  const [location, setLocation] = useState<string>("");
  const [lat, setLat] = useState<string>("");
  const [lng, setLng] = useState<string>("");
  const [isp, setIsp] = useState<string>("");
  const [ispTicket, setIspTicket] = useState<string>("");
  const [maintenanceReason, setMaintenanceReason] = useState<string>("");
  const [impactExpected, setImpactExpected] = useState<boolean>(true);
  const [startDate, setStartDate] = useState<Date | null>(null);
  const [startWarning, setStartWarning] = useState<string | null>(null);
  const [pendingEmergency, setPendingEmergency] = useState<boolean>(false);
  const [showEmergencyDialog, setShowEmergencyDialog] = useState<boolean>(false);
  const [endDate, setEndDate] = useState<Date | null>(null);
  const [startTime, setStartTime] = useState<string>("00:00");
  const [endTime, setEndTime] = useState<string>("00:00");
  const [additionalWindows, setAdditionalWindows] = useState<MaintenanceWindow[]>([]);
  const [userEmail, setUserEmail] = useState<string>("");
  const [cc, setCc] = useState<string>("");

  // simple validation state
  const [fieldErrors, setFieldErrors] = useState<Record<string, string>>({});
  const [showValidation, setShowValidation] = useState<boolean>(false);

  // Try to detect signed-in user's email from App Service/Static Web Apps auth
  useEffect(() => {
    // Read persisted login email (if any) so CC preview is available immediately
    try {
      const stored = localStorage.getItem("loggedInEmail");
      if (stored && stored.length > 3) setUserEmail(stored);
    } catch (e) {}

    const fetchUserEmail = async () => {
      try {
        const res = await fetch("/.auth/me", { credentials: "include" });
        if (!res.ok) return;
        const data = await res.json();
        // Debugging: print the full /.auth/me response so we can see available claims
        try {
          // eslint-disable-next-line no-console
          console.debug("/.auth/me response:", data);
        } catch (e) {}
        // Handle both App Service ([identities]) and Static Web Apps ({clientPrincipal}) shapes
        const identities = Array.isArray(data)
          ? data
          : data?.clientPrincipal
          ? [{ user_claims: data.clientPrincipal?.claims || [] }]
          : [];
        for (const id of identities) {
          const claims = id?.user_claims || [];
          const getClaim = (t: string) => claims.find((c: any) => c?.typ === t)?.val || "";
          const email =
            getClaim("http://schemas.xmlsoap.org/ws/2005/05/identity/claims/emailaddress") ||
            getClaim("emails") || // SWA sometimes
            getClaim("preferred_username") ||
            getClaim("upn") ||
            "";
          if (email) {
            setUserEmail(email);
            try {
              localStorage.setItem("loggedInEmail", email);
            } catch (e) {}
            return;
          }
        }
      } catch {}
    };
    fetchUserEmail();
  }, []);

  const addWindow = () =>
    setAdditionalWindows((w) => [
      ...w,
      { startDate: null, startTime: "00:00", endDate: null, endTime: "00:00" },
    ]);

  const removeWindow = (index: number) =>
    setAdditionalWindows((w) => w.filter((_, i) => i !== index));
  const [sendLoading, setSendLoading] = useState<boolean>(false);
  const [sendSuccess, setSendSuccess] = useState<string | null>(null);
  const [sendError, setSendError] = useState<string | null>(null);
  const [showSendSuccessDialog, setShowSendSuccessDialog] = useState<boolean>(false);

  // === Diversity options ===
  const diversityOptions: IDropdownOption[] = [
    { key: "West", text: "West, West 1, West 2" },
    { key: "East", text: "East, East 1, East 2" },
    { key: "North", text: "North" },
    { key: "South", text: "South" },
    { key: "Y", text: "Y" },
    { key: "Z", text: "Z" },
  ];

  // === Filter DCs based on input ===
  const filteredDcOptions: IComboBoxOption[] = useMemo(() => {
    const base = datacenterOptions.map((d) => ({ key: d.key, text: d.text }));
    const search = dcSearch.toLowerCase().trim();
    const items = !search
      ? base
      : base.filter(
          (opt) =>
            opt.key.toString().toLowerCase().includes(search) ||
            opt.text.toString().toLowerCase().includes(search)
        );
    // Remove explicit (None); clicking selected option will now deselect
    return items;
  }, [dcSearch]);

  // === Submit ===
  const handleSubmit = async () => {
    if (!facilityCodeA) {
      alert("Please select a valid Data Center first.");
      return;
    }

    setLoading(true);
    setError(null);
    setResult([]);
    setSearchDone(false);

    try {
      const logicAppUrl =
        "https://fibertools-dsavavdcfdgnh2cm.westeurope-01.azurewebsites.net:443/api/VSO/triggers/When_an_HTTP_request_is_received/invoke?api-version=2022-05-01&sp=%2Ftriggers%2FWhen_an_HTTP_request_is_received%2Frun&sv=1.0&sig=6ViXNM-TmW5F7Qd9_e4fz3IhRNqmNzKwovWvcmuNJto";

      const payload = {
        FacilityCodeA: facilityCodeA,
        Diversity:
          (() => {
            const raw = (diversity || "").toString();
            if (!raw) return "N";
            // Take the part before any comma, trim spaces, and strip stray punctuation like trailing commas
            const parsed = raw.split(",")[0].trim().replace(/,$/, "");
            // Fallback to alphanumeric-only label if needed
            return parsed.replace(/[^A-Za-z0-9 ]/g, "").trim() || "N";
          })(),
        SpliceRackA: spliceRackA || "N",
        Stage: "1",
      };

      const response = await fetch(logicAppUrl, {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify(payload),
      });

      if (!response.ok) throw new Error(`HTTP error! Status: ${response.status}`);
      const data: LogicAppResponse = await response.json();

      if (data?.Spans) setResult(data.Spans);
      if (data?.RackElevationUrl) setRackUrl(data.RackElevationUrl);
      if (data?.DataCenter) setRackDC(data.DataCenter);
    } catch (err: any) {
      setError(err.message || "Unknown error occurred.");
    } finally {
      setLoading(false);
      setSearchDone(true);
    }
  };

  const toggleSelectSpan = (spanId: string) => {
    setSelectedSpans((prev) =>
      prev.includes(spanId) ? prev.filter((id) => id !== spanId) : [...prev, spanId]
    );
  };

  const filteredResults = showAll
    ? result
    : result.filter((r) => r.Status.toLowerCase() === "inproduction");

  // Accessor for sorting
  const getSortValue = (row: SpanData, key: string): string | number => {
    switch (key) {
      case "diversity":
        return row.Diversity || "";
      case "span":
        return row.SpanID || "";
      case "idf":
        return row.IDF_A || "";
      case "splice":
        return row.SpliceRackA || "";
      case "scope":
        return row.WiringScope || "";
      case "status":
        return row.Status || "";
      default:
        return "";
    }
  };

  const sortedResults = useMemo(() => {
    const rows = [...filteredResults];
    if (!sortBy) return rows;
    rows.sort((a, b) => {
      const av = getSortValue(a, sortBy);
      const bv = getSortValue(b, sortBy);
      // numeric-aware compare where possible
      const an = typeof av === "string" ? Number(av) : (av as number);
      const bn = typeof bv === "string" ? Number(bv) : (bv as number);
      const aIsNum = !isNaN(an as number) && an !== (undefined as any) && an !== null && av !== "";
      const bIsNum = !isNaN(bn as number) && bn !== (undefined as any) && bn !== null && bv !== "";
      if (aIsNum && bIsNum) {
        return (an as number) - (bn as number);
      }
      const as = (av ?? "").toString().toLowerCase();
      const bs = (bv ?? "").toString().toLowerCase();
      return as.localeCompare(bs);
    });
    if (sortDir === "desc") rows.reverse();
    return rows;
  }, [filteredResults, sortBy, sortDir]);

  const handleSort = (key: string) => {
    if (sortBy === key) {
      setSortDir((d) => (d === "asc" ? "desc" : "asc"));
    } else {
      setSortBy(key);
      setSortDir("asc");
    }
  };

  const comboBoxStyles = {
    root: { width: "100%" },
    input: {
      color: "#fff",
      backgroundColor: "#141414",
      height: 42,
      border: "1px solid #333",
      borderRadius: 8,
      paddingLeft: 10,
    },
    callout: { background: "#181818", maxHeight: 240, overflowY: "auto" },
    optionsContainer: { background: "#181818" },
    caretDownWrapper: { color: "#fff" },
  } as const;

  const dropdownStyles = {
    root: { width: "100%" },
    dropdown: { backgroundColor: "#141414", color: "#fff", borderRadius: 8 },
    title: {
      background: "#141414",
      color: "#fff",
      border: "1px solid #333",
      borderRadius: 8,
      height: 42,
      display: "flex",
      alignItems: "center",
      paddingLeft: 10,
    },
    caretDownWrapper: { color: "#fff" },
    dropdownItemsWrapper: { background: "#181818" },
    dropdownItem: { background: "transparent", color: "#fff" },
    dropdownItemSelected: { background: "#003b6f", color: "#fff" },
    callout: { background: "#181818" },
  };

  // Make Diversity placeholder text exactly match TextField placeholder (e.g., AM111)
  const diversityDropdownStyles = {
    ...dropdownStyles,
    title: {
      ...dropdownStyles.title,
    },
    titleIsPlaceHolder: {
      ...dropdownStyles.title,
      color: "#a6b7c6",
      opacity: 0.8,
      fontStyle: "normal",
      fontSize: 14,
      fontWeight: 400,
    },
  } as const;

  const textFieldStyles = {
    fieldGroup: { backgroundColor: "#141414", border: "1px solid #333", borderRadius: 8, height: 42 },
    field: { color: "#fff" },
  };

  // Time dropdown styles (reuse dropdownStyles but allow container width to control it)
  const timeDropdownStyles = {
    ...dropdownStyles,
    root: { width: '100%' },
  } as const;

  // Dark DatePicker styles to avoid white-on-white
  const datePickerStyles: any = {
    root: { width: 220 },
    textField: {
      fieldGroup: { backgroundColor: "#141414", border: "1px solid #333", borderRadius: 8, height: 42 },
      field: { color: "#fff", selectors: { '::placeholder': { color: '#a6b7c6', opacity: 0.8 } } },
    },
    callout: { background: "#181818" },
    // dayPicker (calendar) styles to avoid white popover
    dayPicker: { root: { background: '#181818', color: '#fff' }, monthPickerVisible: {}, showWeekNumbers: {} },
  };

  // Build 30-min interval time options in 24h format
  const timeOptions: IDropdownOption[] = useMemo(() => {
    const opts: IDropdownOption[] = [];
    for (let h = 0; h < 24; h++) {
      for (let m = 0; m < 60; m += 30) {
        const hh = h.toString().padStart(2, "0");
        const mm = m.toString().padStart(2, "0");
        const text = `${hh}:${mm}`;
        opts.push({ key: text, text });
      }
    }
    return opts;
  }, []);

  const spansComma = useMemo(() => selectedSpans.join(","), [selectedSpans]);

  const formatUtcString = (date: Date | null, time: string) => {
    if (!date) return "";
    // Treat selected date + time as UTC and format as MM/DD/YYYY HH:MM UTC
    const [hh, mm] = time.split(":").map((s) => parseInt(s, 10));
    const y = date.getUTCFullYear();
    const m = (date.getUTCMonth() + 1).toString().padStart(2, "0");
    const d = date.getUTCDate().toString().padStart(2, "0");
    const H = (isNaN(hh) ? 0 : hh).toString().padStart(2, "0");
    const M = (isNaN(mm) ? 0 : mm).toString().padStart(2, "0");
    return `${m}/${d}/${y} ${H}:${M} UTC`;
  };

  const isWithinDays = (date: Date | null, days: number) => {
    if (!date) return false;
    const now = new Date();
    // normalize to start of day for comparison convenience
    const diff = date.getTime() - now.getTime();
    return diff < days * 24 * 60 * 60 * 1000;
  };

  const startUtc = useMemo(() => formatUtcString(startDate, startTime), [startDate, startTime]);
  const endUtc = useMemo(() => formatUtcString(endDate, endTime), [endDate, endTime]);

  // Prefill subject when entering compose
  useEffect(() => {
    if (!composeOpen) return;
    if (!subject || subject.trim().length === 0) {
      const region = rackDC || facilityCodeA || "Region";
      setSubject(`[${region}] Maintenance scheduled in <${region}> Contractor:  Lead Engineer:`);
    }
  }, [composeOpen, subject, rackDC, facilityCodeA]);

  const latLongCombined = useMemo(() => (lat && lng ? `${lat},${lng}` : ""), [lat, lng]);

  // Prefill CC from detected user email when compose opens or userEmail changes
  useEffect(() => {
    if (composeOpen && !cc && userEmail) {
      setCc(userEmail);
    }
  }, [composeOpen, userEmail, cc]);

  // Helpers to determine emergency tag state across windows and update subject
  const anyWindowWithin7Days = (primary: Date | null, windows: MaintenanceWindow[]) => {
    if (isWithinDays(primary, 7)) return true;
    for (const w of windows) {
      if (isWithinDays(w.startDate, 7)) return true;
    }
    return false;
  };

  const addEmergencyTag = () => {
    if (!/\[EMERGENCY\]/i.test(subject || "")) {
      setSubject(((subject || "") + " [EMERGENCY]").trim());
    }
  };

  const removeEmergencyTag = () => {
    if (/\[EMERGENCY\]/i.test(subject || "")) {
      setSubject(((subject || "") as string).replace(/\s*\[EMERGENCY\]\s*/i, "").trim());
    }
  };

  const emailBody = useMemo(() => {
    const impactStr = impactExpected ? "Yes/True" : "No/False";

    // Build comma-separated start/end lists: primary first, then any additional windows
    const startList: string[] = [];
    const endList: string[] = [];
    if (startUtc) startList.push(startUtc);
    if (endUtc) endList.push(endUtc);
    additionalWindows.forEach((w) => {
      const s = formatUtcString(w.startDate, w.startTime);
      const e = formatUtcString(w.endDate, w.endTime);
      if (s) startList.push(s);
      if (e) endList.push(e);
    });

    const parts: string[] = [
      `To: ${EMAIL_TO}`,
  `From: Fibervsoassistant@microsoft.com`,
      `CC: ${cc || ""}`,
      `Subject: ${subject}`,
      ``,
      `----------------------------------------`,
      `CircuitIds: ${spansComma}`,
      `StartDatetime: ${startList.join(', ')}`,
      `EndDatetime: ${endList.join(', ')}`,
    ];

    parts.push(
      `NotificationType: ${notificationType}`,
      `MaintenanceReason: ${maintenanceReason}`,
      `Location: ${location}`,
      `ISP: ${isp}`,
      `ISPTicket: ${ispTicket}`,
      `ImpactExpected: ${impactStr}`,
    );

    return parts.map((p) => p || "").join("\n");
  }, [EMAIL_TO, subject, spansComma, startUtc, endUtc, notificationType, location, maintenanceReason, isp, ispTicket, impactExpected, additionalWindows, cc]);

  const canSend = useMemo(() => {
    return (
      selectedSpans.length > 0 &&
      !!subject &&
      !!startDate && !!startTime &&
      !!endDate && !!endTime &&
      !!location &&
      !!isp &&
      !!ispTicket &&
      (impactExpected === true || impactExpected === false) &&
      !!maintenanceReason &&
      !!(cc && cc.trim())
    );
  }, [selectedSpans.length, subject, startDate, startTime, endDate, endTime, location, isp, ispTicket, impactExpected, maintenanceReason, cc]);

  const validateCompose = (): boolean => {
    const errors: Record<string, string> = {};
    if (!startDate) errors.startDate = "Required";
    if (!startTime) errors.startTime = "Required";
    if (!endDate) errors.endDate = "Required";
    if (!endTime) errors.endTime = "Required";
    if (!location?.trim()) errors.location = "Required";
    if (!isp?.trim()) errors.isp = "Required";
    if (!ispTicket?.trim()) errors.ispTicket = "Required";
    if (!(impactExpected === true || impactExpected === false)) errors.impactExpected = "Required";
    if (!maintenanceReason?.trim()) errors.maintenanceReason = "Required";
    if (!cc?.trim()) errors.cc = "Required";
    setFieldErrors(errors);
    return Object.keys(errors).length === 0;
  };

  const friendlyFieldNames: Record<string, string> = {
    startDate: "Start Date",
    startTime: "Start Time",
    endDate: "End Date",
    endTime: "End Time",
    location: "Location",
    isp: "ISP",
    ispTicket: "ISP Ticket",
    impactExpected: "Impact Expected",
    maintenanceReason: "Maintenance Reason",
    cc: "CC",
    subject: "Subject",
  };

  const handleSend = async () => {
    // Enable showing validation UI once the user attempts to send
    setShowValidation(true);
    if (!validateCompose()) {
      // scroll to top so the validation summary is visible
      window.scrollTo({ top: 0, behavior: 'smooth' });
      return;
    }
    // Clear validation UI if passed
    setShowValidation(false);
    setSendError(null);
    setSendSuccess(null);
    setSendLoading(true);
    try {
      // Build comma-separated StartDatetime and EndDatetime (primary + any additional)
      const startList: string[] = [];
      const endList: string[] = [];
      if (startUtc) startList.push(startUtc);
      if (endUtc) endList.push(endUtc);
      additionalWindows.forEach((w) => {
        const s = formatUtcString(w.startDate, w.startTime);
        const e = formatUtcString(w.endDate, w.endTime);
        if (s) startList.push(s);
        if (e) endList.push(e);
      });

      const payload = {
        FacilityCodeA: facilityCodeA || "",
        Diversity:
          (() => {
            const raw = (diversity || "").toString();
            if (!raw) return "";
            const parsed = raw.split(",")[0].trim().replace(/,$/, "");
            return parsed.replace(/[^A-Za-z0-9 ]/g, "").trim() || "";
          })(),
        SpliceRackA: spliceRackA || "",
        Stage: "2",
        CC: cc || "",
        Subject: subject || "",
        CircuitIds: spansComma || "",
        StartDatetime: startList.join(', '),
        EndDatetime: endList.join(', '),
        LatLong: latLongCombined || "",
        NotificationType: notificationType || "",
        MaintenanceReason: maintenanceReason || "",
        Location: location || "",
        ISP: isp || "",
        ISPTicket: ispTicket || "",
        ImpactExpected: impactExpected ? "True" : "False",
      };

      const resp = await fetch(EMAIL_LOGIC_APP_URL, {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify(payload),
      });

      if (!resp.ok) {
        const text = await resp.text().catch(() => ``);
        throw new Error(`HTTP ${resp.status} ${text}`);
      }

  setSendSuccess("Request submitted to Logic App successfully.");
  setShowSendSuccessDialog(true);
    } catch (e: any) {
      setSendError(e?.message || "Failed to send email.");
    } finally {
      setSendLoading(false);
    }
  };

  // Map status text to visual pill styles defined in Theme.css (.status-label.*)
  const getStatusClass = (status?: string) => {
    const t = (status || "").toLowerCase();
    if (t.includes("inproduction") || t === "in production" || t === "production") return "good";
    if (
      t.includes("decom") ||
      t.includes("retired") ||
      t.includes("outofservice") ||
      t.includes("out of service") ||
      t.includes("warning")
    )
      return "warning";
    return "accent";
  };

  // Map diversity text to color-coded pill
  const getDiversityClass = (div?: string) => {
    const t = (div || "").toLowerCase().trim();
    if (t.includes("east 1")) return "accent"; // blue
    if (t.includes("east 2")) return "good"; // green
    if (t === "south") return "accent"; // blue
    if (t === "y") return "accent"; // blue
    if (t.includes("west 1")) return "danger"; // red
    if (t.includes("west 2")) return "warning"; // yellow
    if (t === "north") return "danger"; // red
    if (t === "z") return "danger"; // red
    if (t.startsWith("east")) return "accent";
    if (t.startsWith("west")) return "danger";
    return "accent";
  };

  return (
    <div className="main-content fade-in">
      <div className="vso-form-container glow" style={{ width: "80%", maxWidth: 1000 }}>
        <div className="banner-title">
          <span className="title-text">Fiber VSO Assistant</span>
          <span className="title-sub">Simplifying Span Lookup and VSO Creation.</span>
        </div>

        {!composeOpen && (
          <>
        {/* === Data Center ComboBox === */}
          <Text styles={{ root: { color: "#ccc", fontSize: 15, fontWeight: 500 } }}>
            Data Center <span style={{ color: "#ff4d4d" }}>*</span>
          </Text>

        <ComboBox
          placeholder="Type or select a Data Center"
          options={filteredDcOptions}
          selectedKey={facilityCodeA || undefined}
          allowFreeform={true}
          autoComplete="on"
          useComboBoxAsMenuWidth
          calloutProps={{ className: 'combo-dark-callout' }}
          componentRef={dcComboRef}
          onClick={() => dcComboRef.current?.focus(true)}
          onFocus={() => dcComboRef.current?.focus(true)}
          styles={comboBoxStyles}
          onChange={(_, option, index, value) => {
            // value is typed text, only commit if matches a valid key
            const typed = (value || "").toString().toLowerCase();
            const found = datacenterOptions.find((d) => {
              const keyStr = d.key?.toString().toLowerCase();
              const textStr = d.text?.toString().toLowerCase();
              return textStr === typed || keyStr === typed;
            });

            if (option) {
              const selectedKey = option.key?.toString() ?? "";
              // Toggle off if the selected option is clicked again
              if (selectedKey === facilityCodeA) {
                setFacilityCodeA("");
              } else {
                setFacilityCodeA(selectedKey);
              }
            } else if (found) {
              setFacilityCodeA(found.key.toString());
            } else {
              // reset if invalid typed text
              setFacilityCodeA("");
            }
          }}
          onPendingValueChanged={(option, index, value) => {
            setDcSearch(value || "");
          }}
          onMenuDismiss={() => setDcSearch("")}
        />

        {/* === Diversity Dropdown === */}
        <Text styles={{ root: { color: "#ccc", fontSize: 15, fontWeight: 500, marginTop: 10 } }}>
          Diversity Path (Optional)
        </Text>
        <Dropdown
          placeholder="Select a Diversity Path"
          options={diversityOptions}
          calloutProps={{ className: 'combo-dark-callout' }}
          styles={diversityDropdownStyles}
          selectedKey={diversity === undefined || diversity === "" ? undefined : diversity}
          onChange={(_, option) => {
            if (!option) return;
            const nextKey = option.key?.toString() ?? "";
            // Toggle off if the same diversity option is clicked
            if ((diversity || "") === nextKey) {
              setDiversity(undefined);
            } else {
              setDiversity(nextKey);
            }
          }}
        />

        <TextField
          label="Splice Rack A (Optional)"
          placeholder="e.g. AM111"
          onChange={(_, value) => setSpliceRackA(value)}
          styles={textFieldStyles}
        />

        <div className="form-buttons" style={{ marginTop: 16 }}>
          <button className="submit-btn" onClick={handleSubmit}>
            Submit
          </button>
          <button className="help-btn" onClick={() => alert("Coming soon")}>
            Help
          </button>
        </div>

        {loading && <Spinner label="Loading results..." size={SpinnerSize.medium} styles={{ root: { marginTop: 15 } }} />}

        {error && (
          <MessageBar messageBarType={MessageBarType.error} isMultiline={false}>
            {error}
          </MessageBar>
        )}

        {!loading && !error && searchDone && result.length === 0 && (
          <div className="notice-banner warning">
            <span className="banner-icon">!</span>
            <div className="banner-text">
              There were no results for the selections you made. Try adjusting your search.
            </div>
          </div>
        )}

        {result.length > 0 && (
          <div className="table-container" style={{ marginTop: 14 }}>
            <div style={{ display: "flex", justifyContent: "flex-end", alignItems: "center", marginBottom: 8 }}>
              {(() => {
                const resolvedDC = rackDC || facilityCodeA;
                const href = rackUrl || getRackElevationUrl(resolvedDC);
                return href ? (
                  <div style={{ textAlign: "center", flex: 1 }}>
                    <a
                      href={href}
                      target="_blank"
                      rel="noopener noreferrer"
                      className="rack-btn slim"
                      style={{ display: "inline-block", minWidth: 320 }}
                    >
                      {`View Rack Elevation - ${resolvedDC}`}
                    </a>
                  </div>
                ) : null;
              })()}
              <div style={{ textAlign: "right" }}>
                <button className="sleek-btn optical" onClick={() => setShowAll(!showAll)}>
                  {showAll ? "Show Only Production" : "Show All Spans"}
                </button>
              </div>
            </div>

            <table className="data-table compact">
              <thead>
                <tr>
                  <th></th>
                  <th onClick={() => handleSort("diversity")}>
                    Diversity {sortBy === "diversity" && (sortDir === "asc" ? "▲" : "▼")}
                  </th>
                  <th onClick={() => handleSort("span")}>Span ID {sortBy === "span" && (sortDir === "asc" ? "▲" : "▼")}</th>
                  <th onClick={() => handleSort("idf")}>IDF {sortBy === "idf" && (sortDir === "asc" ? "▲" : "▼")}</th>
                  <th onClick={() => handleSort("splice")}>
                    Splice Rack {sortBy === "splice" && (sortDir === "asc" ? "▲" : "▼")}
                  </th>
                  <th onClick={() => handleSort("scope")}>Scope {sortBy === "scope" && (sortDir === "asc" ? "▲" : "▼")}</th>
                  <th onClick={() => handleSort("status")}>Status {sortBy === "status" && (sortDir === "asc" ? "▲" : "▼")}</th>
                </tr>
              </thead>
              <tbody>
                {sortedResults.map((row, i) => (
                  <tr
                    key={i}
                    className={selectedSpans.includes(row.SpanID) ? "highlight-row" : ""}
                    onClick={() => toggleSelectSpan(row.SpanID)}
                    style={{ cursor: "pointer" }}
                  >
                    <td>
                      <Checkbox
                        checked={selectedSpans.includes(row.SpanID)}
                        onChange={() => toggleSelectSpan(row.SpanID)}
                      />
                    </td>
                    <td>
                      <span className={`status-label ${getDiversityClass(row.Diversity)}`}>
                        {row.Diversity}
                      </span>
                    </td>
                    <td>
                      <a
                        href={row.OpticalLink}
                        target="_blank"
                        rel="noopener noreferrer"
                        className="uid-click"
                        onClick={(e) => e.stopPropagation()}
                      >
                        {row.SpanID}
                      </a>
                    </td>
                    <td>{row.IDF_A}</td>
                    <td>{row.SpliceRackA}</td>
                    <td>{row.WiringScope}</td>
                    <td>
                      <span className={`status-label ${getStatusClass(row.Status)}`}>{row.Status}</span>
                    </td>
                  </tr>
                ))}
              </tbody>
            </table>

            <div style={{ textAlign: "center", marginTop: 12 }}>
              <PrimaryButton
                text={`Continue (${selectedSpans.length} selected)`}
                disabled={selectedSpans.length === 0}
                onClick={() => setComposeOpen(true)}
              />
            </div>
          </div>
        )}
          </>
        )}

        {/* === Compose Section === */}
        {composeOpen && (
          <div className="table-container compose-container" style={{ marginTop: 16 }}>
            <div style={{ display: 'flex', alignItems: 'center', gap: 8, marginBottom: 6 }}>
              <IconButton iconProps={{ iconName: 'Back' }} title="Back" ariaLabel="Back" onClick={() => setComposeOpen(false)} />
              <div className="section-title" style={{ margin: 0 }}>Compose Maintenance Email</div>
            </div>

            {/* policy-info removed per user request */}

            {/* success is shown as a dialog to match Emergency UX */}
            {sendError && (
              <MessageBar messageBarType={MessageBarType.error} isMultiline={false}>
                {sendError}
              </MessageBar>
            )}
            {showValidation && Object.keys(fieldErrors).length > 0 && (
              <MessageBar messageBarType={MessageBarType.error} isMultiline={false}>
                Please complete the required fields: {Object.keys(fieldErrors).map(k => friendlyFieldNames[k] || k).join(', ')}
              </MessageBar>
            )}

            <Dialog
              hidden={!showEmergencyDialog}
              onDismiss={() => { setShowEmergencyDialog(false); setStartWarning(null); setPendingEmergency(false); }}
              dialogContentProps={{
                type: DialogType.normal,
                title: 'Confirm Emergency Maintenance',
                subText: startWarning || '',
              }}
              modalProps={{ isBlocking: true }}
            >
              <DialogFooter>
                <PrimaryButton
                  text="Confirm & Mark Emergency"
                  onClick={() => {
                    if (pendingEmergency) {
                      addEmergencyTag();
                      setPendingEmergency(false);
                    }
                    setShowEmergencyDialog(false);
                    setStartWarning(null);
                  }}
                />
                <DefaultButton
                  text="Cancel"
                  onClick={() => { setShowEmergencyDialog(false); setStartWarning(null); setPendingEmergency(false); }}
                />
              </DialogFooter>
            </Dialog>

            <Dialog
              hidden={!showSendSuccessDialog}
              onDismiss={() => setShowSendSuccessDialog(false)}
              dialogContentProps={{
                type: DialogType.normal,
                title: 'Email Sent',
                subText: sendSuccess || 'Email has been sent successfully.'
              }}
              modalProps={{ isBlocking: false }}
            >
              <DialogFooter>
                <PrimaryButton text="Start Over" onClick={() => { setComposeOpen(false); setShowSendSuccessDialog(false); }} />
                <DefaultButton text="Close" onClick={() => setShowSendSuccessDialog(false)} />
              </DialogFooter>
            </Dialog>

            <div style={{ display: 'flex', gap: 12, alignItems: 'flex-start' }}>
              <div style={{ flex: 1 }}>
                <div style={{ display: 'flex', gap: 12, alignItems: 'center' }}>
                  <div style={{ flex: 2 }}>
                    <TextField
                      label="Subject"
                      placeholder="[region] Maintenance scheduled in <region> Contractor:  Lead Engineer:"
                      value={subject}
                      onChange={(_, v) => setSubject(v || "")}
                      styles={textFieldStyles}
                      required
                      errorMessage={showValidation ? fieldErrors.subject : undefined}
                    />
                  </div>
                  <div style={{ flex: 1 }}>
                    <TextField
                      label="CC"
                      placeholder="name@contoso.com"
                      value={cc}
                      onChange={(_, v) => setCc(v || "")}
                      styles={textFieldStyles}
                      required
                      errorMessage={showValidation ? fieldErrors.cc : undefined}
                    />
                  </div>
                  <div style={{ width: 320, flexShrink: 0 }}>
                    <Dropdown
                      label="Notification Type"
                      options={[
                        { key: "New Maintenance Scheduled", text: "New Maintenance Scheduled" },
                        { key: "Rescheduled", text: "Rescheduled" },
                        { key: "Maintenance Cancelled", text: "Maintenance Cancelled" },
                        { key: "Maintenance Reminder", text: "Maintenance Reminder" },
                      ]}
                      selectedKey={notificationType}
                      onChange={(_, opt) => opt && setNotificationType(opt.key.toString())}
                      styles={dropdownStyles}
                      required
                    />
                  </div>
                </div>

                <div className="compose-datetime-row">
                    <div className="dt-field">
                      <Text styles={{ root: { color: "#ccc", fontSize: 12, fontWeight: 600 } }}>Start Date (UTC)</Text>
                      <DatePicker
                        placeholder="Select start date"
                        value={startDate || undefined}
                        onSelectDate={(d) => {
                          const selected = d || null;
                          setStartDate(selected);
                          if (isWithinDays(selected, 7)) {
                            setStartWarning(
                              "You have selected a date less than 7 days in advance. If this is required please press confirm to continue and the email will be updated with the Emergency Tag."
                            );
                            setPendingEmergency(true);
                            setShowEmergencyDialog(true);
                          } else {
                            setStartWarning(null);
                            setPendingEmergency(false);
                            // If there are no windows (primary + additional) within 7 days, remove the emergency tag
                            if (!anyWindowWithin7Days(selected, additionalWindows)) removeEmergencyTag();
                          }
                        }}
                        styles={datePickerStyles}
                        isRequired
                        aria-label="Start Date (UTC)"
                      />
                      {showValidation && fieldErrors.startDate ? (
                        <Text styles={{ root: { color: '#a80000', fontSize: 12 } }}>Required</Text>
                      ) : null}
                    </div>

                    <div className="dt-time">
                      <Text styles={{ root: { color: "#ccc", fontSize: 12, fontWeight: 600 } }}>Start Time (UTC)</Text>
                      <Dropdown
                        options={timeOptions}
                        selectedKey={startTime}
                        onChange={(_, opt) => opt && setStartTime(opt.key.toString())}
                        styles={timeDropdownStyles}
                        required
                        errorMessage={showValidation ? fieldErrors.startTime : undefined}
                      />
                    </div>

                    <div className="dt-field">
                      <Text styles={{ root: { color: "#ccc", fontSize: 12, fontWeight: 600 } }}>End Date (UTC)</Text>
                      <DatePicker
                        placeholder="Select end date"
                        value={endDate || undefined}
                        onSelectDate={(d) => setEndDate(d || null)}
                        styles={datePickerStyles}
                        isRequired
                        aria-label="End Date (UTC)"
                      />
                      {showValidation && fieldErrors.endDate ? (
                        <Text styles={{ root: { color: '#a80000', fontSize: 12 } }}>Required</Text>
                      ) : null}
                    </div>

                    <div className="dt-time">
                      <Text styles={{ root: { color: "#ccc", fontSize: 12, fontWeight: 600 } }}>End Time (UTC)</Text>
                      <Dropdown
                        options={timeOptions}
                        selectedKey={endTime}
                        onChange={(_, opt) => opt && setEndTime(opt.key.toString())}
                        styles={timeDropdownStyles}
                        required
                        errorMessage={showValidation ? fieldErrors.endTime : undefined}
                      />
                    </div>

                    <div className="dt-actions">
                      <button type="button" className="tiny-icon-btn add-window-btn" aria-label="Add Window" onClick={addWindow} title="Add Window">
                        <span className="glyph">+</span>
                      </button>
                    </div>
                </div>

              {/* Additional maintenance windows */}
              <div style={{ marginTop: 8 }}>
                {additionalWindows.map((w, i) => (
                  <div key={i} className="compose-datetime-row additional-window">
                    <div className="dt-field">
                      <Text styles={{ root: { color: "#ccc", fontSize: 12, fontWeight: 600 } }}>Start Date (UTC)</Text>
                      <DatePicker
                        placeholder="Select start date"
                        value={w.startDate || undefined}
                        onSelectDate={(d) => {
                          const selected = d || null;
                          setAdditionalWindows((arr) => {
                            const next = [...arr];
                            next[i] = { ...next[i], startDate: selected };
                            return next;
                          });

                          if (isWithinDays(selected, 7)) {
                            setStartWarning(
                              "You have selected a date less than 7 days in advance. If this is required please press confirm to continue and the email will be updated with the Emergency Tag."
                            );
                            setPendingEmergency(true);
                            setShowEmergencyDialog(true);
                          } else {
                            // Re-evaluate across primary + updated additional windows
                            const updated = additionalWindows.map((w, idx) => (idx === i ? { ...w, startDate: selected } : w));
                            if (!anyWindowWithin7Days(startDate, updated)) {
                              setStartWarning(null);
                              setPendingEmergency(false);
                              removeEmergencyTag();
                            }
                          }
                        }}
                        styles={datePickerStyles}
                      />
                    </div>
                    <div className="dt-time">
                      <Text styles={{ root: { color: "#ccc", fontSize: 12, fontWeight: 600 } }}>Start Time (UTC)</Text>
                      <Dropdown
                        options={timeOptions}
                        selectedKey={w.startTime}
                        onChange={(_, opt) => opt && setAdditionalWindows((arr) => {
                          const next = [...arr];
                          next[i] = { ...next[i], startTime: opt.key.toString() };
                          return next;
                        })}
                        styles={timeDropdownStyles}
                      />
                    </div>
                    <div className="dt-field">
                      <Text styles={{ root: { color: "#ccc", fontSize: 12, fontWeight: 600 } }}>End Date (UTC)</Text>
                      <DatePicker
                        placeholder="Select end date"
                        value={w.endDate || undefined}
                        onSelectDate={(d) => {
                          setAdditionalWindows((arr) => {
                            const next = [...arr];
                            next[i] = { ...next[i], endDate: d || null };
                            return next;
                          });
                        }}
                        styles={datePickerStyles}
                      />
                    </div>
                    <div className="dt-time">
                      <Text styles={{ root: { color: "#ccc", fontSize: 12, fontWeight: 600 } }}>End Time (UTC)</Text>
                      <Dropdown
                        options={timeOptions}
                        selectedKey={w.endTime}
                        onChange={(_, opt) => opt && setAdditionalWindows((arr) => {
                          const next = [...arr];
                          next[i] = { ...next[i], endTime: opt.key.toString() };
                          return next;
                        })}
                        styles={timeDropdownStyles}
                      />
                    </div>
                    <div className="dt-actions">
                      {/* only show a compact remove button for additional windows */}
                      <button type="button" className="tiny-icon-btn remove-window-btn" aria-label={`Remove window ${i + 2}`} title="Remove Window" onClick={() => removeWindow(i)}>
                        <span className="glyph">−</span>
                      </button>
                    </div>
                  </div>
                ))}
              </div>

                <div style={{ display: 'flex', gap: 12, marginTop: 8 }}>
                <div style={{ flex: 1 }}>
                  <TextField label="Location" value={location} onChange={(_, v) => setLocation(v || "")} styles={textFieldStyles} required errorMessage={showValidation ? fieldErrors.location : undefined} />
                </div>
                <div style={{ flex: 1 }}>
                  <TextField label="Latitude" value={lat} onChange={(_, v) => setLat((v || '').trim())} styles={textFieldStyles} />
                </div>
                <div style={{ flex: 1 }}>
                  <TextField label="Longitude" value={lng} onChange={(_, v) => setLng((v || '').trim())} styles={textFieldStyles} />
                </div>
              </div>

                <div style={{ display: 'flex', gap: 12, marginTop: 8, alignItems: 'flex-end' }}>
                  <div style={{ flex: 1 }}>
                    <TextField label="ISP" value={isp} onChange={(_, v) => setIsp(v || "")} styles={textFieldStyles} required errorMessage={showValidation ? fieldErrors.isp : undefined} />
                  </div>
                  <div style={{ flex: 1 }}>
                    <TextField label="ISP Ticket / Change ID" value={ispTicket} onChange={(_, v) => setIspTicket(v || "")} styles={textFieldStyles} required errorMessage={showValidation ? fieldErrors.ispTicket : undefined} />
                  </div>
                  <div style={{ width: 200 }}>
                    <Dropdown
                      label="Impact Expected"
                      options={[{ key: "true", text: "Yes/True" }, { key: "false", text: "No/False" }]}
                      selectedKey={impactExpected ? "true" : "false"}
                      onChange={(_, opt) => opt && setImpactExpected(opt.key === "true")}
                      styles={dropdownStyles}
                      required
                      errorMessage={showValidation ? fieldErrors.impactExpected : undefined}
                    />
                  </div>
                </div>

                <div style={{ display: 'flex', gap: 8, marginTop: 8, alignItems: 'center' }}>
                  {lat && lng ? (
                    <a className="uid-click" href={`https://www.bing.com/maps?q=${encodeURIComponent(lat+','+lng)}`} target="_blank" rel="noopener noreferrer">Open Map</a>
                  ) : null}
                </div>
              </div>
            </div>

            <div style={{ marginTop: 10 }}>
              {/* Reason wrapper with counter */}
              <div className="reason-wrapper" style={{ position: 'relative' }}>
                <TextField
                  className="reason-field"
                  label="Reason for Maintenance"
                  multiline
                  autoAdjustHeight
                  value={maintenanceReason}
                  onChange={(_, v) => setMaintenanceReason(v || "")}
                  styles={{
                    ...textFieldStyles,
                    field: { ...(textFieldStyles as any).field, minHeight: 220, paddingBottom: 28 },
                    fieldGroup: { ...(textFieldStyles as any).fieldGroup, height: 'auto' },
                  }}
                  // limit to a reasonable amount
                  maxLength={2000}
                  aria-label="Reason for Maintenance"
                  required
                  errorMessage={showValidation ? fieldErrors.maintenanceReason : undefined}
                />

                <div className="reason-counter" aria-hidden style={{ position: 'absolute', right: 10, bottom: 8, fontSize: 12 }}>
                  {`${maintenanceReason.length}/2000`}
                </div>
              </div>
            </div>

            <div className="section-title" style={{ marginTop: 6 }}>Email Body Preview</div>
            <div style={{ background: "#0f0f0f", border: "1px solid #333", borderRadius: 8, padding: 12, whiteSpace: "pre-wrap", color: "#dfefff" }}>
              {emailBody}
            </div>

            <div style={{ display: "flex", justifyContent: "flex-end", marginTop: 12 }}>
                <button className="sleek-btn wan" disabled={!canSend || sendLoading} onClick={handleSend}>
                  {sendLoading ? "Sending..." : "Confirm & Send"}
                </button>
            </div>
          </div>
        )}

        <hr />
        <div className="disclaimer">
        This tool is intended for internal use within Microsoft’s Data Center Operations and Network Delivery environments. Always verify critical data before taking operational action. The information provided is automatically retrieved from validated sources but may not reflect the most recent updates, configurations, or status changes in live systems. Users are responsible for ensuring all details are accurate before proceeding with submitting a VSO. This application is developed and maintained by <b>Josh Maclean</b>, supported by the <b>CIA | Network Delivery</b> team. For for any issues or requests please <a href="https://teams.microsoft.com/l/chat/0/0?users=joshmaclean@microsoft.com" className="uid-click">send a message</a>. 
        </div>
      </div>
    </div>
  );
};

export default VSOAssistant;
