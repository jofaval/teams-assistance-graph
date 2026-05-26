const GENERAL_STATS = {
  UNKNOWN: {
    TITLE: 0,
    AVERAGE_RETENTION: 6,
    TOTAL_TIME: 5,
    TOTAL_ATTENDEES: 1,
    UNKNOWN_ATTENDEES: 2,
  },
  KNOWN: {
    TITLE: 0,
    AVERAGE_RETENTION: 5,
    TOTAL_TIME: 4,
    TOTAL_ATTENDEES: 1,
  },
};

const USER_ENTRY_KEY = {
  FIRST_ENTRY: 1,
  LAST_ENTRY: 2,
  LENGTH: 3,
  EMAIL: 4,
};

class TeamsAttendance {
  attendees = [];
  clipboardAttendees = [];
  generalStats = {};
  activityEntries = [];
  interactionEvents = [];
  reactionEvents = [];
  attendeesByKey = new Map();

  constructor() {
    this.reset();
  }

  parseSpanishDuration(time) {
    const cleanTime = time.replaceAll("\u00a0", " ").trim();
    const matches = [...cleanTime.matchAll(/(\d+)\s*(h|min|s)/g)];

    let hours = 0;
    let minutes = 0;
    let seconds = 0;

    matches.forEach(([, amount, unit]) => {
      if (unit === "h") {
        hours = parseInt(amount);
      } else if (unit === "min") {
        minutes = parseInt(amount);
      } else if (unit === "s") {
        seconds = parseInt(amount);
      }
    });

    return hours * 60 * 60 + minutes * 60 + seconds;
  }

  parseEnglishDuration(time) {
    const parts = time.split(/\s+/);

    let hours = 0,
      minutes = 0,
      seconds = 0;

    parts.forEach((part) => {
      if (part.includes("s")) {
        return (seconds = parseInt(part));
      } else if (part.includes("m")) {
        return (minutes = parseInt(part));
      } else if (part.includes("h")) {
        return (hours = parseInt(part));
      }
    });

    return hours * 60 * 60 + minutes * 60 + seconds;
  }

  /**
   * @param {string} time
   */
  parseDuration(time) {
    if (/\s(s|min|h)/.test(time)) {
      return this.parseSpanishDuration(time);
    } else {
      return this.parseEnglishDuration(time);
    }
  }

  parseDate(date) {
    const parsedDate = new Date(date);
    if (Number.isNaN(parsedDate.getTime())) {
      return null;
    }

    return parsedDate;
  }

  parseFile(content) {
    return content
      .replaceAll("\t", ";")
      .replaceAll("\r", "")
      .replaceAll("\u00a0", " ");
  }

  getSectionContent(content, sectionNumber) {
    const sectionPattern = new RegExp(
      `(?:^|\\n)${sectionNumber}\\. [^\\n]*\\n([\\s\\S]*?)(?=\\n\\d+\\. [^\\n]*\\n|$)`,
    );
    const match = content.match(sectionPattern);

    return match ? match[1].trim() : "";
  }

  getRawAttendees(content) {
    return content
      .trim()
      .split("\n")
      .map((row) => row.trim())
      .filter(Boolean);
  }

  parseAttendees(rawAttendees) {
    return rawAttendees
      .slice(1)
      .map((attendee) => {
        const values = attendee.split(";");
        const participantName = values[0]?.trim();
        const email = values[USER_ENTRY_KEY.EMAIL]?.trim();
        const identityKey = (email || participantName || "").toLowerCase();

        return {
          participantName,
          firstEntry: values[USER_ENTRY_KEY.FIRST_ENTRY],
          lastEntry: values[USER_ENTRY_KEY.LAST_ENTRY],
          length: values[USER_ENTRY_KEY.LENGTH],
          email,
          identityKey,
        };
      })
      .filter((attendee) => attendee.identityKey);
  }

  attendeesWithTimeStats(attendees) {
    return attendees
      .map((attendee) => {
        const start = this.parseDate(attendee.firstEntry);
        const durationSeconds = this.parseDuration(attendee.length);

        if (!start || Number.isNaN(durationSeconds)) {
          return null;
        }

        const end = new Date(start.getTime() + durationSeconds * 1000);

        return {
          ...attendee,
          start,
          end,
          durationSeconds,
          segments: [{ start, end }],
        };
      })
      .filter(Boolean);
  }

  parseMeetingActivities(rawActivities) {
    return rawActivities
      .slice(1)
      .map((activity) => {
        const values = activity.split(";");
        const participantName = values[0]?.trim();
        const start = this.parseDate(values[1]);
        const end = this.parseDate(values[2]);
        const email = values[4]?.trim();
        const identityKey = (email || participantName || "").toLowerCase();

        if (!participantName || !start || !end || !identityKey) {
          return null;
        }

        return {
          participantName,
          email,
          start,
          end,
          identityKey,
        };
      })
      .filter(Boolean);
  }

  aggregateAttendeesFromActivities(activityEntries) {
    const attendeesMap = new Map();

    activityEntries.forEach((entry) => {
      const current = attendeesMap.get(entry.identityKey);
      const durationSeconds = Math.max(0, (entry.end - entry.start) / 1000);

      if (!current) {
        attendeesMap.set(entry.identityKey, {
          participantName: entry.participantName,
          email: entry.email,
          identityKey: entry.identityKey,
          start: entry.start,
          end: entry.end,
          durationSeconds,
          segments: [{ start: entry.start, end: entry.end }],
        });
        return;
      }

      current.start = new Date(Math.min(current.start.getTime(), entry.start.getTime()));
      current.end = new Date(Math.max(current.end.getTime(), entry.end.getTime()));
      current.durationSeconds += durationSeconds;
      current.segments.push({ start: entry.start, end: entry.end });
    });

    return [...attendeesMap.values()];
  }

  normalizeInteractionType(rawType) {
    return rawType.replaceAll('""', '"').replace(/^"|"$/g, "").trim();
  }

  extractReactionType(type) {
    const normalizedType = this.normalizeInteractionType(type).toLowerCase();
    if (!normalizedType.includes("reacción enviada")) {
      return null;
    }

    const reaction = normalizedType
      .replace("reacción enviada", "")
      .replaceAll('"', "")
      .trim();

    return reaction || null;
  }

  parseMeetingInteractions(rawInteractions) {
    return rawInteractions
      .slice(1)
      .map((interaction) => {
        const values = interaction.split(";");
        const participantName = values[0]?.trim();
        const interactionType = values[1]?.trim();
        const timestamp = this.parseDate(values[2]);

        if (!participantName || !interactionType || !timestamp) {
          return null;
        }

        const normalizedType = this.normalizeInteractionType(interactionType);
        const reactionType = this.extractReactionType(interactionType);

        return {
          participantName,
          interactionType: normalizedType,
          reactionType,
          isReaction: Boolean(reactionType),
          timestamp,
        };
      })
      .filter(Boolean);
  }

  getReactionStats({ start, end, totalAttendees }) {
    const reactionsInRange = this.reactionEvents.filter((reaction) => {
      if (!start || !end) {
        return true;
      }

      return this.dateBetweenRange({
        start,
        end,
        needle: reaction.timestamp,
      });
    });

    const participantsReacted = new Set(
      reactionsInRange.map((reaction) => reaction.participantName),
    ).size;

    return {
      totalReactions: reactionsInRange.length,
      participantsReacted,
    };
  }

  getValueFromGeneralStats(line) {
    return line.split(";").at(-1).trim();
  }

  parseGeneralStats(summarySection) {
    const rows = summarySection.trim().split("\n").filter(Boolean);

    if (rows.length === 0) {
      return;
    }

    const hasUnknownAttendees = rows.length === 7;
    const accessor = hasUnknownAttendees
      ? GENERAL_STATS.UNKNOWN
      : GENERAL_STATS.KNOWN;

    const averageRetention = this.getValueFromGeneralStats(
      rows[accessor.AVERAGE_RETENTION],
    );
    const totalTime = this.getValueFromGeneralStats(rows[accessor.TOTAL_TIME]);

    const rawRetention =
      this.parseDuration(averageRetention) / this.parseDuration(totalTime);
    const retentionPercentage = Math.floor(rawRetention * 10_000) / 100;

    this.generalStats = {
      title: this.getValueFromGeneralStats(rows[accessor.TITLE]),
      totalAttendees: Number(
        this.getValueFromGeneralStats(rows[accessor.TOTAL_ATTENDEES]),
      ),
      averageRetention,
      retentionPercentage: retentionPercentage + "%",
      totalTime,
      unknownAttendees: hasUnknownAttendees
        ? Number(
            this.getValueFromGeneralStats(rows[accessor.UNKNOWN_ATTENDEES]),
          )
        : 0,
      totalReactions: 0,
      participantsReacted: 0,
    };
  }

  getHoursMinutesSeconds(duration) {
    const hours = Math.floor(duration / 3600);
    const minutes = Math.floor((duration % 3600) / 60);
    const seconds = Math.floor(duration % 60);

    return { hours, minutes, seconds };
  }

  getTotalWatchTimeStat(attendees) {
    const totalWatchHours = attendees.reduce((acc, attendee) => {
      const duration =
        attendee.durationSeconds != null
          ? attendee.durationSeconds / 60
          : (attendee.end - attendee.start) / 1000 / 60;
      return acc + duration;
    }, 0);
    const hours = Math.floor(totalWatchHours / 60);
    const minutes = Math.floor(totalWatchHours % 60);
    const seconds = Math.floor((totalWatchHours * 60) % 60);

    return { hours, minutes, seconds };
  }

  formatTimeStat({ hours, minutes, seconds }) {
    return [
      hours > 0 ? hours + "h" : undefined,
      minutes > 0 ? minutes + "m" : undefined,
      seconds > 0 ? seconds + "s" : undefined,
    ]
      .filter(Boolean)
      .join(" ");
  }

  parseTotalWatchHoursStat(attendees) {
    const { hours, minutes, seconds } = this.getTotalWatchTimeStat(attendees);

    return this.formatTimeStat({ hours, minutes, seconds });
  }

  processFile(fileContent) {
    const actualContent = this.parseFile(fileContent);

    const summarySection = this.getSectionContent(actualContent, 1);
    const participantsSection = this.getSectionContent(actualContent, 2);
    const activitiesSection = this.getSectionContent(actualContent, 3);
    const interactionsSection = this.getSectionContent(actualContent, 4);

    this.parseGeneralStats(summarySection);

    const rawAttendees = this.getRawAttendees(participantsSection);
    const attendees = this.parseAttendees(rawAttendees);
    const section2Attendees = this.attendeesWithTimeStats(attendees);

    const rawActivities = this.getRawAttendees(activitiesSection);
    this.activityEntries = this.parseMeetingActivities(rawActivities);
    this.attendees =
      this.activityEntries.length > 0
        ? this.aggregateAttendeesFromActivities(this.activityEntries)
        : section2Attendees;

    this.attendeesByKey = new Map(
      this.attendees.map((attendee) => [attendee.identityKey, attendee]),
    );

    const rawInteractions = this.getRawAttendees(interactionsSection);
    this.interactionEvents = this.parseMeetingInteractions(rawInteractions);
    this.reactionEvents = this.interactionEvents.filter(
      (interaction) => interaction.isReaction,
    );

    this.generalStats.totalWatchHours = this.parseTotalWatchHoursStat(
      this.attendees,
    );

    const attendeesRange = this.getAttendeesRange();
    const reactionStats = this.getReactionStats({
      start: attendeesRange.start,
      end: attendeesRange.end,
      totalAttendees: this.attendees.length,
    });

    this.generalStats.totalReactions = reactionStats.totalReactions;
    this.generalStats.participantsReacted = reactionStats.participantsReacted;

    this.clipboardAttendees = this.attendees;

    return this.attendees;
  }

  getTimelineEntries() {
    if (this.activityEntries.length > 0) {
      return this.activityEntries;
    }

    return this.attendees.map((attendee) => ({
      start: attendee.start,
      end: attendee.end,
      identityKey: attendee.identityKey,
    }));
  }

  getTimeSeriesConstraints(entries) {
    const start = [...entries].sort((a, b) => {
      const startA = a.start.getTime();
      const startB = b.start.getTime();

      if (startA < startB) return -1;
      if (startA > startB) return 1;

      return 0;
    })[0].start;

    const end = [...entries].sort((a, b) => {
      const endA = a.end.getTime();
      const endB = b.end.getTime();

      if (endA < endB) return 1;
      if (endA > endB) return -1;

      return 0;
    })[0].end;

    return { start, end };
  }

  dateBetweenRange({ start, end, needle }) {
    return needle >= start && needle <= end;
  }

  getAttendeesAsTimeSeries(resolution = "minute", rangeStart = null, rangeEnd = null) {
    const entries = this.getTimelineEntries();
    if (entries.length === 0) {
      return [];
    }

    let { start, end } = this.getTimeSeriesConstraints(entries);

    // Si se proporciona un rango personalizado, ajustar los límites
    if (rangeStart && rangeEnd) {
      start = new Date(Math.max(start.getTime(), rangeStart.getTime()));
      end = new Date(Math.min(end.getTime(), rangeEnd.getTime()));
    }

    const stepMs = resolution === "second" ? 1_000 : 60_000;
    const interval = Math.floor((end.getTime() - start.getTime()) / stepMs);

    const timeSeries = [];
    for (let index = 0; index <= interval; index++) {
      const now = new Date(start.getTime() + index * stepMs);

      const activeEntries = entries.filter(({ start, end }) => {
        return this.dateBetweenRange({ start, end, needle: now });
      });

      const count = new Set(
        activeEntries.map((entry) => entry.identityKey),
      ).size;

      timeSeries.push({ x: now, y: count });
    }

    return timeSeries;
  }

  getAttendeesOverXMinutes(minutes) {
    return this.clipboardAttendees.filter((attendee) => {
      const duration =
        attendee.durationSeconds != null
          ? attendee.durationSeconds / 60
          : (attendee.end - attendee.start) / 1000 / 60;
      return duration > minutes;
    });
  }

  getAttendeesAtPointInTime(date) {
    const needle = new Date(date);
    const entries = this.getTimelineEntries();

    const keys = new Set(
      entries
        .filter(({ start, end }) => {
          return this.dateBetweenRange({ start, end, needle });
        })
        .map((entry) => entry.identityKey),
    );

    return [...keys]
      .map((key) => this.attendeesByKey.get(key))
      .filter(Boolean);
  }

  getAttendeesBetweenDates({ start: first, end: last }) {
    const entries = this.getTimelineEntries();

    const keys = new Set(
      entries
        .filter((entry) => {
          return (
            (entry.start >= first && entry.start <= last) ||
            (entry.end >= first && entry.end <= last) ||
            (first >= entry.start && last <= entry.end)
          );
        })
        .map((entry) => entry.identityKey),
    );

    return [...keys]
      .map((key) => this.attendeesByKey.get(key))
      .filter(Boolean);
  }

  getAttendeeOverlapInRange({ attendee, start, end }) {
    if (attendee.segments?.length) {
      return attendee.segments.reduce((acc, segment) => {
        const overlap = Math.max(
          0,
          Math.min(end, segment.end) - Math.max(start, segment.start),
        );

        return acc + overlap;
      }, 0);
    }

    return Math.max(0, Math.min(end, attendee.end) - Math.max(start, attendee.start));
  }

  getTimeStatsFromAttendees({ attendees, start, end }) {
    const totalTime = (end - start) / 1000;

    if (attendees.length === 0 || totalTime <= 0) {
      return {
        averageRetention: "0s",
        retentionPercentage: "0%",
        totalTime: this.formatTimeStat(this.getHoursMinutesSeconds(totalTime)),
      };
    }

    const averageTime =
      attendees.reduce((acc, attendee) => {
        const duration = this.getAttendeeOverlapInRange({ attendee, start, end });
        return acc + duration;
      }, 0) /
      attendees.length /
      1000;

    return {
      averageRetention: this.formatTimeStat(
        this.getHoursMinutesSeconds(averageTime),
      ),
      retentionPercentage:
        Math.floor((averageTime / totalTime) * 10000) / 100 + "%",
      totalTime: this.formatTimeStat(this.getHoursMinutesSeconds(totalTime)),
    };
  }

  getGeneralStatsFromAttendeesInRange({ attendees, start, end }) {
    const { averageRetention, retentionPercentage, totalTime } =
      this.getTimeStatsFromAttendees({ attendees, start, end });

    const totalAttendees = attendees.length;
    const unknownAttendees = this.generalStats.unknownAttendees;
    const totalWatchTime = this.parseTotalWatchHoursStat(attendees);
    const reactionsStats = this.getReactionStats({
      start,
      end,
      totalAttendees,
    });

    return {
      averageRetention,
      retentionPercentage,
      title: this.generalStats.title,
      totalAttendees,
      totalTime,
      unknownAttendees,
      totalWatchHours: totalWatchTime,
      totalReactions: reactionsStats.totalReactions,
      participantsReacted: reactionsStats.participantsReacted,
    };
  }

  getAttendeesRange() {
    const entries = this.getTimelineEntries();
    if (entries.length === 0) {
      const now = new Date();
      return {
        start: now,
        end: now,
      };
    }

    const { start, end } = this.getTimeSeriesConstraints(entries);

    return {
      start,
      end,
    };
  }

  setClipboardAttendees(attendees) {
    this.clipboardAttendees = attendees;
  }

  reset() {
    this.attendees = [];
    this.clipboardAttendees = [];
    this.activityEntries = [];
    this.interactionEvents = [];
    this.reactionEvents = [];
    this.attendeesByKey = new Map();
    this.generalStats = {
      averageRetention: "",
      retentionPercentage: "",
      title: "",
      totalAttendees: 0,
      totalTime: "",
      unknownAttendees: 0,
      totalWatchHours: "",
      totalReactions: 0,
      participantsReacted: 0,
    };
  }
}

/**
 * @type {TeamsAttendance}
 */
let teamsAttendanceManager;
const CRITICAL_DROP_THRESHOLD = 0.2;
const TOOLTIP_MOVING_AVERAGE_WINDOW = 5;
const DISTRIBUTION_VIEW = {
  MILESTONES: "milestones",
  DISTRIBUTION: "distribution",
};

const dropArea = document.getElementById("dropArea");
const dropAreaView = document.getElementById("dropAreaView");
const graphView = document.getElementById("graphView");
const generalStats = document.getElementById("generalStats");

const attendeesDurationArea = document.querySelector(
  ".attendees-duration-area",
);
const attendeesDurationResult = document.querySelector(
  ".attendees-duration-result",
);
const timeRangeSelector = document.getElementById("timeRangeSelector");
const distributionToggle = document.getElementById("distributionToggle");
const distributionChart = document.querySelector("#distribution-chart");
const executiveSummary = document.getElementById("executiveSummary");
let currentDistributionView = DISTRIBUTION_VIEW.MILESTONES;
let lastDistributionRange = null;
let currentChart = null;
let currentTimeseries = [];
let currentRangeContext = {
  start: null,
  end: null,
  totalAttendees: 0,
  averageAttendees: 0,
};
let lastSelectedAttendee = null;

function setHeaderHeightVar() {
  const header = document.querySelector("header");
  const headerHeight = header ? Math.ceil(header.getBoundingClientRect().height) : 0;
  document.documentElement.style.setProperty("--header-height", `${headerHeight}px`);

  // enforce max heights via inline styles to be robust across browsers
  const dashboard = document.getElementById("dashboard");
  if (dashboard) {
    dashboard.style.maxHeight = `calc(100vh - ${headerHeight}px)`;
  }

  const sidebar = document.getElementById("attendeeDetail");
  if (sidebar) {
    sidebar.style.top = `${headerHeight}px`;
    sidebar.style.height = `calc(100vh - ${headerHeight}px)`;
  }
}

bootstrap();

function bootstrap() {
  teamsAttendanceManager = new TeamsAttendance();
  console.log({ teamsAttendanceManager });
  window.dev = { teamsAttendanceManager };

  prepareDistributionToggle();
  prepareSidebarToggle();
  setHeaderHeightVar();
  window.addEventListener('resize', debounce(() => setHeaderHeightVar(), 120));

  enableDropArea();
}

function enableGraphView() {
  document.body.classList.remove("drop-area-active");
  dropAreaView.style.display = "none";
  graphView.style.display = "block";
  attendeesDurationArea.style.display = "flex";
  timeRangeSelector.style.display = "flex";
  executiveSummary.style.display = "grid";
  generalStats.style.display = "grid";
  distributionToggle.style.display = "flex";
  distributionChart.style.display = "grid";
  const listView = document.getElementById("attendeesListView");
  if (listView) listView.classList.remove("hidden");
  const restoreBtn = document.getElementById("restoreListButton");
  if (restoreBtn) restoreBtn.classList.add("hidden");
  const toggleBtn = document.getElementById("toggleList");
  if (toggleBtn) toggleBtn.textContent = "Ocultar lista";
  const dashboard = document.getElementById("dashboard");
  if (dashboard) dashboard.classList.add("two-columns");
}

function resetEventListeners(element) {
  element.outerHTML = element.outerHTML;
}

function enableDropArea() {
  document.body.classList.add("drop-area-active");
  dropAreaView.style.display = "block";
  graphView.style.display = "none";
  graphView.innerHTML = "";
  attendeesDurationArea.style.display = "none";
  timeRangeSelector.style.display = "none";
  executiveSummary.style.display = "none";
  generalStats.style.display = "none";
  distributionToggle.style.display = "none";
  distributionChart.style.display = "none";

  const fileInput = document.getElementById("fileInput");
  fileInput.value = "";

  attendeesDurationResult.innerHTML = "";
  document.getElementById("attendeesDurationInput").value = "0";
  // remove all attached events
  resetEventListeners(document.getElementById("attendeesDurationInput"));
  resetEventListeners(
    document.getElementById("copyAttendeesOverXMinutesClipboard"),
  );
  resetEventListeners(document.getElementById("timeResolution"));
  resetEventListeners(document.getElementById("timeStart"));
  resetEventListeners(document.getElementById("timeEnd"));
  resetEventListeners(document.getElementById("resetTimeRange"));

  document.getElementById("timeStart").value = "";
  document.getElementById("timeEnd").value = "";

  teamsAttendanceManager.reset();

  generalStats.querySelectorAll(".stat-value").forEach((element) => {
    element.textContent = "...";
  });

  document.getElementById("summaryTotalAttendees").textContent = "...";
  document.getElementById("summaryAverageRetention").textContent = "...";
  document.getElementById("summaryTotalTime").textContent = "...";

  setDistributionView(DISTRIBUTION_VIEW.MILESTONES, { render: false });
  lastDistributionRange = null;
  currentRangeContext = {
    start: null,
    end: null,
    totalAttendees: 0,
    averageAttendees: 0,
  };
  const listView = document.getElementById("attendeesListView");
  if (listView) listView.classList.add("hidden");
  const sidebar = document.getElementById("attendeeDetail");
  if (sidebar) {
    sidebar.classList.remove("open");
    sidebar.setAttribute("aria-hidden", "true");
  }
  const restoreBtn = document.getElementById("restoreListButton");
  if (restoreBtn) restoreBtn.classList.add("hidden");
  const dashboard = document.getElementById("dashboard");
  if (dashboard) dashboard.classList.remove("two-columns");
}

function getTimelineResolution() {
  const selector = document.getElementById("timeResolution");
  return selector?.value === "second" ? "second" : "minute";
}

function getTimeRange() {
  const startInput = document.getElementById("timeStart");
  const endInput = document.getElementById("timeEnd");

  const startValue = startInput?.value;
  const endValue = endInput?.value;

  if (!startValue || !endValue) {
    return { start: null, end: null };
  }

  // Obtener la fecha base del rango de asistentes actual
  const attendeesRange = teamsAttendanceManager.getAttendeesRange();
  const baseDate = attendeesRange.start;

  // Parsear HH:MM y crear objetos Date para hoy
  const [startHours, startMinutes] = startValue.split(":").map(Number);
  const [endHours, endMinutes] = endValue.split(":").map(Number);

  const start = new Date(baseDate);
  start.setHours(startHours, startMinutes, 0, 0);

  const end = new Date(baseDate);
  end.setHours(endHours, endMinutes, 59, 999);

  // Validar que end no sea menor que start
  if (end < start) {
    console.warn("End time must be after start time");
    return { start: null, end: null };
  }

  return { start, end };
}

function redrawGraphUsingSelectedResolution() {
  const { start: rangeStart, end: rangeEnd } = getTimeRange();
  const timeseries = teamsAttendanceManager.getAttendeesAsTimeSeries(
    getTimelineResolution(),
    rangeStart,
    rangeEnd,
  );

  buildGraph(timeseries);

  const attendeesRange = teamsAttendanceManager.getAttendeesRange();
  const finalStart = rangeStart || attendeesRange.start;
  const finalEnd = rangeEnd || attendeesRange.end;

  updateDashboardForRange({
    start: finalStart,
    end: finalEnd,
  });
}

function prepareTimeResolutionSelector() {
  const selector = document.getElementById("timeResolution");
  selector.value = "minute";

  selector.addEventListener("change", () => {
    redrawGraphUsingSelectedResolution();
  });
}

function prepareTimeRangeSelector() {
  const startInput = document.getElementById("timeStart");
  const endInput = document.getElementById("timeEnd");
  const resetButton = document.getElementById("resetTimeRange");

  startInput?.addEventListener("change", () => {
    redrawGraphUsingSelectedResolution();
  });

  endInput?.addEventListener("change", () => {
    redrawGraphUsingSelectedResolution();
  });

  resetButton?.addEventListener("click", () => {
    startInput.value = "";
    endInput.value = "";
    redrawGraphUsingSelectedResolution();
  });
}

function setDistributionQuartileValue({ qLabel, label, value, number, total }) {
  const quartile = document.getElementById(`distribution-quartile-${number}`);
  if (!quartile) {
    return;
  }

  const percentage = total > 0 ? Math.floor(((value / total) * 10000) / 100) : 0;
  quartile.innerHTML = `<span class="distribution-title">${qLabel} - ${label}</span><br/><strong>${value}</strong> attendees (${percentage}%)`;
  quartile.title = `${qLabel} - ${label}: ${value} attendees`;
}

function getAttendeePresenceRatio({ attendee, start, end }) {
  const rangeMs = end - start;
  if (rangeMs <= 0) {
    return 0;
  }

  const overlapMs = teamsAttendanceManager.getAttendeeOverlapInRange({
    attendee,
    start,
    end,
  });

  return overlapMs / rangeMs;
}

function getRetentionMilestones({ attendees, start, end }) {
  return attendees.reduce(
    (acc, attendee) => {
      const ratio = getAttendeePresenceRatio({ attendee, start, end });

      if (ratio >= 0.25) {
        acc.q1 += 1;
      }

      if (ratio >= 0.5) {
        acc.q2 += 1;
      }

      if (ratio >= 0.75) {
        acc.q3 += 1;
      }

      if (ratio >= 0.999) {
        acc.q4 += 1;
      }

      return acc;
    },
    { q1: 0, q2: 0, q3: 0, q4: 0 },
  );
}

function getStayDistribution({ attendees, start, end }) {
  return attendees.reduce(
    (acc, attendee) => {
      const ratio = getAttendeePresenceRatio({ attendee, start, end });

      if (ratio <= 0.25) {
        acc.q1 += 1;
      } else if (ratio <= 0.5) {
        acc.q2 += 1;
      } else if (ratio <= 0.75) {
        acc.q3 += 1;
      } else {
        acc.q4 += 1;
      }

      return acc;
    },
    { q1: 0, q2: 0, q3: 0, q4: 0 },
  );
}

function fillDistributionChart({ attendees, start, end, view }) {
  const isMilestones = view === DISTRIBUTION_VIEW.MILESTONES;
  const { q1, q2, q3, q4 } = isMilestones
    ? getRetentionMilestones({ attendees, start, end })
    : getStayDistribution({ attendees, start, end });
  const total = attendees.length;

  setDistributionQuartileValue({
    qLabel: "Q1",
    label: isMilestones ? "25% reached" : "0-25% stay",
    value: q1,
    number: 1,
    total,
  });
  setDistributionQuartileValue({
    qLabel: "Q2",
    label: isMilestones ? "50% reached" : "25-50% stay",
    value: q2,
    number: 2,
    total,
  });
  setDistributionQuartileValue({
    qLabel: "Q3",
    label: isMilestones ? "75% reached" : "50-75% stay",
    value: q3,
    number: 3,
    total,
  });
  setDistributionQuartileValue({
    qLabel: "Q4",
    label: isMilestones ? "100% reached" : "75-100% stay",
    value: q4,
    number: 4,
    total,
  });
}

function setDistributionView(view, { render = true } = {}) {
  currentDistributionView = view;
  distributionChart.dataset.view = view;

  const milestonesButton = document.getElementById("toggleRetentionMilestones");
  const distributionButton = document.getElementById("toggleStayDistribution");
  const isMilestones = view === DISTRIBUTION_VIEW.MILESTONES;

  milestonesButton.classList.toggle("active", isMilestones);
  distributionButton.classList.toggle("active", !isMilestones);
  milestonesButton.setAttribute("aria-pressed", String(isMilestones));
  distributionButton.setAttribute("aria-pressed", String(!isMilestones));

  if (render && lastDistributionRange) {
    fillDistributionChart({
      attendees: lastDistributionRange.attendees,
      start: lastDistributionRange.start,
      end: lastDistributionRange.end,
      view: currentDistributionView,
    });
  }
}

function prepareDistributionToggle() {
  const milestonesButton = document.getElementById("toggleRetentionMilestones");
  const distributionButton = document.getElementById("toggleStayDistribution");

  milestonesButton.addEventListener("click", () => {
    setDistributionView(DISTRIBUTION_VIEW.MILESTONES);
  });

  distributionButton.addEventListener("click", () => {
    setDistributionView(DISTRIBUTION_VIEW.DISTRIBUTION);
  });
}

function getSeriesAverage(data) {
  if (data.length === 0) {
    return 0;
  }

  return data.reduce((acc, point) => acc + point.y, 0) / data.length;
}

function getPointsWithinRange({ data, start, end }) {
  if (!Array.isArray(data) || data.length === 0 || !start || !end) {
    return [];
  }

  const startMs = start.getTime();
  const endMs = end.getTime();

  return data.filter((point) => point.x >= startMs && point.x <= endMs);
}

function updateRangeContext({ start, end, totalAttendees }) {
  const pointsInRange = getPointsWithinRange({
    data: currentTimeseries,
    start,
    end,
  });

  currentRangeContext = {
    start,
    end,
    totalAttendees,
    averageAttendees: getSeriesAverage(pointsInRange),
  };
}

function getMovingAverage(data, index, windowSize) {
  const start = Math.max(0, index - windowSize + 1);
  const window = data.slice(start, index + 1);

  return getSeriesAverage(window);
}

function getTrendLabel({ currentAverage, previousAverage }) {
  const delta = currentAverage - previousAverage;
  if (delta > 0.1) {
    return "up";
  }

  if (delta < -0.1) {
    return "down";
  }

  return "flat";
}

function getTrendMeta(trend) {
  if (trend === "up") {
    return {
      text: "Up trend",
      arrow: "&uarr;",
      className: "is-up",
    };
  }

  if (trend === "down") {
    return {
      text: "Down trend",
      arrow: "&darr;",
      className: "is-down",
    };
  }

  return {
    text: "Flat trend",
    arrow: "&rarr;",
    className: "is-flat",
  };
}

function getSignedValue(value, digits = 0) {
  const safeValue = Number.isFinite(value) ? value : 0;
  const normalizedValue = Math.abs(safeValue) < 0.001 ? 0 : safeValue;
  const fixedValue = normalizedValue.toFixed(digits);

  if (normalizedValue > 0) {
    return `+${fixedValue}`;
  }

  return fixedValue;
}

function getDeltaClass(value) {
  if (value > 0) {
    return "is-up";
  }

  if (value < 0) {
    return "is-down";
  }

  return "is-flat";
}

function getInstantRetentionPercentage({ attendeesAtPoint, totalAttendeesInRange }) {
  if (!totalAttendeesInRange || totalAttendeesInRange <= 0) {
    return 0;
  }

  return (attendeesAtPoint / totalAttendeesInRange) * 100;
}

function getTooltipContent({ point, pointIndex, data }) {
  const previousPoint = pointIndex > 0 ? data[pointIndex - 1] : null;
  const moment = Highcharts.dateFormat("%Y-%m-%d %H:%M:%S", point.x);
  const totalAttendeesInRange = currentRangeContext.totalAttendees || 0;
  const retention = getInstantRetentionPercentage({
    attendeesAtPoint: point.y,
    totalAttendeesInRange,
  });
  const deltaPrevious = previousPoint ? point.y - previousPoint.y : 0;
  const movingAverage = getMovingAverage(
    data,
    pointIndex,
    TOOLTIP_MOVING_AVERAGE_WINDOW,
  );
  const previousMovingAverage = getMovingAverage(
    data,
    Math.max(0, pointIndex - 1),
    TOOLTIP_MOVING_AVERAGE_WINDOW,
  );
  const trend = getTrendLabel({
    currentAverage: movingAverage,
    previousAverage: previousMovingAverage,
  });
  const trendMeta = getTrendMeta(trend);
  const deltaVsRangeAverage = point.y - (currentRangeContext.averageAttendees || 0);
  const deltaPreviousClass = getDeltaClass(deltaPrevious);
  const deltaRangeClass = getDeltaClass(deltaVsRangeAverage);

  return `
    <div class="tooltip-card">
      <div class="tooltip-time">${moment}</div>
      <div class="tooltip-main-row">
        <div class="tooltip-main-value">${point.y}</div>
        <div class="tooltip-main-caption">attendees</div>
      </div>
      <div class="tooltip-retention">${retention.toFixed(2)}% retention in visible range</div>
      <div class="tooltip-trend ${trendMeta.className}">
        <span class="tooltip-trend-arrow">${trendMeta.arrow}</span>
        <span>${trendMeta.text}</span>
      </div>
      <div class="tooltip-deltas">
        <span class="tooltip-delta ${deltaPreviousClass}">${getSignedValue(deltaPrevious)} vs prev</span>
        <span class="tooltip-delta ${deltaRangeClass}">${getSignedValue(deltaVsRangeAverage, 2)} vs range avg</span>
      </div>
    </div>
  `;
}

function detectCriticalMoments(data, threshold = CRITICAL_DROP_THRESHOLD) {
  const moments = [];

  for (let index = 1; index < data.length; index++) {
    const previous = data[index - 1];
    const current = data[index];

    if (previous.y <= 0) {
      continue;
    }

    const drop = (previous.y - current.y) / previous.y;
    if (drop >= threshold) {
      moments.push({
        x: new Date(current.x).getTime(),
        drop,
      });
    }
  }

  return moments;
}

function getCriticalMomentLines(data) {
  return detectCriticalMoments(data).map((moment) => ({
    color: "#c64747",
    width: 1,
    dashStyle: "ShortDot",
    zIndex: 3,
    value: moment.x,
    label: {
      text: `-${Math.round(moment.drop * 100)}%`,
      align: "left",
      x: 4,
      style: {
        color: "#8f2222",
        fontSize: "10px",
      },
    },
  }));
}

function updateDashboardForRange({ start, end }) {
  const attendees = teamsAttendanceManager.getAttendeesBetweenDates({
    start,
    end,
  });
  lastDistributionRange = { attendees, start, end };

  teamsAttendanceManager.setClipboardAttendees(attendees);

  const input = document.getElementById("attendeesDurationInput");
  input.dispatchEvent(new Event("input"));

  const stats = teamsAttendanceManager.getGeneralStatsFromAttendeesInRange({
    attendees,
    start,
    end,
  });

  updateRangeContext({
    start,
    end,
    totalAttendees: attendees.length,
  });

  fillGeneralStats(stats);
  fillDistributionChart({
    attendees,
    start,
    end,
    view: currentDistributionView,
  });
  // Populate attendees list view for the selected range
  try {
    fillAttendeesList(attendees, { start, end });
  } catch (err) {
    console.warn("Error filling attendees list:", err);
  }
}

function buildGraph(data) {
  currentTimeseries = data;
  const criticalMomentLines = getCriticalMomentLines(data);

  if (currentChart) {
    currentChart.destroy();
    currentChart = null;
  }

  let chart = Highcharts.chart("graphView", {
    chart: {
      type: "area",
      events: {
        click: function (event) {
          const attendees = teamsAttendanceManager.getAttendeesAtPointInTime(
            event.xAxis[0].value,
          );
          console.log("Asistentes en el momento:", attendees);
          copyAttendeesToClipboard(attendees);
        },
      },
      zooming: {
        type: "x",
        resetButton: {
          position: {
            align: "center",
            x: -30,
            y: 35,
          },
          relativeTo: "chart",
        },
      },
    },
    tooltip: {
      enabled: true,
      shared: true,
      useHTML: true,
      className: "attendance-tooltip",
      formatter: function () {
        const point = this.points?.[0]?.point || this.point;
        const pointIndex = point.index;
        return getTooltipContent({ point, pointIndex, data });
      },
    },
    title: {
      text: teamsAttendanceManager.generalStats.title,
    },
    xAxis: {
      type: "datetime",
      plotLines: criticalMomentLines,
      title: {
        text: "Time",
      },
      events: {
        setExtremes: function (e) {
          let start, end;
          const isReset = typeof e.min == "undefined" && typeof e.max == "undefined";

          if (isReset) {
            const attendeesRange = teamsAttendanceManager.getAttendeesRange();
            start = attendeesRange.start;
            end = attendeesRange.end;

            const startInput = document.getElementById("timeStart");
            const endInput = document.getElementById("timeEnd");
            if (startInput) startInput.value = "";
            if (endInput) endInput.value = "";
          } else {
            start = new Date(e.min);
            end = new Date(e.max);

            const pad = (n) => String(n).padStart(2, "0");
            const startHHMM = `${pad(start.getHours())}:${pad(start.getMinutes())}`;
            const endHHMM = `${pad(end.getHours())}:${pad(end.getMinutes())}`;

            const startInput = document.getElementById("timeStart");
            const endInput = document.getElementById("timeEnd");
            if (startInput) startInput.value = startHHMM;
            if (endInput) endInput.value = endHHMM;
          }

          const rangeStart = start.getTime();
          const rangeEnd = end.getTime();
          const visibleData = currentTimeseries.filter(
            (point) => point.x >= rangeStart && point.x <= rangeEnd,
          );
          const newCriticalLines = getCriticalMomentLines(visibleData);
          const axis = this;
          setTimeout(() => {
            axis.update({ plotLines: newCriticalLines }, true);
          }, 0);

          updateDashboardForRange({ start, end });
        },
      },
    },
    yAxis: {
      title: {
        text: "Number of Attendees",
      },
    },
    series: [
      {
        name: "Attendees",
        data,
      },
    ],
    time: {
      timezoneOffset: new Date().getTimezoneOffset(),
    },
  });

  window.requestAnimationFrame(() => {
    chart.reflow();
  });

  chart.redraw();
  currentChart = chart;
}

function copyToClipboard(text) {
  navigator.clipboard.writeText(text).then(() => {
    console.log("Texto copiado al portapapeles:", text);
  });
}

function copyAttendeesToClipboard(attendees) {
  const text = attendees
    .map((attendee) => attendee.email)
    .filter(Boolean)
    .map((email) => `${email};`)
    .join("\n");

  copyToClipboard(text);
}

function prepareAttendeesDurationQuestion() {
  const input = document.getElementById("attendeesDurationInput");

  input.addEventListener("input", (e) => {
    console.log(`Asistentes con duración mayor a ${e.target.value} minutos:`);
    const attendeesDurationOverX =
      teamsAttendanceManager.getAttendeesOverXMinutes(parseInt(e.target.value));

    console.log(
      `Asistentes con duración mayor a ${e.target.value} minutos:`,
      attendeesDurationOverX,
    );

    attendeesDurationResult.innerHTML = attendeesDurationOverX.length;
  });

  input.value = "0";
  input.dispatchEvent(new Event("input"));

  const copyAttendeesOverXMinutesClipboard = document.getElementById(
    "copyAttendeesOverXMinutesClipboard",
  );

  copyAttendeesOverXMinutesClipboard.addEventListener("click", () => {
    const minutes = parseInt(input.value);
    copyAttendeesToClipboard(
      teamsAttendanceManager.getAttendeesOverXMinutes(minutes),
    );
  });
}

function fillGeneralStats(generalStats) {
  document.querySelector(".total-attendees .stat-value").textContent =
    generalStats.totalAttendees;
  document.querySelector(".average-retention .stat-value").textContent =
    generalStats.averageRetention;
  document.querySelector(".retention-percentage .stat-value").textContent =
    generalStats.retentionPercentage;
  document.querySelector(".total-duration .stat-value").textContent =
    generalStats.totalTime;
  document.querySelector(".unknown-attendees .stat-value").textContent =
    generalStats.unknownAttendees;
  document.querySelector(".total-watch-hours .stat-value").textContent =
    generalStats.totalWatchHours;
  document.querySelector(".total-reactions .stat-value").textContent =
    generalStats.totalReactions;
  document.querySelector(".participants-reacted .stat-value").textContent =
    generalStats.participantsReacted;

  document.getElementById("summaryTotalAttendees").textContent =
    generalStats.totalAttendees;
  document.getElementById("summaryAverageRetention").textContent =
    generalStats.averageRetention;
  document.getElementById("summaryTotalTime").textContent =
    generalStats.totalTime;
}

function decodeFileContent(arrayBuffer) {
  const bytes = new Uint8Array(arrayBuffer);
  const isUtf16Le = bytes[0] === 0xff && bytes[1] === 0xfe;
  const isUtf16Be = bytes[0] === 0xfe && bytes[1] === 0xff;
  const encoding = isUtf16Le ? "utf-16le" : isUtf16Be ? "utf-16be" : "utf-8";

  return new TextDecoder(encoding).decode(bytes);
}

function debounce(fn, wait = 200) {
  let timer;
  return (...args) => {
    clearTimeout(timer);
    timer = setTimeout(() => fn.apply(this, args), wait);
  };
}

function prepareSidebarToggle() {
  const btn = document.getElementById("toggleSidebarButton");
  const sidebar = document.getElementById("attendeeDetail");
  if (!btn) return;

  btn.addEventListener("click", () => {
    if (!sidebar) return;
    const isOpen = sidebar.classList.contains("open");
    if (isOpen) {
      sidebar.classList.remove("open");
      sidebar.setAttribute("aria-hidden", "true");
      btn.setAttribute("aria-pressed", "false");
    } else {
      sidebar.classList.add("open");
      sidebar.setAttribute("aria-hidden", "false");
      btn.setAttribute("aria-pressed", "true");
      // If there's a previously selected attendee, render its details
      if (lastSelectedAttendee) {
        const { start, end } = currentRangeContext.start && currentRangeContext.end ? currentRangeContext : teamsAttendanceManager.getAttendeesRange();
        showAttendeeDetail(lastSelectedAttendee, { start, end });
      } else {
        const content = document.getElementById("attendeeDetailContent");
        if (content) content.innerHTML = `<div class="meta">Selecciona un asistente en la lista para ver detalles.</div>`;
      }
    }
  });
}

function prepareAttendeesListView() {
  const search = document.getElementById("attendeeSearch");
  const minDuration = document.getElementById("attendeeMinDuration");
  const sortSelect = document.getElementById("attendeeSort");
  const toggleBtn = document.getElementById("toggleList");
  const closeBtn = document.getElementById("closeAttendeeDetail");

  // Initialize toggle button label according to current visibility
  const initialListView = document.getElementById("attendeesListView");
  if (toggleBtn && initialListView) {
    const isHidden = initialListView.classList.contains("hidden") || window.getComputedStyle(initialListView).display === "none";
    toggleBtn.textContent = isHidden ? "Mostrar lista" : "Ocultar lista";
  }

  const onInputChange = debounce(() => {
    const attendees = lastDistributionRange?.attendees || [];
    const rangeStart = lastDistributionRange?.start;
    const rangeEnd = lastDistributionRange?.end;
    fillAttendeesList(attendees, { start: rangeStart, end: rangeEnd });
  }, 200);

  search?.addEventListener("input", onInputChange);
  minDuration?.addEventListener("input", onInputChange);
  sortSelect?.addEventListener("change", onInputChange);

  toggleBtn?.addEventListener("click", () => {
    const listView = document.getElementById("attendeesListView");
    if (!listView) return;
    const nowHidden = listView.classList.toggle("hidden");
    toggleBtn.textContent = nowHidden ? "Mostrar lista" : "Ocultar lista";
    // show/hide the restore button in header
    const restoreBtn = document.getElementById("restoreListButton");
    if (restoreBtn) {
      if (nowHidden) restoreBtn.classList.remove("hidden");
      else restoreBtn.classList.add("hidden");
    }
    // switch dashboard columns
    const dashboard = document.getElementById("dashboard");
    if (dashboard) {
      if (nowHidden) dashboard.classList.remove("two-columns");
      else dashboard.classList.add("two-columns");
    }
  });

  const restoreBtn = document.getElementById("restoreListButton");
  restoreBtn?.addEventListener("click", () => {
    const listView = document.getElementById("attendeesListView");
    if (!listView) return;
    listView.classList.remove("hidden");
    restoreBtn.classList.add("hidden");
    if (toggleBtn) toggleBtn.textContent = "Ocultar lista";
    const dashboard = document.getElementById("dashboard");
    if (dashboard) dashboard.classList.add("two-columns");
  });

  closeBtn?.addEventListener("click", () => {
    const sidebar = document.getElementById("attendeeDetail");
    sidebar?.classList.remove("open");
    sidebar?.setAttribute("aria-hidden", "true");
  });
}

function fillAttendeesList(attendees, { start, end } = {}) {
  const tbody = document.querySelector("#attendeesTable tbody");
  if (!tbody) return;

  const searchValue = (document.getElementById("attendeeSearch")?.value || "").toLowerCase();
  const minDurationValue = parseFloat(document.getElementById("attendeeMinDuration")?.value || "0");
  const sortValue = document.getElementById("attendeeSort")?.value || "name";

  const filtered = (attendees || []).filter((attendee) => {
    if (!attendee) return false;
    const name = attendee.participantName || "";
    const email = attendee.email || "";
    const durationMin = (attendee.durationSeconds != null ? attendee.durationSeconds : Math.round((attendee.end - attendee.start)/1000))/60;
    if (minDurationValue && durationMin < minDurationValue) return false;
    if (searchValue) {
      return (name.toLowerCase().includes(searchValue) || email.toLowerCase().includes(searchValue));
    }
    return true;
  });

  filtered.sort((a,b) => {
    if (sortValue === "name") {
      return (a.participantName || "").localeCompare(b.participantName || "");
    }
    if (sortValue === "duration") {
      const da = a.durationSeconds != null ? a.durationSeconds : (a.end - a.start)/1000;
      const db = b.durationSeconds != null ? b.durationSeconds : (b.end - b.start)/1000;
      return db - da;
    }
    if (sortValue === "retention") {
      if (!start || !end) return 0;
      const ra = teamsAttendanceManager.getAttendeeOverlapInRange({attendee:a, start, end}) ;
      const rb = teamsAttendanceManager.getAttendeeOverlapInRange({attendee:b, start, end}) ;
      return rb - ra;
    }
    return 0;
  });

  tbody.innerHTML = "";

  filtered.forEach((attendee) => {
    const overlapMs = start && end ? teamsAttendanceManager.getAttendeeOverlapInRange({ attendee, start, end }) : (attendee.durationSeconds != null ? attendee.durationSeconds*1000 : (attendee.end - attendee.start));
    const retention = start && end ? Math.floor((overlapMs / (end - start)) * 10000) / 100 : 0;
    const durationSeconds = attendee.durationSeconds != null ? attendee.durationSeconds : Math.round((attendee.end - attendee.start)/1000);
    const durationLabel = teamsAttendanceManager.formatTimeStat(teamsAttendanceManager.getHoursMinutesSeconds(durationSeconds));
    const tr = document.createElement("tr");
    tr.className = "list-row";
    tr.dataset.identityKey = attendee.identityKey || "";

    tr.innerHTML = `<td>${attendee.participantName || ""}${attendee.email ? ` <div class="meta">${attendee.email}</div>` : ""}</td><td>${retention}%</td><td>${durationLabel}</td>`;
    tr.addEventListener("click", () => {
      showAttendeeDetail(attendee, { start: start || teamsAttendanceManager.getAttendeesRange().start, end: end || teamsAttendanceManager.getAttendeesRange().end });
    });

    tbody.appendChild(tr);
  });
}

function showAttendeeDetail(attendee, { start, end } = {}) {
  const sidebar = document.getElementById("attendeeDetail");
  const content = document.getElementById("attendeeDetailContent");
  if (!sidebar || !content || !attendee) return;

  lastSelectedAttendee = attendee;

  const overlapMs = start && end ? teamsAttendanceManager.getAttendeeOverlapInRange({ attendee, start, end }) : (attendee.durationSeconds != null ? attendee.durationSeconds*1000 : (attendee.end - attendee.start));
  const retentionPct = start && end ? Math.floor((overlapMs / (end - start)) * 10000) / 100 : 0;
  const durationSeconds = attendee.durationSeconds != null ? attendee.durationSeconds : Math.round((attendee.end - attendee.start)/1000);
  const durationLabel = teamsAttendanceManager.formatTimeStat(teamsAttendanceManager.getHoursMinutesSeconds(durationSeconds));

  const segmentsHtml = (attendee.segments || []).map(s => `<li>${new Date(s.start).toLocaleString()} → ${new Date(s.end).toLocaleString()}</li>`).join("");
  const interactions = teamsAttendanceManager.interactionEvents.filter(ev => {
    const norm = (ev.participantName || "").toLowerCase();
    const name = (attendee.participantName || "").toLowerCase();
    const email = (attendee.email || "").toLowerCase();
    return norm === name || (email && norm.includes(email));
  });
  const reactions = interactions.filter(i => i.isReaction);

  content.innerHTML = `\n    <h3>${attendee.participantName || "(sin nombre)"}</h3>\n    <div class="meta">${attendee.email || ""}</div>\n    <div><strong>Duración:</strong> ${durationLabel}</div>\n    <div><strong>Retención:</strong> ${retentionPct}%</div>\n    <div><strong>Segments:</strong><ul>${segmentsHtml}</ul></div>\n    <div><strong>Total interacciones:</strong> ${interactions.length}</div>\n    <div><strong>Total reacciones:</strong> ${reactions.length}</div>\n  `;

  sidebar.classList.add("open");
  sidebar.setAttribute("aria-hidden", "false");
}

// Prevent default behavior for drag events
["dragenter", "dragover", "dragleave", "drop"].forEach((eventName) => {
  dropArea.addEventListener(eventName, (e) => e.preventDefault());
  dropArea.addEventListener(eventName, (e) => e.stopPropagation());
});

// Highlight drop area when file is dragged over
["dragenter", "dragover"].forEach((eventName) => {
  dropArea.addEventListener(eventName, () => {
    dropArea.classList.add("drag-over");
  });
});

// Remove highlight when file is dragged out
["dragleave", "drop"].forEach((eventName) => {
  dropArea.addEventListener(eventName, () => {
    dropArea.classList.remove("drag-over");
  });
});

dropArea.addEventListener("drop", (e) => {
  const files = e.dataTransfer.files;
  console.log("Dropped files:", files);

  const file = files[0];
  if (!file) {
    console.log("No file dropped");
    return;
  }

  enableGraphView();

  // Example: Display file names in the console
  console.log(`File name: ${file.name}`);

  const reader = new FileReader();
  reader.onload = function (event) {
    const decodedContent = decodeFileContent(event.target.result);

    teamsAttendanceManager.processFile(decodedContent);

    const timeseries = teamsAttendanceManager.getAttendeesAsTimeSeries(
      getTimelineResolution(),
    );
    console.log(`Contenido del archivo (${file.name}):`, timeseries);

    buildGraph(timeseries);
    prepareAttendeesDurationQuestion();
    prepareTimeResolutionSelector();
    prepareTimeRangeSelector();
    prepareAttendeesListView();
    const attendeesRange = teamsAttendanceManager.getAttendeesRange();
    updateDashboardForRange({
      start: attendeesRange.start,
      end: attendeesRange.end,
    });
  };

  reader.readAsArrayBuffer(file);
});

