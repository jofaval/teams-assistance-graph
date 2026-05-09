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

  getAttendeesAsTimeSeries(resolution = "minute") {
    const entries = this.getTimelineEntries();
    if (entries.length === 0) {
      return [];
    }

    const { start, end } = this.getTimeSeriesConstraints(entries);

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
const distributionChart = document.querySelector("#distribution-chart");
const executiveSummary = document.getElementById("executiveSummary");

bootstrap();

function bootstrap() {
  teamsAttendanceManager = new TeamsAttendance();
  console.log({ teamsAttendanceManager });
  window.dev = { teamsAttendanceManager };

  enableDropArea();
}

function enableGraphView() {
  dropAreaView.style.display = "none";
  graphView.style.display = "block";
  attendeesDurationArea.style.display = "flex";
  executiveSummary.style.display = "grid";
  generalStats.style.display = "grid";
  distributionChart.style.display = "grid";
}

function resetEventListeners(element) {
  element.outerHTML = element.outerHTML;
}

function enableDropArea() {
  dropAreaView.style.display = "block";
  graphView.style.display = "none";
  graphView.innerHTML = "";
  attendeesDurationArea.style.display = "none";
  executiveSummary.style.display = "none";
  generalStats.style.display = "none";
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

  teamsAttendanceManager.reset();

  generalStats.querySelectorAll(".stat-value").forEach((element) => {
    element.textContent = "...";
  });

  document.getElementById("summaryTotalAttendees").textContent = "...";
  document.getElementById("summaryAverageRetention").textContent = "...";
  document.getElementById("summaryTotalTime").textContent = "...";
}

function getTimelineResolution() {
  const selector = document.getElementById("timeResolution");
  return selector?.value === "second" ? "second" : "minute";
}

function redrawGraphUsingSelectedResolution() {
  const timeseries = teamsAttendanceManager.getAttendeesAsTimeSeries(
    getTimelineResolution(),
  );

  buildGraph(timeseries);

  const attendeesRange = teamsAttendanceManager.getAttendeesRange();
  updateDashboardForRange({
    start: attendeesRange.start,
    end: attendeesRange.end,
  });
}

function prepareTimeResolutionSelector() {
  const selector = document.getElementById("timeResolution");
  selector.value = "minute";

  selector.addEventListener("change", () => {
    redrawGraphUsingSelectedResolution();
  });
}

function setDistributionQuartileValue({ label, value, number, total }) {
  const quartile = document.getElementById(`distribution-quartile-${number}`);
  if (!quartile) {
    return;
  }

  const percentage = total > 0 ? Math.floor(((value / total) * 10000) / 100) : 0;
  quartile.innerHTML = `${label}<br/><strong>${value}</strong> attendees (${percentage}%)`;
  quartile.title = `${label}: ${value} attendees`;
}

function getRetentionCohorts({ attendees, start, end }) {
  const rangeMs = end - start;

  if (rangeMs <= 0) {
    return { q1: 0, q2: 0, q3: 0, q4: 0 };
  }

  return attendees.reduce(
    (acc, attendee) => {
      const overlap = teamsAttendanceManager.getAttendeeOverlapInRange({
        attendee,
        start,
        end,
      });
      const ratio = overlap / rangeMs;

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

function fillDistributionChart({ attendees, start, end }) {
  const { q1, q2, q3, q4 } = getRetentionCohorts({ attendees, start, end });
  const total = attendees.length;

  setDistributionQuartileValue({
    label: "0-25% stay",
    value: q1,
    number: 1,
    total,
  });
  setDistributionQuartileValue({
    label: "25-50% stay",
    value: q2,
    number: 2,
    total,
  });
  setDistributionQuartileValue({
    label: "50-75% stay",
    value: q3,
    number: 3,
    total,
  });
  setDistributionQuartileValue({
    label: "75-100% stay",
    value: q4,
    number: 4,
    total,
  });
}

function getSeriesAverage(data) {
  if (data.length === 0) {
    return 0;
  }

  return data.reduce((acc, point) => acc + point.y, 0) / data.length;
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
  teamsAttendanceManager.setClipboardAttendees(attendees);

  const input = document.getElementById("attendeesDurationInput");
  input.dispatchEvent(new Event("input"));

  const stats = teamsAttendanceManager.getGeneralStatsFromAttendeesInRange({
    attendees,
    start,
    end,
  });

  fillGeneralStats(stats);
  fillDistributionChart({ attendees, start, end });
}

function buildGraph(data) {
  const meetingAverage = getSeriesAverage(data);
  const criticalMomentLines = getCriticalMomentLines(data);

  const chart = Highcharts.chart("graphView", {
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
      formatter: function () {
        const point = this.points?.[0]?.point || this.point;
        const pointIndex = point.index;
        const previousPoint = pointIndex > 0 ? data[pointIndex - 1] : null;
        const moment = Highcharts.dateFormat("%Y-%m-%d %H:%M:%S", point.x);
        const totalAttendees = teamsAttendanceManager.generalStats.totalAttendees || 1;
        const retention = ((point.y / totalAttendees) * 100).toFixed(2);
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
        const deltaVsMeetingAverage = point.y - meetingAverage;
        const deltaPrefix = deltaPrevious >= 0 ? "+" : "";
        const meanPrefix = deltaVsMeetingAverage >= 0 ? "+" : "";

        return `${moment}<br/><strong>${point.y}</strong> attendees | <strong>${retention}</strong>% retention<br/>Delta vs previous: <strong>${deltaPrefix}${deltaPrevious}</strong><br/>Delta vs meeting avg: <strong>${meanPrefix}${deltaVsMeetingAverage.toFixed(2)}</strong><br/>5-point trend: <strong>${trend}</strong> (${movingAverage.toFixed(2)} avg)`;
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

          if (typeof e.min == "undefined" && typeof e.max == "undefined") {
            const attendeesRange = teamsAttendanceManager.getAttendeesRange();
            start = attendeesRange.start;
            end = attendeesRange.end;
          } else {
            start = new Date(e.min);
            end = new Date(e.max);
          }

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
    const attendeesRange = teamsAttendanceManager.getAttendeesRange();
    updateDashboardForRange({
      start: attendeesRange.start,
      end: attendeesRange.end,
    });
  };

  reader.readAsArrayBuffer(file);
});
