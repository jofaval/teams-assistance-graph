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

  constructor() {
    this.reset();
  }

  parseSpanishDuration(time) {
    const parts = time.split(/\s+/);

    let hours = 0,
      minutes = 0,
      seconds = 0;

    for (let index = 0; index < parts.length; index += 2) {
      const element = parts[index];

      const type = parts[index + 1];
      if (type.includes("s")) {
        seconds = parseInt(element);
      } else if (type.includes("min")) {
        minutes = parseInt(element);
      } else if (type.includes("h")) {
        hours = parseInt(element);
      }
    }

    if (parts.length >= 2) {
      seconds = parseInt(parts.at(-2));
    }

    if (parts.length >= 4) {
      minutes = parseInt(parts.at(-4));
    }

    if (parts.length >= 6) {
      hours = parseInt(parts.at(-6));
    }

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
    return new Date(date);
  }

  parseFile(content) {
    return content.replaceAll("\t", ";").replaceAll("\r", "");
  }

  getRawAttendees(content) {
    return content.split("2. ")[1].split("3. ")[0].trim().split("\n").slice(1);
  }

  parseAttendees(rawAttendees) {
    return rawAttendees.slice(1).map((attendee) => {
      const values = attendee.split(";");

      return {
        firstEntry: values[USER_ENTRY_KEY.FIRST_ENTRY],
        lastEntry: values[USER_ENTRY_KEY.FIRST_ENTRY + 1],
        length: values[USER_ENTRY_KEY.LENGTH],
        email: values[USER_ENTRY_KEY.EMAIL],
      };
    });
  }

  attendeesWithTimeStats(attendees) {
    return attendees.map((attendee) => {
      const start = this.parseDate(attendee.firstEntry);
      const duration = this.parseDuration(attendee.length);

      const end = new Date(start.getTime() + duration * 1000);
      return { ...attendee, start, end };
    });
  }

  getValueFromGeneralStats(line) {
    return line.split(";").at(-1).trim();
  }

  parseGeneralStats(actualContent) {
    const rows = actualContent
      .trim()
      .split("2. ")[0]
      .trim()
      .split("\n")
      .slice(1);

    const hasUnknownAttendees = rows.length === 7;
    const accessor = hasUnknownAttendees
      ? GENERAL_STATS.UNKNOWN
      : GENERAL_STATS.KNOWN;

    const averageRetention = this.getValueFromGeneralStats(
      rows[accessor.AVERAGE_RETENTION]
    );
    const totalTime = this.getValueFromGeneralStats(rows[accessor.TOTAL_TIME]);

    const rawRetention =
      this.parseDuration(averageRetention) / this.parseDuration(totalTime);
    const retentionPercentage = Math.floor(rawRetention * 10_000) / 100;

    this.generalStats = {
      title: this.getValueFromGeneralStats(rows[accessor.TITLE]),
      totalAttendees: Number(
        this.getValueFromGeneralStats(rows[accessor.TOTAL_ATTENDEES])
      ),
      averageRetention,
      retentionPercentage: retentionPercentage + "%",
      totalTime,
      unknownAttendees: hasUnknownAttendees
        ? Number(
            this.getValueFromGeneralStats(rows[accessor.UNKNOWN_ATTENDEES])
          )
        : 0,
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
      const duration = (attendee.end - attendee.start) / 1000 / 60;
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
    this.parseGeneralStats(actualContent);

    const rawAttendees = this.getRawAttendees(actualContent);
    const attendees = this.parseAttendees(rawAttendees);

    this.attendees = this.attendeesWithTimeStats(attendees);
    this.generalStats.totalWatchHours = this.parseTotalWatchHoursStat(
      this.attendees
    );
    this.clipboardAttendees = this.attendees;

    return this.attendees;
  }

  getTimeSeriesConstraints(attendees) {
    const start = attendees.sort((a, b) => {
      const startA = a.start.getTime();
      const startB = b.start.getTime();

      if (startA < startB) return -1;
      if (startA > startB) return 1;

      return 0;
    })[0].start;

    const end = attendees.sort((a, b) => {
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

  getAttendeesAsTimeSeries() {
    const { start, end } = this.getTimeSeriesConstraints(this.attendees);

    const interval = (end.getTime() - start.getTime()) / 1_000;

    const timeSeries = [];
    for (let index = 1; index <= interval; index++) {
      const now = new Date(start.getTime() + index * 1000);

      const count = this.attendees.filter(({ start, end }) => {
        return this.dateBetweenRange({ start, end, needle: now });
      }).length;

      timeSeries.push({ x: now, y: count });
    }

    return timeSeries;
  }

  getAttendeesOverXMinutes(minutes) {
    return this.clipboardAttendees.filter((attendee) => {
      const duration = (attendee.end - attendee.start) / 1000 / 60;
      return duration > minutes;
    });
  }

  getAttendeesAtPointInTime(date) {
    return this.attendees.filter(({ start, end }) => {
      return this.dateBetweenRange({ start, end, needle: date });
    });
  }

  getAttendeesBetweenDates({ start: first, end: last }) {
    return this.attendees.filter((attendee) => {
      return (
        (attendee.start >= first && attendee.start <= last) ||
        (attendee.end >= first && attendee.end <= last) ||
        (first >= attendee.start && last <= attendee.end)
      );
    });
  }

  getTimeStatsFromAttendees({ attendees, start, end }) {
    const totalTime = (end - start) / 1000;

    const averageTime =
      attendees.reduce((acc, attendee) => {
        const duration =
          Math.min(end, attendee.end) - Math.max(start, attendee.start);
        return acc + duration;
      }, 0) /
      attendees.length /
      1000;

    return {
      averageRetention: this.formatTimeStat(
        this.getHoursMinutesSeconds(averageTime)
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

    return {
      averageRetention,
      retentionPercentage,
      title: this.generalStats.title,
      totalAttendees,
      totalTime,
      unknownAttendees,
      totalWatchHours: totalWatchTime,
    };
  }

  getAttendeesRange() {
    const { start, end } = this.getTimeSeriesConstraints(this.attendees);

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
    this.generalStats = {
      averageRetention: "",
      retentionPercentage: "",
      title: "",
      totalAttendees: 0,
      totalTime: "",
      unknownAttendees: 0,
      totalWatchHours: "",
    };
  }
}

/**
 * @type {TeamsAttendance}
 */
let teamsAttendanceManager;

const dropArea = document.getElementById("dropArea");
const dropAreaView = document.getElementById("dropAreaView");
const graphView = document.getElementById("graphView");
const generalStats = document.getElementById("generalStats");

const attendeesDurationArea = document.querySelector(
  ".attendees-duration-area"
);
const attendeesDurationResult = document.querySelector(
  ".attendees-duration-result"
);

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
  generalStats.style.display = "flex";
}

function resetEventListeners(element) {
  element.outerHTML = element.outerHTML;
}

function enableDropArea() {
  dropAreaView.style.display = "block";
  graphView.style.display = "none";
  graphView.innerHTML = "";
  attendeesDurationArea.style.display = "none";
  generalStats.style.display = "none";

  const fileInput = document.getElementById("fileInput");
  fileInput.value = "";

  attendeesDurationResult.innerHTML = "";
  document.getElementById("attendeesDurationInput").value = "0";
  // remove all attached events
  resetEventListeners(document.getElementById("attendeesDurationInput"));
  resetEventListeners(
    document.getElementById("copyAttendeesOverXMinutesClipboard")
  );

  teamsAttendanceManager.reset();

  generalStats.querySelectorAll(".stat-value").forEach((element) => {
    element.textContent = "...";
  });
}

function buildGraph(data) {
  const chart = Highcharts.chart("graphView", {
    chart: {
      type: "area",
      width: window.innerWidth,
      height: window.innerHeight,
      events: {
        click: function (event) {
          const attendees = teamsAttendanceManager.getAttendeesAtPointInTime(
            event.xAxis[0].value
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
    title: {
      text: teamsAttendanceManager.generalStats.title,
    },
    xAxis: {
      type: "datetime",
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

          const attendees = teamsAttendanceManager.getAttendeesBetweenDates({
            start,
            end,
          });
          teamsAttendanceManager.setClipboardAttendees(attendees);

          console.log("Asistentes en el rango de fechas:", {
            start,
            end,
            attendees,
            atStart: teamsAttendanceManager.getAttendeesAtPointInTime(start),
          });

          const input = document.getElementById("attendeesDurationInput");
          input.dispatchEvent(new Event("input"));

          fillGeneralStats(
            teamsAttendanceManager.getGeneralStatsFromAttendeesInRange({
              attendees,
              start,
              end,
            })
          );
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

  chart.redraw();
}

function copyToClipboard(text) {
  navigator.clipboard.writeText(text).then(() => {
    console.log("Texto copiado al portapapeles:", text);
  });
}

function copyAttendeesToClipboard(attendees) {
  const text = attendees.map((attendee) => `${attendee.email};`).join("\n");

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
      attendeesDurationOverX
    );

    attendeesDurationResult.innerHTML = attendeesDurationOverX.length;
  });

  input.value = "0";
  input.dispatchEvent(new Event("input"));

  const copyAttendeesOverXMinutesClipboard = document.getElementById(
    "copyAttendeesOverXMinutesClipboard"
  );

  copyAttendeesOverXMinutesClipboard.addEventListener("click", () => {
    const minutes = parseInt(input.value);
    copyAttendeesToClipboard(
      teamsAttendanceManager.getAttendeesOverXMinutes(minutes)
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

  // Leer el contenido del archivo (si es necesario)
  const reader = new FileReader();
  reader.onload = function (event) {
    teamsAttendanceManager.processFile(event.target.result);

    const timeseries = teamsAttendanceManager.getAttendeesAsTimeSeries();
    console.log(`Contenido del archivo (${file.name}):`, timeseries);

    buildGraph(timeseries);
    prepareAttendeesDurationQuestion();
    fillGeneralStats(teamsAttendanceManager.generalStats);
  };

  reader.readAsText(file); // Puedes usar readAsText, readAsDataURL, etc.
});
