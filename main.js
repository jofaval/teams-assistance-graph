const dropArea = document.getElementById("dropArea");
const dropAreaView = document.getElementById("dropAreaView");
const graphView = document.getElementById("graphView");

const attendeesDurationArea = document.querySelector(
  ".attendees-duration-area"
);
const attendeesDurationResult = document.querySelector(
  ".attendees-duration-result"
);

function enableGraphView() {
  dropAreaView.style.display = "none";
  graphView.style.display = "block";
  attendeesDurationArea.style.display = "flex";
}

function enableDropArea() {
  dropAreaView.style.display = "block";
  graphView.style.display = "none";
  attendeesDurationArea.style.display = "none";

  const fileInput = document.getElementById("fileInput");
  fileInput.value = "";

  attendeesDurationResult.innerHTML = "";
  document.getElementById("attendeesDurationInput").value = "0";
  // remove all attached events
  document.getElementById("attendeesDurationInput").outerHTML =
    document.getElementById("attendeesDurationInput").outerHTML;
}
enableDropArea();

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

function processDuration(time) {
  const parts = time.split(/\s+/);

  let hours = 0,
    minutes = 0,
    seconds = 0;

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

function processDate(date) {
  return new Date(date);
}

function processFile(fileContent) {
  const actualContent = fileContent
    .replaceAll("\t", ";")
    .replaceAll("\r", "")
    .split("2. ")[1]
    .split("3. ")[0]
    .trim()
    .split("\n")
    .slice(1);

  const head = actualContent[0].split(";");
  const attendance = actualContent
    .slice(1)
    .map((row) =>
      Object.fromEntries(
        row.split(";").map((item, index) => [head[index], item])
      )
    );

  const timeReportAttendance = attendance.map((row) => {
    const start = processDate(row["Primera entrada"]);
    const duration = processDuration(row["Duración de la reunión"]);

    const end = new Date(start.getTime() + duration * 1000);
    return { ...row, start, end, duration };
  });

  return timeReportAttendance;
}

function betweenRange({ start, end, needle }) {
  const startTime = start;
  const endTime = end;
  const needleTime = needle;

  return needleTime >= startTime && needleTime <= endTime;
}

function asTimeSeriesCount(rows) {
  const start = rows.sort((a, b) => {
    const startA = a.start.getTime();
    const startB = b.start.getTime();

    if (startA < startB) return -1;
    if (startA > startB) return 1;

    return 0;
  })[0].start;

  const end = rows.sort((a, b) => {
    const endA = a.end.getTime();
    const endB = b.end.getTime();

    if (endA < endB) return 1;
    if (endA > endB) return -1;

    return 0;
  })[0].end;

  const interval = (end.getTime() - start.getTime()) / 1_000;

  const timeSeries = [];
  for (let index = 1; index <= interval; index++) {
    const count = rows.filter((row) => {
      return betweenRange({
        start: row.start,
        end: row.end,
        needle: new Date(start.getTime() + index * 1000),
      });
    }).length;

    timeSeries.push({
      x: new Date(start.getTime() + index * 1000),
      y: count,
    });
  }

  return timeSeries;
}

function buildGraph(data) {
  Highcharts.chart("graphView", {
    chart: {
      type: "area",
      width: window.innerWidth,
      height: window.innerHeight,
    },
    title: {
      text: "Attendance Over Time",
    },
    xAxis: {
      type: "datetime",
      title: {
        text: "Time",
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
}

function prepareAttendeesDurationQuestion(attendees) {
  const attendeesDuration = attendees.map((row) => {
    return {
      name: row["Nombre"],
      duration: (row.end - row.start) / 1000 / 60,
    };
  });

  const input = document.getElementById("attendeesDurationInput");

  input.addEventListener("input", (e) => {
    console.log(`Asistentes con duración mayor a ${e.target.value} minutos:`);
    const targetDuration = parseInt(e.target.value);
    const attendeesDurationOverX = attendeesDuration.filter((row) => {
      return row.duration > targetDuration;
    });

    console.log(
      `Asistentes con duración mayor a ${targetDuration} minutos:`,
      attendeesDurationOverX
    );

    attendeesDurationResult.innerHTML = attendeesDurationOverX.length;
  });

  input.value = "0";
  input.dispatchEvent(new Event("input"));

  console.log("Duración de los asistentes:", attendeesDuration);
}

// Handle dropped files
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
    const attendees = processFile(event.target.result);
    const timeSeriesAttendance = asTimeSeriesCount(attendees);

    console.log(`Contenido del archivo (${file.name}):`, timeSeriesAttendance);

    buildGraph(timeSeriesAttendance);
    prepareAttendeesDurationQuestion(attendees);
  };
  reader.readAsText(file); // Puedes usar readAsText, readAsDataURL, etc.
});
