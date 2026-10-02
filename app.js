let latestRecommendations = [];
let latestGrid = [];

document.getElementById("runButton").addEventListener("click", async function () {
    const file = document.getElementById("fileInput").files[0];

    if (!file) {
        alert("Please upload the de-identified schedule file created by the File Prep Tool.");
        return;
    }

    const data = await file.arrayBuffer();
    const workbook = XLSX.read(data);
    const sheetName = workbook.SheetNames[0];
    const worksheet = workbook.Sheets[sheetName];
    const rows = XLSX.utils.sheet_to_json(worksheet);

    if (rows.length === 0) {
        alert("The uploaded file does not contain any schedule rows.");
        return;
    }

    const uploadedHeaders = Object.keys(rows[0]);

    const containsUCID = uploadedHeaders.some(function (header) {
        const normalized = normalizeHeader(header);
        return normalized === "ucid";
    });

    if (containsUCID) {
        alert(
            "This file contains a UC ID column. Do not upload the original CaRT export. Please use the Schedule Heat Map File Prep Tool and upload the de-identified CSV it creates."
        );
        return;
    }

    const anonymousIDHeader = uploadedHeaders.find(function (header) {
        return normalizeHeader(header) === "anonymousstudentid";
    });

    if (!anonymousIDHeader) {
        alert(
            "The file does not contain an Anonymous Student ID column. Please use the Schedule Heat Map File Prep Tool to prepare the CaRT export before uploading it."
        );
        return;
    }

    const studentIDs = new Set();

    rows.forEach(function (row) {
        const id = String(row[anonymousIDHeader] ?? "").trim();

        if (id !== "") {
            studentIDs.add(id);
        }
    });

    const studentCount = studentIDs.size;

    if (studentCount === 0) {
        alert("No Anonymous Student IDs were found in the uploaded file.");
        return;
    }

    const settings = getSettings();

    const searchStart = timeTextToMinutes(settings.searchStart);
    const searchEnd = timeTextToMinutes(settings.searchEnd);

    if (searchEnd <= searchStart) {
        alert("The meeting search end time must be later than the start time.");
        return;
    }

    if (settings.meetingLength > searchEnd - searchStart) {
        alert("The selected meeting length is longer than the meeting search window.");
        return;
    }

    const cleanRows = cleanScheduleRows(
        rows,
        settings,
        anonymousIDHeader
    );

    if (cleanRows.length === 0) {
        alert("No usable class meeting records were found in the uploaded file.");
        return;
    }

    const availabilityGrid = buildMeetingAvailabilityGrid(
        cleanRows,
        studentIDs,
        settings
    );

    const recommendedWindows = findRecommendedWindows(
        availabilityGrid,
        settings
    );

    const collapsedWindows = collapseRecommendedWindows(
        recommendedWindows
    );

    latestRecommendations = collapsedWindows;
    latestGrid = availabilityGrid;

    const bestWindow =
        collapsedWindows.length > 0
            ? collapsedWindows[0]
            : null;

    const tier1Options = availabilityGrid.filter(function (row) {
        return row.pctAvailable >= 0.90;
    }).length;

    document.getElementById("results").innerHTML = `
        <div class="summary-grid">

            <div class="summary-box">
                <strong>${studentCount}</strong>
                students analyzed
            </div>

            <div class="summary-box">
                <strong>${rows.length}</strong>
                class rows processed
            </div>

            <div class="summary-box">
                <strong>${tier1Options}</strong>
                Tier 1 meeting options
            </div>

            <div class="summary-box">
                <strong>${bestWindow ? bestWindow.pctAvailableLabel : "N/A"}</strong>
                best ${settings.meetingLength}-minute window
            </div>

        </div>

        ${
            bestWindow
                ? `
                    <p>
                        <strong>Best meeting option:</strong>
                        ${dayName(bestWindow.day)},
                        ${minutesToTimeText(bestWindow.startMinutes)}–${minutesToTimeText(bestWindow.endMinutes)}
                        (${bestWindow.pctAvailableLabel} of students available for the full meeting).
                    </p>
                `
                : `
                    <p>
                        <strong>No meeting windows found</strong>
                        at the selected availability threshold.
                    </p>
                `
        }

        <div class="export-buttons">
            <button onclick="downloadRecommendationsCSV()">
                Download Recommendations CSV
            </button>

            <button onclick="downloadHeatMapCSV()">
                Download Heat Map CSV
            </button>

            <button onclick="window.print()">
                Print / Save as PDF
            </button>
        </div>

        ${buildRecommendedWindowsTable(collapsedWindows)}
        ${buildAvailabilityTable(availabilityGrid, settings)}
    `;
});


function normalizeHeader(header) {
    return String(header || "")
        .toLowerCase()
        .replace(/[\s_-]/g, "");
}


function getSettings() {
    return {
        meetingLength:
            Number(document.getElementById("meetingLength").value),

        minimumAvailability:
            Number(document.getElementById("minimumAvailability").value),

        searchStart:
            document.getElementById("searchStart").value,

        searchEnd:
            document.getElementById("searchEnd").value,

        useDefaultBuffers:
            document.getElementById("useDefaultBuffers").checked,

        bufferWestOnline:
            Number(document.getElementById("bufferWestOnline").value),

        bufferEastVictory:
            Number(document.getElementById("bufferEastVictory").value),

        bufferRegionalOffCampus:
            Number(document.getElementById("bufferRegionalOffCampus").value),

        bufferOther:
            Number(document.getElementById("bufferOther").value)
    };
}


function cleanScheduleRows(rows, settings, anonymousIDHeader) {
    const cleanRows = [];

    rows.forEach(function (row) {
        const studentId =
            String(row[anonymousIDHeader] ?? "").trim();

        const dayText =
            String(row["Days of Week"] || "").trim();

        if (studentId === "" || dayText === "") {
            return;
        }

        const startMinutes =
            convertExcelTimeToMinutes(
                row["Meeting Start Time"]
            );

        const endMinutes =
            convertExcelTimeToMinutes(
                row["Meeting End Time"]
            );

        if (
            !Number.isFinite(startMinutes) ||
            !Number.isFinite(endMinutes) ||
            endMinutes <= startMinutes
        ) {
            return;
        }

        const buffer =
            getBufferMinutes(
                row["Location"],
                settings
            );

        const days = parseMeetingDays(dayText);

        days.forEach(function (day) {
            cleanRows.push({
                studentId: studentId,
                subject: row["Subject Code"],
                catalog: row["Catalog Number"],
                location: row["Location"],
                day: day,
                startMinutes: startMinutes,
                endMinutes: endMinutes,
                bufferMinutes: buffer,
                startBuffered: startMinutes - buffer,
                endBuffered: endMinutes + buffer
            });
        });
    });

    return cleanRows;
}


function parseMeetingDays(dayText) {
    const validDays = new Set(["M", "T", "W", "R", "F"]);

    return String(dayText)
        .trim()
        .split(/\s+/)
        .filter(function (day) {
            return validDays.has(day);
        });
}


function buildMeetingAvailabilityGrid(
    cleanRows,
    studentIDs,
    settings
) {
    const days = ["M", "T", "W", "R", "F"];
    const stepMinutes = 10;

    const searchStart =
        timeTextToMinutes(settings.searchStart);

    const searchEnd =
        timeTextToMinutes(settings.searchEnd);

    const latestStart =
        searchEnd - settings.meetingLength;

    const studentCount = studentIDs.size;

    const rowsByDay = new Map();

    days.forEach(function (day) {
        rowsByDay.set(
            day,
            cleanRows.filter(function (row) {
                return row.day === day;
            })
        );
    });

    const grid = [];

    for (
        let startMinutes = searchStart;
        startMinutes <= latestStart;
        startMinutes += stepMinutes
    ) {
        const endMinutes =
            startMinutes + settings.meetingLength;

        days.forEach(function (day, index) {
            const busyStudents = new Set();

            const dayRows =
                rowsByDay.get(day) || [];

            dayRows.forEach(function (row) {
                const overlapsMeeting =
                    row.startBuffered < endMinutes &&
                    row.endBuffered > startMinutes;

                if (overlapsMeeting) {
                    busyStudents.add(row.studentId);
                }
            });

            const busyCount =
                busyStudents.size;

            const availableCount =
                Math.max(
                    studentCount - busyCount,
                    0
                );

            const pctAvailable =
                studentCount === 0
                    ? 0
                    : availableCount / studentCount;

            grid.push({
                day: day,
                dayOrder: index + 1,
                slotMinutes: startMinutes,
                slotLabel: minutesToTimeText(startMinutes),
                startMinutes: startMinutes,
                endMinutes: endMinutes,
                busyCount: busyCount,
                availableCount: availableCount,
                pctAvailable: pctAvailable,
                pctAvailableLabel:
                    Math.round(pctAvailable * 100) + "%",
                tier: getTier(pctAvailable)
            });
        });
    }

    return grid;
}


function findRecommendedWindows(grid, settings) {
    return grid
        .filter(function (row) {
            return row.pctAvailable >=
                settings.minimumAvailability;
        })
        .map(function (row) {
            return {
                ...row
            };
        })
        .sort(function (a, b) {
            return (
                b.pctAvailable -
                    a.pctAvailable ||
                a.dayOrder -
                    b.dayOrder ||
                getStartTimePreference(a.startMinutes) -
                    getStartTimePreference(b.startMinutes) ||
                a.startMinutes -
                    b.startMinutes
            );
        });
}


function collapseRecommendedWindows(windows) {
    const selected = [];
    const days = ["M", "T", "W", "R", "F"];

    days.forEach(function (day) {
        const dayWindows = windows
            .filter(function (window) {
                return window.day === day;
            })
            .sort(function (a, b) {
                return (
                    b.pctAvailable -
                        a.pctAvailable ||
                    getStartTimePreference(a.startMinutes) -
                        getStartTimePreference(b.startMinutes) ||
                    a.startMinutes -
                        b.startMinutes
                );
            });

        dayWindows.forEach(function (window) {
            const alreadyRepresented =
                selected.some(function (existing) {
                    return (
                        existing.day === window.day &&
                        existing.pctAvailable ===
                            window.pctAvailable
                    );
                });

            if (!alreadyRepresented) {
                selected.push({
                    ...window
                });
            }
        });
    });

    return selected
        .sort(function (a, b) {
            return (
                b.pctAvailable -
                    a.pctAvailable ||
                a.dayOrder -
                    b.dayOrder ||
                getStartTimePreference(a.startMinutes) -
                    getStartTimePreference(b.startMinutes) ||
                a.startMinutes -
                    b.startMinutes
            );
        })
        .slice(0, 25);
}


function getStartTimePreference(minutes) {
    const minuteOfHour =
        minutes % 60;

    if (minuteOfHour === 0) {
        return 0;
    }

    if (minuteOfHour === 30) {
        return 1;
    }

    return 2;
}


function buildRecommendedWindowsTable(windows) {
    let html = `
        <h3>Recommended Meeting Windows</h3>

        <table>
            <tr>
                <th>Rank</th>
                <th>Day</th>
                <th>Best Window</th>
                <th>Students Available</th>
                <th>Tier</th>
            </tr>
    `;

    if (windows.length === 0) {
        html += `
            <tr>
                <td colspan="5">
                    No meeting windows found at the selected threshold.
                </td>
            </tr>
        `;
    }

    windows.forEach(
        function (window, index) {
            html += `
                <tr>
                    <td>${index + 1}</td>
                    <td>${dayName(window.day)}</td>
                    <td>
                        ${minutesToTimeText(window.startMinutes)}
                        –
                        ${minutesToTimeText(window.endMinutes)}
                    </td>
                    <td>
                        ${window.pctAvailableLabel}
                    </td>
                    <td>
                        ${window.tier}
                    </td>
                </tr>
            `;
        }
    );

    html += `</table>`;

    return html;
}


function buildAvailabilityTable(grid, settings) {
    let html = `
        <h3>Availability Heat Map</h3>

        <p class="results-note">
            Each cell shows the percentage of students available for the full
            ${settings.meetingLength}-minute meeting beginning at that time.
        </p>

        <div class="table-wrap">
            <table>
                <tr>
                    <th>Meeting Start</th>
                    <th>Monday</th>
                    <th>Tuesday</th>
                    <th>Wednesday</th>
                    <th>Thursday</th>
                    <th>Friday</th>
                </tr>
    `;

    const times =
        [...new Set(
            grid.map(
                row => row.slotMinutes
            )
        )];

    times.forEach(function (time) {
        html += `<tr>`;

        html += `
            <td>
                ${minutesToTimeText(time)}
            </td>
        `;

        ["M", "T", "W", "R", "F"]
            .forEach(function (day) {
                const row =
                    grid.find(
                        r =>
                            r.day === day &&
                            r.slotMinutes === time
                    );

                if (!row) {
                    html += `<td></td>`;
                    return;
                }

                const cssClass =
                    getHeatClass(
                        row.pctAvailable
                    );

                html += `
                    <td
                        class="${cssClass}"
                        title="${row.availableCount} of ${row.availableCount + row.busyCount} students available for the full meeting"
                    >
                        ${row.pctAvailableLabel}
                    </td>
                `;
            });

        html += `</tr>`;
    });

    html += `
            </table>
        </div>
    `;

    return html;
}


function downloadRecommendationsCSV() {
    const rows = [
        [
            "Rank",
            "Day",
            "Start",
            "End",
            "Students Available",
            "Tier"
        ]
    ];

    latestRecommendations.forEach(
        function (row, index) {
            rows.push([
                index + 1,
                dayName(row.day),
                minutesToTimeText(
                    row.startMinutes
                ),
                minutesToTimeText(
                    row.endMinutes
                ),
                row.pctAvailableLabel,
                row.tier
            ]);
        }
    );

    downloadCSV(
        rows,
        "recommended-meeting-windows.csv"
    );
}


function downloadHeatMapCSV() {
    const rows = [
        [
            "Meeting Start",
            "Monday",
            "Tuesday",
            "Wednesday",
            "Thursday",
            "Friday"
        ]
    ];

    const times =
        [...new Set(
            latestGrid.map(
                row => row.slotMinutes
            )
        )];

    times.forEach(function (time) {
        const line = [
            minutesToTimeText(time)
        ];

        ["M", "T", "W", "R", "F"]
            .forEach(function (day) {
                const row =
                    latestGrid.find(
                        r =>
                            r.day === day &&
                            r.slotMinutes === time
                    );

                line.push(
                    row
                        ? row.pctAvailableLabel
                        : ""
                );
            });

        rows.push(line);
    });

    downloadCSV(
        rows,
        "availability-heat-map.csv"
    );
}


function downloadCSV(rows, filename) {
    const csv = rows
        .map(
            row =>
                row
                    .map(
                        value =>
                            `"${String(value).replaceAll(
                                '"',
                                '""'
                            )}"`
                    )
                    .join(",")
        )
        .join("\n");

    const blob =
        new Blob(
            [csv],
            {
                type:
                    "text/csv;charset=utf-8;"
            }
        );

    const link =
        document.createElement("a");

    link.href =
        URL.createObjectURL(blob);

    link.download = filename;

    link.click();

    URL.revokeObjectURL(link.href);
}


function getBufferMinutes(location, settings) {
    const loc =
        String(location || "").trim();

    if (
        loc === "Uptown Campus West" ||
        loc === "Online"
    ) {
        return settings.bufferWestOnline;
    }

    if (
        loc === "Uptown Campus East" ||
        loc === "Victory Parkway"
    ) {
        return settings.bufferEastVictory;
    }

    if (
        loc === "Blue Ash College" ||
        loc === "Clermont College" ||
        loc === "Off Campus"
    ) {
        return settings.bufferRegionalOffCampus;
    }

    return settings.bufferOther;
}


function getTier(pct) {
    if (pct >= 0.90) {
        return "Tier 1 – Excellent";
    }

    if (pct >= 0.80) {
        return "Tier 2 – Strong";
    }

    if (pct >= 0.70) {
        return "Tier 3 – Viable";
    }

    return "Below threshold";
}


function getHeatClass(pct) {
    if (pct >= 0.90) {
        return "heat-excellent";
    }

    if (pct >= 0.80) {
        return "heat-strong";
    }

    if (pct >= 0.70) {
        return "heat-viable";
    }

    return "heat-low";
}


function dayName(day) {
    if (day === "M") return "Monday";
    if (day === "T") return "Tuesday";
    if (day === "W") return "Wednesday";
    if (day === "R") return "Thursday";
    if (day === "F") return "Friday";

    return day;
}


function convertExcelTimeToMinutes(value) {
    if (value === null || value === undefined || value === "") {
        return NaN;
    }

    if (typeof value === "number") {
        const fractionOfDay =
            value % 1;

        return Math.round(
            fractionOfDay *
            24 *
            60
        );
    }

    const text =
        String(value).trim();

    const timeMatch =
        text.match(
            /^(\d{1,2}):(\d{2})(?::\d{2})?\s*(AM|PM)?$/i
        );

    if (timeMatch) {
        let hours =
            Number(timeMatch[1]);

        const minutes =
            Number(timeMatch[2]);

        const period =
            timeMatch[3]
                ? timeMatch[3].toUpperCase()
                : null;

        if (period === "AM" && hours === 12) {
            hours = 0;
        }

        if (period === "PM" && hours !== 12) {
            hours += 12;
        }

        return hours * 60 + minutes;
    }

    const date =
        new Date(value);

    if (Number.isNaN(date.getTime())) {
        return NaN;
    }

    return (
        date.getHours() * 60 +
        date.getMinutes()
    );
}


function timeTextToMinutes(timeText) {
    const parts =
        timeText.split(":");

    return (
        Number(parts[0]) * 60 +
        Number(parts[1])
    );
}


function minutesToTimeText(minutes) {
    const hours24 =
        Math.floor(minutes / 60);

    const mins =
        minutes % 60;

    const period =
        hours24 >= 12
            ? "PM"
            : "AM";

    let hours12 =
        hours24 % 12;

    if (hours12 === 0) {
        hours12 = 12;
    }

    return (
        hours12 +
        ":" +
        String(mins).padStart(2, "0") +
        " " +
        period
    );
}