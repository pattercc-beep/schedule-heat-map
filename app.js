let latestRecommendations = [];
let latestGrid = [];

document.getElementById("runButton").addEventListener("click", async function () {
    const file = document.getElementById("fileInput").files[0];
    const studentCount = Number(document.getElementById("studentCount").value);

    if (!file) {
        alert("Please upload a de-identified schedule file first.");
        return;
    }

    if (!Number.isInteger(studentCount) || studentCount < 1) {
        alert("Please enter the number of students submitted in CaRT.");
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

    /*
        Privacy safeguard:
        The uploaded file should NOT contain a UC ID column.
    */
    const uploadedHeaders = Object.keys(rows[0]);

    const containsUCID = uploadedHeaders.some(function (header) {
        const normalized = String(header)
            .toLowerCase()
            .replace(/[\s_-]/g, "");

        return normalized === "ucid";
    });

    if (containsUCID) {
        alert(
            "This file appears to contain a UC ID column. Please remove Column A containing student UCIDs, save the file, and upload the de-identified version."
        );
        return;
    }

    const settings = getSettings();
    const cleanRows = cleanScheduleRows(rows, settings);
    const busySlots = buildBusySlots(cleanRows, settings);

    const availabilityGrid = buildAvailabilityGrid(
        busySlots,
        studentCount,
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
                <strong>${cleanRows.length}</strong>
                day/class rows after cleanup
            </div>

            <div class="summary-box">
                <strong>${bestWindow ? bestWindow.minPctAvailableLabel : "N/A"}</strong>
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
                        (${bestWindow.minPctAvailableLabel} minimum availability).
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
        ${buildAvailabilityTable(availabilityGrid)}
    `;
});


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


function cleanScheduleRows(rows, settings) {
    const cleanRows = [];

    rows.forEach(function (row) {
        const dayText =
            String(row["Days of Week"] || "").trim();

        if (dayText === "") {
            return;
        }

        const days = dayText.split(/\s+/);

        const startMinutes =
            convertExcelTimeToMinutes(
                row["Meeting Start Time"]
            );

        const endMinutes =
            convertExcelTimeToMinutes(
                row["Meeting End Time"]
            );

        const buffer =
            getBufferMinutes(
                row["Location"],
                settings
            );

        days.forEach(function (day) {
            cleanRows.push({
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


function buildBusySlots(cleanRows, settings) {
    const stepMinutes = 10;

    const searchStart =
        timeTextToMinutes(settings.searchStart);

    const searchEnd =
        timeTextToMinutes(settings.searchEnd);

    const busySlots = [];

    cleanRows.forEach(function (row) {
        const startRounded =
            roundDown(
                row.startBuffered,
                stepMinutes
            );

        const endRounded =
            roundUp(
                row.endBuffered,
                stepMinutes
            );

        for (
            let slot = startRounded;
            slot < endRounded;
            slot += stepMinutes
        ) {
            if (
                slot >= searchStart &&
                slot < searchEnd
            ) {
                busySlots.push({
                    day: row.day,
                    slotMinutes: slot,
                    slotLabel: minutesToTimeText(slot)
                });
            }
        }
    });

    return busySlots;
}


function buildAvailabilityGrid(
    busySlots,
    studentCount,
    settings
) {
    const days = ["M", "T", "W", "R", "F"];
    const stepMinutes = 10;

    const searchStart =
        timeTextToMinutes(settings.searchStart);

    const searchEnd =
        timeTextToMinutes(settings.searchEnd);

    const busyMap = new Map();

    /*
        Each schedule record occupying a slot
        contributes 1 to BusyCount.
    */
    busySlots.forEach(function (slot) {
        const key =
            slot.day + "|" + slot.slotMinutes;

        busyMap.set(
            key,
            (busyMap.get(key) || 0) + 1
        );
    });

    const grid = [];

    for (
        let time = searchStart;
        time < searchEnd;
        time += stepMinutes
    ) {
        days.forEach(function (day, index) {
            const key =
                day + "|" + time;

            const rawBusyCount =
                busyMap.get(key) || 0;

            /*
                Prevent overlapping records from
                producing negative availability.
            */
            const busyCount =
                Math.min(
                    rawBusyCount,
                    studentCount
                );

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
                slotMinutes: time,
                slotLabel: minutesToTimeText(time),
                busyCount: busyCount,
                availableCount: availableCount,
                pctAvailable: pctAvailable,
                pctAvailableLabel:
                    Math.round(
                        pctAvailable * 100
                    ) + "%"
            });
        });
    }

    return grid;
}


function findRecommendedWindows(grid, settings) {
    const days = ["M", "T", "W", "R", "F"];
    const stepMinutes = 10;

    const slotsNeeded =
        settings.meetingLength / stepMinutes;

    const searchEnd =
        timeTextToMinutes(settings.searchEnd);

    const results = [];

    days.forEach(function (day, dayIndex) {
        const dayRows = grid
            .filter(
                row => row.day === day
            )
            .sort(
                (a, b) =>
                    a.slotMinutes -
                    b.slotMinutes
            );

        for (
            let i = 0;
            i <= dayRows.length - slotsNeeded;
            i++
        ) {
            const startMinutes =
                dayRows[i].slotMinutes;

            const endMinutes =
                startMinutes +
                settings.meetingLength;

            /*
                Only consider meetings that fit
                completely inside the selected
                search window.
            */
            if (endMinutes > searchEnd) {
                continue;
            }

            const windowRows =
                dayRows.slice(
                    i,
                    i + slotsNeeded
                );

            const minPct =
                Math.min(
                    ...windowRows.map(
                        row =>
                            row.pctAvailable
                    )
                );

            if (
                minPct >=
                settings.minimumAvailability
            ) {
                results.push({
                    day: day,
                    dayOrder: dayIndex + 1,
                    startMinutes: startMinutes,
                    endMinutes: endMinutes,
                    minPctAvailable: minPct,
                    minPctAvailableLabel:
                        Math.round(
                            minPct * 100
                        ) + "%",
                    tier: getTier(minPct)
                });
            }
        }
    });

    return results.sort(
        (a, b) =>
            b.minPctAvailable -
                a.minPctAvailable ||
            a.dayOrder -
                b.dayOrder ||
            getStartTimePreference(a.startMinutes) -
                getStartTimePreference(b.startMinutes) ||
            a.startMinutes -
                b.startMinutes
    );
}


/*
    Reduce repetitive recommendations.

    For each day and availability percentage,
    keep one representative meeting time.

    Preference:
    1. Start on the hour (:00)
    2. Start on the half hour (:30)
    3. Earliest remaining start time

    Availability always remains the primary
    ranking factor.
*/
function collapseRecommendedWindows(windows) {
    const selected = [];
    const days = ["M", "T", "W", "R", "F"];

    days.forEach(function (day) {
        const dayWindows = windows
            .filter(
                window => window.day === day
            )
            .sort((a, b) =>
                b.minPctAvailable -
                    a.minPctAvailable ||
                getStartTimePreference(a.startMinutes) -
                    getStartTimePreference(b.startMinutes) ||
                a.startMinutes -
                    b.startMinutes
            );

        dayWindows.forEach(function (window) {
            const alreadyRepresented =
                selected.some(function (existing) {
                    return (
                        existing.day === window.day &&
                        existing.minPctAvailable ===
                            window.minPctAvailable
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
        .sort((a, b) =>
            b.minPctAvailable -
                a.minPctAvailable ||
            a.dayOrder -
                b.dayOrder ||
            getStartTimePreference(a.startMinutes) -
                getStartTimePreference(b.startMinutes) ||
            a.startMinutes -
                b.startMinutes
        )
        .slice(0, 25);
}


/*
    Lower numbers are preferred.

    :00 = first choice
    :30 = second choice
    everything else = third choice
*/
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
                <th>Minimum Available</th>
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
                        ${window.minPctAvailableLabel}
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


function buildAvailabilityTable(grid) {
    let html = `
        <h3>Availability Heat Map</h3>

        <div class="table-wrap">
            <table>
                <tr>
                    <th>Time</th>
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

                const cssClass =
                    getHeatClass(
                        row.pctAvailable
                    );

                html += `
                    <td
                        class="${cssClass}"
                        title="${row.availableCount} students available"
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
            "Minimum Available",
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
                row.minPctAvailableLabel,
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
            "Time",
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
    let totalMinutes;

    if (typeof value === "number") {
        const fractionOfDay =
            value % 1;

        totalMinutes =
            Math.round(
                fractionOfDay *
                24 *
                60
            );
    } else {
        const date =
            new Date(value);

        totalMinutes =
            date.getHours() * 60 +
            date.getMinutes();
    }

    return totalMinutes;
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
    const hours =
        Math.floor(minutes / 60);

    const mins =
        minutes % 60;

    return (
        String(hours).padStart(2, "0") +
        ":" +
        String(mins).padStart(2, "0")
    );
}


function roundDown(value, step) {
    return (
        Math.floor(value / step) *
        step
    );
}


function roundUp(value, step) {
    return (
        Math.ceil(value / step) *
        step
    );
}