document.addEventListener("DOMContentLoaded", function () {
    let currentChart;

    fetch("db1.xlsx")
        .then((response) => response.arrayBuffer())
        .then((data) => {
            const workbook = XLSX.read(data, { type: "array" });
            const sheetName = workbook.SheetNames[2]; // Assuming the data is on the 6th sheet (index 5)
            const worksheet = workbook.Sheets[sheetName];
            const json = XLSX.utils.sheet_to_json(worksheet);

            currentChart = createBarChart(json,"this week",["thisw open all", "thisw new all", "thisw closed all",], ["Open", "New", "Closed"]);

            document.getElementById("timescale").addEventListener("change", () => {
                const timescale = document.getElementById("timescale").value;
                const region = document.getElementById("region").value;

                    if (currentChart) {
                        currentChart.destroy();
                    }

                    if (timescale === "thisweek" && region === "all") {
                        currentChart = createBarChart(
                            json,
                            "this week",
                            [
                                "thisw open all",
                                "thisw new all",
                                "thisw closed all",
                            ],
                            ["Open", "New", "Closed"]
                        );
                    }
                    if (timescale === "lastweek" && region === "all") {
                        currentChart = createBarChart(
                            json,
                            "last week",
                            [
                                "lastw open all",
                                "lastw new all",
                                "lastw closed all",
                            ],
                            ["Open", "New", "Closed"]
                        );
                    }
                    if (timescale === "month" && region === "all") {
                        currentChart = createBarChart(
                            json,
                            "month",
                            [
                                "month open all",
                                "month new all",
                                "month closed all",
                            ],
                            ["Open", "New", "Closed"]
                        );
                    }
                    if (timescale === "thisweek" && region === "se") {
                        currentChart = createBarChart(
                            json,
                            "this week",
                            [
                                "thisw open se",
                                "thisw new se",
                                "thisw closed se",
                            ],
                            ["Open", "New", "Closed"]
                        );
                    }
                    if (timescale === "lastweek" && region === "se") {
                        currentChart = createBarChart(
                            json,
                            "last week",
                            [
                                "lastw open se",
                                "lastw new se",
                                "lastw closed se",
                            ],
                            ["Open", "New", "Closed"]
                        );
                    }
                    if (timescale === "month" && region === "se") {
                        currentChart = createBarChart(
                            json,
                            "month",
                            [
                                "month open se",
                                "month new se",
                                "month closed se",
                            ],
                            ["Open", "New", "Closed"]
                        );
                    }
                    if (timescale === "thisweek" && region === "sw") {
                        currentChart = createBarChart(
                            json,
                            "this week",
                            [
                                "thisw open sw",
                                "thisw new sw",
                                "thisw closed sw",
                            ],
                            ["Open", "New", "Closed"]
                        );
                    }
                    if (timescale === "lastweek" && region === "sw") {
                        currentChart = createBarChart(
                            json,
                            "last week",
                            [
                                "lastw open sw",
                                "lastw new sw",
                                "lastw closed sw",
                            ],
                            ["Open", "New", "Closed"]
                        );
                    }
                    if (timescale === "month" && region === "sw") {
                        currentChart = createBarChart(
                            json,
                            "month",
                            [
                                "month open sw",
                                "month new sw",
                                "month closed sw",
                            ],
                            ["Open", "New", "Closed"]
                        );
                    }
                });

            document.getElementById("region").addEventListener("change", () => {
                
                const timescale = document.getElementById("timescale").value;
                const region = document.getElementById("region").value;

                if (currentChart) {
                    currentChart.destroy();
                }

                if (timescale === "thisweek" && region === "all") {
                    currentChart = createBarChart(
                        json,
                        "this week",
                        [
                            "thisw open all",
                            "thisw new all",
                            "thisw closed all",
                        ],
                        ["Open", "New", "Closed"]
                    );
                }
                if (timescale === "lastweek" && region === "all") {
                    currentChart = createBarChart(
                        json,
                        "last week",
                        [
                            "lastw open all",
                            "lastw new all",
                            "lastw closed all",
                        ],
                        ["Open", "New", "Closed"]
                    );
                }
                if (timescale === "month" && region === "all") {
                    currentChart = createBarChart(
                        json,
                        "month",
                        [
                            "month open all",
                            "month new all",
                            "month closed all",
                        ],
                        ["Open", "New", "Closed"]
                    );
                }
                if (timescale === "thisweek" && region === "se") {
                    currentChart = createBarChart(
                        json,
                        "this week",
                        [
                            "thisw open se",
                            "thisw new se",
                            "thisw closed se",
                        ],
                        ["Open", "New", "Closed"],
                        (region === 'all' || region === 'se') ? 4 : 4
                    );
                }
                if (timescale === "lastweek" && region === "se") {
                    currentChart = createBarChart(
                        json,
                        "last week",
                        [
                            "lastw open se",
                            "lastw new se",
                            "lastw closed se",
                        ],
                        ["Open", "New", "Closed"],
                        (region === 'all' || region === 'se') ? 7 : 7
                    );
                }
                if (timescale === "month" && region === "se") {
                    currentChart = createBarChart(
                        json,
                        "month",
                        [
                            "month open se",
                            "month new se",
                            "month closed se",
                        ],
                        ["Open", "New", "Closed"],
                        (region === 'all' || region === 'se') ? 20 : 20
                    );
                }
                if (timescale === "thisweek" && region === "sw") {
                    currentChart = createBarChart(
                        json,
                        "this week",
                        [
                            "thisw open sw",
                            "thisw new sw",
                            "thisw closed sw",
                        ],
                        ["Open", "New", "Closed"],
                        (region === 'all' || region === 'sw') ? 4 : 4
                    );
                }
                if (timescale === "lastweek" && region === "sw") {
                    currentChart = createBarChart(
                        json,
                        "last week",
                        [
                            "lastw open sw",
                            "lastw new sw",
                            "lastw closed sw",
                        ],
                        ["Open", "New", "Closed"],
                        (region === 'all' || region === 'sw') ? 7 : 7
                    );
                }
                if (timescale === "month" && region === "sw") {
                    currentChart = createBarChart(
                        json,
                        "month",
                        [
                            "month open sw",
                            "month new sw",
                            "month closed sw",
                        ],
                        ["Open", "New", "Closed"],
                        (region === 'all' || region === 'sw') ? 20 : 20
                    );
                }
            });
        });
});

function createBarChart(data, labelKey, valueKeys, labels, maxY) {
    const labelsData = data.map((row) => row[labelKey]);
    const colors = [
        'rgba(5, 35, 229, 0.2)', 
        'rgba(75, 192, 192, 0.2)',  
        'rgba(255, 99, 132, 0.2)', 
    ];
    const borderColors = [
        'rgba(5, 35, 229, 1)',
        'rgba(75, 192, 192, 1)',  
        'rgba(255, 99, 132, 1)', 
    ];
    const datasets = valueKeys.map((key, index) => {
        return {
            label: labels[index],
            data: data.map((row) => parseFloat(row[key])),
            backgroundColor: colors[index],
            borderColor: borderColors[index],
            borderWidth: 1,
        };
    });

    const ctx = document.getElementById("myChart").getContext("2d");
    return new Chart(ctx, {
        type: "bar",
        data: {
            labels: labelsData,
            datasets: datasets,
        },
        options: {
            responsive: true,
            maintainAspectRatio: false,
            scales: {
                x: {
                    title: {
                        display: false,
                        text:
                            labelKey.charAt(0).toUpperCase() +
                            labelKey.slice(1),
                        color: "#000",
                        font: {
                            family: "",
                            size: 20,
                            weight: "bold",
                            lineHeight: 1.2,
                        },
                        padding: { top: 20, left: 0, right: 0, bottom: 0 },
                    },
                    grid: {
                        display: true,
                        color: "rgba(200, 200, 200, 0.3)",
                    },
                    ticks: {
                        beginAtZero: false,
                        color: "#000",
                        font: {
                            size: 14,
                        },
                    },
                },
                y: {
                    title: {
                        display: true,
                        text: "Complaints",
                        color: "#000",
                        font: {
                            family: "",
                            size: 20,
                            weight: "bold",
                            lineHeight: 1.2,
                        },
                        padding: { top: 0, left: 0, right: 0, bottom: 20 },
                    },
                    grid: {
                        display: true,
                        color: "rgba(200, 200, 200, 0.3)",
                    },
                    ticks: {
                        beginAtZero: true,
                        color: "#000",
                        font: {
                            size: 14,
                        },
                        callback: function(value) {
                            return Number.isInteger(value) ? value : null;
                        }
                    },
                    max: maxY
                }
            }
        }
    });
}
