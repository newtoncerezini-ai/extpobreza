const workbookPath = "Ranking_PE_25.xlsx";

const peCriteria = {
  current: {
    label: "PPC 2021 • US$ 3/dia",
    title: "Critério atual: PPC 2021",
    note:
      "A linha atual considera aproximadamente R$ 280 mensais por pessoa. Mesmo com uma régua mais rigorosa, Pernambuco reduz a extrema pobreza entre 2022 e 2025.",
    population: [1519889, 1325134, 1162913, 1338673, 1583781, 1611153, 1568470, 1698610, 1626931, 2303089, 1521944, 1119245, 929279, 895796],
    percent: [16.94, 14.66, 12.79, 14.62, 17.21, 17.44, 16.89, 18.19, 17.35, 24.47, 16.13, 11.84, 9.81, 9.43],
  },
  old: {
    label: "PPC 2017 • US$ 2,15/dia",
    title: "Critério antigo: PPC 2017",
    note:
      "A linha antiga considera aproximadamente R$ 232 mensais por pessoa. A série também mostra queda relevante entre 2022 e 2025.",
    population: [1146733, 980337, 882061, 1064629, 1286446, 1317598, 1255782, 1404513, 1215507, 1923735, 1266403, 893117, 621883, 626097],
    percent: [12.78, 10.85, 9.7, 11.63, 13.98, 14.26, 13.52, 15.04, 12.96, 20.44, 13.42, 9.44, 6.56, 6.59],
  },
};

const validation = {
  years: [2012, 2013, 2014, 2015, 2016, 2017, 2018, 2019, 2020, 2021, 2022, 2023, 2024, 2025],
  processed: [18.34, 18.14, 14.78, 16.21, 18.39, 17.13, 16.77, 16.61, 13.11, 19.98, 14.55, 12.54, 10.44, 9.44],
  ipece: [18.3, 18.1, 14.6, 15.9, 18.1, 17.1, 16.8, 16.5, 13.0, 19.8, 14.5, 12.4, 10.3, 9.4],
  oldProcessed: [14.64, 13.92, 12.03, 12.34, 14.21, 14.04, 13.85, 13.69, 9.93, 16.32, 11.68, 9.39, 8.02, 6.39],
  oldIpece: [14.6, 13.8, 11.9, 12.2, 14.2, 14.0, 13.8, 13.6, 10.0, 16.2, 11.0, 9.4, 7.9, 6.4],
};

const charts = {};
let stateRows = [];
let rankingRows = [];

const years = peCriteria.current.percent.map((_, index) => 2012 + index);
const brNumber = new Intl.NumberFormat("pt-BR");
const brPercent = new Intl.NumberFormat("pt-BR", {
  minimumFractionDigits: 1,
  maximumFractionDigits: 1,
});
const brPercent2 = new Intl.NumberFormat("pt-BR", {
  minimumFractionDigits: 2,
  maximumFractionDigits: 2,
});

document.addEventListener("DOMContentLoaded", async () => {
  setupChartDefaults();
  wireControls();
  await loadWorkbook();
  renderAll();
});

async function loadWorkbook() {
  const response = await fetch(workbookPath);
  const data = await response.arrayBuffer();
  const workbook = XLSX.read(data, { type: "array" });

  stateRows = sheetToRows(workbook.Sheets.Todos_Estados, 4);
  rankingRows = sheetToRows(workbook.Sheets.Ranking_2025_vs_2022, 4);

  populateRegions();
}

function sheetToRows(sheet, headerRow) {
  const raw = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: null });
  const headers = raw[headerRow - 1].filter(Boolean);
  return raw.slice(headerRow).filter((row) => row.some((value) => value !== null)).map((row) => {
    return Object.fromEntries(headers.map((header, index) => [header, row[index]]));
  });
}

function setupChartDefaults() {
  Chart.defaults.font.family = "Arial, Helvetica, sans-serif";
  Chart.defaults.color = "#4d5753";
  Chart.defaults.plugins.tooltip.backgroundColor = "#17221f";
  Chart.defaults.plugins.tooltip.padding = 12;
  Chart.defaults.plugins.tooltip.cornerRadius = 6;
}

function renderAll() {
  const peRows = stateRows.filter((row) => row.UF === "PE").sort((a, b) => a.Ano - b.Ano);
  renderMainLine(peRows);
  renderSeries(peRows);
  renderRanking();
  renderCriterion("current");
  renderValidation();
  updateSeriesStats(peRows);
}

function lineDataset(label, data, color, yAxisID = "y") {
  return {
    label,
    data,
    borderColor: color,
    backgroundColor: color,
    borderWidth: 3,
    pointRadius: 3.5,
    pointHoverRadius: 6,
    tension: 0.28,
    yAxisID,
  };
}

function renderMainLine(peRows) {
  const ctx = document.getElementById("mainLineChart");
  charts.mainLine = new Chart(ctx, {
    type: "line",
    data: {
      labels: peRows.map((row) => row.Ano),
      datasets: [
        {
          ...lineDataset("% da população em extrema pobreza", peRows.map((row) => row.Percentual_Extrema_Pobreza), "#0e6f5c"),
          pointRadius: peRows.map((row) => (row.Ano === 2021 || row.Ano === 2025 ? 7 : 3.5)),
          pointBackgroundColor: peRows.map((row) => (row.Ano === 2021 ? "#b84d5a" : row.Ano === 2025 ? "#c9902e" : "#0e6f5c")),
        },
      ],
    },
    options: baseLineOptions("%", (value) => `${brPercent2.format(value)}%`),
  });
}

function renderSeries(peRows) {
  const ctx = document.getElementById("seriesChart");
  charts.series = new Chart(ctx, {
    type: "line",
    data: {
      labels: peRows.map((row) => row.Ano),
      datasets: [
        lineDataset("% da população", peRows.map((row) => row.Percentual_Extrema_Pobreza), "#0e6f5c", "y"),
        lineDataset("Pessoas", peRows.map((row) => row.Populacao_Extrema_Pobreza), "#265d83", "y1"),
      ],
    },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      interaction: { mode: "index", intersect: false },
      plugins: {
        legend: { position: "bottom" },
        tooltip: {
          callbacks: {
            label(context) {
              if (context.dataset.yAxisID === "y1") return `${context.dataset.label}: ${brNumber.format(context.raw)}`;
              return `${context.dataset.label}: ${brPercent2.format(context.raw)}%`;
            },
          },
        },
      },
      scales: {
        x: { grid: { display: false } },
        y: {
          title: { display: true, text: "Percentual" },
          ticks: { callback: (value) => `${value}%` },
          grid: { color: "#edf0ed" },
        },
        y1: {
          position: "right",
          title: { display: true, text: "Pessoas" },
          ticks: { callback: (value) => compact(value) },
          grid: { display: false },
        },
      },
    },
  });
}

function renderRanking() {
  const metric = document.getElementById("rankingMetric").value;
  const region = document.getElementById("regionFilter").value;
  const filtered = rankingRows
    .filter((row) => region === "Todas" || row.Regiao === region)
    .sort((a, b) => Number(b[metric]) - Number(a[metric]));

  const labels = filtered.map((row) => row.UF);
  const values = filtered.map((row) => Number(row[metric]));
  const peIndex = filtered.findIndex((row) => row.UF === "PE");
  const colors = filtered.map((row) => (row.UF === "PE" ? "#c9902e" : "#0e6f5c"));

  if (charts.ranking) charts.ranking.destroy();
  charts.ranking = new Chart(document.getElementById("rankingChart"), {
    type: "bar",
    data: {
      labels,
      datasets: [
        {
          label: metricLabel(metric),
          data: values,
          backgroundColor: colors,
          borderRadius: 4,
          barPercentage: 0.76,
          categoryPercentage: 0.78,
        },
      ],
    },
    options: {
      indexAxis: "y",
      responsive: true,
      maintainAspectRatio: false,
      plugins: {
        legend: { display: false },
        tooltip: {
          callbacks: {
            title(items) {
              const row = filtered[items[0].dataIndex];
              return `${row.Estado} (${row.UF})`;
            },
            label(context) {
              return `${metricLabel(metric)}: ${formatMetric(metric, context.raw)}`;
            },
          },
        },
      },
      scales: {
        x: {
          grid: { color: "#edf0ed" },
          ticks: { callback: (value) => formatAxisMetric(metric, value) },
        },
        y: { grid: { display: false } },
      },
    },
  });

  updateRankingNote(metric, filtered, peIndex);
}

function renderCriterion(key) {
  const criterion = peCriteria[key];
  if (charts.criterion) charts.criterion.destroy();
  charts.criterion = new Chart(document.getElementById("criterionChart"), {
    type: "line",
    data: {
      labels: years,
      datasets: [lineDataset("% em extrema pobreza", criterion.percent, key === "current" ? "#0e6f5c" : "#265d83")],
    },
    options: baseLineOptions("%", (value) => `${brPercent2.format(value)}%`),
  });

  const reduction = criterion.percent[10] - criterion.percent[13];
  document.getElementById("criterionTitle").textContent = criterion.title;
  document.getElementById("criterionText").textContent = criterion.note;
  document.getElementById("criterionStart").textContent = `${brPercent2.format(criterion.percent[10])}%`;
  document.getElementById("criterionEnd").textContent = `${brPercent2.format(criterion.percent[13])}%`;
  document.getElementById("criterionReduction").textContent = `-${brPercent2.format(reduction)} p.p.`;
}

function renderValidation() {
  charts.validation = new Chart(document.getElementById("validationChart"), {
    type: "line",
    data: {
      labels: validation.years,
      datasets: [
        lineDataset("Processado", validation.processed, "#0e6f5c"),
        lineDataset("IPECE", validation.ipece, "#c9902e"),
      ],
    },
    options: baseLineOptions("%", (value) => `${brPercent2.format(value)}%`),
  });
}

function baseLineOptions(unit, formatter) {
  return {
    responsive: true,
    maintainAspectRatio: false,
    interaction: { mode: "index", intersect: false },
    plugins: {
      legend: { position: "bottom" },
      tooltip: {
        callbacks: {
          label(context) {
            return `${context.dataset.label}: ${formatter(context.raw)}`;
          },
        },
      },
    },
    scales: {
      x: { grid: { display: false } },
      y: {
        ticks: { callback: (value) => `${value}${unit}` },
        grid: { color: "#edf0ed" },
      },
    },
  };
}

function updateSeriesStats(peRows) {
  const peak = peRows.reduce((max, row) =>
    row.Percentual_Extrema_Pobreza > max.Percentual_Extrema_Pobreza ? row : max
  );
  const lowest = peRows.reduce((min, row) =>
    row.Percentual_Extrema_Pobreza < min.Percentual_Extrema_Pobreza ? row : min
  );
  document.getElementById("peakYear").textContent = peak.Ano;
  document.getElementById("lowestYear").textContent = lowest.Ano;
  document.getElementById("seriesDrop").textContent = `-${brPercent2.format(
    peak.Percentual_Extrema_Pobreza - lowest.Percentual_Extrema_Pobreza
  )} p.p.`;
}

function populateRegions() {
  const select = document.getElementById("regionFilter");
  [...new Set(rankingRows.map((row) => row.Regiao))].sort().forEach((region) => {
    const option = document.createElement("option");
    option.value = region;
    option.textContent = region;
    select.appendChild(option);
  });
}

function wireControls() {
  document.getElementById("rankingMetric").addEventListener("change", renderRanking);
  document.getElementById("regionFilter").addEventListener("change", renderRanking);
  document.querySelectorAll("[data-criterion]").forEach((button) => {
    button.addEventListener("click", () => {
      document.querySelectorAll("[data-criterion]").forEach((item) => item.classList.remove("active"));
      button.classList.add("active");
      renderCriterion(button.dataset.criterion);
    });
  });
}

function updateRankingNote(metric, rows, peIndex) {
  const pe = rows[peIndex];
  const note = document.getElementById("rankingNote");
  if (!pe) {
    note.textContent = "Pernambuco não aparece no recorte regional selecionado.";
    return;
  }
  note.textContent = `No recorte selecionado, Pernambuco ocupa a ${peIndex + 1}ª posição em ${metricLabel(
    metric
  ).toLowerCase()}, com ${formatMetric(metric, pe[metric])} no período 2022-2025.`;
}

function metricLabel(metric) {
  return {
    Reducao_Pessoas: "Redução em pessoas",
    Reducao_pp: "Redução em pontos percentuais",
    Reducao_Percentual_Taxa: "Redução percentual da taxa",
    Reducao_Percentual_Pessoas: "Redução percentual de pessoas",
  }[metric];
}

function formatMetric(metric, value) {
  const number = Number(value);
  if (metric === "Reducao_Pessoas") return `${brNumber.format(number)} pessoas`;
  if (metric === "Reducao_pp") return `${brPercent2.format(number)} p.p.`;
  return `${brPercent.format(number)}%`;
}

function formatAxisMetric(metric, value) {
  if (metric === "Reducao_Pessoas") return compact(value);
  if (metric === "Reducao_pp") return `${value} p.p.`;
  return `${value}%`;
}

function compact(value) {
  return new Intl.NumberFormat("pt-BR", {
    notation: "compact",
    maximumFractionDigits: 1,
  }).format(value);
}
