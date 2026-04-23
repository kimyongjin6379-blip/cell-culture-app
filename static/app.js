/* ── State ───────────────────────────────────────────────────────────── */
let currentFileId = null;
let currentData   = null;
let activeTab     = null;

const COLORS = [
  "#1F77B4","#FF7F0E","#2CA02C","#D62728","#9467BD",
  "#8C564B","#E377C2","#7F7F7F","#BCBD22","#17BECF",
  "#AEC7E8","#FFBB78","#98DF8A","#FF9896","#C5B0D5",
];

/* ── DOM refs ────────────────────────────────────────────────────────── */
const dropZone     = document.getElementById("drop-zone");
const fileInput    = document.getElementById("file-input");
const fileNameEl   = document.getElementById("file-name");
const uploadForm   = document.getElementById("upload-form");
const submitBtn    = document.getElementById("submit-btn");
const btnText      = document.getElementById("btn-text");
const btnSpinner   = document.getElementById("btn-spinner");
const showMeanCb   = document.getElementById("show-mean");
const infoBox      = document.getElementById("info-box");
const downloadBtn  = document.getElementById("download-btn");
const placeholder  = document.getElementById("placeholder");
const tabBar       = document.getElementById("tab-bar");
const chartsContainer = document.getElementById("charts-container");

/* ── File drag & drop ────────────────────────────────────────────────── */
dropZone.addEventListener("click", () => fileInput.click());
dropZone.addEventListener("dragover", e => { e.preventDefault(); dropZone.classList.add("drag-over"); });
dropZone.addEventListener("dragleave", () => dropZone.classList.remove("drag-over"));
dropZone.addEventListener("drop", e => {
  e.preventDefault();
  dropZone.classList.remove("drag-over");
  const f = e.dataTransfer.files[0];
  if (f) selectFile(f);
});
fileInput.addEventListener("change", () => {
  if (fileInput.files[0]) selectFile(fileInput.files[0]);
});

function selectFile(f) {
  fileInput._selectedFile = f;
  fileNameEl.textContent = f.name;
  fileNameEl.classList.remove("hidden");
  submitBtn.disabled = false;
}

/* ── Form submit ─────────────────────────────────────────────────────── */
uploadForm.addEventListener("submit", async e => {
  e.preventDefault();
  const file = fileInput._selectedFile;
  if (!file) return;

  setLoading(true);

  const fd = new FormData();
  fd.append("file", file);

  try {
    const res = await fetch("/api/process", { method: "POST", body: fd });
    if (!res.ok) {
      const err = await res.json();
      alert("오류: " + (err.detail || "처리 실패"));
      return;
    }
    currentData = await res.json();
    currentFileId = currentData.file_id;
    renderAll(currentData);
  } catch (err) {
    alert("서버 오류: " + err.message);
  } finally {
    setLoading(false);
  }
});

/* ── Show-mean toggle re-renders active chart ────────────────────────── */
showMeanCb.addEventListener("change", () => {
  if (currentData && activeTab) renderChart(activeTab, currentData);
});

/* ── Download ────────────────────────────────────────────────────────── */
downloadBtn.addEventListener("click", () => {
  if (currentFileId) window.location.href = `/api/download/${currentFileId}`;
});

/* ── Render all ──────────────────────────────────────────────────────── */
function renderAll(data) {
  // Info box
  document.getElementById("info-cell-line").textContent   = data.cell_line   || "—";
  document.getElementById("info-culture-mode").textContent = data.culture_mode || "—";
  document.getElementById("info-feeding").textContent =
    data.feeding_days && data.feeding_days.length
      ? data.feeding_days.map(d => `Day ${d}`).join(", ")
      : "해당 없음";
  document.getElementById("info-sections").textContent =
    Object.keys(data.sections).length + "개";
  infoBox.classList.remove("hidden");
  downloadBtn.classList.remove("hidden");

  // Build tabs
  tabBar.innerHTML = "";
  chartsContainer.innerHTML = "";

  const secKeys = Object.keys(data.sections);
  secKeys.forEach((key, idx) => {
    const sec = data.sections[key];

    // Tab button
    const btn = document.createElement("button");
    btn.className = "tab-btn" + (idx === 0 ? " active" : "");
    btn.textContent = sec.label.split("(")[0].trim(); // short label
    btn.dataset.key = key;
    btn.addEventListener("click", () => switchTab(key, data));
    tabBar.appendChild(btn);

    // Chart panel
    const panel = document.createElement("div");
    panel.id = `panel-${key}`;
    panel.className = "chart-panel" + (idx === 0 ? "" : " hidden");
    const chartDiv = document.createElement("div");
    chartDiv.id = `chart-${key}`;
    chartDiv.className = "chart-div";
    panel.appendChild(chartDiv);
    chartsContainer.appendChild(panel);
  });

  placeholder.classList.add("hidden");
  tabBar.classList.remove("hidden");
  chartsContainer.classList.remove("hidden");

  activeTab = secKeys[0];
  if (activeTab) renderChart(activeTab, data);
}

function switchTab(key, data) {
  document.querySelectorAll(".tab-btn").forEach(b => b.classList.remove("active"));
  document.querySelector(`[data-key="${key}"]`).classList.add("active");

  document.querySelectorAll(".chart-panel").forEach(p => p.classList.add("hidden"));
  document.getElementById(`panel-${key}`).classList.remove("hidden");

  activeTab = key;
  renderChart(key, data);
}

/* ── Render one chart ────────────────────────────────────────────────── */
function renderChart(key, data) {
  const sec         = data.sections[key];
  const showMean    = showMeanCb.checked;
  const feedingDays = data.feeding_days || [];
  const isBar       = sec.chart_type === "bar";
  const treatments  = sec.treatments;
  const xLabels     = sec.x_labels;
  const days        = sec.days;

  const traces = [];
  let colorIdx = 0;

  for (const [name, stat] of Object.entries(treatments)) {
    const color = COLORS[colorIdx % COLORS.length];
    colorIdx++;

    if (showMean) {
      // ── Mean ± SD trace ──────────────────────────────────────────
      const meanTrace = {
        name,
        x: xLabels,
        y: stat.mean,
        type: isBar ? "bar" : "scatter",
        mode: isBar ? undefined : "lines+markers",
        marker: { color, size: 6 },
        line: { color, width: 2 },
        error_y: {
          type: "data",
          array: stat.std,
          visible: true,
          color: color,
          thickness: 1.5,
          width: 4,
        },
        legendgroup: name,
      };
      if (isBar) {
        meanTrace.error_y.color = "#555";
        meanTrace.marker = { color };
      }
      traces.push(meanTrace);
    } else {
      // ── Individual replicates ────────────────────────────────────
      stat.replicates.forEach((rep, ri) => {
        traces.push({
          name: `${name} (rep ${ri + 1})`,
          x: xLabels,
          y: rep,
          type: isBar ? "bar" : "scatter",
          mode: isBar ? undefined : "lines+markers",
          marker: { color, size: 5, opacity: 0.75 },
          line: { color, width: 1.5, dash: ri === 0 ? "solid" : "dash" },
          legendgroup: name,
          showlegend: ri === 0,
        });
      });
    }
  }

  // Feeding day shapes (only for line charts)
  const shapes = [];
  if (!isBar && feedingDays.length) {
    feedingDays.forEach(fd => {
      const lbl = `D${fd}`;
      if (xLabels.includes(lbl)) {
        shapes.push({
          type: "line",
          x0: lbl, x1: lbl,
          y0: 0, y1: 1,
          yref: "paper",
          line: { color: "#E74C3C", width: 1.5, dash: "dot" },
        });
      }
    });
  }

  const layout = {
    title: {
      text: `<b>${sec.label}</b>  <span style="font-size:12px;color:#666">(${sec.unit})</span>`,
      font: { family: "Arial", size: 16 },
      x: 0.04,
    },
    xaxis: {
      title: isBar ? "측정 구간" : "배양 일 (Day)",
      tickfont: { family: "Arial", size: 12 },
      gridcolor: "#E8E8E8",
    },
    yaxis: {
      title: sec.unit,
      tickfont: { family: "Arial", size: 12 },
      gridcolor: "#E8E8E8",
      rangemode: "tozero",
    },
    legend: {
      orientation: "h",
      x: 0, y: -0.18,
      font: { family: "Arial", size: 11 },
    },
    shapes,
    plot_bgcolor:  "#FAFAFA",
    paper_bgcolor: "#FFFFFF",
    margin: { t: 60, b: 80, l: 70, r: 30 },
    barmode: "group",
    hovermode: "x unified",
    font: { family: "Arial" },
  };

  // Feeding day annotations
  if (!isBar && feedingDays.length) {
    layout.annotations = feedingDays
      .filter(fd => xLabels.includes(`D${fd}`))
      .map(fd => ({
        x: `D${fd}`, y: 1, yref: "paper",
        text: `Feed D${fd}`, showarrow: false,
        font: { size: 10, color: "#E74C3C" },
        xanchor: "center", yanchor: "bottom",
      }));
  }

  const config = {
    responsive: true,
    displayModeBar: true,
    modeBarButtonsToRemove: ["sendDataToCloud", "select2d", "lasso2d"],
    toImageButtonOptions: { format: "png", filename: `${key}_chart`, scale: 2 },
  };

  Plotly.react(`chart-${key}`, traces, layout, config);
}

/* ── Utilities ───────────────────────────────────────────────────────── */
function setLoading(on) {
  submitBtn.disabled = on;
  btnText.textContent = on ? "분석 중..." : "분석 시작";
  btnSpinner.classList.toggle("hidden", !on);
}
