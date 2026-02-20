/*
  Glucose Log + Predictor (Offline-First)
  ======================================
  - Local-first glucose logging (add/edit/delete/undo)
  - Summary KPIs (TIR, avg, median, std dev, fasting avg)
  - Trend chart from logged readings
  - Food sugar prediction curve (explainable forward model)
  - Synthetic test set (~90 days, ~500 points) generator
  - Backtest: baseline vs improved alert logic (FP reduction)

  SAFETY / SCOPE:
  - Educational decision-support only; not medical advice.
  - Do not dose insulin or change treatment based on this tool.
*/

"use strict";

/* ---------------------------
   Storage schema + versioning
---------------------------- */

const SCHEMA_VERSION = 3;
const STORAGE_KEY = "glucose_log_records_v3";
const META_KEY = "glucose_log_meta_v3";

/* ---------------------------
   DOM helper
---------------------------- */
const el = (id) => document.getElementById(id);

/* ---------------------------
   Meta + Records persistence
---------------------------- */

function loadMeta() {
  try {
    const meta = JSON.parse(localStorage.getItem(META_KEY) || "null");
    return meta || { schemaVersion: SCHEMA_VERSION };
  } catch {
    return { schemaVersion: SCHEMA_VERSION };
  }
}

function saveMeta(meta) {
  localStorage.setItem(META_KEY, JSON.stringify(meta));
}

function loadRecords() {
  try {
    const arr = JSON.parse(localStorage.getItem(STORAGE_KEY) || "[]");
    return Array.isArray(arr) ? arr : [];
  } catch {
    return [];
  }
}

function saveRecords(records) {
  localStorage.setItem(STORAGE_KEY, JSON.stringify(records));
}

function migrateIfNeeded() {
  const meta = loadMeta();
  if (!meta.schemaVersion) meta.schemaVersion = SCHEMA_VERSION;
  // Future migrations go here (kept simple for now)
  meta.schemaVersion = SCHEMA_VERSION;
  saveMeta(meta);
}

/* ---------------------------
   Small utilities
---------------------------- */

function uid() {
  return Math.random().toString(16).slice(2) + Date.now().toString(16);
}

function pad2(n) {
  return String(n).padStart(2, "0");
}

function toLocalInputValue(date) {
  // Date -> "YYYY-MM-DDTHH:mm"
  return `${date.getFullYear()}-${pad2(date.getMonth() + 1)}-${pad2(date.getDate())}T${pad2(date.getHours())}:${pad2(date.getMinutes())}`;
}

function formatLocal(iso) {
  return new Date(iso).toLocaleString();
}

function clampNumber(n, min, max) {
  if (Number.isNaN(n)) return null;
  return Math.max(min, Math.min(max, n));
}

function escapeHtml(s) {
  return String(s || "")
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;");
}

function withinDays(dtIso, days) {
  if (days >= 99999) return true;
  const t = new Date(dtIso).getTime();
  const cutoff = Date.now() - days * 24 * 60 * 60 * 1000;
  return t >= cutoff;
}

/* ---------------------------
   Stats (KPIs)
---------------------------- */

function mean(values) {
  if (!values.length) return null;
  return values.reduce((a, b) => a + b, 0) / values.length;
}

function median(values) {
  if (!values.length) return null;
  const sorted = [...values].sort((a, b) => a - b);
  const mid = Math.floor(sorted.length / 2);
  return (sorted.length % 2 === 0) ? (sorted[mid - 1] + sorted[mid]) / 2 : sorted[mid];
}

function stdDev(values) {
  if (values.length < 2) return null;
  const m = mean(values);
  const variance = values.reduce((acc, v) => acc + (v - m) ** 2, 0) / values.length;
  return Math.sqrt(variance);
}

function statusFor(value, low, high) {
  if (value < low) return { label: "LOW", cls: "badge--low" };
  if (value > high) return { label: "HIGH", cls: "badge--high" };
  return { label: "IN RANGE", cls: "badge--in" };
}

function computeSummary(records, low, high) {
  const values = records.map(r => r.value);
  const total = values.length;

  const inRange = records.filter(r => r.value >= low && r.value <= high).length;
  const lowCount = records.filter(r => r.value < low).length;
  const highCount = records.filter(r => r.value > high).length;

  const tirPct = total ? Math.round((inRange / total) * 100) : null;

  const fastingVals = records.filter(r => r.context === "Fasting").map(r => r.value);

  return {
    total,
    tirPct,
    inRange,
    lowCount,
    highCount,
    avg: mean(values),
    med: median(values),
    sd: stdDev(values),
    fastingAvg: mean(fastingVals),
    fastingCount: fastingVals.length
  };
}

/* ---------------------------
   Canvas charts (no libraries)
---------------------------- */

function drawSeriesChart(canvas, records, low, high) {
  const ctx = canvas.getContext("2d");
  ctx.clearRect(0, 0, canvas.width, canvas.height);

  const pts = records
    .slice()
    .sort((a, b) => new Date(a.dtIso) - new Date(b.dtIso))
    .slice(-60);

  if (pts.length < 2) {
    ctx.font = "16px system-ui";
    ctx.fillText("Add at least 2 readings to see a trend.", 16, 40);
    return;
  }

  const values = pts.map(r => r.value);
  const minV = Math.min(...values, low - 20);
  const maxV = Math.max(...values, high + 20);

  const padL = 56, padR = 16, padT = 16, padB = 28;
  const W = canvas.width - padL - padR;
  const H = canvas.height - padT - padB;

  // Axes
  ctx.globalAlpha = 0.35;
  ctx.strokeStyle = "#888";
  ctx.beginPath();
  ctx.moveTo(padL, padT);
  ctx.lineTo(padL, padT + H);
  ctx.lineTo(padL + W, padT + H);
  ctx.stroke();
  ctx.globalAlpha = 1;

  // Target band
  const yHigh = padT + (1 - (high - minV) / (maxV - minV || 1)) * H;
  const yLow = padT + (1 - (low - minV) / (maxV - minV || 1)) * H;
  ctx.fillStyle = "rgba(17,115,75,0.10)";
  ctx.fillRect(padL, yHigh, W, (yLow - yHigh));

  // Labels
  ctx.fillStyle = "#777";
  ctx.font = "12px system-ui";
  ctx.fillText(String(Math.round(maxV)), 10, padT + 12);
  ctx.fillText(String(Math.round(minV)), 10, padT + H);

  // Line
  ctx.strokeStyle = "#111";
  ctx.lineWidth = 2;
  ctx.beginPath();
  pts.forEach((r, i) => {
    const x = padL + (i / (pts.length - 1)) * W;
    const y = padT + (1 - (r.value - minV) / (maxV - minV || 1)) * H;
    if (i === 0) ctx.moveTo(x, y);
    else ctx.lineTo(x, y);
  });
  ctx.stroke();

  // Points
  ctx.fillStyle = "#111";
  pts.forEach((r, i) => {
    const x = padL + (i / (pts.length - 1)) * W;
    const y = padT + (1 - (r.value - minV) / (maxV - minV || 1)) * H;
    ctx.beginPath();
    ctx.arc(x, y, 3, 0, Math.PI * 2);
    ctx.fill();
  });
}

function drawPredictionChart(canvas, curvePoints, low, high) {
  const ctx = canvas.getContext("2d");
  ctx.clearRect(0, 0, canvas.width, canvas.height);

  if (!curvePoints || curvePoints.length < 2) {
    ctx.font = "16px system-ui";
    ctx.fillText("Run a prediction to see the curve.", 16, 40);
    return;
  }

  const values = curvePoints.map(p => p.bg);
  const minV = Math.min(...values, low - 20);
  const maxV = Math.max(...values, high + 20);

  const padL = 56, padR = 16, padT = 16, padB = 28;
  const W = canvas.width - padL - padR;
  const H = canvas.height - padT - padB;

  // Axes
  ctx.globalAlpha = 0.35;
  ctx.strokeStyle = "#888";
  ctx.beginPath();
  ctx.moveTo(padL, padT);
  ctx.lineTo(padL, padT + H);
  ctx.lineTo(padL + W, padT + H);
  ctx.stroke();
  ctx.globalAlpha = 1;

  // Target band
  const yHigh = padT + (1 - (high - minV) / (maxV - minV || 1)) * H;
  const yLow = padT + (1 - (low - minV) / (maxV - minV || 1)) * H;
  ctx.fillStyle = "rgba(17,115,75,0.10)";
  ctx.fillRect(padL, yHigh, W, (yLow - yHigh));

  // Labels
  ctx.fillStyle = "#777";
  ctx.font = "12px system-ui";
  ctx.fillText(String(Math.round(maxV)), 10, padT + 12);
  ctx.fillText(String(Math.round(minV)), 10, padT + H);

  const tMax = curvePoints[curvePoints.length - 1].tMin || 1;
  ctx.fillText("0m", padL, padT + H + 18);
  ctx.fillText(`${tMax}m`, padL + W - 24, padT + H + 18);

  // Line
  ctx.strokeStyle = "#111";
  ctx.lineWidth = 2;
  ctx.beginPath();
  curvePoints.forEach((p, i) => {
    const x = padL + (p.tMin / tMax) * W;
    const y = padT + (1 - (p.bg - minV) / (maxV - minV || 1)) * H;
    if (i === 0) ctx.moveTo(x, y);
    else ctx.lineTo(x, y);
  });
  ctx.stroke();

  // Points
  ctx.fillStyle = "#111";
  curvePoints.forEach((p) => {
    const x = padL + (p.tMin / tMax) * W;
    const y = padT + (1 - (p.bg - minV) / (maxV - minV || 1)) * H;
    ctx.beginPath();
    ctx.arc(x, y, 2.5, 0, Math.PI * 2);
    ctx.fill();
  });
}

/* ---------------------------
   Food sugar prediction model
---------------------------- */

/**
 * Explainable “absorption curve” forward model:
 * - peakRise = gramsSugar * mgdlPerGram * activityFactor
 * - delta(t) = peakRise * (t/peakMin) * exp(1 - t/peakMin)
 *   (gamma-like curve: rises to peak at t=peakMin then decays)
 */
function predictSugarCurve(currentBg, gramsSugar, mgdlPerGram, peakMin, horizonMin, activityFactor) {
  const peakRise = gramsSugar * mgdlPerGram * activityFactor;
  const step = 10; // minutes per point
  const pts = [];

  for (let t = 0; t <= horizonMin; t += step) {
    const x = t / Math.max(1, peakMin);
    const shape = x * Math.exp(1 - x); // peaks at x=1
    const delta = peakRise * shape;
    pts.push({ tMin: t, bg: currentBg + delta });
  }
  return pts;
}

/* ---------------------------
   Synthetic test set generator
---------------------------- */

function generateSyntheticSeries(days, approxPoints) {
  const totalMinutes = days * 24 * 60;
  const stepMin = Math.max(5, Math.round(totalMinutes / approxPoints));

  const base = 115;
  const noiseSd = 8;
  const dailyAmp = 12;

  const mealsPerDay = 3;
  const mgdlPerGram = 1.4;
  const peakMin = 45;

  // Box-Muller gaussian noise
  const randn = () => {
    const u = Math.max(1e-9, Math.random());
    const v = Math.max(1e-9, Math.random());
    return Math.sqrt(-2 * Math.log(u)) * Math.cos(2 * Math.PI * v);
  };

  // Meal schedule
  const mealTimes = [];
  for (let d = 0; d < days; d++) {
    const dayStart = d * 24 * 60;
    mealTimes.push(dayStart + 8 * 60 + Math.round((Math.random() - 0.5) * 45));
    mealTimes.push(dayStart + 13 * 60 + Math.round((Math.random() - 0.5) * 60));
    mealTimes.push(dayStart + 19 * 60 + Math.round((Math.random() - 0.5) * 60));
  }
  const mealSugar = mealTimes.map(() => 20 + Math.round(Math.random() * 60)); // 20–80g

  function mealDeltaAtMinute(minute) {
    let total = 0;
    for (let i = 0; i < mealTimes.length; i++) {
      const dt = minute - mealTimes[i];
      if (dt < 0) continue;
      const x = dt / Math.max(1, peakMin);
      const shape = x * Math.exp(1 - x);
      total += (mealSugar[i] * mgdlPerGram) * shape;
    }
    return total;
  }

  // Random hypo/hyper events
  const hypoEvents = [];
  const hyperEvents = [];
  const numHypo = Math.max(1, Math.round(days / 18));
  const numHyper = Math.max(1, Math.round(days / 12));

  for (let i = 0; i < numHypo; i++) hypoEvents.push(Math.round(Math.random() * totalMinutes));
  for (let i = 0; i < numHyper; i++) hyperEvents.push(Math.round(Math.random() * totalMinutes));

  function eventDeltaAtMinute(minute) {
    let delta = 0;
    for (const t0 of hypoEvents) {
      const dt = minute - t0;
      const sigma = 35;
      delta += -35 * Math.exp(-(dt * dt) / (2 * sigma * sigma));
    }
    for (const t0 of hyperEvents) {
      const dt = minute - t0;
      const sigma = 50;
      delta += 55 * Math.exp(-(dt * dt) / (2 * sigma * sigma));
    }
    return delta;
  }

  const startMs = Date.now() - days * 24 * 60 * 60 * 1000;
  const series = [];

  for (let m = 0; m <= totalMinutes; m += stepMin) {
    const tMs = startMs + m * 60 * 1000;
    const dayPhase = (m % (24 * 60)) / (24 * 60) * 2 * Math.PI;
    const daily = dailyAmp * Math.sin(dayPhase - 0.8);
    const meal = mealDeltaAtMinute(m);
    const events = eventDeltaAtMinute(m);
    const noise = noiseSd * randn();
    const bg = base + daily + meal + events + noise;
    series.push({ tMs, bg: Math.max(40, Math.min(400, bg)) });
  }

  return series;
}

/* ---------------------------
   Backtest: baseline vs improved alert logic
---------------------------- */

/**
 * We define a “true event” in the future horizon:
 * - HYPO true if any future point < low
 * - HYPER true if any future point > high
 *
 * Baseline rule:
 * - Alert HYPO if current < low OR projected (linear trend) crosses low in horizon
 * - Alert HYPER if current > high OR projected crosses high in horizon
 *
 * Improved rule:
 * - Baseline + smoothing + persistence + cooldown + hysteresis
 * - This usually reduces noisy false alerts.
 */
function backtest(series, low, high, horizonMin) {
  if (!series || series.length < 20) return null;

  const stepMin = (series[1].tMs - series[0].tMs) / 60000;
  const horizonSteps = Math.max(1, Math.round(horizonMin / stepMin));

  // Quick slope estimate over last k points using endpoints
  function slopePerMin(i, k) {
    const start = Math.max(0, i - k + 1);
    const t0 = series[start].tMs;
    const t1 = series[i].tMs;
    const dt = (t1 - t0) / 60000;
    if (dt <= 0) return 0;
    return (series[i].bg - series[start].bg) / dt;
  }

  function projectedBg(i) {
    const m = slopePerMin(i, 6);
    return series[i].bg + m * horizonMin;
  }

  // Baseline counts
  let baseTP = 0, baseFP = 0, baseFN = 0;

  // Improved counts
  let impTP = 0, impFP = 0, impFN = 0;

  // Improved state (persistence + cooldown + hysteresis)
  let active = "NONE";
  let persist = 0;
  let cooldownUntilMs = 0;

  // Hysteresis thresholds
  const hypoExit = low + 5;
  const hyperExit = high - 10;

  for (let i = 10; i < series.length - horizonSteps; i++) {
    const now = series[i];

    // Future truth labels
    const future = series.slice(i + 1, i + 1 + horizonSteps);
    const trueHypo = future.some(p => p.bg < low);
    const trueHyper = future.some(p => p.bg > high);

    // ---------- Baseline decision ----------
    const proj = projectedBg(i);
    const baseHypoAlert = (now.bg < low) || (proj < low);
    const baseHyperAlert = (now.bg > high) || (proj > high);

    // Evaluate baseline (count hypo+hyper alerts as “an alert”)
    const baseAlert = baseHypoAlert || baseHyperAlert;
    const trueEvent = trueHypo || trueHyper;

    if (baseAlert && trueEvent) baseTP++;
    if (baseAlert && !trueEvent) baseFP++;
    if (!baseAlert && trueEvent) baseFN++;

    // ---------- Improved decision ----------
    // Smooth by averaging last 3 points (very lightweight smoothing)
    const s1 = series[i].bg;
    const s2 = series[i - 1].bg;
    const s3 = series[i - 2].bg;
    const smoothBg = (s1 + s2 + s3) / 3;

    // Use smoothed value for “current”
    const m = slopePerMin(i, 6);
    const smoothProj = smoothBg + m * horizonMin;

    const rawHypo = (smoothBg < low) || (smoothProj < low);
    const rawHyper = (smoothBg > high) || (smoothProj > high);
    const rawRisk = rawHypo ? "HYPO" : (rawHyper ? "HYPER" : "NONE");

    // Persistence: risk must repeat twice before alerting
    if (rawRisk !== "NONE") {
      persist = (rawRisk === active) ? (persist + 1) : (persist + 1);
      // Above is intentionally simple; we’ll enforce “2 consecutive risks” via a buffer:
      // Better: track lastRisk; here we just require persist>=2 while risk present
    } else {
      persist = 0;
    }

    // Cooldown
    const inCooldown = now.tMs < cooldownUntilMs;

    // Trigger improved alert if not in cooldown and we have persistent risk
    let impAlert = false;
    if (!inCooldown && rawRisk !== "NONE" && persist >= 2) {
      active = rawRisk;
      impAlert = true;
      cooldownUntilMs = now.tMs + 60 * 60000; // 60 min cooldown
    }

    // Clear active via hysteresis (optional display logic)
    if (active === "HYPO" && smoothProj > hypoExit) active = "NONE";
    if (active === "HYPER" && smoothProj < hyperExit) active = "NONE";

    const impTrueEvent = trueEvent;
    if (impAlert && impTrueEvent) impTP++;
    if (impAlert && !impTrueEvent) impFP++;
    if (!impAlert && impTrueEvent) impFN++;
  }

  function precision(tp, fp) {
    const denom = tp + fp;
    return denom ? tp / denom : 0;
  }
  function recall(tp, fn) {
    const denom = tp + fn;
    return denom ? tp / denom : 0;
  }

  const fpReduction = baseFP ? ((baseFP - impFP) / baseFP) * 100 : 0;

  return {
    baseline: { TP: baseTP, FP: baseFP, FN: baseFN, precision: precision(baseTP, baseFP), recall: recall(baseTP, baseFN) },
    improved: { TP: impTP, FP: impFP, FN: impFN, precision: precision(impTP, impFP), recall: recall(impTP, impFN) },
    fpReductionPct: fpReduction
  };
}

/* ---------------------------
   UI bindings
---------------------------- */

const ui = {
  patientName: el("patientName"),
  glucose: el("glucose"),
  context: el("context"),
  dt: el("dt"),
  notes: el("notes"),
  editingId: el("editingId"),

  saveBtn: el("saveBtn"),
  cancelEditBtn: el("cancelEditBtn"),
  exportBtn: el("exportBtn"),
  reportBtn: el("reportBtn"),
  clearBtn: el("clearBtn"),

  windowDays: el("windowDays"),
  targetLow: el("targetLow"),
  targetHigh: el("targetHigh"),
  contextFilter: el("contextFilter"),
  sortOrder: el("sortOrder"),

  modePill: el("modePill"),
  undoPill: el("undoPill"),
  undoBtn: el("undoBtn"),

  kpiTir: el("kpiTir"),
  kpiTirSub: el("kpiTirSub"),
  kpiAvg: el("kpiAvg"),
  kpiAvgSub: el("kpiAvgSub"),
  kpiStd: el("kpiStd"),
  kpiStdSub: el("kpiStdSub"),
  kpiMedian: el("kpiMedian"),
  kpiMedianSub: el("kpiMedianSub"),
  kpiFastingAvg: el("kpiFastingAvg"),
  kpiFastingSub: el("kpiFastingSub"),
  kpiEvents: el("kpiEvents"),
  kpiEventsSub: el("kpiEventsSub"),

  mainChart: el("chart"),
  predChart: el("predChart"),

  // Prediction panel
  predCurrent: el("predCurrent"),
  predSugar: el("predSugar"),
  predHorizon: el("predHorizon"),
  mgdlPerGram: el("mgdlPerGram"),
  absorbMin: el("absorbMin"),
  activityFactor: el("activityFactor"),
  predictBtn: el("predictBtn"),
  useLastBtn: el("useLastBtn"),
  predText: el("predText"),

  // Synthetic / backtest
  synDays: el("synDays"),
  synPoints: el("synPoints"),
  btHorizon: el("btHorizon"),
  genTestBtn: el("genTestBtn"),
  runBacktestBtn: el("runBacktestBtn"),
  backtestOut: el("backtestOut"),

  tbody: el("tbody"),
};

let lastDeleted = null;
let syntheticSeries = null;

/* ---------------------------
   Preferences + render pipeline
---------------------------- */

function getPrefs() {
  const low = clampNumber(Number(ui.targetLow.value), 40, 300) ?? 70;
  const high = clampNumber(Number(ui.targetHigh.value), 40, 400) ?? 180;
  const days = Number(ui.windowDays.value);
  const ctxFilter = ui.contextFilter.value;
  const sortOrder = ui.sortOrder.value;
  const patient = (ui.patientName.value || "").trim();

  const fixedLow = Math.min(low, high - 1);
  const fixedHigh = Math.max(high, fixedLow + 1);

  return { low: fixedLow, high: fixedHigh, days, ctxFilter, sortOrder, patient };
}

function filterRecords(records, prefs) {
  return records
    .filter(r => withinDays(r.dtIso, prefs.days))
    .filter(r => prefs.ctxFilter === "ALL" ? true : r.context === prefs.ctxFilter);
}

function sortRecords(records, sortOrder) {
  const dir = (sortOrder === "ASC") ? 1 : -1;
  return records.slice().sort((a, b) => (new Date(a.dtIso) - new Date(b.dtIso)) * dir);
}

function renderKPIs(summary, prefs) {
  if (!summary.total) {
    ui.kpiTir.textContent = "—";
    ui.kpiTirSub.textContent = "No readings in selected window";
    ui.kpiAvg.textContent = "—";
    ui.kpiAvgSub.textContent = "—";
    ui.kpiStd.textContent = "—";
    ui.kpiStdSub.textContent = "—";
    ui.kpiMedian.textContent = "—";
    ui.kpiMedianSub.textContent = "—";
    ui.kpiFastingAvg.textContent = "—";
    ui.kpiFastingSub.textContent = "—";
    ui.kpiEvents.textContent = "—";
    ui.kpiEventsSub.textContent = `Target ${prefs.low}–${prefs.high}`;
    return;
  }

  ui.kpiTir.textContent = `${summary.tirPct}%`;
  ui.kpiTirSub.textContent = `${summary.inRange}/${summary.total} in range • ${summary.lowCount} low • ${summary.highCount} high`;

  ui.kpiAvg.textContent = summary.avg ? summary.avg.toFixed(1) : "—";
  ui.kpiAvgSub.textContent = `mg/dL • ${prefs.days >= 99999 ? "all-time" : `${prefs.days}d`}`;

  ui.kpiStd.textContent = summary.sd ? summary.sd.toFixed(1) : "—";
  ui.kpiStdSub.textContent = "std dev";

  ui.kpiMedian.textContent = summary.med ? summary.med.toFixed(1) : "—";
  ui.kpiMedianSub.textContent = "median";

  ui.kpiFastingAvg.textContent = summary.fastingAvg ? summary.fastingAvg.toFixed(1) : "—";
  ui.kpiFastingSub.textContent = summary.fastingCount ? `${summary.fastingCount} fasting` : "No fasting readings";

  ui.kpiEvents.textContent = `${summary.lowCount + summary.highCount}`;
  ui.kpiEventsSub.textContent = `Target ${prefs.low}–${prefs.high}`;
}

function renderTable(records, prefs) {
  ui.tbody.innerHTML = "";

  records.forEach(r => {
    const st = statusFor(r.value, prefs.low, prefs.high);

    const tr = document.createElement("tr");
    tr.className = "clickRow";
    tr.dataset.editId = r.id;

    tr.innerHTML = `
      <td>${escapeHtml(formatLocal(r.dtIso))}</td>
      <td><strong>${r.value}</strong> <span class="muted">mg/dL</span></td>
      <td><span class="badge ${st.cls}">${st.label}</span></td>
      <td>${escapeHtml(r.context)}</td>
      <td class="muted">${escapeHtml(r.notes)}</td>
      <td><button class="btn" data-del="${r.id}" type="button">Delete</button></td>
    `;

    ui.tbody.appendChild(tr);
  });
}

function render() {
  const prefs = getPrefs();

  // Persist prefs to meta (polished UX)
  const meta = loadMeta();
  meta.schemaVersion = SCHEMA_VERSION;
  meta.patientName = prefs.patient;
  meta.targetLow = prefs.low;
  meta.targetHigh = prefs.high;
  meta.windowDays = prefs.days;
  meta.contextFilter = prefs.ctxFilter;
  meta.sortOrder = prefs.sortOrder;
  saveMeta(meta);

  const all = loadRecords();
  const filtered = filterRecords(all, prefs);
  const summary = computeSummary(filtered, prefs.low, prefs.high);

  renderKPIs(summary, prefs);
  drawSeriesChart(ui.mainChart, filtered, prefs.low, prefs.high);

  const sortedForTable = sortRecords(all, prefs.sortOrder);
  renderTable(sortedForTable, prefs);

  ui.undoBtn.disabled = !lastDeleted;
  ui.undoPill.hidden = !lastDeleted;
}

/* ---------------------------
   Form mode: Add vs Edit
---------------------------- */

function setModeAdd() {
  ui.editingId.value = "";
  ui.saveBtn.textContent = "Add";
  ui.cancelEditBtn.disabled = true;
  ui.modePill.textContent = "Mode: Add";
}

function setModeEdit(record) {
  ui.editingId.value = record.id;
  ui.glucose.value = record.value;
  ui.context.value = record.context;
  ui.dt.value = toLocalInputValue(new Date(record.dtIso));
  ui.notes.value = record.notes || "";

  ui.saveBtn.textContent = "Save Changes";
  ui.cancelEditBtn.disabled = false;
  ui.modePill.textContent = "Mode: Edit";
}

function readForm() {
  const value = clampNumber(Number(ui.glucose.value), 20, 600);
  const context = ui.context.value;
  const dtLocal = ui.dt.value;
  const notes = (ui.notes.value || "").trim();

  if (!value) return { ok: false, msg: "Enter glucose 20–600." };
  if (!dtLocal) return { ok: false, msg: "Pick a date/time." };

  // Outlier confirm (safety-minded UX)
  if (value < 55 || value > 350) {
    const ok = confirm("This value is outside typical ranges. Confirm it’s correct?");
    if (!ok) return { ok: false, msg: "Cancelled." };
  }

  const dtIso = new Date(dtLocal).toISOString();
  return { ok: true, record: { value, context, dtIso, notes } };
}

/* ---------------------------
   CRUD actions
---------------------------- */

function addRecord(record) {
  const records = loadRecords();
  records.push({ id: uid(), ...record });
  saveRecords(records);
}

function updateRecord(id, patch) {
  const records = loadRecords();
  const idx = records.findIndex(r => r.id === id);
  if (idx === -1) return;
  records[idx] = { ...records[idx], ...patch };
  saveRecords(records);
}

function deleteRecord(id) {
  const records = loadRecords();
  const idx = records.findIndex(r => r.id === id);
  if (idx === -1) return;
  lastDeleted = records[idx];
  records.splice(idx, 1);
  saveRecords(records);
}

function undoDelete() {
  if (!lastDeleted) return;
  const records = loadRecords();
  records.push(lastDeleted);
  saveRecords(records);
  lastDeleted = null;
}

/* ---------------------------
   Export / Report
---------------------------- */

function exportCSV() {
  const prefs = getPrefs();
  const records = loadRecords().slice().sort((a, b) => new Date(a.dtIso) - new Date(b.dtIso));

  const headerLines = [
    "# Glucose Log Export",
    `# exported_at=${new Date().toISOString()}`,
    `# patient=${prefs.patient || "N/A"}`,
    `# target_range_mgdl=${prefs.low}-${prefs.high}`,
    `# notes=manual entry; not medical advice`,
    "datetime_iso,datetime_local,value_mgdl,context,notes"
  ];

  const rows = records.map(r => {
    const local = new Date(r.dtIso).toLocaleString();
    const safeNotes = String(r.notes || "").replaceAll('"', '""');
    return `${r.dtIso},"${local}",${r.value},"${r.context}","${safeNotes}"`;
  });

  const csv = headerLines.join("\n") + "\n" + rows.join("\n");
  const blob = new Blob([csv], { type: "text/csv;charset=utf-8" });
  const url = URL.createObjectURL(blob);

  const a = document.createElement("a");
  a.href = url;
  a.download = "glucose_log.csv";
  a.click();
  URL.revokeObjectURL(url);
}

function printReport() {
  const prefs = getPrefs();
  const all = loadRecords();
  const filtered = filterRecords(all, prefs).sort((a, b) => new Date(a.dtIso) - new Date(b.dtIso));
  const summary = computeSummary(filtered, prefs.low, prefs.high);

  const meanTxt = summary.avg ? summary.avg.toFixed(1) : "—";
  const medTxt = summary.med ? summary.med.toFixed(1) : "—";
  const sdTxt = summary.sd ? summary.sd.toFixed(1) : "—";
  const tirTxt = summary.tirPct != null ? `${summary.tirPct}%` : "—";

  const html = `
  <html>
    <head>
      <title>Glucose Summary Report</title>
      <meta charset="utf-8" />
      <style>
        body { font-family: system-ui, -apple-system, Segoe UI, Roboto, Arial; margin: 24px; }
        h1 { margin: 0 0 6px; }
        .muted { color: #555; }
        .grid { display: grid; grid-template-columns: 1fr 1fr; gap: 12px; margin: 14px 0; }
        .card { border: 1px solid #ddd; border-radius: 12px; padding: 12px; }
        .big { font-size: 28px; margin-top: 6px; }
        table { width: 100%; border-collapse: collapse; margin-top: 10px; }
        th, td { text-align:left; border-bottom: 1px solid #eee; padding: 8px; font-size: 12px; vertical-align: top; }
        th { color: #555; }
        @media print { button { display:none; } }
      </style>
    </head>
    <body>
      <h1>Glucose Summary Report</h1>
      <div class="muted">
        Generated: ${new Date().toLocaleString()} • Patient: ${escapeHtml(prefs.patient || "N/A")}<br/>
        Window: ${prefs.days >= 99999 ? "All time" : `Last ${prefs.days} days`} • Context: ${escapeHtml(prefs.ctxFilter)} • Target: ${prefs.low}–${prefs.high} mg/dL
      </div>

      <div class="grid">
        <div class="card">
          <div class="muted">Time in Range</div>
          <div class="big">${tirTxt}</div>
          <div class="muted">${summary.inRange || 0}/${summary.total || 0} in range • ${summary.lowCount || 0} low • ${summary.highCount || 0} high</div>
        </div>
        <div class="card">
          <div class="muted">Average / Median / Std Dev</div>
          <div class="big">${meanTxt}</div>
          <div class="muted">avg mg/dL • median ${medTxt} • std dev ${sdTxt}</div>
        </div>
        <div class="card">
          <div class="muted">Fasting</div>
          <div class="big">${summary.fastingAvg ? summary.fastingAvg.toFixed(1) : "—"}</div>
          <div class="muted">${summary.fastingCount || 0} fasting readings</div>
        </div>
        <div class="card">
          <div class="muted">Events</div>
          <div class="big">${(summary.lowCount || 0) + (summary.highCount || 0)}</div>
          <div class="muted">low + high events</div>
        </div>
      </div>

      <h2 style="margin:18px 0 6px;">Readings</h2>
      <table>
        <thead><tr><th>Date</th><th>Value</th><th>Context</th><th>Notes</th></tr></thead>
        <tbody>
          ${filtered.map(r => `
            <tr>
              <td>${escapeHtml(new Date(r.dtIso).toLocaleString())}</td>
              <td>${r.value} mg/dL</td>
              <td>${escapeHtml(r.context)}</td>
              <td>${escapeHtml(r.notes)}</td>
            </tr>
          `).join("")}
        </tbody>
      </table>

      <p class="muted" style="margin-top:14px;">
        Not medical advice. Data source: manual entry. Share with a clinician for interpretation.
      </p>

      <button onclick="window.print()">Print / Save as PDF</button>
    </body>
  </html>`;

  const w = window.open("", "_blank");
  w.document.open();
  w.document.write(html);
  w.document.close();
}

/* ---------------------------
   Prediction panel handlers
---------------------------- */

function runPrediction() {
  const prefs = getPrefs();

  const current = clampNumber(Number(ui.predCurrent.value), 40, 600);
  const grams = clampNumber(Number(ui.predSugar.value), 0, 300);
  const horizon = Number(ui.predHorizon.value);
  const mgdlPerGram = clampNumber(Number(ui.mgdlPerGram.value), 0, 10) ?? 1.6;
  const peakMin = clampNumber(Number(ui.absorbMin.value), 15, 180) ?? 45;
  const activityFactor = Number(ui.activityFactor.value) || 1.0;

  if (current == null) {
    ui.predText.textContent = "Enter current glucose.";
    drawPredictionChart(ui.predChart, null, prefs.low, prefs.high);
    return;
  }
  if (grams == null) {
    ui.predText.textContent = "Enter sugar grams.";
    drawPredictionChart(ui.predChart, null, prefs.low, prefs.high);
    return;
  }

  const curve = predictSugarCurve(current, grams, mgdlPerGram, peakMin, horizon, activityFactor);

  // Summarize key values
  const peak = curve.reduce((best, p) => (p.bg > best.bg ? p : best), curve[0]);
  const end = curve[curve.length - 1];

  ui.predText.textContent =
    `Peak ≈ ${peak.bg.toFixed(0)} mg/dL at ~${peak.tMin} min • End ≈ ${end.bg.toFixed(0)} mg/dL at ${horizon} min (educational estimate)`;

  drawPredictionChart(ui.predChart, curve, prefs.low, prefs.high);
}

function useLastLoggedAsCurrent() {
  const records = loadRecords().slice().sort((a, b) => new Date(b.dtIso) - new Date(a.dtIso));
  if (!records.length) {
    ui.predText.textContent = "No logged readings yet.";
    return;
  }
  ui.predCurrent.value = String(records[0].value);
  ui.predText.textContent = `Using last logged value: ${records[0].value} mg/dL`;
}

/* ---------------------------
   Synthetic + backtest handlers
---------------------------- */

function generateTestSet() {
  const days = clampNumber(Number(ui.synDays.value), 7, 180) ?? 90;
  const pts = clampNumber(Number(ui.synPoints.value), 100, 5000) ?? 500;

  syntheticSeries = generateSyntheticSeries(days, pts);

  ui.backtestOut.textContent =
    `Generated synthetic series:\n` +
    `- days: ${days}\n` +
    `- points: ${syntheticSeries.length}\n` +
    `- approx interval: ${((syntheticSeries[1].tMs - syntheticSeries[0].tMs)/60000).toFixed(1)} min\n\n` +
    `Next: click "Run Backtest" to compare baseline vs improved alert logic.`;
}

function runBacktest() {
  const prefs = getPrefs();
  const horizon = Number(ui.btHorizon.value);

  if (!syntheticSeries) {
    ui.backtestOut.textContent = "No synthetic series yet. Click 'Generate Test Set' first.";
    return;
  }

  const result = backtest(syntheticSeries, prefs.low, prefs.high, horizon);
  if (!result) {
    ui.backtestOut.textContent = "Backtest failed (insufficient series length).";
    return;
  }

  const b = result.baseline;
  const im = result.improved;

  ui.backtestOut.textContent =
    `Backtest (synthetic)\n` +
    `Thresholds: low=${prefs.low}, high=${prefs.high} • horizon=${horizon} min\n\n` +
    `BASELINE:\n` +
    `  TP=${b.TP} FP=${b.FP} FN=${b.FN}\n` +
    `  precision=${(b.precision*100).toFixed(1)}% recall=${(b.recall*100).toFixed(1)}%\n\n` +
    `IMPROVED (smoothing+persistence+cooldown+hysteresis):\n` +
    `  TP=${im.TP} FP=${im.FP} FN=${im.FN}\n` +
    `  precision=${(im.precision*100).toFixed(1)}% recall=${(im.recall*100).toFixed(1)}%\n\n` +
    `False-positive reduction vs baseline: ${result.fpReductionPct.toFixed(1)}%`;
}

/* ---------------------------
   Event wiring + init
---------------------------- */

function init() {
  migrateIfNeeded();

  // Default dt to now
  ui.dt.value = toLocalInputValue(new Date());

  // Load meta prefs
  const meta = loadMeta();
  if (meta.patientName) ui.patientName.value = meta.patientName;
  if (meta.targetLow) ui.targetLow.value = meta.targetLow;
  if (meta.targetHigh) ui.targetHigh.value = meta.targetHigh;
  if (meta.windowDays) ui.windowDays.value = String(meta.windowDays);
  if (meta.contextFilter) ui.contextFilter.value = meta.contextFilter;
  if (meta.sortOrder) ui.sortOrder.value = meta.sortOrder;

  // Main save/add button
  ui.saveBtn.addEventListener("click", () => {
    const result = readForm();
    if (!result.ok) {
      if (result.msg !== "Cancelled.") alert(result.msg);
      return;
    }

    const editingId = ui.editingId.value;
    if (editingId) {
      updateRecord(editingId, result.record);
      setModeAdd();
    } else {
      addRecord(result.record);
    }

    // Clear for quick next entry
    ui.glucose.value = "";
    ui.notes.value = "";
    ui.dt.value = toLocalInputValue(new Date());
    render();
  });

  ui.cancelEditBtn.addEventListener("click", () => {
    setModeAdd();
    ui.glucose.value = "";
    ui.notes.value = "";
    ui.dt.value = toLocalInputValue(new Date());
    render();
  });

  ui.exportBtn.addEventListener("click", exportCSV);
  ui.reportBtn.addEventListener("click", printReport);

  ui.clearBtn.addEventListener("click", () => {
    if (confirm("Delete ALL readings from this browser?")) {
      localStorage.removeItem(STORAGE_KEY);
      lastDeleted = null;
      setModeAdd();
      render();
    }
  });

  ui.undoBtn.addEventListener("click", () => {
    undoDelete();
    render();
  });

  // Filters re-render
  ["windowDays", "targetLow", "targetHigh", "contextFilter", "sortOrder", "patientName"].forEach(id => {
    el(id).addEventListener("change", render);
    el(id).addEventListener("input", render);
  });

  // Table click: delete or edit
  ui.tbody.addEventListener("click", (e) => {
    const delId = e.target?.dataset?.del;
    if (delId) {
      deleteRecord(delId);
      render();
      return;
    }

    const tr = e.target.closest("tr");
    const editId = tr?.dataset?.editId;
    if (!editId) return;

    const records = loadRecords();
    const rec = records.find(r => r.id === editId);
    if (!rec) return;

    setModeEdit(rec);
    render();
  });

  // Prediction buttons
  ui.predictBtn.addEventListener("click", runPrediction);
  ui.useLastBtn.addEventListener("click", useLastLoggedAsCurrent);

  // Synthetic/backtest buttons
  ui.genTestBtn.addEventListener("click", generateTestSet);
  ui.runBacktestBtn.addEventListener("click", runBacktest);

  // First render
  setModeAdd();
  render();
  drawPredictionChart(ui.predChart, null, getPrefs().low, getPrefs().high);
}

init();
