/* NHT Operations Dashboard V3
   Supabase Auth + browser-only Excel dashboard
*/
const SUPABASE_URL = "https://goypnlxygamrcwuedshz.supabase.co";
const SUPABASE_PUBLISHABLE_KEY = "sb_publishable_dFI0KenxorgJlbspXucwQg_tzefgC7e";

const supabaseClient = window.supabase.createClient(
  SUPABASE_URL,
  SUPABASE_PUBLISHABLE_KEY,
  {
    auth: {
      persistSession: true,
      autoRefreshToken: true,
      detectSessionInUrl: true
    }
  }
);

let rawRows = [];
let charts = {};
const STORAGE_KEY = "nhtDashboardDataV3";
const META_KEY = "nhtDashboardMetaV3";
const $ = id => document.getElementById(id);

const num = value => {
  if (value === null || value === undefined || value === "" || value === "-") return 0;
  if (typeof value === "number") return Number.isFinite(value) ? value : 0;
  const parsed = parseFloat(String(value).replace(/,/g, "").replace(/%/g, "").trim());
  return Number.isFinite(parsed) ? parsed : 0;
};

const pct = value => `${num(value).toFixed(1)}%`;
const sum = (rows, key) => rows.reduce((total, row) => total + num(row[key]), 0);

function cleanHeader(value) {
  return String(value || "")
    .trim()
    .toLowerCase()
    .replace(/\s+/g, " ")
    .replace(/%/g, "");
}

function canonical(row) {
  const result = {};
  Object.keys(row).forEach(key => {
    result[cleanHeader(key)] = row[key];
  });
  return result;
}

function get(row, names) {
  for (const name of names) {
    if (row[name] !== undefined) return row[name];
  }
  return "";
}

function isTotalText(value) {
  return /\b(total|grand total|overall)\b/i.test(String(value || ""));
}

function monthBase(value) {
  return String(value || "")
    .trim()
    .replace(/\s+(total|grand total|overall)\s*$/i, "")
    .trim();
}

function normalizeRows(rows) {
  let currentMonth = "";
  const output = [];

  rows.forEach(row => {
    const x = canonical(row);
    const rawMonth = String(get(x, ["month"]) || "").trim();
    const location = String(get(x, ["location"]) || "").trim();
    const batchRaw = get(x, ["total batch conducted", "batch", "total batches"]);

    // The workbook contains monthly total rows such as "June Total".
    // Exclude them so they cannot duplicate months or double-count KPIs.
    if (rawMonth && !isTotalText(rawMonth)) {
      currentMonth = monthBase(rawMonth);
    }

    const marker = `${rawMonth} ${location} ${batchRaw}`;
    if (isTotalText(marker)) return;
    if (!currentMonth || !location) return;

    output.push({
      month: currentMonth,
      location,
      batch: num(batchRaw),
      inflow: num(get(x, ["total inflow", "inflow"])),
      outflow: num(get(x, ["total outflow", "outflow"])),
      hr: num(get(x, ["hr attrition"])),
      training: num(get(x, ["training attrition"])),
      throughput: num(get(x, ["throughput"])),
      joined: num(get(x, ["total joined", "joined"])),
      joiningThroughput: num(get(x, ["joining throughput"]))
    });
  });

  return output;
}

function saveData(fileName) {
  try {
    localStorage.setItem(STORAGE_KEY, JSON.stringify(rawRows));
    localStorage.setItem(
      META_KEY,
      JSON.stringify({ fileName, savedAt: new Date().toISOString() })
    );
    return true;
  } catch (error) {
    console.warn("Could not save dashboard data locally", error);
    return false;
  }
}

function restoreData() {
  try {
    const saved = localStorage.getItem(STORAGE_KEY);
    if (!saved) return false;

    const parsed = JSON.parse(saved);
    if (!Array.isArray(parsed) || !parsed.length) return false;

    rawRows = parsed;
    populateFilters();
    render();

    const meta = JSON.parse(localStorage.getItem(META_KEY) || "{}");
    if (meta.fileName) {
      $("fileStatus").textContent = `✓ ${meta.fileName} restored`;
    }
    if (meta.savedAt) {
      $("updatedAt").textContent = `Last update: ${new Date(meta.savedAt).toLocaleString()}`;
    }

    $("periodNote").textContent =
      `${rawRows.length} location-level records restored • monthly total rows excluded`;

    return true;
  } catch (error) {
    console.warn("Could not restore saved dashboard data", error);
    return false;
  }
}

function parseFile(file) {
  $("fileStatus").textContent = `Reading ${file.name}…`;

  const reader = new FileReader();
  reader.onload = event => {
    try {
      const workbook = XLSX.read(new Uint8Array(event.target.result), { type: "array" });
      const sheet = workbook.Sheets[workbook.SheetNames[0]];
      const rows = XLSX.utils.sheet_to_json(sheet, { defval: "" });
      const normalized = normalizeRows(rows);

      if (!normalized.length) {
        throw new Error("No location-level records found");
      }

      rawRows = normalized;
      saveData(file.name);
      populateFilters();
      render();

      $("fileStatus").textContent = `✓ ${file.name} loaded & saved`;
      $("updatedAt").textContent = `Last update: ${new Date().toLocaleString()}`;
      $("periodNote").textContent =
        `${rawRows.length} location-level records loaded • monthly total rows excluded • upload again only when the Excel is updated`;
    } catch (error) {
      console.error(error);
      $("fileStatus").textContent = "Could not read this workbook.";
      alert("The Excel file could not be read. Please check the workbook format and headers.");
    }
  };

  reader.readAsArrayBuffer(file);
}

function populateFilters() {
  const months = [...new Set(rawRows.map(row => row.month).filter(Boolean))];
  const locations = [...new Set(rawRows.map(row => row.location).filter(Boolean))];
  const currentMonth = $("monthFilter").value;
  const currentLocation = $("locationFilter").value;

  $("monthFilter").innerHTML =
    '<option value="ALL">All months</option>' +
    months.map(month => `<option value="${escapeHtml(month)}">${escapeHtml(month)}</option>`).join("");

  $("locationFilter").innerHTML =
    '<option value="ALL">All locations</option>' +
    locations.map(location => `<option value="${escapeHtml(location)}">${escapeHtml(location)}</option>`).join("");

  if (months.includes(currentMonth)) $("monthFilter").value = currentMonth;
  if (locations.includes(currentLocation)) $("locationFilter").value = currentLocation;
}

function escapeHtml(value) {
  return String(value).replace(/[&<>"']/g, character => ({
    "&": "&amp;",
    "<": "&lt;",
    ">": "&gt;",
    '"': "&quot;",
    "'": "&#039;"
  }[character]));
}

function filteredRows() {
  const month = $("monthFilter").value;
  const location = $("locationFilter").value;

  return rawRows.filter(row =>
    (month === "ALL" || row.month === month) &&
    (location === "ALL" || row.location === location)
  );
}

function monthlyRows() {
  const groups = {};

  filteredRows().forEach(row => {
    const key = row.month;
    if (!groups[key]) {
      groups[key] = {
        month: key,
        inflow: 0,
        outflow: 0,
        hr: 0,
        training: 0,
        joined: 0,
        joiningWeighted: 0
      };
    }

    groups[key].inflow += row.inflow;
    groups[key].outflow += row.outflow;
    groups[key].hr += row.hr;
    groups[key].training += row.training;
    groups[key].joined += row.joined;
    groups[key].joiningWeighted += row.joiningThroughput * row.inflow;
  });

  return Object.values(groups);
}

function throughputRate(rows) {
  const inflow = sum(rows, "inflow");
  return inflow ? (sum(rows, "joined") / inflow) * 100 : 0;
}

function joiningThroughputRate(rows) {
  const inflow = sum(rows, "inflow");
  if (!inflow) return 0;
  return rows.reduce((total, row) => total + row.joiningThroughput * row.inflow, 0) / inflow;
}

function render() {
  const rows = filteredRows();
  const months = monthlyRows();

  const batches = sum(rows, "batch");
  const inflow = sum(rows, "inflow");
  const outflow = sum(rows, "outflow");
  const joined = sum(rows, "joined");
  const hr = sum(rows, "hr");
  const training = sum(rows, "training");
  const throughput = throughputRate(rows);
  const joiningFinal = joiningThroughputRate(rows);

  $("kpiBatches").textContent = format(batches);
  $("kpiInflow").textContent = format(inflow);
  $("kpiOutflow").textContent = format(outflow);
  $("kpiJoined").textContent = format(joined);
  $("kpiHr").textContent = format(hr);
  $("kpiTraining").textContent = format(training);
  $("kpiThroughput").textContent = pct(throughput);
  $("kpiJoiningThroughput").textContent = pct(joiningFinal);

  const month = $("monthFilter").value;
  const location = $("locationFilter").value;
  $("viewTitle").textContent =
    (month === "ALL" ? "All months" : month) +
    (location === "ALL" ? "" : ` • ${location}`);

  const best = locationGroups(rows).sort((a, b) => b.rate - a.rate)[0];
  $("insightText").textContent = rows.length
    ? `${rows.length} location-level records are in view. ${best ? `${best.location} currently has the highest calculated throughput at ${pct(best.rate)}.` : ""}`
    : "No data matches the current filters.";

  $("periodNote").textContent = rawRows.length
    ? `${rawRows.length} location-level records loaded • monthly total rows excluded to prevent double-counting`
    : "Upload the NHT summary Excel once. Your last processed data is saved in this browser.";

  drawThroughput(months);
  drawLocation(locationGroups(rows));
  drawMovement(months);
  drawAttrition(months);
  drawTable(locationGroups(rows));
}

function format(value) {
  return Math.round(value).toLocaleString("en-IN");
}

function locationGroups(rows) {
  const groups = {};

  rows.forEach(row => {
    if (!groups[row.location]) {
      groups[row.location] = {
        location: row.location,
        inflow: 0,
        outflow: 0,
        hr: 0,
        training: 0,
        joined: 0,
        joiningWeighted: 0
      };
    }

    const group = groups[row.location];
    group.inflow += row.inflow;
    group.outflow += row.outflow;
    group.hr += row.hr;
    group.training += row.training;
    group.joined += row.joined;
    group.joiningWeighted += row.joiningThroughput * row.inflow;
  });

  return Object.values(groups).map(group => ({
    ...group,
    rate: group.inflow ? (group.joined / group.inflow) * 100 : 0,
    joinRate: group.inflow ? group.joiningWeighted / group.inflow : 0
  }));
}

function destroy(name) {
  if (charts[name]) charts[name].destroy();
}

function commonScales() {
  return {
    responsive: true,
    maintainAspectRatio: false,
    plugins: {
      legend: {
        labels: { usePointStyle: true, boxWidth: 8, font: { size: 11 } }
      }
    },
    scales: {
      x: { grid: { display: false }, ticks: { font: { size: 10 } } },
      y: { beginAtZero: true, ticks: { font: { size: 10 } } }
    }
  };
}

function drawThroughput(rows) {
  destroy("throughput");
  charts.throughput = new Chart($("throughputChart"), {
    type: "line",
    data: {
      labels: rows.map(row => row.month),
      datasets: [
        {
          label: "Throughput",
          data: rows.map(row => row.inflow ? (row.joined / row.inflow) * 100 : 0),
          tension: 0.35,
          borderWidth: 3,
          pointRadius: 3
        },
        {
          label: "Joining Throughput",
          data: rows.map(row => row.inflow ? (row.joiningWeighted / row.inflow) : 0),
          tension: 0.35,
          borderWidth: 2,
          pointRadius: 3,
          borderDash: [6, 5]
        }
      ]
    },
    options: {
      ...commonScales(),
      scales: {
        ...commonScales().scales,
        y: { beginAtZero: true, max: 100, ticks: { callback: value => `${value}%` } }
      }
    }
  });
}

function drawLocation(rows) {
  destroy("location");
  const sorted = [...rows].sort((a, b) => b.rate - a.rate);

  charts.location = new Chart($("locationChart"), {
    type: "bar",
    data: {
      labels: sorted.map(row => row.location),
      datasets: [{ label: "Throughput", data: sorted.map(row => row.rate), borderRadius: 6 }]
    },
    options: {
      indexAxis: "y",
      responsive: true,
      maintainAspectRatio: false,
      plugins: { legend: { display: false } },
      scales: {
        x: { beginAtZero: true, max: 100, ticks: { callback: value => `${value}%` }, grid: { display: false } },
        y: { grid: { display: false }, ticks: { font: { size: 10 } } }
      }
    }
  });
}

function drawMovement(rows) {
  destroy("movement");
  charts.movement = new Chart($("movementChart"), {
    type: "bar",
    data: {
      labels: rows.map(row => row.month),
      datasets: [
        { label: "Inflow", data: rows.map(row => row.inflow), borderRadius: 5 },
        { label: "Outflow", data: rows.map(row => row.outflow), borderRadius: 5 }
      ]
    },
    options: commonScales()
  });
}

function drawAttrition(rows) {
  destroy("attrition");
  charts.attrition = new Chart($("attritionChart"), {
    type: "line",
    data: {
      labels: rows.map(row => row.month),
      datasets: [
        { label: "HR Attrition", data: rows.map(row => row.hr), tension: 0.3, borderWidth: 3, pointRadius: 3 },
        { label: "Training Attrition", data: rows.map(row => row.training), tension: 0.3, borderWidth: 3, pointRadius: 3 }
      ]
    },
    options: commonScales()
  });
}

function drawTable(rows) {
  const body = $("summaryTable");
  if (!rows.length) {
    body.innerHTML = '<tr><td colspan="8" class="empty">No records match the selected filters.</td></tr>';
    return;
  }

  body.innerHTML = [...rows]
    .sort((a, b) => b.rate - a.rate)
    .map(row => `
      <tr>
        <td>${escapeHtml(row.location)}</td>
        <td>${format(row.inflow)}</td>
        <td>${format(row.outflow)}</td>
        <td>${format(row.hr)}</td>
        <td>${format(row.training)}</td>
        <td>${format(row.joined)}</td>
        <td>${pct(row.rate)}</td>
        <td>${pct(row.joinRate)}</td>
      </tr>`)
    .join("");
}

function showDashboard(user) {
  $("authLoading").classList.add("hidden");
  $("loginScreen").classList.add("hidden");
  $("appShell").classList.remove("hidden");
  $("signedInAs").textContent = user?.email || "Signed in";

  if (!window.__dashboardInitialized) {
    window.__dashboardInitialized = true;
    $("excelFile").addEventListener("change", event => {
      if (event.target.files[0]) parseFile(event.target.files[0]);
    });
    $("monthFilter").addEventListener("change", render);
    $("locationFilter").addEventListener("change", render);
    restoreData();
  }
}

function showLogin(message = "") {
  $("authLoading").classList.add("hidden");
  $("appShell").classList.add("hidden");
  $("loginScreen").classList.remove("hidden");
  $("loginError").textContent = message;
}

async function signIn(email, password) {
  $("loginButton").disabled = true;
  $("loginButton").textContent = "Signing in…";
  $("loginError").textContent = "";

  const { error } = await supabaseClient.auth.signInWithPassword({ email, password });

  if (error) {
    $("loginError").textContent = "Unable to sign in. Please check your email and password.";
  }

  $("loginButton").disabled = false;
  $("loginButton").textContent = "Sign in";
}

async function startAuth() {
  try {
    const { data, error } = await supabaseClient.auth.getSession();

    if (error) {
      showLogin("Unable to check your secure session. Please try again.");
      return;
    }

    if (data.session?.user) showDashboard(data.session.user);
    else showLogin();

    supabaseClient.auth.onAuthStateChange((_event, session) => {
      if (session?.user) showDashboard(session.user);
      else showLogin();
    });
  } catch (error) {
    console.error(error);
    showLogin("Unable to connect to secure authentication.");
  }
}

$("loginForm").addEventListener("submit", async event => {
  event.preventDefault();
  await signIn($("loginEmail").value.trim(), $("loginPassword").value);
});

$("signOutButton").addEventListener("click", async () => {
  await supabaseClient.auth.signOut();
});

startAuth();
