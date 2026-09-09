/* NHT Operations Dashboard V4
   Supabase Auth + shared Supabase database + realtime updates
   Excel is parsed in-browser; processed summary rows are stored centrally.
*/
const SUPABASE_URL = "https://goypnlxygamrcwuedshz.supabase.co";
const SUPABASE_PUBLISHABLE_KEY = "sb_publishable_dFI0KenxorgJlbspXucwQg_tzefgC7e";

const supabaseClient = window.supabase.createClient(
  SUPABASE_URL,
  SUPABASE_PUBLISHABLE_KEY,
  { auth: { persistSession: true, autoRefreshToken: true, detectSessionInUrl: true } }
);

let rawRows = [];
let charts = {};
let currentUser = null;
let realtimeChannel = null;
const $ = id => document.getElementById(id);

const num = value => {
  if (value === null || value === undefined || value === "" || value === "-") return 0;
  if (typeof value === "number") return Number.isFinite(value) ? value : 0;
  const parsed = parseFloat(String(value).replace(/,/g, "").replace(/%/g, "").trim());
  return Number.isFinite(parsed) ? parsed : 0;
};

const pct = value => `${num(value).toFixed(1)}%`;
const rateFraction = value => {
  const n = num(value);
  return n > 1 ? n / 100 : n;
};
const safeRatePercent = value => rateFraction(value) * 100;
const sum = (rows, key) => rows.reduce((total, row) => total + num(row[key]), 0);

function cleanHeader(value) {
  return String(value || "")
    .trim().toLowerCase()
    .replace(/\s+/g, " ")
    .replace(/%/g, "");
}

function canonical(row) {
  const result = {};
  Object.keys(row).forEach(key => { result[cleanHeader(key)] = row[key]; });
  return result;
}

function get(row, names) {
  for (const name of names) if (row[name] !== undefined) return row[name];
  return "";
}

function isTotalText(value) { return /\b(total|grand total|overall)\b/i.test(String(value || "")); }
function monthBase(value) { return String(value || "").trim().replace(/\s+(total|grand total|overall)\s*$/i, "").trim(); }

function normalizeRows(rows) {
  let currentMonth = "";
  const output = [];
  rows.forEach(row => {
    const x = canonical(row);
    const rawMonth = String(get(x, ["month"]) || "").trim();
    const location = String(get(x, ["location"]) || "").trim();
    const batchRaw = get(x, ["total batch conducted", "batch", "total batches"]);
    if (rawMonth && !isTotalText(rawMonth)) currentMonth = monthBase(rawMonth);
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

function setSync(text, state = "") {
  const el = $("syncStatus");
  if (!el) return;
  el.textContent = text;
  el.className = `sync-status ${state}`;
}

function parseDbRow(row) {
  return {
    month: row.month,
    location: row.location,
    batch: num(row.batch),
    inflow: num(row.inflow),
    outflow: num(row.outflow),
    hr: num(row.hr),
    training: num(row.training),
    throughput: num(row.throughput),
    joined: num(row.joined),
    joiningThroughput: num(row.joining_throughput)
  };
}

async function loadSharedData() {
  setSync("Loading shared data…", "loading");
  const { data, error } = await supabaseClient
    .from("nht_dashboard_data")
    .select("month,location,batch,inflow,outflow,hr,training,throughput,joined,joining_throughput")
    .order("month", { ascending: true })
    .order("location", { ascending: true });

  if (error) {
    console.error(error);
    setSync("Database setup required", "error");
    $("fileStatus").textContent = "Shared data unavailable";
    $("periodNote").textContent = "The dashboard is connected to login, but the shared data table still needs to be created in Supabase.";
    return false;
  }

  rawRows = (data || []).map(parseDbRow);
  populateFilters();
  render();
  setSync(`Live • ${rawRows.length} records`, "live");
  $("fileStatus").textContent = rawRows.length ? "✓ Shared data loaded" : "No shared data yet";
  $("updatedAt").textContent = "Shared data loaded: " + new Date().toLocaleString();
  return true;
}

async function uploadToSharedDatabase(file) {
  $("fileStatus").textContent = `Reading ${file.name}…`;
  setSync("Processing Excel…", "loading");

  try {
    const workbook = await readWorkbook(file);
    const sheet = workbook.Sheets[workbook.SheetNames[0]];
    const rows = XLSX.utils.sheet_to_json(sheet, { defval: "" });
    const normalized = normalizeRows(rows);
    if (!normalized.length) throw new Error("No location-level records found");

    const payload = normalized.map(row => ({
      month: row.month,
      location: row.location,
      batch: row.batch,
      inflow: row.inflow,
      outflow: row.outflow,
      hr: row.hr,
      training: row.training,
      throughput: row.throughput,
      joined: row.joined,
      joining_throughput: row.joiningThroughput
    }));

    const { error } = await supabaseClient.rpc("replace_nht_dashboard_data", { p_rows: payload });
    if (error) throw error;

    rawRows = normalized;
    populateFilters();
    render();
    $("fileStatus").textContent = `✓ ${file.name} uploaded & synced`;
    $("updatedAt").textContent = "Last Excel update: " + new Date().toLocaleString();
    setSync(`Live • ${rawRows.length} records`, "live");
    $("periodNote").textContent = `${rawRows.length} location-level records • shared centrally • monthly total rows excluded`;
  } catch (error) {
    console.error(error);
    $("fileStatus").textContent = "Could not update shared data";
    setSync("Update failed", "error");
    alert("The Excel file could not be processed or saved. Please check the workbook and confirm the Supabase database setup is complete.");
  }
}

function readWorkbook(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = event => {
      try { resolve(XLSX.read(new Uint8Array(event.target.result), { type: "array" })); }
      catch (error) { reject(error); }
    };
    reader.onerror = reject;
    reader.readAsArrayBuffer(file);
  });
}

function populateFilters() {
  const months = [...new Set(rawRows.map(row => row.month).filter(Boolean))];
  const locations = [...new Set(rawRows.map(row => row.location).filter(Boolean))];
  const currentMonth = $("monthFilter").value;
  const currentLocation = $("locationFilter").value;

  $("monthFilter").innerHTML = '<option value="ALL">All months</option>' +
    months.map(month => `<option value="${escapeHtml(month)}">${escapeHtml(month)}</option>`).join("");
  $("locationFilter").innerHTML = '<option value="ALL">All locations</option>' +
    locations.map(location => `<option value="${escapeHtml(location)}">${escapeHtml(location)}</option>`).join("");

  if (months.includes(currentMonth)) $("monthFilter").value = currentMonth;
  if (locations.includes(currentLocation)) $("locationFilter").value = currentLocation;
}

function escapeHtml(value) {
  return String(value).replace(/[&<>"']/g, character => ({"&":"&amp;","<":"&lt;",">":"&gt;",'"':"&quot;","'":"&#039;"}[character]));
}

function filteredRows() {
  const month = $("monthFilter").value;
  const location = $("locationFilter").value;
  return rawRows.filter(row => (month === "ALL" || row.month === month) && (location === "ALL" || row.location === location));
}

function monthlyRows() {
  const groups = {};
  filteredRows().forEach(row => {
    if (!groups[row.month]) groups[row.month] = { month: row.month, inflow: 0, outflow: 0, hr: 0, training: 0, joined: 0, joiningWeighted: 0 };
    const g = groups[row.month];
    g.inflow += row.inflow;
    g.outflow += row.outflow;
    g.hr += row.hr;
    g.training += row.training;
    g.joined += row.joined;
    g.joiningWeighted += rateFraction(row.joiningThroughput) * row.inflow;
  });
  return Object.values(groups).sort((a, b) => monthIndex(a.month) - monthIndex(b.month));
}

function monthIndex(month) {
  const order = ["January","February","March","April","May","June","July","August","September","October","November","December"];
  const i = order.indexOf(month);
  return i === -1 ? 999 : i;
}

function throughputRate(rows) {
  const inflow = sum(rows, "inflow");
  return inflow ? (sum(rows, "joined") / inflow) * 100 : 0;
}

function joiningThroughputRate(rows) {
  const inflow = sum(rows, "inflow");
  if (!inflow) return 0;
  return rows.reduce((total, row) => total + rateFraction(row.joiningThroughput) * row.inflow, 0) / inflow * 100;
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
  $("viewTitle").textContent = (month === "ALL" ? "All months" : month) + (location === "ALL" ? "" : ` • ${location}`);

  const best = locationGroups(rows).sort((a, b) => b.rate - a.rate)[0];
  $("insightText").textContent = rows.length
    ? `${rows.length} location-level records are in view. ${best ? `${best.location} currently has the highest calculated throughput at ${pct(best.rate)}.` : ""}`
    : "No data matches the current filters.";

  $("periodNote").textContent = rawRows.length
    ? `${rawRows.length} location-level records • shared centrally • monthly total rows excluded`
    : "No shared NHT data has been uploaded yet.";

  drawThroughput(months);
  drawLocation(locationGroups(rows));
  drawMovement(months);
  drawAttrition(months);
  drawTable(locationGroups(rows));
}

function format(value) { return Math.round(value).toLocaleString("en-IN"); }

function locationGroups(rows) {
  const groups = {};
  rows.forEach(row => {
    if (!groups[row.location]) groups[row.location] = { location: row.location, inflow: 0, outflow: 0, hr: 0, training: 0, joined: 0, joiningWeighted: 0 };
    const g = groups[row.location];
    g.inflow += row.inflow;
    g.outflow += row.outflow;
    g.hr += row.hr;
    g.training += row.training;
    g.joined += row.joined;
    g.joiningWeighted += rateFraction(row.joiningThroughput) * row.inflow;
  });
  return Object.values(groups).map(g => ({
    ...g,
    rate: g.inflow ? (g.joined / g.inflow) * 100 : 0,
    joinRate: g.inflow ? (g.joiningWeighted / g.inflow) * 100 : 0
  }));
}

function destroy(name) { if (charts[name]) { charts[name].destroy(); charts[name] = null; } }

function commonScales() {
  return {
    responsive: true,
    maintainAspectRatio: false,
    animation: { duration: 500 },
    plugins: { legend: { labels: { usePointStyle: true, boxWidth: 8, font: { size: 11 } } } },
    scales: { x: { grid: { display: false }, ticks: { font: { size: 10 } } }, y: { beginAtZero: true, ticks: { font: { size: 10 } } } }
  };
}

function drawThroughput(rows) {
  destroy("throughput");
  const labels = rows.map(r => r.month);
  const throughputData = rows.map(r => r.inflow ? (r.joined / r.inflow) * 100 : null);
  const joiningData = rows.map(r => r.inflow ? (r.joiningWeighted / r.inflow) * 100 : null);
  charts.throughput = new Chart($("throughputChart"), {
    type: "line",
    data: { labels, datasets: [
      { label: "Throughput", data: throughputData, tension: 0.35, borderWidth: 3, pointRadius: 3, spanGaps: true, borderColor: "#0b5ed7", backgroundColor: "rgba(11,94,215,.10)" },
      { label: "Joining Throughput", data: joiningData, tension: 0.35, borderWidth: 3, pointRadius: 3, borderDash: [6,5], spanGaps: true, borderColor: "#e21d2f", backgroundColor: "rgba(226,29,47,.08)" }
    ] },
    options: {
      ...commonScales(),
      scales: { ...commonScales().scales, y: { beginAtZero: true, max: 100, ticks: { callback: value => `${value}%` } } },
      plugins: {
        ...commonScales().plugins,
        tooltip: { callbacks: { label: ctx => `${ctx.dataset.label}: ${num(ctx.raw).toFixed(1)}%` } }
      }
    }
  });
}

function drawLocation(rows) {
  destroy("location");
  const sorted = [...rows].sort((a, b) => b.rate - a.rate);
  charts.location = new Chart($("locationChart"), { type: "bar", data: { labels: sorted.map(r => r.location), datasets: [{ label: "Throughput", data: sorted.map(r => r.rate), borderRadius: 6, backgroundColor: "#0b5ed7" }] }, options: { indexAxis: "y", responsive: true, maintainAspectRatio: false, plugins: { legend: { display: false } }, scales: { x: { beginAtZero: true, max: 100, ticks: { callback: value => `${value}%` }, grid: { display: false } }, y: { grid: { display: false }, ticks: { font: { size: 10 } } } } } });
}

function drawMovement(rows) {
  destroy("movement");
  charts.movement = new Chart($("movementChart"), { type: "bar", data: { labels: rows.map(r => r.month), datasets: [{ label: "Inflow", data: rows.map(r => r.inflow), borderRadius: 5, backgroundColor: "#0b5ed7" }, { label: "Outflow", data: rows.map(r => r.outflow), borderRadius: 5, backgroundColor: "#e21d2f" }] }, options: commonScales() });
}

function drawAttrition(rows) {
  destroy("attrition");
  charts.attrition = new Chart($("attritionChart"), { type: "line", data: { labels: rows.map(r => r.month), datasets: [{ label: "HR Attrition", data: rows.map(r => r.hr), tension: 0.3, borderWidth: 3, pointRadius: 3, spanGaps: true, borderColor: "#e21d2f" }, { label: "Training Attrition", data: rows.map(r => r.training), tension: 0.3, borderWidth: 3, pointRadius: 3, spanGaps: true, borderColor: "#0b5ed7" }] }, options: { ...commonScales(), plugins: { ...commonScales().plugins, tooltip: { callbacks: { label: ctx => `${ctx.dataset.label}: ${num(ctx.raw).toLocaleString("en-IN")}` } } } } });
}

function drawTable(rows) {
  const body = $("summaryTable");
  if (!rows.length) { body.innerHTML = '<tr><td colspan="8" class="empty">No records match the selected filters.</td></tr>'; return; }
  body.innerHTML = [...rows].sort((a,b) => b.rate-a.rate).map(r => `<tr><td>${escapeHtml(r.location)}</td><td>${format(r.inflow)}</td><td>${format(r.outflow)}</td><td>${format(r.hr)}</td><td>${format(r.training)}</td><td>${format(r.joined)}</td><td>${pct(r.rate)}</td><td>${pct(r.joinRate)}</td></tr>`).join("");
}

function subscribeToRealtime() {
  if (realtimeChannel) supabaseClient.removeChannel(realtimeChannel);
  realtimeChannel = supabaseClient.channel("nht-dashboard-live")
    .on("postgres_changes", { event: "*", schema: "public", table: "nht_dashboard_data" }, async () => {
      setSync("Updating…", "loading");
      await loadSharedData();
    })
    .subscribe(status => {
      if (status === "SUBSCRIBED") setSync(`Live • ${rawRows.length} records`, "live");
    });
}

function showDashboard(user) {
  $("authLoading").classList.add("hidden");
  $("loginScreen").classList.add("hidden");
  $("appShell").classList.remove("hidden");
  $("signedInAs").textContent = user?.email || "Signed in";
  if (!window.__dashboardInitialized) {
    window.__dashboardInitialized = true;
    currentUser = user;
    $("excelFile").addEventListener("change", event => { if (event.target.files[0]) { uploadToSharedDatabase(event.target.files[0]); event.target.value = ""; } });
    $("monthFilter").addEventListener("change", render);
    $("locationFilter").addEventListener("change", render);
    loadSharedData().then(subscribeToRealtime);
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
  if (error) $("loginError").textContent = "Unable to sign in. Please check your email and password.";
  $("loginButton").disabled = false;
  $("loginButton").textContent = "Sign in";
}

async function startAuth() {
  try {
    const { data, error } = await supabaseClient.auth.getSession();
    if (error) { showLogin("Unable to check your secure session. Please try again."); return; }
    if (data.session?.user) showDashboard(data.session.user); else showLogin();
    supabaseClient.auth.onAuthStateChange((_event, session) => { if (session?.user) showDashboard(session.user); else showLogin(); });
  } catch (error) { console.error(error); showLogin("Unable to connect to secure authentication."); }
}

$("loginForm").addEventListener("submit", async event => { event.preventDefault(); await signIn($("loginEmail").value.trim(), $("loginPassword").value); });
$("signOutButton").addEventListener("click", async () => { if (realtimeChannel) await supabaseClient.removeChannel(realtimeChannel); await supabaseClient.auth.signOut(); });
startAuth();
