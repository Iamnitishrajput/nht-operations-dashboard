/* NHT Operations Dashboard V8
   Supabase Auth + role-based upload access + shared Supabase database + realtime updates.
   Excel is parsed in-browser; validated summary rows are stored centrally.
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
let currentRole = "viewer";
let realtimeChannel = null;
let inactivityTimer = null;
let lastActivityAt = 0;
const INACTIVITY_MS = 5 * 60 * 1000;
let inactivityInterval = null;
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

function validateWorkbookRows(rows) {
  if (!rows.length) throw new Error("The worksheet is empty.");
  const headers = Object.keys(canonical(rows[0]));
  const required = [
    "month", "location", "total batch conducted", "total inflow", "total outflow",
    "hr attrition", "training attrition", "throughput", "total joined", "joining throughput"
  ];
  const missing = required.filter(h => !headers.includes(h));
  if (missing.length) throw new Error(`Missing required columns: ${missing.join(", ")}`);

  const numericFields = [
    ["Total Batch Conducted", "total batch conducted"],
    ["Total Inflow", "total inflow"],
    ["Total Outflow", "total outflow"],
    ["HR Attrition", "hr attrition"],
    ["Training Attrition", "training attrition"],
    ["Throughput", "throughput"],
    ["Total Joined", "total joined"],
    ["Joining Throughput", "joining throughput"]
  ];

  const problems = [];
  let currentMonth = "";
  rows.forEach((row, i) => {
    const x = canonical(row);
    const rawMonth = String(get(x, ["month"]) || "").trim();
    const location = String(get(x, ["location"]) || "").trim();
    if (rawMonth && !isTotalText(rawMonth)) currentMonth = monthBase(rawMonth);
    const effectiveMonth = rawMonth || currentMonth;
    const marker = `${effectiveMonth} ${location}`;
    if (isTotalText(marker)) return;
    if (!effectiveMonth || !location) problems.push(`Row ${i + 2}: Month and Location are required.`);
    numericFields.forEach(([label, key]) => {
      const value = get(x, [key]);
      if (value !== "" && value !== "-" && value !== null && value !== undefined && !Number.isFinite(num(value))) {
        problems.push(`Row ${i + 2}: ${label} must be numeric.`);
      }
    });
  });
  if (problems.length) throw new Error(problems.slice(0, 5).join("\n"));
}

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

  const duplicateKeys = new Set();
  const duplicates = [];
  output.forEach(row => {
    const key = `${row.month}|||${row.location}`.toLowerCase();
    if (duplicateKeys.has(key)) duplicates.push(`${row.month} / ${row.location}`);
    duplicateKeys.add(key);
  });
  if (duplicates.length) throw new Error(`Duplicate Month + Location rows found: ${duplicates.slice(0, 5).join(", ")}`);
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

async function loadRole(user) {
  currentRole = "viewer";
  try {
    // V10: read the role through a SECURITY DEFINER RPC instead of querying
    // nht_user_roles directly. This avoids PostgREST table-privilege/RLS 403s
    // while still returning only the signed-in user's own role.
    const { data, error } = await supabaseClient.rpc("get_nht_role");
    if (error) throw error;
    if (data === "admin" || data === "uploader" || data === "viewer") {
      currentRole = data;
    } else {
      console.warn("No role assigned to this account; defaulting to viewer.");
    }
  } catch (error) {
    console.error("Role lookup failed:", error);
    const hint = $("uploadHint");
    if (hint) hint.textContent = "Role could not be verified. Please contact the administrator.";
  }
  const canUpload = ["admin", "uploader"].includes(currentRole);
  const uploadButton = $("uploadButton");
  const uploadHint = $("uploadHint");
  if (uploadButton) {
    uploadButton.classList.toggle("disabled", !canUpload);
    uploadButton.setAttribute("aria-disabled", String(!canUpload));
    uploadButton.textContent = canUpload ? "↥  Upload / Update Excel" : "🔒  Upload Restricted";
    uploadButton.style.pointerEvents = canUpload ? "auto" : "none";
  }
  if (uploadHint) uploadHint.textContent = canUpload ? "Authorized upload access • changes sync centrally" : "Viewer access • contact the administrator for upload rights";
  $("roleBadge").textContent = currentRole === "admin" ? "ADMIN" : currentRole === "uploader" ? "UPLOADER" : "VIEWER";
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
    setSync("Data unavailable", "error");
    $("fileStatus").textContent = "Shared data unavailable";
    $("periodNote").textContent = "Unable to load shared dashboard data.";
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
  if (!["admin", "uploader"].includes(currentRole)) {
    alert("You have view-only access. Please contact the dashboard administrator for upload access.");
    return;
  }

  $("fileStatus").textContent = `Reading ${file.name}…`;
  setSync("Validating Excel…", "loading");

  try {
    const extension = file.name.toLowerCase().split(".").pop();
    if (!["xlsx", "xls", "csv"].includes(extension)) throw new Error("Unsupported file type. Please use .xlsx, .xls or .csv.");

    const workbook = await readWorkbook(file);
    if (!workbook.SheetNames?.length) throw new Error("No worksheet found in the file.");
    const sheet = workbook.Sheets[workbook.SheetNames[0]];
    const rows = XLSX.utils.sheet_to_json(sheet, { defval: "" });
    validateWorkbookRows(rows);
    const normalized = normalizeRows(rows);
    if (!normalized.length) throw new Error("No location-level records found after excluding monthly total rows.");

    const months = [...new Set(normalized.map(row => row.month))];
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
      joining_throughput: row.joiningThroughput,
      updated_by: currentUser?.id || null
    }));

    // Important: replace only the months present in this upload.
    // The database function makes delete + insert one atomic operation.
    // Uploading June again updates June without deleting July/August.
    setSync("Saving shared data…", "loading");
    const { error: syncError } = await supabaseClient.rpc("replace_nht_months", {
      p_rows: payload,
      p_months: months
    });
    if (syncError) throw syncError;

    rawRows = normalized;
    populateFilters();
    render();
    $("fileStatus").textContent = `✓ ${file.name} uploaded & synced`;
    $("updatedAt").textContent = "Last Excel update: " + new Date().toLocaleString();
    setSync(`Live • ${rawRows.length} records`, "live");
    $("periodNote").textContent = `${rawRows.length} location-level records • shared centrally • monthly total rows excluded`;
  } catch (error) {
    console.error(error);
    $("fileStatus").textContent = "Update rejected — existing data kept";
    setSync("Update failed", "error");
    alert(`The Excel update was rejected.\n\n${error.message}\n\nNo replacement was made until the file passed validation.`);
    await loadSharedData();
  }
}

function readWorkbook(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = event => {
      try { resolve(XLSX.read(new Uint8Array(event.target.result), { type: "array" })); }
      catch (error) { reject(new Error("The workbook could not be read. Please check that it is a valid Excel file.")); }
    };
    reader.onerror = () => reject(new Error("The browser could not read the selected file."));
    reader.readAsArrayBuffer(file);
  });
}

function populateFilters() {
  const months = [...new Set(rawRows.map(row => row.month).filter(Boolean))].sort((a,b)=>monthIndex(a)-monthIndex(b));
  const locations = [...new Set(rawRows.map(row => row.location).filter(Boolean))].sort();
  const quarters = [...new Set(months.map(quarterOf).filter(Boolean))].sort();
  const currentMonth = $("monthFilter").value;
  const currentQuarter = $("quarterFilter").value;
  const currentLocation = $("locationFilter").value;

  $("monthFilter").innerHTML = '<option value="ALL">All months</option>' + months.map(month => `<option value="${escapeHtml(month)}">${escapeHtml(month)}</option>`).join("");
  $("quarterFilter").innerHTML = '<option value="ALL">All quarters</option>' + quarters.map(q => `<option value="${q}">${q}</option>`).join("");
  $("locationFilter").innerHTML = '<option value="ALL">All locations</option>' + locations.map(location => `<option value="${escapeHtml(location)}">${escapeHtml(location)}</option>`).join("");

  if (months.includes(currentMonth)) $("monthFilter").value = currentMonth;
  if (quarters.includes(currentQuarter)) $("quarterFilter").value = currentQuarter;
  if (locations.includes(currentLocation)) $("locationFilter").value = currentLocation;
}

function escapeHtml(value) { return String(value).replace(/[&<>"']/g, character => ({"&":"&amp;","<":"&lt;",">":"&gt;",'"':"&quot;","'":"&#039;"}[character])); }
function filteredRows() {
  const month = $("monthFilter").value;
  const quarter = $("quarterFilter").value;
  const location = $("locationFilter").value;
  const qMonths = quarterMonths(quarter);
  return rawRows.filter(row =>
    (month === "ALL" || row.month === month) &&
    (!qMonths || qMonths.has(row.month)) &&
    (location === "ALL" || row.location === location)
  );
}
function monthlyRows() {
  const groups = {};
  filteredRows().forEach(row => {
    if (!groups[row.month]) groups[row.month] = {
      month: row.month,
      inflow: 0,
      outflow: 0,
      hr: 0,
      training: 0,
      joined: 0,
      joiningWeighted: 0,
      batchCandidates: []
    };
    const g = groups[row.month];
    g.inflow += num(row.inflow);
    g.outflow += num(row.outflow);
    g.hr += num(row.hr);
    g.training += num(row.training);
    g.joined += num(row.joined);
    g.joiningWeighted += rateFraction(row.joiningThroughput) * num(row.outflow);
    if (num(row.batch) > 0) g.batchCandidates.push(num(row.batch));
  });
  return Object.values(groups)
    .map(g => ({ ...g, batch: g.batchCandidates.length ? Math.max(...g.batchCandidates) : 0 }))
    .sort((a, b) => monthIndex(a.month) - monthIndex(b.month));
}

// Batch conducted is a monthly metric in the source workbook. It must not be
// summed across locations because one batch can contain employees from several
// locations. The dashboard therefore derives one batch count per month and keeps
// that same monthly count when a location filter is applied.
function selectedBatchCount() {
  const month = $("monthFilter").value;
  const quarter = $("quarterFilter").value;
  const qMonths = quarterMonths(quarter);
  const monthGroups = {};
  rawRows.forEach(row => {
    if (!monthGroups[row.month]) monthGroups[row.month] = [];
    if (num(row.batch) > 0) monthGroups[row.month].push(num(row.batch));
  });
  const months = Object.keys(monthGroups).filter(m =>
    (month === "ALL" || m === month) && (!qMonths || qMonths.has(m))
  );
  return months.reduce((total, m) => {
    const values = monthGroups[m] || [];
    return total + (values.length ? Math.max(...values) : 0);
  }, 0);
}

function monthIndex(month) {
  const order = ["January","February","March","April","May","June","July","August","September","October","November","December"];
  const i = order.indexOf(month); return i === -1 ? 999 : i;
}
function quarterOf(month) {
  const index = monthIndex(month);
  if (index < 0 || index > 11) return "";
  return `Q${Math.floor(index / 3) + 1}`;
}
function quarterMonths(quarter) {
  if (quarter === "ALL") return null;
  const q = Number(String(quarter).replace("Q", ""));
  if (![1,2,3,4].includes(q)) return null;
  const start = (q - 1) * 3;
  const order = ["January","February","March","April","May","June","July","August","September","October","November","December"];
  return new Set(order.slice(start, start + 3));
}
function throughputRate(rows) { const inflow = sum(rows, "inflow"); return inflow ? (sum(rows, "joined") / inflow) * 100 : 0; }
function joiningThroughputRate(rows) { const outflow = sum(rows, "outflow"); return outflow ? rows.reduce((t, r) => t + rateFraction(r.joiningThroughput) * r.outflow, 0) / outflow * 100 : 0; }
function locationFilterLabel() {
  return $("locationFilter").value === "ALL" ? "monthly batches • shared metric" : "monthly batches • same monthly metric";
}

function render() {
  const rows = filteredRows();
  const months = monthlyRows();
  $("kpiBatches").textContent = format(selectedBatchCount());
  $("kpiBatches").nextElementSibling.textContent = locationFilterLabel();
  $("kpiInflow").textContent = format(sum(rows, "inflow"));
  $("kpiOutflow").textContent = format(sum(rows, "outflow"));
  $("kpiJoined").textContent = format(sum(rows, "joined"));
  $("kpiHr").textContent = format(sum(rows, "hr"));
  $("kpiTraining").textContent = format(sum(rows, "training"));
  $("kpiThroughput").textContent = pct(throughputRate(rows));
  $("kpiJoiningThroughput").textContent = pct(joiningThroughputRate(rows));

  const month = $("monthFilter").value;
  const quarter = $("quarterFilter").value;
  const location = $("locationFilter").value;
  const periodLabel = month !== "ALL" ? month : quarter !== "ALL" ? quarter : "All months";
  $("viewTitle").textContent = periodLabel + (location === "ALL" ? "" : ` • ${location}`);
  const throughputPanelNote = document.querySelector("#throughputChart")?.closest(".panel")?.querySelector(".panel-head p");
  const attritionPanelNote = document.querySelector("#attritionChart")?.closest(".panel")?.querySelector(".panel-head p");
  if (throughputPanelNote) throughputPanelNote.textContent = location === "ALL" ? "Monthly throughput and joining throughput." : `Monthly throughput and joining throughput • ${location}.`;
  if (attritionPanelNote) attritionPanelNote.textContent = location === "ALL" ? "HR and training attrition by month." : `HR and training attrition by month • ${location}.`;
  $("periodNote").textContent = rawRows.length ? `${rawRows.length} location-level records • shared centrally • monthly total rows excluded` : "No shared NHT data has been uploaded yet.";

  const groups = locationGroups(rows);
  const best = [...groups].sort((a,b) => b.rate-a.rate)[0];
  $("insightText").textContent = rows.length ? `${rows.length} location-level records are in view. ${best ? `${best.location} currently has the highest calculated throughput at ${pct(best.rate)}.` : ""}` : "No data matches the current filters.";

  drawThroughput(months); drawLocation(groups); drawMovement(months); drawAttrition(months); drawTable(groups); drawLiveInsights(groups, rows);
}
function format(value) { return Math.round(value).toLocaleString("en-IN"); }
function locationGroups(rows) {
  const groups = {};
  rows.forEach(row => {
    if (!groups[row.location]) groups[row.location] = { location: row.location, inflow: 0, outflow: 0, hr: 0, training: 0, joined: 0, joiningWeighted: 0 };
    const g = groups[row.location];
    g.inflow += row.inflow; g.outflow += row.outflow; g.hr += row.hr; g.training += row.training; g.joined += row.joined;
    g.joiningWeighted += rateFraction(row.joiningThroughput) * row.outflow;
  });
  return Object.values(groups).map(g => ({ ...g, rate: g.inflow ? g.joined/g.inflow*100 : 0, joinRate: g.outflow ? g.joiningWeighted/g.outflow*100 : 0, attrition: g.hr + g.training, attritionRate: g.inflow ? (g.hr+g.training)/g.inflow*100 : 0, dropOff: Math.max(g.inflow-g.outflow,0) }));
}
function destroy(name) { if (charts[name]) { charts[name].destroy(); charts[name] = null; } }
function commonScales() { return { responsive:true, maintainAspectRatio:false, animation:{duration:500}, plugins:{legend:{labels:{usePointStyle:true,boxWidth:8,font:{size:11}}}}, scales:{x:{grid:{display:false},ticks:{font:{size:10}}},y:{beginAtZero:true,ticks:{font:{size:10}}}}}; }

function drawThroughput(rows) {
  destroy("throughput");
  const labels = rows.map(r => r.month);
  const monthSelected = $("monthFilter").value !== "ALL";
  const quarterSelected = $("quarterFilter").value !== "ALL";
  const locationSelected = $("locationFilter").value !== "ALL";
  const throughputData = rows.map(r => r.inflow > 0 ? (r.joined / r.inflow) * 100 : null);
  const joiningData = rows.map(r => r.outflow > 0 ? (r.joiningWeighted / r.outflow) * 100 : null);
  const singlePeriod = monthSelected;
  const chartType = singlePeriod ? "bar" : "line";
  const title = monthSelected && locationSelected ? "Selected month • selected location" : monthSelected ? "Selected month" : quarterSelected && locationSelected ? `${$("quarterFilter").value} trend • selected location` : quarterSelected ? `${$("quarterFilter").value} monthly trend` : locationSelected ? "Monthly trend for selected location" : "Monthly throughput trend";
  charts.throughput = new Chart($("throughputChart"), {
    type: chartType,
    data: { labels, datasets: [
      {label:"Throughput", data:throughputData, tension:.35, borderWidth:3, pointRadius:4, pointHoverRadius:6, borderRadius:5, borderColor:"#0b5ed7", backgroundColor:"rgba(11,94,215,.85)"},
      {label:"Joining Throughput", data:joiningData, tension:.35, borderWidth:3, pointRadius:4, pointHoverRadius:6, borderRadius:5, borderDash: singlePeriod ? [] : [6,5], borderColor:"#e21d2f", backgroundColor:"rgba(226,29,47,.85)"}
    ]},
    options:{...commonScales(), scales:{...commonScales().scales, y:{beginAtZero:true,max:100,ticks:{callback:value=>`${value}%`}}}, plugins:{...commonScales().plugins, title:{display:true,text:title,align:"start",font:{size:12,weight:"600"},padding:{bottom:10}}, tooltip:{callbacks:{label:ctx=>`${ctx.dataset.label}: ${num(ctx.raw).toFixed(1)}%`}}}}
  });
}

function drawLocation(rows) { destroy("location"); const sorted=[...rows].sort((a,b)=>b.rate-a.rate); charts.location=new Chart($("locationChart"),{type:"bar",data:{labels:sorted.map(r=>r.location),datasets:[{label:"Throughput",data:sorted.map(r=>r.rate),borderRadius:6,backgroundColor:"#0b5ed7"}]},options:{indexAxis:"y",responsive:true,maintainAspectRatio:false,plugins:{legend:{display:false}},scales:{x:{beginAtZero:true,max:100,ticks:{callback:v=>`${v}%`},grid:{display:false}},y:{grid:{display:false},ticks:{font:{size:10}}}}}}); }
function drawMovement(rows) {
  destroy("movement");
  const monthSelected = $("monthFilter").value !== "ALL";
  const quarterSelected = $("quarterFilter").value !== "ALL";
  const locationSelected = $("locationFilter").value !== "ALL";
  const title = monthSelected && locationSelected ? "Selected month • selected location" : monthSelected ? "Selected month" : quarterSelected && locationSelected ? `${$("quarterFilter").value} inflow vs outflow • selected location` : quarterSelected ? `${$("quarterFilter").value} inflow vs outflow` : locationSelected ? "Monthly inflow vs outflow • selected location" : "Monthly inflow vs outflow";
  charts.movement=new Chart($("movementChart"),{type:"bar",data:{labels:rows.map(r=>r.month),datasets:[{label:"Inflow",data:rows.map(r=>r.inflow),borderRadius:5,backgroundColor:"#0b5ed7"},{label:"Outflow",data:rows.map(r=>r.outflow),borderRadius:5,backgroundColor:"#e21d2f"}]},options:{...commonScales(),plugins:{...commonScales().plugins,title:{display:true,text:title,align:"start",font:{size:12,weight:"600"},padding:{bottom:10}}}}});
}
function drawAttrition(rows) {
  destroy("attrition");
  const monthSelected = $("monthFilter").value !== "ALL";
  const quarterSelected = $("quarterFilter").value !== "ALL";
  const locationSelected = $("locationFilter").value !== "ALL";
  const title = monthSelected && locationSelected ? "Selected month • selected location" : monthSelected ? "Selected month" : quarterSelected && locationSelected ? `${$("quarterFilter").value} attrition • selected location` : quarterSelected ? `${$("quarterFilter").value} attrition trend` : locationSelected ? "Monthly attrition for selected location" : "Monthly attrition trend";
  const chartType = monthSelected ? "bar" : "line";
  charts.attrition = new Chart($("attritionChart"), {
    type: chartType,
    data: { labels: rows.map(r => r.month), datasets: [
      {label:"HR Attrition", data:rows.map(r=>r.hr), borderWidth:3, pointRadius:4, pointHoverRadius:6, borderRadius:5, borderColor:"#e21d2f", backgroundColor:"rgba(226,29,47,.85)"},
      {label:"Training Attrition", data:rows.map(r=>r.training), borderWidth:3, pointRadius:4, pointHoverRadius:6, borderRadius:5, borderColor:"#0b5ed7", backgroundColor:"rgba(11,94,215,.85)"}
    ]},
    options:{...commonScales(), scales:{...commonScales().scales, y:{beginAtZero:true,ticks:{precision:0}}}, plugins:{...commonScales().plugins, title:{display:true,text:title,align:"start",font:{size:12,weight:"600"},padding:{bottom:10}}, tooltip:{callbacks:{label:ctx=>`${ctx.dataset.label}: ${num(ctx.raw).toLocaleString("en-IN")}`}}}}
  });
}

function drawTable(rows) { const body=$("summaryTable"); if(!rows.length){body.innerHTML='<tr><td colspan="8" class="empty">No records match the selected filters.</td></tr>';return;} body.innerHTML=[...rows].sort((a,b)=>b.rate-a.rate).map(r=>`<tr><td>${escapeHtml(r.location)}</td><td>${format(r.inflow)}</td><td>${format(r.outflow)}</td><td>${format(r.hr)}</td><td>${format(r.training)}</td><td>${format(r.joined)}</td><td>${pct(r.rate)}</td><td>${pct(r.joinRate)}</td></tr>`).join(""); }

function drawLiveInsights(groups, rows) {
  const el = $("analysisGrid");
  if (!el) return;
  if (!groups.length) { el.innerHTML = '<div class="analysis-empty">No live insights available for the current filters.</div>'; return; }
  const highAttr = [...groups].sort((a,b)=>b.attrition-a.attrition)[0];
  const lowJoin = [...groups].sort((a,b)=>a.joinRate-b.joinRate)[0];
  const highThroughput = [...groups].sort((a,b)=>b.rate-a.rate)[0];
  const highInflow = [...groups].sort((a,b)=>b.inflow-a.inflow)[0];
  const largestDrop = [...groups].sort((a,b)=>b.dropOff-a.dropOff)[0];
  const cards = [
    {tag:"HIGH ATTRITION", title:highAttr.location, value:`${format(highAttr.attrition)} exits`, text:`HR + training attrition is ${pct(highAttr.attritionRate)} of inflow.`},
    {tag:"LOW JOINING THROUGHPUT", title:lowJoin.location, value:pct(lowJoin.joinRate), text:`Lowest joining throughput in the current selection.`},
    {tag:"BEST THROUGHPUT", title:highThroughput.location, value:pct(highThroughput.rate), text:`Highest joined-to-inflow conversion in view.`},
    {tag:"LARGEST INFLOW", title:highInflow.location, value:format(highInflow.inflow), text:`Largest candidate inflow in the current selection.`},
    {tag:"BIGGEST DROP-OFF", title:largestDrop.location, value:format(largestDrop.dropOff), text:`Inflow minus outflow, highlighting the largest operational drop-off.`}
  ];
  el.innerHTML = cards.map((c,i)=>`<article class="analysis-card" style="--delay:${i*90}ms"><span>${c.tag}</span><h4>${escapeHtml(c.title)}</h4><strong>${c.value}</strong><p>${c.text}</p></article>`).join("");
  startInsightTicker();
}


function startInsightTicker() {
  if (window.__insightTickerStarted) return;
  window.__insightTickerStarted = true;
  setInterval(() => {
    const cards = [...document.querySelectorAll(".analysis-card")];
    if (!cards.length) return;
    cards.forEach(card => card.classList.remove("insight-active"));
    const index = (Number(document.body.dataset.insightIndex || 0) + 1) % cards.length;
    document.body.dataset.insightIndex = index;
    cards[index].classList.add("insight-active");
  }, 2800);
}

function downloadChart(canvasId, name) {
  const chart = charts[canvasId.replace("Chart", "")];
  if (!chart) return;
  const link = document.createElement("a");
  link.download = `NHT-${name}-${new Date().toISOString().slice(0,10)}.png`;
  link.href = chart.toBase64Image("image/png", 1);
  link.click();
}

function subscribeToRealtime() {
  if (realtimeChannel) supabaseClient.removeChannel(realtimeChannel);
  realtimeChannel = supabaseClient.channel("nht-dashboard-live")
    .on("postgres_changes", {event:"*",schema:"public",table:"nht_dashboard_data"}, async()=>{setSync("Updating…","loading");await loadSharedData();})
    .subscribe(status=>{if(status==="SUBSCRIBED")setSync(`Live • ${rawRows.length} records`,"live");});
}

function updateIdleIndicator() {
  const el = $("idleStatus");
  if (!el || !currentUser || !lastActivityAt) return;
  const remaining = Math.max(0, INACTIVITY_MS - (Date.now() - lastActivityAt));
  const mins = Math.floor(remaining / 60000);
  const secs = Math.floor((remaining % 60000) / 1000).toString().padStart(2,"0");
  el.textContent = `Auto sign-out in ${mins}:${secs} idle`;
  el.classList.toggle("idle-warning", remaining <= 60000);
}
function resetInactivityTimer() {
  if (!currentUser) return;
  const now = Date.now();
  if (now - lastActivityAt < 1000) return;
  lastActivityAt = now;
  clearTimeout(inactivityTimer);
  inactivityTimer = setTimeout(async()=>{
    if (!currentUser) return;
    if (realtimeChannel) await supabaseClient.removeChannel(realtimeChannel);
    await supabaseClient.auth.signOut();
    alert("You have been signed out after 5 minutes of inactivity.");
  }, INACTIVITY_MS);
  updateIdleIndicator();
}
function startInactivityMonitor() {
  if (window.__idleMonitorStarted) { resetInactivityTimer(); return; }
  window.__idleMonitorStarted = true;
  ["click","keydown","mousemove","scroll","touchstart"].forEach(eventName => window.addEventListener(eventName, resetInactivityTimer, {passive:true}));
  inactivityInterval = setInterval(updateIdleIndicator, 1000);
  resetInactivityTimer();
}

function showDashboard(user) {
  $("authLoading").classList.add("hidden"); $("loginScreen").classList.add("hidden"); $("appShell").classList.remove("hidden");
  $("signedInAs").textContent = user?.email || "Signed in";
  currentUser = user;
  if (!window.__dashboardInitialized) {
    window.__dashboardInitialized = true;
    $("excelFile").addEventListener("change",event=>{if(event.target.files[0]){uploadToSharedDatabase(event.target.files[0]);event.target.value="";}});
    $("monthFilter").addEventListener("change",()=>{ if ($("monthFilter").value !== "ALL") $("quarterFilter").value = quarterOf($("monthFilter").value); render(); });
    $("quarterFilter").addEventListener("change",()=>{ if ($("quarterFilter").value !== "ALL") $("monthFilter").value = "ALL"; render(); });
    $("locationFilter").addEventListener("change",render);
    document.querySelectorAll(".chart-download").forEach(button=>button.addEventListener("click",()=>downloadChart(button.dataset.chart,button.dataset.name)));
  }
  startInactivityMonitor();
  resetInactivityTimer();
  loadRole(user).then(()=>loadSharedData().then(subscribeToRealtime));
}
function showLogin(message="") { clearTimeout(inactivityTimer); currentUser=null; $("authLoading").classList.add("hidden"); $("appShell").classList.add("hidden"); $("loginScreen").classList.remove("hidden"); $("loginError").textContent=message; }
async function signIn(email,password){$("loginButton").disabled=true;$("loginButton").textContent="Signing in…";$("loginError").textContent="";const{error}=await supabaseClient.auth.signInWithPassword({email,password});if(error)$("loginError").textContent="Unable to sign in. Please check your email and password.";$("loginButton").disabled=false;$("loginButton").textContent="Sign in";}
async function startAuth(){try{const{data,error}=await supabaseClient.auth.getSession();if(error){showLogin("Unable to check your secure session. Please try again.");return;}if(data.session?.user)showDashboard(data.session.user);else showLogin();supabaseClient.auth.onAuthStateChange((_event,session)=>{if(session?.user)showDashboard(session.user);else showLogin();});}catch(error){console.error(error);showLogin("Unable to connect to secure authentication.");}}

$("loginForm").addEventListener("submit",async event=>{event.preventDefault();await signIn($("loginEmail").value.trim(),$("loginPassword").value);});
$("signOutButton").addEventListener("click",async()=>{clearTimeout(inactivityTimer);if(realtimeChannel)await supabaseClient.removeChannel(realtimeChannel);await supabaseClient.auth.signOut();});
startAuth();
