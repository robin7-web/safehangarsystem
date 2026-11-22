/* ============================================================
   SafeHangar - Google Sheets edition
   Data layer: talks to a Google Apps Script Web App that
   reads/writes a Google Sheet. Falls back to in-page demo
   data when no API is configured, so the file still runs
   locally (e.g. with the Live Server extension).
   ============================================================ */

/* 1. CONFIG - paste your Web App /exec URL here after deploying. */
const API_URL = "https://script.google.com/macros/s/AKfycbyZZ7QUloInlCQ6l7YCXDlHnqQyg8ocucqf-CLieulIdueNKmG4HGD2EvZiYF0QVM0i/exec"; // e.g. "https://script.google.com/macros/s/AKfycb.../exec"

/* 2. App state */
const state = {
  aircraft: [], inspections: [], hazards: [], facilities: [], users: [],
  online: false
};

/* 3. Generic API helper. POSTs JSON (avoids most CORS preflight
      quirks); falls back to the bundled demo data on failure. */
async function api(action, payload = {}) {
  if (!API_URL) throw new Error("API_URL not configured");
  const res = await fetch(API_URL, {
    method: "POST",
    // text/plain avoids the CORS preflight that Apps Script cannot answer
    headers: { "Content-Type": "text/plain;charset=utf-8" },
    body: JSON.stringify(Object.assign({ action }, payload))
  });
  const json = await res.json();
  if (!json.ok) throw new Error(json.error || "API error");
  return json.data;
}

/* ---------- Page navigation (unchanged behaviour) ---------- */
const pageNames = { dashboard: "Dashboard", aircraft: "Aircraft", inspections: "Inspections", reports: "Reports", alerts: "Safety Alerts", hazards: "Hazard Reports", facilities: "Hangar Directory", users: "User Management", settings: "Settings" };
function showPage(name) {
  document.querySelectorAll(".page").forEach(p => p.classList.remove("active-page"));
  const target = document.getElementById(name);
  if (target) target.classList.add("active-page");
  document.querySelectorAll(".nav-link").forEach(a => a.classList.toggle("active", a.dataset.page === name));
  document.getElementById("pageTitle").textContent = pageNames[name] || "Dashboard";
  window.location.hash = name; 
  window.scrollTo({
  top: 0,
  behavior: "instant"
});
}

document.querySelectorAll(".nav-link").forEach(a => a.addEventListener("click", e => { e.preventDefault(); showPage(a.dataset.page); }));

/* ---------- Reference data (14 hangar hazard categories) ---------- */
const HAZARDS = [
  ["Fire Hazard", "Flammable materials stored near welding or electrical equipment"],
  ["Fuel or Chemical Spillage", "Aviation fuel or hydraulic fluid leaking onto the hangar floor"],
  ["Slip, Trip and Fall", "Wet floor, scattered tools, or cables obstructing walkways"],
  ["Falling Objects", "Tools or aircraft components falling from elevated maintenance platforms"],
  ["Electrical Hazard", "Exposed electrical wires, damaged extension cords, or overloaded sockets"],
  ["Foreign Object Debris (FOD)", "Loose bolts, screws, or tools left near or inside an aircraft"],
  ["Equipment Malfunction", "Defective aircraft jacks, hydraulic lifts, or maintenance equipment"],
  ["Aircraft Ground Collision", "Aircraft wing contacting a hangar structure or ground equipment"],
  ["Vehicle-Related Incident", "Aircraft towing vehicle or forklift nearly hitting a worker"],
  ["Improper PPE Usage", "Worker handling hazardous chemicals without appropriate protective equipment"],
  ["Hazardous Chemical Exposure", "Inhalation of paint fumes, cleaning solvents, or other harmful substances"],
  ["Unsafe Working at Heights", "Technician working on an aircraft without adequate fall protection"],
  ["Improper Maintenance Procedures", "Incorrect installation, missing components, or incomplete maintenance documentation"],
  ["Unsafe Aircraft Movement", "Aircraft being moved without adequate clearance or coordination"]
];

const todayISO = () => new Date().toISOString().slice(0, 10);
const sevClass = v => v === "Low" ? "success" : (v === "Medium" ? "warning" : "danger");
const resultClass = v => v === "Passed" ? "success" : (v === "Failed" ? "danger" : "warning");
const fmtDate = iso => {
  if (!iso) return "—";
  const d = new Date(iso);
  return isNaN(d) ? String(iso) : d.toLocaleDateString("en-US", { month: "short", day: "2-digit", year: "numeric" });
};

/* ---------- Demo fallback data (mirrors the seed rows in Code.gs) ---------- */
const DEMO = {
  facilities: [
    { ID: "FA-0001", Name: "Lufthansa Technik Philippines", Address: "MacroAsia Special Economic Zone, Villamor Air Base, Pasay City, Metro Manila 1309", Region: "Metro Manila" },
    { ID: "FA-0002", Name: "Aviation Partnership (Philippines) Corporation (Aplus)", Address: "Building 8006, Domestic Road, Pasay City, Metro Manila 1301", Region: "Metro Manila" },
    { ID: "FA-0003", Name: "INAEC Aviation Corporation — Hangars A & B", Address: "General Aviation Area, Manila Domestic Airport, Pasay City, Metro Manila", Region: "Metro Manila" },
    { ID: "FA-0004", Name: "INAEC Aviation Corporation — Hangar D", Address: "Andrews Avenue, Manila Domestic Airport, Pasay City, Metro Manila", Region: "Metro Manila" },
    { ID: "FA-0005", Name: "PhilJets Aero Services Inc.", Address: "AAOP Hangar 2, Manila Domestic Airport, Domestic Road corner MIAA Road, Pasay City, Metro Manila 1300", Region: "Metro Manila" },
    { ID: "FA-0006", Name: "Asian Aerospace Corporation", Address: "Lima Gate, Andrews Avenue, NAIA Complex, Pasay City, Metro Manila 1300", Region: "Metro Manila" },
    { ID: "FA-0007", Name: "SIA Engineering (Philippines) Corporation", Address: "Building 7592, A. Bonifacio Avenue, Civil Aviation Complex, Clark Freeport Zone, Pampanga 2023", Region: "Clark, Pampanga" },
    { ID: "FA-0008", Name: "Metrojet Engineering (Clark) Ltd.", Address: "Building N7727, M.A. Roxas Highway, Clark Civil Aviation Complex, Clark Freeport Zone, Pampanga 2023", Region: "Clark, Pampanga" },
    { ID: "FA-0009", Name: "Dornier Technology, Inc.", Address: "Hangar 7224, Clark International Airport, Clark Freeport Zone, Pampanga", Region: "Clark, Pampanga" },
    { ID: "FA-0010", Name: "Asian Aerospace Corporation — Clark Facility", Address: "Clark Freeport Zone, Pampanga", Region: "Clark, Pampanga" },
    { ID: "FA-0011", Name: "INAEC North", Address: "Building 7216, Civil Aviation Complex, Clark International Airport, Pampanga", Region: "Clark, Pampanga" },
    { ID: "FA-0012", Name: "First Aviation Academy, Inc. (FAA)", Address: "Building 8303, Southwest Apron B, Argonaut Highway, Subic Bay International Airport, Subic Bay Freeport Zone, Zambales 2200", Region: "Subic, Zambales" }
  ],
  aircraft: [
    { ID: "AI-0001", Registration: "RP-C1234", Model: "Cessna 172", "Serial Number": "172-8841", Operator: "Training Fleet", Status: "Active" },
    { ID: "AI-0002", Registration: "RP-C5678", Model: "Piper PA-28", "Serial Number": "PA28-2291", Operator: "Private", Status: "Active" },
    { ID: "AI-0003", Registration: "RP-C9012", Model: "Cessna 152", "Serial Number": "C152-3377", Operator: "Training Fleet", Status: "Active" },
    { ID: "AI-0004", Registration: "RP-C3456", Model: "Cessna 208 Caravan", "Serial Number": "208B-0102", Operator: "Private", Status: "Active" },
    { ID: "AI-0005", Registration: "RP-C7890", Model: "Beechcraft Baron 58", "Serial Number": "BE-58-4410", Operator: "Private", Status: "Maintenance" }
  ],
  inspections: [
    { ID: "IN-0001", Aircraft: "RP-C1234", "Inspection Type": "Pre-flight inspection", Inspector: "J. Mendoza", Date: "2026-11-15", "Overall Result": "Passed" },
    { ID: "IN-0002", Aircraft: "RP-C5678", "Inspection Type": "Engine inspection", Inspector: "R. Santos", Date: "2026-11-14", "Overall Result": "Needs review" },
    { ID: "IN-0003", Aircraft: "RP-C9012", "Inspection Type": "Structural inspection", Inspector: "A. Cruz", Date: "2026-11-13", "Overall Result": "Failed" }
  ],
  hazards: [
    { ID: "HZ-0001", Category: "Fuel or Chemical Spillage", Severity: "High", Facility: "Lufthansa Technik Philippines", Location: "Bay 2, near the wing stand", "Reported By": "M. Reyes", "Date Observed": "2026-11-15", Status: "Open", Description: "Hydraulic fluid puddle left unmarked on the hangar floor." },
    { ID: "HZ-0002", Category: "Foreign Object Debris (FOD)", Severity: "Medium", Facility: "SIA Engineering (Philippines) Corporation", Location: "Apron walkway B", "Reported By": "K. Lim", "Date Observed": "2026-11-12", Status: "Closed", Description: "Loose bolts and rivets found near the tow path." }
  ],
  users: [
    { ID: "US-0001", Name: "System Admin", Role: "Administrator", Email: "admin@safehangar.example", Status: "Active" },
    { ID: "US-0002", Name: "J. Mendoza", Role: "Inspector", Email: "jmendoza@safehangar.example", Status: "Active" },
    { ID: "US-0003", Name: "R. Santos", Role: "Inspector", Email: "rsantos@safehangar.example", Status: "Active" }
  ]
};

/* ---------- Rendering ---------- */
function setStatus(online, msg) {
  state.online = online;
  const dot = document.getElementById("statusDot");
  dot.className = "status-dot" + (online ? "" : " offline");
  const txt = document.getElementById("statusText");
  txt.textContent = msg;
  txt.className = "online" + (online ? "" : " offline-text");
  const footer = document.getElementById("footerMode");
  if (footer) footer.textContent = online ? "Connected to Google Sheets" : "Offline demo mode";
  const banner = document.getElementById("offlineBanner");
  if (banner) banner.classList.toggle("hidden", online);
  const pill = document.getElementById("connPill");
  if (pill) {
    if (online) { pill.textContent = "Live — Google Sheets"; pill.className = "pill success"; }
    else if (API_URL) { pill.textContent = "Offline — API unreachable"; pill.className = "pill danger"; }
    else { pill.textContent = "Offline demo mode"; pill.className = "pill warning"; }
  }
}

function setConnPanel(mode, summary, errorText) {
  // mode: "setup" (no URL), "live", "error" (URL set but unreachable)
  const setup = document.getElementById("connSetup");
  const live = document.getElementById("connLive");
  const err = document.getElementById("connError");
  if (!setup || !live || !err) return;
  setup.classList.toggle("hidden", mode !== "setup");
  live.classList.toggle("hidden", mode !== "live");
  err.classList.toggle("hidden", mode !== "error");
  if (summary) document.getElementById("connSummary").textContent = summary;
  if (errorText) document.getElementById("connErrorText").textContent = errorText;
}

function renderAll() {
  renderStats(); renderRecent(); renderDashAlerts(); renderAircraft();
  renderReports(); renderHazards(); renderAlerts(); renderUsers(); renderFacilities();
  updateAlertCounts();
}

function renderStats() {
  document.getElementById("todayLabel").textContent =
    new Date().toLocaleDateString("en-US", { weekday: "long", month: "long", day: "2-digit", year: "numeric" }).toUpperCase();
  document.getElementById("aircraftCount").textContent = state.aircraft.length;
  const pending = state.inspections.filter(i => i["Overall Result"] === "Needs review" || i["Overall Result"] === "Failed").length;
  const thisMonth = new Date().toISOString().slice(0, 7);
  const completed = state.inspections.filter(i => i["Overall Result"] === "Passed" && String(i.Date).slice(0, 7) === thisMonth).length;
  document.getElementById("pendingCount").textContent = String(pending).padStart(2, "0");
  document.getElementById("completedCount").textContent = String(completed).padStart(2, "0");
}

function openAlertItems() {
  // Open alerts = high/critical open hazards + failed / needs-review inspections
  const items = [];
  state.hazards.filter(h => h.Status !== "Closed" && (h.Severity === "High" || h.Severity === "Critical")).forEach(h => {
    items.push({ title: h.Category + " reported", sub: h.Facility + " · " + h.Location, when: "Open · " + h.Severity + " · " + fmtDate(h["Date Observed"]), cls: h.Severity === "Critical" ? "red-bg" : "amber-bg" });
  });
  state.inspections.filter(i => i["Overall Result"] === "Failed" || i["Overall Result"] === "Needs review").forEach(i => {
    items.push({ title: i["Overall Result"] === "Failed" ? "Structural defect recorded" : "Engine component requires review", sub: i.Aircraft + " · " + i["Inspection Type"], when: "Open · " + fmtDate(i.Date), cls: i["Overall Result"] === "Failed" ? "red-bg" : "amber-bg" });
  });
  return items;
}

function updateAlertCounts() {
  const n = openAlertItems().length;
  document.getElementById("alertCount").textContent = String(n).padStart(2, "0");
  document.getElementById("alertNavCount").textContent = n;
  document.getElementById("alertBadge").textContent = n + " OPEN";
}

function alertItemHTML(it) {
  return '<div class="alert-item"><span class="alert-symbol ' + it.cls + '">!</span><div><strong>' +
    escapeHTML(it.title) + '</strong><p>' + escapeHTML(it.sub) + '</p><small>' + escapeHTML(it.when) + '</small></div></div>';
}

function renderDashAlerts() {
  const items = openAlertItems().slice(0, 3);
  document.getElementById("dashAlerts").innerHTML = items.length ? items.map(alertItemHTML).join("") :
    '<p class="helper">No open alerts. All clear.</p>';
}

function renderAlerts() {
  const items = openAlertItems();
  document.getElementById("alertList").innerHTML = items.length ? items.map(alertItemHTML).join("") :
    '<p class="helper">No open alerts. All clear.</p>';
}

function renderRecent() {
  const rows = state.inspections.slice().sort((a, b) => String(b.Date).localeCompare(String(a.Date))).slice(0, 5);
  document.getElementById("recentRows").innerHTML = rows.map(r => {
    const ac = state.aircraft.find(a => a.Registration === r.Aircraft);
    return "<tr><td><strong>" + escapeHTML(r.Aircraft) + "</strong><small>" + escapeHTML(ac ? ac.Model : "") + "</small></td><td>" +
      escapeHTML(r["Inspection Type"]) + "</td><td>" + fmtDate(r.Date) + "</td><td>" +
      '<span class="pill ' + resultClass(r["Overall Result"]) + '">' + escapeHTML(r["Overall Result"]) + "</span></td></tr>";
  }).join("") || '<tr><td colspan="4">No inspections recorded yet.</td></tr>';
}

function renderAircraft() {
  const q = document.getElementById("aircraftSearch").value.toLowerCase();
  const list = state.aircraft.filter(a => (a.Registration + " " + a.Model + " " + a["Serial Number"] + " " + a.Operator).toLowerCase().includes(q));
  document.getElementById("aircraftRows").innerHTML = list.map(a =>
    "<tr><td><strong>" + escapeHTML(a.Registration) + "</strong></td><td>" + escapeHTML(a.Model) + "</td><td>" +
    escapeHTML(a["Serial Number"] || "—") + "</td><td>" + escapeHTML(a.Operator || "—") + "</td><td>" +
    '<span class="pill ' + (a.Status === "Active" ? "success" : "warning") + '">' + escapeHTML(a.Status || "Active") + "</span></td></tr>"
  ).join("") || '<tr><td colspan="5">No aircraft match your search.</td></tr>';
  document.getElementById("aircraftDatalist").innerHTML = state.aircraft.map(a =>
    '<option value="' + escapeHTML(a.Registration) + '"></option>').join("");
}

function renderReports() {
  const rows = state.inspections.slice().sort((a, b) => String(b.Date).localeCompare(String(a.Date)));
  document.getElementById("reportRows").innerHTML = rows.map(r =>
    "<tr><td>" + fmtDate(r.Date) + "</td><td>" + escapeHTML(r.Aircraft) + "</td><td>" + escapeHTML(r["Inspection Type"]) +
    "</td><td>" + escapeHTML(r.Inspector) + '</td><td><span class="pill ' + resultClass(r["Overall Result"]) + '">' +
    escapeHTML(r["Overall Result"]) + "</span></td></tr>"
  ).join("") || '<tr><td colspan="5">No inspection records yet.</td></tr>';
}

function renderHazards() {
  const rows = state.hazards.slice().sort((a, b) => String(b["Date Observed"]).localeCompare(String(a["Date Observed"])));
  document.getElementById("hazardRows").innerHTML = rows.map(h =>
    "<tr><td>" + fmtDate(h["Date Observed"]) + "</td><td>" + escapeHTML(h.Category) + "</td><td>" + escapeHTML(h.Facility) +
    "</td><td>" + escapeHTML(h.Location) + "</td><td>" + escapeHTML(h["Reported By"]) + '</td><td><span class="pill ' +
    sevClass(h.Severity) + '">' + escapeHTML(h.Severity) + "</span></td></tr>"
  ).join("") || '<tr><td colspan="6">No hazard reports yet. Use the form above to submit one.</td></tr>';
}

function renderUsers() {
  document.getElementById("userRows").innerHTML = state.users.map(u =>
    "<tr><td><strong>" + escapeHTML(u.Name) + "</strong></td><td>" + escapeHTML(u.Role) + "</td><td>" +
    escapeHTML(u.Email || "—") + '</td><td><span class="pill ' + (u.Status === "Active" ? "success" : "warning") + '">' +
    escapeHTML(u.Status || "Active") + "</span></td></tr>"
  ).join("") || '<tr><td colspan="4">No users.</td></tr>';
}

/* ---------- Facility directory ---------- */
let activeRegion = "All";
function renderFacilities() {
  const q = document.getElementById("facilitySearch").value.toLowerCase();
  const regions = ["All", ...new Set(state.facilities.map(f => f.Region))];
  const list = document.getElementById("facilityList"); list.replaceChildren();
  let shown = 0;
  regions.slice(1).forEach(r => {
    if (activeRegion !== "All" && activeRegion !== r) return;
    const items = state.facilities.filter(f => f.Region === r && (f.Name + " " + f.Address).toLowerCase().includes(q));
    if (!items.length) return;
    const h = document.createElement("h4"); h.className = "region-heading"; h.textContent = r; list.appendChild(h);
    const grid = document.createElement("div"); grid.className = "facility-grid";
    items.forEach(f => { shown++;
      const card = document.createElement("div"); card.className = "facility-card";
      const t = document.createElement("strong"); t.textContent = f.Name;
      const a = document.createElement("p"); a.textContent = f.Address;
      const b = document.createElement("button"); b.className = "text-btn"; b.type = "button"; b.textContent = "Report a hazard here";
      b.addEventListener("click", () => { document.getElementById("hazardFacility").value = f.Name; showPage("hazards"); });
      card.append(t, a, b); grid.appendChild(card);
    });
    list.appendChild(grid);
  });
  if (!shown) { const p = document.createElement("p"); p.className = "helper"; p.textContent = "No facilities match your search. Try a different name or region."; list.appendChild(p); }
  document.getElementById("facilityCount").textContent = shown + (shown === 1 ? " facility" : " facilities");
}

function buildRegionChips() {
  const regions = ["All", ...new Set(state.facilities.map(f => f.Region))];
  const chips = document.getElementById("regionChips"); chips.replaceChildren();
  regions.forEach(r => {
    const c = document.createElement("button"); c.type = "button";
    c.className = "chip" + (r === "All" ? " active" : ""); c.textContent = r === "All" ? "All regions" : r;
    c.addEventListener("click", () => { activeRegion = r; chips.querySelectorAll(".chip").forEach(x => x.classList.toggle("active", x === c)); renderFacilities(); });
    chips.appendChild(c);
  });
}

/* ---------- Hazard reference table + form ---------- */
function renderHazardRef() {
  const q = document.getElementById("hazardSearch").value.toLowerCase();
  const body = document.getElementById("hazardRef"); body.replaceChildren();
  HAZARDS.filter(h => (h[0] + " " + h[1]).toLowerCase().includes(q)).forEach(([c, ex]) => {
    const tr = document.createElement("tr");
    const a = document.createElement("td"); const s = document.createElement("strong"); s.textContent = c; a.appendChild(s);
    const b = document.createElement("td"); b.textContent = ex; b.style.whiteSpace = "normal";
    const d = document.createElement("td"); const btn = document.createElement("button"); btn.type = "button"; btn.className = "text-btn"; btn.textContent = "Report this";
    btn.addEventListener("click", () => { document.getElementById("hazardCategory").value = c; document.getElementById("hazardCategory").dispatchEvent(new Event("change")); document.getElementById("hazardForm").scrollIntoView({ behavior: "smooth", block: "start" }); });
    d.appendChild(btn); tr.append(a, b, d); body.appendChild(tr);
  });
}

/* ---------- Small util ---------- */
function escapeHTML(s) {
  return String(s == null ? "" : s).replace(/[&<>"']/g, m => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[m]));
}

/* ---------- Forms ---------- */
document.getElementById("aircraftForm").addEventListener("submit", async e => {
  e.preventDefault();
  const rec = {
    Registration: document.getElementById("regNumber").value.trim().toUpperCase(),
    Model: document.getElementById("model").value.trim(),
    "Serial Number": document.getElementById("serial").value.trim() || "—",
    Operator: document.getElementById("operator").value.trim() || "—",
    Status: "Active"
  };
  try {
    if (state.online) { const saved = await api("create", { entity: "aircraft", record: rec }); state.aircraft.push(saved); }
    else { rec.ID = "AI-" + String(state.aircraft.length + 1).padStart(4, "0"); state.aircraft.push(rec); }
    e.target.reset(); e.target.classList.add("hidden");
    renderAll();
    notify("Aircraft saved" + (state.online ? " to Google Sheets." : " in offline demo mode."));
  } catch (err) { notify("Save failed: " + err.message, true); }
});

document.getElementById("inspectionForm").addEventListener("submit", async e => {
  e.preventDefault();
  const checks = {};
  document.querySelectorAll(".result-select").forEach(s => checks[s.dataset.item] = s.value);
  const results = Object.values(checks);
  const overall = results.includes("Fail") ? "Failed" : results.includes("Needs review") ? "Needs review" : "Passed";
  const rec = {
    Aircraft: document.getElementById("inspectionAircraft").value.trim().toUpperCase(),
    "Inspection Type": document.getElementById("inspectionType").value,
    Inspector: document.getElementById("inspector").value.trim(),
    Date: document.getElementById("inspectionDate").value,
    Fuselage: checks.Fuselage, Wings: checks.Wings, Engine: checks.Engine, Gear: checks.Gear,
    "Overall Result": overall,
    Remarks: document.getElementById("remarks").value.trim()
  };
  try {
    if (state.online) { const saved = await api("create", { entity: "inspections", record: rec }); state.inspections.push(saved); }
    else { rec.ID = "IN-" + String(state.inspections.length + 1).padStart(4, "0"); state.inspections.push(rec); }
    e.target.reset();
    document.getElementById("inspectionDate").value = todayISO();
    renderAll(); showPage("reports");
    notify("Inspection saved" + (state.online ? " to Google Sheets." : " in offline demo mode."));
  } catch (err) { notify("Save failed: " + err.message, true); }
});

document.getElementById("hazardForm").addEventListener("submit", async e => {
  e.preventDefault();
  const cat = document.getElementById("hazardCategory").value;
  const rec = {
    Category: cat,
    Severity: document.getElementById("hazardSeverity").value,
    Facility: document.getElementById("hazardFacility").value,
    Location: document.getElementById("hazardLocation").value.trim(),
    "Reported By": document.getElementById("hazardReporter").value.trim(),
    "Date Observed": document.getElementById("hazardDate").value,
    Description: document.getElementById("hazardDescription").value.trim(),
    Status: "Open"
  };
  const raised = rec.Severity === "High" || rec.Severity === "Critical";
  try {
    if (state.online) { const saved = await api("create", { entity: "hazards", record: rec }); state.hazards.push(saved); }
    else { rec.ID = "HZ-" + String(state.hazards.length + 1).padStart(4, "0"); state.hazards.push(rec); }
    e.target.reset();
    document.getElementById("hazardCategory").value = "";
    document.getElementById("hazardExample").textContent = "Select a category to see an example of what to report.";
    document.getElementById("hazardDate").value = todayISO();
    renderAll();
    notify("Hazard report saved" + (state.online ? " to Google Sheets." : " in offline demo mode.") + (raised ? " A safety alert was created." : ""));
  } catch (err) { notify("Save failed: " + err.message, true); }
});

function notify(msg, isError) {
  alert(msg);
  if (isError) console.error(msg);
}

/* ---------- Search inputs ---------- */
document.getElementById("aircraftSearch").addEventListener("input", renderAircraft);
document.getElementById("facilitySearch").addEventListener("input", renderFacilities);
document.getElementById("hazardSearch").addEventListener("input", renderHazardRef);

/* ---------- Static form setup ---------- */
const catSel = document.getElementById("hazardCategory"), facSel = document.getElementById("hazardFacility"), exampleBox = document.getElementById("hazardExample");
const ph = document.createElement("option"); ph.value = ""; ph.textContent = "Select a category"; ph.disabled = true; ph.selected = true; catSel.appendChild(ph);
HAZARDS.forEach(([c]) => { const o = document.createElement("option"); o.textContent = c; catSel.appendChild(o); });
catSel.addEventListener("change", () => {
  const h = HAZARDS.find(x => x[0] === catSel.value);
  exampleBox.textContent = h ? "Example: " + h[1] : "Select a category to see an example of what to report.";
});
document.getElementById("inspectionDate").value = todayISO();
document.getElementById("hazardDate").value = todayISO();

/* ---------- Boot: load everything ---------- */
function applyData(data) {
  state.aircraft = data.aircraft; state.inspections = data.inspections;
  state.hazards = data.hazards; state.facilities = data.facilities; state.users = data.users;
}

function countsLine() {
  const part = (n, s) => n + " " + s + (n === 1 ? "" : "s");
  return part(state.aircraft.length, "aircraft") + ", " +
    part(state.inspections.length, "inspection") + ", " +
    part(state.hazards.length, "hazard") + ", " +
    part(state.facilities.length, "facility") + ", " +
    part(state.users.length, "user");
}

async function boot() {
  if (API_URL) {
    try {
      applyData(await api("bootstrap"));
      setStatus(true, "System online");
      setConnPanel("live",
        "Connected to the Apps Script Web App (" + API_URL.slice(0, 60) + "…). All records are read from and saved to your Google Sheet. Counts: " + countsLine() + ".");
      document.getElementById("settingsApiInfo").textContent = "Connected to the Apps Script Web App. Data is stored in the linked Google Sheet.";
    } catch (err) {
      // Keep whatever is in state (previously loaded live data if this is a
      // refresh); only fall back to demo data when there is nothing to keep.
      if (!state.aircraft.length && !state.facilities.length) loadDemo();
      setStatus(false, "Offline — API unreachable");
      document.getElementById("offlineReason").textContent = "Reason: " + err.message;
      setConnPanel("error", null,
        "The Web App URL is set but could not be reached (" + err.message + "). Check that the deployment exists, access is 'Anyone', and that you deployed the latest version. The data below may be stale.");
      document.getElementById("settingsApiInfo").textContent = "API error: " + err.message + ".";
    }
  } else {
    loadDemo();
    setStatus(false, "Offline demo mode");
    document.getElementById("offlineReason").textContent = "Reason: API_URL in script.js is empty.";
    setConnPanel("setup");
  }
  // facility dropdown is filled after data load
  facSel.replaceChildren();
  state.facilities.forEach(f => { const o = document.createElement("option"); o.textContent = f.Name; facSel.appendChild(o); });
  buildRegionChips(); renderHazardRef(); renderAll();
}

/* Reload everything from the Sheet (Settings button). Never swaps in
   demo data: if the fetch fails, live data stays on screen. */
async function refreshData() {
  if (!API_URL) { setConnPanel("setup"); return; }
  const btn = document.getElementById("connPill");
  if (btn) { btn.textContent = "Refreshing…"; btn.className = "pill warning"; }
  try {
    applyData(await api("bootstrap"));
    setStatus(true, "System online");
    setConnPanel("live", "Refreshed from Google Sheets. Counts: " + countsLine() + ".");
    facSel.replaceChildren();
    state.facilities.forEach(f => { const o = document.createElement("option"); o.textContent = f.Name; facSel.appendChild(o); });
    buildRegionChips(); renderAll();
    notify("Data refreshed from Google Sheets.");
  } catch (err) {
    setStatus(false, "Offline — API unreachable");
    setConnPanel("error", null, "Refresh failed: " + err.message + ". The data on screen is unchanged.");
  }
}

function loadDemo() {
  state.aircraft = DEMO.aircraft.slice();
  state.inspections = DEMO.inspections.slice();
  state.hazards = DEMO.hazards.slice();
  state.facilities = DEMO.facilities.slice();
  state.users = DEMO.users.slice();
}

/* ---------- Deep link on load ---------- */
const initial = location.hash.replace("#", "");
if (pageNames[initial]) showPage(initial);
boot();
const loginForm = document.getElementById("loginForm");
const loginScreen = document.getElementById("loginScreen");
const dashboardApp = document.getElementById("dashboardApp");
const loginError = document.getElementById("loginError");

const togglePassword = document.getElementById("togglePassword");
const passwordField = document.getElementById("loginPassword");

togglePassword.addEventListener("click", () => {
  passwordField.type =
    passwordField.type === "password" ? "text" : "password";

  togglePassword.textContent =
    passwordField.type === "password" ? "Show" : "Hide";
});

loginForm.addEventListener("submit", function(e) {
  e.preventDefault();

  const email = document.getElementById("loginEmail").value.trim();
  const password = passwordField.value;

  if (
  email === "admin@safehangar.com" &&
  password === "SafeHangar123"
) {
  loginError.textContent = "";

  loginScreen.style.display = "none";
  dashboardApp.style.display = "block";

  window.location.hash = "dashboard";

  if (typeof showPage === "function") {
    showPage("dashboard");
  }

  window.scrollTo(0, 0);

} else {
  loginError.textContent =
    "Invalid email or password. Please try again.";
}
});