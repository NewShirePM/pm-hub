// ============================================================
// PM Hub — daily recurring-task generator (server-side)
// ============================================================
// Runs unattended in GitHub Actions on a schedule. Replaces the
// Power Automate 6AM flow. App-only Microsoft Graph auth (client
// credentials). Mirrors the in-app generateMissingTasks logic in
// index.html, generalized to run for every active staff member.
//
// Env (set as GitHub Actions secrets):
//   PMHUB_TENANT_ID      Azure AD tenant id
//   PMHUB_CLIENT_ID      app registration (application) client id
//   PMHUB_CLIENT_SECRET  client secret for that app reg
// Optional:
//   PMHUB_SITE_ID        Graph site id (defaults to the PM Hub site)
//
// IMPORTANT: this repo is PUBLIC and Actions logs are world-readable.
// Never log employee emails or names — only counts and task text.
// ============================================================

const TENANT = process.env.PMHUB_TENANT_ID;
const CLIENT_ID = process.env.PMHUB_CLIENT_ID;
const CLIENT_SECRET = process.env.PMHUB_CLIENT_SECRET;
const SITE_ID = process.env.PMHUB_SITE_ID ||
  "newshirepmcom.sharepoint.com,f5d74a99-8b23-477c-aed0-c1682efa5de1,0aa4ab90-0ddc-4a81-a803-06d4f2e1a7d8";

if (!TENANT || !CLIENT_ID || !CLIENT_SECRET) {
  // Not configured yet — no-op rather than fail, so scheduled runs don't
  // send daily failure emails before the secrets/app reg are set up.
  console.log("Secrets not set (PMHUB_TENANT_ID / PMHUB_CLIENT_ID / PMHUB_CLIENT_SECRET). Skipping — see scripts/SETUP.md.");
  process.exit(0);
}

const GRAPH = "https://graph.microsoft.com/v1.0";
const SITE = `${GRAPH}/sites/${SITE_ID}`;
const L = {
  recurring: "PM_RecurringTasks",
  activity: "PM_Activity",
  employees: "Employees",
  holidays: "PM_Holidays",
};
const lUrl = name => `${SITE}/lists/${name}/items`;

// Business days are evaluated in Eastern time (South Carolina), DST-aware.
const ET_TZ = "America/New_York";

// Roles that count as "all PMs" for AppliesToAll tasks (mirrors index.html).
const PM_ROLES = ["pm-onsite", "pm-remote", "apm"];
// Collapse the legacy "pm" value into "pm-onsite" (mirrors index.html normRole).
const normRole = r => { const k = (r || "").trim().toLowerCase(); return k === "pm" ? "pm-onsite" : k; };

// ---------- auth ----------
async function getToken() {
  const body = new URLSearchParams({
    client_id: CLIENT_ID,
    client_secret: CLIENT_SECRET,
    scope: "https://graph.microsoft.com/.default",
    grant_type: "client_credentials",
  });
  const r = await fetch(`https://login.microsoftonline.com/${TENANT}/oauth2/v2.0/token`, {
    method: "POST",
    headers: { "Content-Type": "application/x-www-form-urlencoded" },
    body,
  });
  if (!r.ok) throw new Error(`token ${r.status}: ${await r.text()}`);
  return (await r.json()).access_token;
}

// ---------- graph helpers ----------
async function gGet(token, url) {
  const r = await fetch(url, { headers: { Authorization: `Bearer ${token}` } });
  if (!r.ok) throw new Error(`GET ${r.status}: ${url}`);
  return r.json();
}
async function gAll(token, url) {
  let acc = [], next = url;
  while (next) {
    const d = await gGet(token, next);
    acc = acc.concat(d.value || []);
    next = d["@odata.nextLink"] || null;
  }
  return acc;
}
async function gPost(token, url, fields) {
  const r = await fetch(url, {
    method: "POST",
    headers: { Authorization: `Bearer ${token}`, "Content-Type": "application/json" },
    body: JSON.stringify({ fields }),
  });
  if (!r.ok) throw new Error(`POST ${r.status}: ${await r.text()}`);
  return r.json();
}
// PATCH a list item's fields (url ends in /items/{id}/fields; body is the field set).
async function gPatch(token, url, fieldSet) {
  const r = await fetch(url, {
    method: "PATCH",
    headers: { Authorization: `Bearer ${token}`, "Content-Type": "application/json" },
    body: JSON.stringify(fieldSet),
  });
  if (!r.ok) throw new Error(`PATCH ${r.status}: ${await r.text()}`);
  return r.json();
}

// Load company holidays into a Map of "YYYY-MM-DD" -> name. Resilient: if the
// PM_Holidays list doesn't exist yet, log and continue (no holiday skipping).
async function loadHolidays(token) {
  const map = new Map();
  try {
    const rows = await gAll(token, `${lUrl(L.holidays)}?expand=fields&$top=500`);
    for (const r of rows) {
      const f = r.fields || {};
      if (f.IsActive === false) continue;
      const key = String(f.HolidayDate || f.Date || "").slice(0, 10);
      if (/^\d{4}-\d{2}-\d{2}$/.test(key)) map.set(key, f.Title || "Holiday");
    }
  } catch (e) {
    console.warn(`Could not load ${L.holidays} (continuing without holiday skip): ${e.message}`);
  }
  return map;
}

// Auto-miss sweep: any still-open recurring task whose due date has passed (and
// which is NOT one of today's just-generated tasks) is marked Missed. This is
// what makes yesterday's undone dailies — and last period's undone weekly/
// monthly — roll to Missed instead of lingering as stale "Queued" forever.
async function sweepMissed(token, activity, now, todayStr) {
  let missed = 0, failed = 0;
  for (const a of activity) {
    const status = getChoiceVal(a.Status) || a.Status || "";
    if (status !== "Queued" && status !== "In Progress") continue;
    if (a.Source && a.Source !== "Recurring") continue;        // leave ad-hoc/assigned tasks alone
    const actDay = (a.ActivityDate || "").slice(0, 10);
    if (!actDay || actDay >= todayStr) continue;                // never miss a task dated today
    const due = a.DueDate ? new Date(a.DueDate) : null;
    if (!due || isNaN(due) || due >= now) continue;             // only once genuinely past due
    try {
      await gPatch(token, `${lUrl(L.activity)}/${a._id}/fields`, { Status: "Missed", MissedAt: now.toISOString() });
      a.Status = "Missed";
      missed++;
    } catch (e) { failed++; console.warn(`Auto-miss failed for id=${a._id}: ${e.message}`); }
  }
  return { missed, failed };
}

// ---------- date/period helpers (ported from index.html) ----------
function today() { return new Date().toISOString().slice(0, 10); }
function getChoiceVal(val) {
  if (!val) return "";
  if (typeof val === "string") {
    if (val.startsWith("{")) { try { return JSON.parse(val)?.Value || ""; } catch {} }
    return val;
  }
  if (typeof val === "object") return val?.Value || "";
  return String(val);
}
// Start (midnight) of the 14-day bi-weekly cycle that contains `now`, or null
// before the configured start date. Used to generate/dedup once PER CYCLE
// instead of only on the exact 14-day-multiple day (which a weekend/holiday
// skip would miss entirely).
function biWeeklyCycleStart(def, now) {
  if (!def.BiWeeklyStartDate) return null;
  const start = new Date(def.BiWeeklyStartDate); start.setHours(0, 0, 0, 0);
  const diff = Math.floor((now - start) / 864e5);
  if (diff < 0) return null;
  const cs = new Date(start); cs.setDate(start.getDate() + Math.floor(diff / 14) * 14);
  return cs;
}
function periodKey(cadence, dayOfWeek, dayOfMonth, def, now) {
  now = now || new Date();
  if (cadence === "Daily") return now.toISOString().slice(0, 10);
  if (cadence === "Weekly") {
    const mon = new Date(now); mon.setDate(mon.getDate() - (mon.getDay() + 6) % 7);
    return now.toISOString().slice(0, 4) + "-W" + String(Math.ceil((mon.getDate()) / 7)).padStart(2, "0") + "-" + (dayOfWeek ?? 0);
  }
  if (cadence === "Bi-Weekly") {
    const cs = def && biWeeklyCycleStart(def, now);
    return cs ? "BW-" + cs.toISOString().slice(0, 10) : today();
  }
  if (cadence === "Monthly") return now.toISOString().slice(0, 7) + "-" + (dayOfMonth ?? 1);
  return today();
}
function computeDueDate(def, cad, now, dow) {
  if (cad === "Daily") {
    const [hh, mm] = (def.OnTimeCutoff || "17:00").split(":").map(Number);
    const d = new Date(now); d.setHours(hh, mm, 0, 0); return d.toISOString();
  }
  if (cad === "Weekly") {
    const dueDay = (def.DueDayOfWeek ?? 1) - 1;
    const diff = (dueDay - dow + 7) % 7;
    const d = new Date(now); d.setDate(d.getDate() + diff); d.setHours(17, 0, 0, 0); return d.toISOString();
  }
  if (cad === "Bi-Weekly") {
    // Due at the end of the 14-day cycle, so it stays active (and isn't
    // auto-missed) for the whole cycle rather than just the generation week.
    const cs = biWeeklyCycleStart(def, now) || now;
    const d = new Date(cs); d.setDate(d.getDate() + 13); d.setHours(17, 0, 0, 0); return d.toISOString();
  }
  if (cad === "Monthly" || cad === "Bi-Monthly" || cad === "Quarterly") {
    const dueDay = def.DueDayOfMonth ?? 5;
    const d = new Date(now.getFullYear(), now.getMonth(), dueDay, 17, 0, 0, 0); return d.toISOString();
  }
  return now.toISOString();
}

// Should a def generate today, given its cadence? (ported from index.html)
function shouldGenerateToday(def, cad, ctx) {
  const { now, dow, dom, monthNum, weekStart } = ctx;
  switch (cad) {
    case "Daily": return true;
    case "Bi-Weekly":
      // Eligible on every run from the cycle's start onward; the per-cycle
      // dedup (below) ensures exactly one task per 14-day cycle, created on the
      // first business-day run within the cycle.
      return biWeeklyCycleStart(def, now) != null;
    case "Weekly": {
      const scheduledDay = (def.DueDayOfWeek ?? 1) - 1;
      const dueThisWeek = new Date(weekStart); dueThisWeek.setDate(weekStart.getDate() + scheduledDay);
      return dueThisWeek <= now;
    }
    case "Bi-Monthly": {
      const cycle = def.DueMonthCycle ?? 1;
      const parity = monthNum % 2;
      const cycleMatch = (cycle === 1 && parity === 1) || (cycle === 2 && parity === 0);
      return cycleMatch && dom >= (def.DueDayOfMonth ?? 1);
    }
    case "Monthly": return dom >= (def.DueDayOfMonth ?? 1);
    case "Quarterly": {
      const monthOfQ = ((monthNum - 1) % 3) + 1;
      return monthOfQ === (def.DueMonthOfQuarter ?? 1) && dom >= (def.DueDayOfMonth ?? 1);
    }
    default: return false;
  }
}

// Resolve which active employees a recurring def applies to.
function resolveRecipients(def, staff) {
  const appliesToAll = def.AppliesToAll === true || def.AppliesToAll === "true" || def.AppliesToAll === 1;
  const roles = (def.AppliesToRoles || "").split(",").map(normRole).filter(Boolean);
  const emails = (def.AppliesToEmails || "").split(",").map(e => e.trim().toLowerCase()).filter(Boolean);
  const legacy = (def.AppliesToEmail || "").toLowerCase();

  return staff.filter(e => {
    const role = normRole(e.PMHubRole);
    const email = (e.Email || "").toLowerCase();
    if (appliesToAll && PM_ROLES.includes(role)) return true;
    if (roles.length && roles.includes(role)) return true;
    if (emails.includes(email)) return true;
    if (!appliesToAll && !roles.length && !emails.length && legacy && legacy === email) return true;
    return false;
  });
}

async function main() {
  const token = await getToken();

  const now = new Date();
  const todayStr = today();

  // ---- Business-day gate: skip weekends and company holidays (Eastern time) ----
  const etParts = new Intl.DateTimeFormat("en-US", {
    timeZone: ET_TZ, weekday: "short", year: "numeric", month: "2-digit", day: "2-digit",
  }).formatToParts(now).reduce((o, p) => (o[p.type] = p.value, o), {});
  const etDate = `${etParts.year}-${etParts.month}-${etParts.day}`;
  if (etParts.weekday === "Sat" || etParts.weekday === "Sun") {
    console.log(`Today is ${etParts.weekday} (${etDate}) — weekend, skipping generation.`);
    return;
  }
  const holidays = await loadHolidays(token);
  if (holidays.has(etDate)) {
    console.log(`Today (${etDate}) is a company holiday: "${holidays.get(etDate)}" — skipping generation.`);
    return;
  }

  const dow = (now.getDay() + 6) % 7;       // 0=Mon..6=Sun
  const dom = now.getDate();
  const monthNum = now.getMonth() + 1;
  const weekStart = new Date(now); weekStart.setDate(now.getDate() - dow); weekStart.setHours(0, 0, 0, 0);
  const ctx = { now, dow, dom, monthNum, weekStart };

  // Load lists
  const [recRaw, empRaw, actRaw] = await Promise.all([
    gAll(token, `${lUrl(L.recurring)}?expand=fields&$top=500`),
    gAll(token, `${lUrl(L.employees)}?expand=fields&$top=500`),
    gAll(token, `${lUrl(L.activity)}?expand=fields&$top=2000`),
  ]);

  const defs = recRaw.map(r => ({ ...r.fields, _odataId: r.id, _fieldsId: r.fields?.id }));
  const staff = empRaw.map(r => r.fields).filter(e => e && e.EmployeeActive !== false);
  const activity = actRaw.map(r => ({ ...r.fields, _id: r.id })).filter(Boolean);

  console.log(`Loaded ${defs.length} recurring defs, ${staff.length} active staff, ${activity.length} activity rows.`);

  // Roll any past-due, still-open recurring tasks to Missed before generating today's.
  const sweep = await sweepMissed(token, activity, now, todayStr);
  console.log(`Auto-miss: marked ${sweep.missed} past-due task(s) Missed (${sweep.failed} failed).`);

  let created = 0, excused = 0, skippedExisting = 0, failed = 0;
  const activeDefs = defs.filter(d => d.IsActive !== false);

  for (const def of activeDefs) {
    const cad = getChoiceVal(def.Cadence) || def.Cadence || "Daily";
    if (!shouldGenerateToday(def, cad, ctx)) continue;

    const recipients = resolveRecipients(def, staff);
    if (!recipients.length) continue;

    const pk = periodKey(cad, def.DueDayOfWeek, def.DueDayOfMonth, def, now);
    const cycleStart = cad === "Bi-Weekly" ? biWeeklyCycleStart(def, now) : null;
    const dueDate = computeDueDate(def, cad, now, dow);
    // Match dedup against either id form, since the old Power Automate flow
    // and the in-app generator stored RecurringTaskId differently.
    const defIds = [String(def._odataId || ""), String(def._fieldsId || "")].filter(Boolean);
    const defDesc = (def.Description || "").trim().toLowerCase();

    for (const emp of recipients) {
      const email = (emp.Email || "").toLowerCase();
      const exists = activity.some(a => {
        if ((a.PMEmail || "").toLowerCase() !== email) return false;
        const periodMatch = cad === "Daily"
          ? a.ActivityDate?.slice(0, 10) === todayStr
          : cad === "Weekly"
            ? new Date(a.ActivityDate) >= weekStart
            : cad === "Bi-Weekly"
              ? (cycleStart && new Date(a.ActivityDate) >= cycleStart)
              : a.PeriodKey === pk;
        if (!periodMatch) return false;
        const idMatch = defIds.includes(String(a.RecurringTaskId));
        const descMatch = defDesc && (a.Description || "").trim().toLowerCase() === defDesc;
        return idMatch || descMatch;
      });
      if (exists) { skippedExisting++; continue; }

      // Auto-excuse: if the person is marked Out, the task is born "Not
      // Applicable" so it never counts as Missed on a day they're away.
      const isOut = (emp.PMTrackerStatus || "") === "Out";
      const fields = {
        Title: def.Description?.slice(0, 255) || "Checklist Task",
        ActivityType: "Checklist",
        RecurringTaskId: String(def._odataId || ""),
        PMEmail: emp.Email,
        PMName: emp.Name || emp.Email,
        Category: def.Category || "Admin/Other",
        Cadence: cad,
        Status: isOut ? "Not Applicable" : "Queued",
        Priority: getChoiceVal(def.Priority) || def.Priority || "Normal",
        Description: def.Description || "",
        ContextNotes: def.ContextNotes || "",
        OnTimeCutoff: def.OnTimeCutoff || "17:00",
        DueDayOfWeek: def.DueDayOfWeek ?? null,
        DueDayOfMonth: def.DueDayOfMonth ?? null,
        DueDate: dueDate,
        PeriodKey: pk,
        ActivityDate: new Date().toISOString(),
        AssignedByEmail: "system@newshirepm.com",
        AssignedByName: "Recurring Schedule",
        Source: "Recurring",
        IsActive: true,
      };
      try {
        await gPost(token, lUrl(L.activity), fields);
        created++;
        if (isOut) excused++;
      } catch (e) {
        failed++;
        console.warn(`Create failed for "${def.Description}": ${e.message}`);
      }
    }
  }

  console.log(`Done. Created ${created} (${excused} auto-excused for Out staff), skipped ${skippedExisting} already-existing, ${failed} failed.`);
  if (failed > 0) process.exit(1);
}

main().catch(e => { console.error("Fatal:", e.message); process.exit(1); });
