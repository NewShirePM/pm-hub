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
  "vanrockre.sharepoint.com,a02c1cd8-9f1f-4827-8286-7b6b7ce74232,01202419-6625-4499-b0d5-8ceb1cffdba3";

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
};
const lUrl = name => `${SITE}/lists/${name}/items`;

// Roles that count as "all PMs" for AppliesToAll tasks (mirrors index.html).
const PM_ROLES = ["pm", "pm-onsite", "pm-remote", "apm"];

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
function periodKey(cadence, dayOfWeek, dayOfMonth) {
  const now = new Date();
  if (cadence === "Daily") return now.toISOString().slice(0, 10);
  if (cadence === "Weekly") {
    const mon = new Date(now); mon.setDate(mon.getDate() - (mon.getDay() + 6) % 7);
    return now.toISOString().slice(0, 4) + "-W" + String(Math.ceil((mon.getDate()) / 7)).padStart(2, "0") + "-" + (dayOfWeek ?? 0);
  }
  if (cadence === "Monthly") return now.toISOString().slice(0, 7) + "-" + (dayOfMonth ?? 1);
  return today();
}
function computeDueDate(def, cad, now, dow) {
  if (cad === "Daily") {
    const [hh, mm] = (def.OnTimeCutoff || "17:00").split(":").map(Number);
    const d = new Date(now); d.setHours(hh, mm, 0, 0); return d.toISOString();
  }
  if (cad === "Weekly" || cad === "Bi-Weekly") {
    const dueDay = (def.DueDayOfWeek ?? 1) - 1;
    const diff = (dueDay - dow + 7) % 7;
    const d = new Date(now); d.setDate(d.getDate() + diff); d.setHours(17, 0, 0, 0); return d.toISOString();
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
    case "Bi-Weekly": {
      if (!def.BiWeeklyStartDate) return false;
      const start = new Date(def.BiWeeklyStartDate); start.setHours(0, 0, 0, 0);
      const diff = Math.floor((now - start) / 864e5);
      return diff >= 0 && diff % 14 === 0;
    }
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
  const roles = (def.AppliesToRoles || "").split(",").map(r => r.trim().toLowerCase()).filter(Boolean);
  const emails = (def.AppliesToEmails || "").split(",").map(e => e.trim().toLowerCase()).filter(Boolean);
  const legacy = (def.AppliesToEmail || "").toLowerCase();

  return staff.filter(e => {
    const role = (e.PMHubRole || "").trim().toLowerCase();
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
  const activity = actRaw.map(r => r.fields).filter(Boolean);

  console.log(`Loaded ${defs.length} recurring defs, ${staff.length} active staff, ${activity.length} activity rows.`);

  let created = 0, skippedExisting = 0, failed = 0;
  const activeDefs = defs.filter(d => d.IsActive !== false);

  for (const def of activeDefs) {
    const cad = getChoiceVal(def.Cadence) || def.Cadence || "Daily";
    if (!shouldGenerateToday(def, cad, ctx)) continue;

    const recipients = resolveRecipients(def, staff);
    if (!recipients.length) continue;

    const pk = periodKey(cad, def.DueDayOfWeek, def.DueDayOfMonth);
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
          : (cad === "Weekly" || cad === "Bi-Weekly")
            ? new Date(a.ActivityDate) >= weekStart
            : a.PeriodKey === pk;
        if (!periodMatch) return false;
        const idMatch = defIds.includes(String(a.RecurringTaskId));
        const descMatch = defDesc && (a.Description || "").trim().toLowerCase() === defDesc;
        return idMatch || descMatch;
      });
      if (exists) { skippedExisting++; continue; }

      const fields = {
        Title: def.Description?.slice(0, 255) || "Checklist Task",
        ActivityType: "Checklist",
        RecurringTaskId: String(def._odataId || ""),
        PMEmail: emp.Email,
        PMName: emp.Name || emp.Email,
        Category: def.Category || "Admin/Other",
        Cadence: cad,
        Status: "Queued",
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
      } catch (e) {
        failed++;
        console.warn(`Create failed for "${def.Description}": ${e.message}`);
      }
    }
  }

  console.log(`Done. Created ${created}, skipped ${skippedExisting} already-existing, ${failed} failed.`);
  if (failed > 0) process.exit(1);
}

main().catch(e => { console.error("Fatal:", e.message); process.exit(1); });
