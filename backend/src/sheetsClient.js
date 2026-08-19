const { google } = require("googleapis");

const SCOPES = ["https://www.googleapis.com/auth/spreadsheets.readonly"];
const NAME_PREFIX = "KPI_";
const EXTRA_NAMES = new Set(["PM_Workload_Table", "Heartbeat_Cell"]);

let sheetsApi = null;
let cache = { data: null, fetchedAt: 0 };

function isTrackedName(name) {
  return name.startsWith(NAME_PREFIX) || EXTRA_NAMES.has(name);
}

async function getSheetsApi() {
  if (sheetsApi) return sheetsApi;

  const inlineKey = process.env.GOOGLE_SERVICE_ACCOUNT_KEY_JSON;
  const keyFile = process.env.GOOGLE_SERVICE_ACCOUNT_KEY_PATH;

  if (!inlineKey && !keyFile) {
    throw new Error(
      "No Google credentials configured. Set GOOGLE_SERVICE_ACCOUNT_KEY_PATH " +
        "or GOOGLE_SERVICE_ACCOUNT_KEY_JSON in .env."
    );
  }

  const auth = new google.auth.GoogleAuth(
    inlineKey
      ? { credentials: JSON.parse(inlineKey), scopes: SCOPES }
      : { keyFile, scopes: SCOPES }
  );

  sheetsApi = google.sheets({ version: "v4", auth });
  return sheetsApi;
}

function requireSpreadsheetId() {
  const id = process.env.SPREADSHEET_ID;
  if (!id) throw new Error("SPREADSHEET_ID is not set in .env.");
  return id;
}

/** Single-cell range -> scalar. Multi-cell range -> array of row arrays. */
function collapseValues(values) {
  if (!values || values.length === 0) return null;
  if (values.length === 1 && values[0].length === 1) return values[0][0];
  return values;
}

/** Fetches every tracked (KPI_ prefixed, or listed in EXTRA_NAMES) named range's value. */
async function fetchAllKpis() {
  const sheets = await getSheetsApi();
  const spreadsheetId = requireSpreadsheetId();

  const meta = await sheets.spreadsheets.get({
    spreadsheetId,
    fields: "namedRanges(name)",
  });
  const names = (meta.data.namedRanges || [])
    .map((nr) => nr.name)
    .filter(isTrackedName);

  if (names.length === 0) {
    return {};
  }

  const res = await sheets.spreadsheets.values.batchGet({
    spreadsheetId,
    ranges: names,
  });

  const result = {};
  (res.data.valueRanges || []).forEach((vr, i) => {
    result[names[i]] = collapseValues(vr.values);
  });
  return result;
}

/** Fetches a single named range's value by name. Returns null if unknown. */
async function fetchOneKpi(name) {
  const sheets = await getSheetsApi();
  const spreadsheetId = requireSpreadsheetId();

  const meta = await sheets.spreadsheets.get({
    spreadsheetId,
    fields: "namedRanges(name)",
  });
  const known = new Set((meta.data.namedRanges || []).map((nr) => nr.name));
  if (!known.has(name)) return undefined;

  const res = await sheets.spreadsheets.values.batchGet({
    spreadsheetId,
    ranges: [name],
  });
  return collapseValues((res.data.valueRanges || [])[0]?.values);
}

/** Cached wrapper around fetchAllKpis — see CACHE_TTL_SECONDS in .env. */
async function getAllKpisCached() {
  const ttlMs = Number(process.env.CACHE_TTL_SECONDS || 20) * 1000;
  const age = Date.now() - cache.fetchedAt;
  if (cache.data && age < ttlMs) {
    return { data: cache.data, cached: true, ageSeconds: Math.round(age / 1000) };
  }
  const data = await fetchAllKpis();
  cache = { data, fetchedAt: Date.now() };
  return { data, cached: false, ageSeconds: 0 };
}

module.exports = { fetchAllKpis, fetchOneKpi, getAllKpisCached };
