/**
 * Property Management Tracker — Dashboard API
 *
 * Exposes every named range in this spreadsheet (see the "KPI_*" and
 * "PM_Workload_Table" names created by build_dashboard.py) as JSON over
 * HTTP GET. Generic by design: add or rename a named range in the sheet
 * and this script picks it up automatically — nothing here is hardcoded
 * to a specific KPI or cell address.
 *
 * SETUP
 * 1. Open the Google Sheet -> Extensions -> Apps Script.
 * 2. Replace/paste this file's contents as Code.gs.
 * 3. Project Settings -> Script Properties -> add a property named
 *    API_TOKEN with a long random value. Requests must pass the same
 *    value as the "token" query param. (Skip this only if the backend
 *    calling in already authenticates some other way, e.g. it's the
 *    same Google Cloud project using OAuth against the Sheets API
 *    instead of this endpoint.)
 * 4. Deploy -> New deployment -> type "Web app".
 *      Execute as: Me
 *      Who has access: Anyone (the API_TOKEN is the actual gate, since
 *      Apps Script web apps don't support custom auth headers)
 * 5. Copy the deployment URL. Your backend calls, e.g.:
 *      GET {url}?token=XXXX                     -> every named range
 *      GET {url}?token=XXXX&name=KPI_TotalUnits  -> just that one
 *
 * NOTE ON FRESHNESS: Google Sheets recalculates formulas server-side as
 * soon as a source row changes, so a call right after an edit already
 * sees the new value. Volatile functions (TODAY(), used by the
 * Leasing Pipeline / Growth sections) only re-evaluate on open, on
 * edit, or on Sheets' periodic recalculation — set File > Settings >
 * Calculation > Recalculation to "On change and every hour" so those
 * stay fresh even on days nobody edits the sheet.
 */

function doGet(e) {
  try {
    var params = (e && e.parameter) || {};
    var authError = checkAuth_(params);
    if (authError) {
      return jsonOutput_({ error: authError }, 401);
    }

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var namedRanges = ss.getNamedRanges();

    if (params.name) {
      var match = namedRanges.filter(function (nr) {
        return nr.getName() === params.name;
      })[0];
      if (!match) {
        return jsonOutput_(
          { error: "Unknown named range: " + params.name }, 404
        );
      }
      return jsonOutput_({
        name: params.name,
        value: readRange_(match.getRange()),
        generatedAt: new Date().toISOString(),
      });
    }

    var result = {};
    namedRanges.forEach(function (nr) {
      // Skip anything not produced by build_dashboard.py (e.g. stray
      // legacy names already present in the workbook).
      var name = nr.getName();
      if (name.indexOf("KPI_") !== 0 && name !== "PM_Workload_Table") {
        return;
      }
      result[name] = readRange_(nr.getRange());
    });

    return jsonOutput_({
      data: result,
      generatedAt: new Date().toISOString(),
      spreadsheetId: ss.getId(),
    });
  } catch (err) {
    return jsonOutput_({ error: String(err) }, 500);
  }
}

/** Returns an error string if auth fails, or null if it's fine. */
function checkAuth_(params) {
  var expected = PropertiesService.getScriptProperties().getProperty("API_TOKEN");
  if (!expected) {
    // No token configured — treat as intentionally open. Remove this
    // branch (make the token mandatory) once API_TOKEN is set up.
    return null;
  }
  if (params.token !== expected) {
    return "Missing or invalid token.";
  }
  return null;
}

/** Single cell -> scalar. Multi-cell range -> array of row arrays. */
function readRange_(range) {
  if (range.getNumRows() === 1 && range.getNumColumns() === 1) {
    return normalizeValue_(range.getValue());
  }
  return range.getValues().map(function (row) {
    return row.map(normalizeValue_);
  });
}

function normalizeValue_(v) {
  if (Object.prototype.toString.call(v) === "[object Date]") {
    return v.toISOString();
  }
  return v;
}

/**
 * Apps Script's ContentService can't set a real HTTP status code on a
 * web app response (doGet always returns 200 to the client) — the
 * status param is included here for callers who log it, and the
 * error/success shape of the body is what your backend should branch
 * on instead of the transport-level status.
 */
function jsonOutput_(body, status) {
  if (status) {
    body._status = status;
  }
  return ContentService.createTextOutput(JSON.stringify(body)).setMimeType(
    ContentService.MimeType.JSON
  );
}
