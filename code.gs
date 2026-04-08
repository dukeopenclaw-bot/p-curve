/*************************************************
 * CONFIG
 *************************************************/
const SPREADSHEET_ID = "1wbLVVeLsEg0UXdhXbn73cs20Bmv8y5cuShvaI8tPOkM";


/*************************************************
 * COMMON UTIL
 *************************************************/
function getSheet(name) {
  return SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(name);
}

function getJson(sheetName) {
  const sheet = getSheet(sheetName);
  if (!sheet) return [];
  const data = sheet.getDataRange().getValues();
  if (data.length < 2) return [];
  const headers = data.shift();
  return data
    .filter(row => row[0] !== '' && row[0] !== null)
    .map(row => {
      let obj = {};
      headers.forEach((h, i) => obj[h] = row[i]);
      return obj;
    });
}

function createResponse(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

function generateId() {
  return Utilities.getUuid();
}


/*************************************************
 * CURVE DEFINITIONS
 *************************************************/
const CURVE_MAP = {
  "IEC NI":  { k: 0.14,   a: 0.02 },
  "IEC VI":  { k: 13.5,   a: 1    },
  "IEC EI":  { k: 80,     a: 2    },
  "IEC LTI": { k: 120,    a: 1    },
  "IEEE MI": { k: 0.0515, a: 0.02 },
  "IEEE VI": { k: 19.61,  a: 2    },
  "IEEE EI": { k: 28.2,   a: 2    },
  "DT":      "DT"
};


/*************************************************
 * CORE CALCULATION
 *************************************************/
function calculateCurveFull(setting, relay, model) {
  const curveType = setting.CurveType;

  const supported = model.SupportedCurves.split(",").map(c => c.trim());
  if (!supported.includes(curveType)) return null;

  const curveInfo = CURVE_MAP[curveType];

  let Is  = Number(setting.Is);
  let TMS = Number(setting.TMS);

  if (model.Is_Min && model.Is_Min !== "-" && Is < Number(model.Is_Min)) Is = Number(model.Is_Min);
  if (model.Is_Max && model.Is_Max !== "-" && Is > Number(model.Is_Max)) Is = Number(model.Is_Max);
  if (model.TMS_Min && TMS < Number(model.TMS_Min)) TMS = Number(model.TMS_Min);
  if (model.TMS_Max && TMS > Number(model.TMS_Max)) TMS = Number(model.TMS_Max);

  const CT_ratio = Number(relay.CT_Ratio) ||
                   (Number(relay.CT_Primary) / Number(relay.CT_Secondary));
  const In = Number(relay.In) || 1;

  let result = [];

  for (let I = 10; I <= 20000; I *= 1.2) {
    let multiple;
    if (model.IsUnit === "pu") {
      multiple = (I / CT_ratio) / (Is * In);
    } else {
      multiple = (I / CT_ratio) / Is;
    }

    if (multiple <= 1) continue;

    let t;
    if (curveType === "DT") {
      t = Number(setting.t) || 1;
    } else {
      t = TMS * (curveInfo.k / (Math.pow(multiple, curveInfo.a) - 1));
    }

    if (setting.t_min && Number(setting.t_min) > 0) t = Math.max(t, Number(setting.t_min));
    if (setting.t_max && Number(setting.t_max) > 0) t = Math.min(t, Number(setting.t_max));

    if (t <= 0) continue;
    result.push({ I: I, t: t });
  }

  return result;
}


/*************************************************
 * RELAY CURVE
 *************************************************/
function getRelayCurveFull(relayTag) {
  const relays   = getJson("Relays");
  const settings = getJson("Settings");
  const models   = getJson("RelayModels");

  const relay = relays.find(r => r.RelayTag == relayTag);
  if (!relay) return [];

  const model = models.find(m => m.ModelID == relay.ModelID);
  if (!model) return [];

  const relaySettings = settings.filter(s =>
    s.RelayTag == relayTag && (s.Enable === true || s.Enable === "TRUE")
  );

  let curves = [];
  relaySettings.forEach(s => {
    const data = calculateCurveFull(s, relay, model);
    if (!data) return;
    curves.push({
      relay:     relayTag,
      model:     model.Model,
      element:   s.Element,
      stage:     s.Stage,
      curveType: s.CurveType,
      color:     relay.Color || "#333333",
      data:      data
    });
  });

  return curves;
}


/*************************************************
 * MULTI RELAY
 *************************************************/
function getMultiRelayCurve(tags) {
  const tagList = tags.split(",");
  let result = [];
  tagList.forEach(tag => {
    result.push(...getRelayCurveFull(tag.trim()));
  });
  return result;
}


/*************************************************
 * GRAPH
 *************************************************/
function getGraphConfig(graphId) {
  const graphs = getJson("Graphs");
  return graphs.find(g => g.GraphID == graphId) || null;
}

function getGraphWithCurves(graphId) {
  const graph = getGraphConfig(graphId);
  if (!graph) return { error: "Graph not found: " + graphId };

  const items = getJson("GraphItems")
    .filter(i => i.GraphID == graphId && (i.Show === true || i.Show === "TRUE"));

  let curves = [];
  items.forEach(item => {
    const relayCurves = getRelayCurveFull(item.RelayTag);
    relayCurves.forEach(c => {
      if (String(c.element) == String(item.Element) &&
          String(c.stage)   == String(item.Stage)) {
        curves.push({
          label: item.DisplayName || (item.RelayTag + " " + item.Element + " S" + item.Stage),
          color: c.color,
          data:  c.data
        });
      }
    });
  });

  return { graph, curves };
}


/*************************************************
 * SHEET WRITE HELPERS
 *************************************************/
function saveRow(sheetName, data, keyCol) {
  const sheet = getSheet(sheetName);
  if (!sheet) throw new Error("Sheet not found: " + sheetName);

  const rows    = sheet.getDataRange().getValues();
  const headers = rows[0];
  const keyIdx  = headers.indexOf(keyCol);
  if (keyIdx < 0) throw new Error("Key column not found: " + keyCol);

  const newRow = headers.map(h => data[h] !== undefined ? data[h] : "");

  let targetRow = -1;
  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][keyIdx]) === String(data[keyCol])) {
      targetRow = i + 1;
      break;
    }
  }

  if (targetRow > 0) {
    sheet.getRange(targetRow, 1, 1, newRow.length).setValues([newRow]);
  } else {
    sheet.appendRow(newRow);
  }
  return { success: true };
}

function deleteRowByKey(sheetName, keyCol, keyVal) {
  const sheet = getSheet(sheetName);
  if (!sheet) throw new Error("Sheet not found: " + sheetName);

  const data    = sheet.getDataRange().getValues();
  const headers = data[0];
  const keyIdx  = headers.indexOf(keyCol);

  for (let i = data.length - 1; i >= 1; i--) {
    if (String(data[i][keyIdx]) === String(keyVal)) {
      sheet.deleteRow(i + 1);
      return { success: true };
    }
  }
  return { error: "Row not found" };
}


/*************************************************
 * API HANDLER (GET)
 *************************************************/
function doGet(e) {
  try {
    const type = e.parameter.type;
    let result;

    switch (type) {
      case "models":        result = getJson("RelayModels");                          break;
      case "relays":        result = getJson("Relays");                               break;
      case "settings":      result = getJson("Settings");                             break;
      case "graphs":        result = getJson("Graphs");                               break;
      case "graphItems":    result = getJson("GraphItems");                           break;
      case "curveFull":     result = getRelayCurveFull(e.parameter.relay);           break;
      case "curveCompare":  result = getMultiRelayCurve(e.parameter.relays);         break;
      case "graphFull":     result = getGraphWithCurves(e.parameter.graphId);        break;
      default:              result = { error: "Invalid type: " + type };
    }

    return createResponse(result);
  } catch (err) {
    return createResponse({ error: err.toString() });
  }
}


/*************************************************
 * API HANDLER (POST)
 *************************************************/
function doPost(e) {
  try {
    const body   = JSON.parse(e.postData.contents);
    const action = body.action;
    let result;

    switch (action) {
      case "saveRelay":       result = saveRow("Relays",     body.data, "RelayTag");   break;
      case "saveSetting":     result = saveRow("Settings",   body.data, "SettingID");  break;
      case "saveGraph":       result = saveRow("Graphs",     body.data, "GraphID");    break;
      case "saveGraphItem":   result = saveRow("GraphItems", body.data, "ItemID");     break;
      case "deleteRelay":     result = deleteRowByKey("Relays",     "RelayTag",  body.RelayTag);   break;
      case "deleteSetting":   result = deleteRowByKey("Settings",   "SettingID", body.SettingID);  break;
      case "deleteGraph":     result = deleteRowByKey("Graphs",     "GraphID",   body.GraphID);    break;
      case "deleteGraphItem": result = deleteRowByKey("GraphItems", "ItemID",    body.ItemID);     break;
      default:                result = { error: "Unknown action: " + action };
    }

    return createResponse(result);
  } catch (err) {
    return createResponse({ error: err.toString() });
  }
}
