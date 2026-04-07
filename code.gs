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
  const data = sheet.getDataRange().getValues();

  if (data.length === 0) return [];

  const headers = data.shift();

  return data.map(row => {
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

  "DT": "DT"
};


/*************************************************
 * CORE CALCULATION
 *************************************************/
function calculateCurveFull(setting, relay, model) {

  const curveType = setting.CurveType;

  // 지원 Curve 체크
  const supported = model.SupportedCurves.split(",").map(c => c.trim());
  if (!supported.includes(curveType)) return null;

  const curveInfo = CURVE_MAP[curveType];

  let Is  = Number(setting.Is);
  let TMS = Number(setting.TMS);

  // Model 제한 적용
  if (model.Is_Min && model.Is_Min !== "-") {
    if (Is < Number(model.Is_Min)) Is = Number(model.Is_Min);
  }
  if (model.Is_Max && model.Is_Max !== "-") {
    if (Is > Number(model.Is_Max)) Is = Number(model.Is_Max);
  }

  if (model.TMS_Min && TMS < Number(model.TMS_Min)) TMS = Number(model.TMS_Min);
  if (model.TMS_Max && TMS > Number(model.TMS_Max)) TMS = Number(model.TMS_Max);

  // CT 계산
  const CT_ratio = Number(relay.CT_Ratio) ||
                   (Number(relay.CT_Primary) / Number(relay.CT_Secondary));

  const In = Number(relay.In) || 1;

  let result = [];

  // 로그 스케일용 전류 증가
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

    // t 제한
    if (setting.t_min) t = Math.max(t, Number(setting.t_min));
    if (setting.t_max) t = Math.min(t, Number(setting.t_max));

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
    s.RelayTag == relayTag && s.Enable == true
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
  return graphs.find(g => g.GraphID == graphId);
}

function getGraphWithCurves(graphId) {

  const graph = getGraphConfig(graphId);
  if (!graph) return {};

  const items = getJson("GraphItems")
    .filter(i => i.GraphID == graphId && i.Show == true);

  let curves = [];

  items.forEach(item => {
    const relayCurves = getRelayCurveFull(item.RelayTag);

    relayCurves.forEach(c => {
      if (c.element == item.Element && c.stage == item.Stage) {
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
 * API HANDLER
 *************************************************/
function doGet(e) {
  try {
    const type = e.parameter.type;
    let result;

    switch (type) {
      case "models":
        result = getJson("RelayModels");
        break;
      case "relays":
        result = getJson("Relays");
        break;
      case "settings":
        result = getJson("Settings");
        break;
      case "graphs":
        result = getJson("Graphs");
        break;
      case "graphItems":
        result = getJson("GraphItems");
        break;
      case "curveFull":
        result = getRelayCurveFull(e.parameter.relay);
        break;
      case "curveCompare":
        result = getMultiRelayCurve(e.parameter.relays);
        break;
      case "graphFull":
        result = getGraphWithCurves(e.parameter.graphId);
        break;
      default:
        result = { error: "Invalid type" };
    }

    return createResponse(result);

  } catch (err) {
    return createResponse({ error: err.toString() });
  }
}


/*************************************************
 * POST
 *************************************************/
function doPost(e) {
  try {
    const body = JSON.parse(e.postData.contents);
    const action = body.action;
    let result;

    switch (action) {
      case "addRelay":
        result = appendRow("Relays", body);
        break;
      case "deleteRelay":
        result = deleteRowByKey("Relays", "RelayTag", body.RelayTag);
        break;
      case "addSetting":
        result = appendRow("Settings", body);
        break;
      case "deleteSetting":
        result = deleteRowByKey("Settings", "SettingID", body.SettingID);
        break;
      case "addGraph":
        result = appendRow("Graphs", body);
        break;
      case "deleteGraph":
        result = deleteRowByKey("Graphs", "GraphID", body.GraphID);
        break;
      case "addGraphItem":
        result = appendRow("GraphItems", body);
        break;
      case "deleteGraphItem":
        result = deleteRowByKey("GraphItems", "ItemID", body.ItemID);
        break;
      default:
        result = { error: "Unknown action: " + action };
    }

    return createResponse(result);

  } catch (err) {
    return createResponse({ error: err.toString() });
  }
}


/*************************************************
 * SHEET WRITE HELPERS
 *************************************************/
function appendRow(sheetName, data) {
  const sheet   = getSheet(sheetName);
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const row     = headers.map(h => data[h] !== undefined ? data[h] : "");
  sheet.appendRow(row);
  return { success: true };
}

function deleteRowByKey(sheetName, keyCol, keyVal) {
  const sheet = getSheet(sheetName);
  const data  = sheet.getDataRange().getValues();
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
