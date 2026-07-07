/***************************************
 * SHADOW MODEL v0.1.0
 *
 * Purpose:
 * - Add/maintain a Test Weight column on Settings
 * - Score MODEL_MATRIX with Test Weight values
 * - Write shadow model outputs beside live model outputs
 * - Does NOT change live Weight values or live picks
 *
 * Main runner:
 *   scoreShadowModelMatrix()
 *
 * Recommended use:
 *   Run after scoreModelMatrix() and before HISTORY snapshot.
 ***************************************/

const SHADOW_MODEL = {
  SETTINGS_SHEET: "Settings",
  MODEL_MATRIX_SHEET: "MODEL_MATRIX",
  TEST_WEIGHT_HEADER: "Test Weight",

  OUTPUT_HEADERS: [
    "Shadow Away Model Score",
    "Shadow Home Model Score",
    "Shadow Model Pick",
    "Shadow Confidence"
  ]
};


function setupShadowModelTestWeights() {
  const ss = SpreadsheetApp.getActive();
  const settingsSheet = ss.getSheetByName(SHADOW_MODEL.SETTINGS_SHEET);
  if (!settingsSheet) throw new Error("Missing Settings sheet.");

  const info = ensureShadowTestWeightColumn_(settingsSheet);
  copySuggestedWeightsToTestWeights_(settingsSheet, info);
}


function scoreShadowModelMatrix() {
  const ss = SpreadsheetApp.getActive();
  const modelSheet = ss.getSheetByName(SHADOW_MODEL.MODEL_MATRIX_SHEET) || ss.getSheetByName("Model_Matrix");
  if (!modelSheet) throw new Error("Missing MODEL_MATRIX / Model_Matrix sheet.");

  const settings = getShadowModelSettings_();
  if (!settings.length) throw new Error("No active Test Weight settings found. Run setupShadowModelTestWeights() or fill Test Weight values.");

  ensureShadowModelMatrixColumns_(modelSheet);

  const values = modelSheet.getDataRange().getValues();
  if (values.length < 2) return;

  const headers = values[0];
  const rows = values.slice(1);
  const edgeStats = calculateEdgeStats(rows, headers, settings);

  const shadowAwayCol = headers.indexOf("Shadow Away Model Score");
  const shadowHomeCol = headers.indexOf("Shadow Home Model Score");
  const shadowPickCol = headers.indexOf("Shadow Model Pick");
  const shadowConfidenceCol = headers.indexOf("Shadow Confidence");

  if ([shadowAwayCol, shadowHomeCol, shadowPickCol, shadowConfidenceCol].some(col => col === -1)) {
    throw new Error("Shadow output columns were not found after creation.");
  }

  const updates = [];

  rows.forEach(row => {
    const result = scoreModelRowWithSettings(row, headers, settings, edgeStats);
    updates.push([
      result.awayScore,
      result.homeScore,
      result.pick,
      result.confidence
    ]);
  });

  if (updates.length) {
    modelSheet.getRange(2, shadowAwayCol + 1, updates.length, 4).setValues(updates);
  }
}


function runShadowModelUpdate() {
  scoreShadowModelMatrix();
}


function ensureShadowTestWeightColumn_(settingsSheet) {
  const values = settingsSheet.getDataRange().getValues();
  if (!values.length) throw new Error("Settings sheet is empty.");

  const headerInfo = findShadowSettingsHeaderRow_(values);
  const headerRowNumber = headerInfo.rowIndex + 1;
  let headers = headerInfo.headers;

  if (headers[SHADOW_MODEL.TEST_WEIGHT_HEADER] === undefined) {
    settingsSheet.getRange(headerRowNumber, settingsSheet.getLastColumn() + 1).setValue(SHADOW_MODEL.TEST_WEIGHT_HEADER);
  }

  const refreshed = settingsSheet.getDataRange().getValues();
  headers = mapShadowHeaders_(refreshed[headerRowNumber - 1]);

  return {
    headerRowNumber,
    headers
  };
}


function copySuggestedWeightsToTestWeights_(settingsSheet, info) {
  const headers = info.headers;
  const h = {
    stat: "Stat",
    weight: "Weight",
    suggestedWeight: "Suggested Weight",
    testWeight: SHADOW_MODEL.TEST_WEIGHT_HEADER
  };

  const weightCol = headers[h.weight];
  const suggestedCol = headers[h.suggestedWeight];
  const testCol = headers[h.testWeight];

  if (weightCol === undefined) throw new Error("Settings missing Weight column.");
  if (testCol === undefined) throw new Error("Settings missing Test Weight column.");

  for (let rowNumber = info.headerRowNumber + 1; rowNumber <= settingsSheet.getLastRow(); rowNumber++) {
    const currentWeight = settingsSheet.getRange(rowNumber, weightCol + 1).getValue();
    let testWeight = currentWeight;

    if (suggestedCol !== undefined) {
      const suggested = settingsSheet.getRange(rowNumber, suggestedCol + 1).getValue();
      const parsedSuggested = parseShadowWeight_(suggested);
      if (parsedSuggested !== "") testWeight = parsedSuggested;
    }

    settingsSheet.getRange(rowNumber, testCol + 1).setValue(testWeight);
  }
}


function ensureShadowModelMatrixColumns_(modelSheet) {
  let values = modelSheet.getDataRange().getValues();
  if (!values.length) throw new Error("MODEL_MATRIX is empty.");

  let headers = values[0].map(value => String(value || "").trim());

  SHADOW_MODEL.OUTPUT_HEADERS.forEach(header => {
    if (headers.indexOf(header) === -1) {
      modelSheet.getRange(1, modelSheet.getLastColumn() + 1).setValue(header);
      headers.push(header);
    }
  });
}


function getShadowModelSettings_() {
  const ss = SpreadsheetApp.getActive();
  const settingsSheet = ss.getSheetByName(SHADOW_MODEL.SETTINGS_SHEET);
  if (!settingsSheet) throw new Error("Missing Settings sheet.");

  ensureShadowTestWeightColumn_(settingsSheet);

  const values = settingsSheet.getDataRange().getValues();
  const headerInfo = findShadowSettingsHeaderRow_(values);
  const headers = headerInfo.headers;
  const rows = values.slice(headerInfo.rowIndex + 1);

  const statCol = headers["Stat"];
  const activeCol = headers["Active"];
  const testWeightCol = headers[SHADOW_MODEL.TEST_WEIGHT_HEADER];

  if (statCol === undefined) throw new Error("Settings missing Stat column.");
  if (activeCol === undefined) throw new Error("Settings missing Active column.");
  if (testWeightCol === undefined) throw new Error("Settings missing Test Weight column.");

  const settings = [];

  rows.forEach(row => {
    const stat = cleanShadowText_(row[statCol]);
    const active = parseShadowBool_(row[activeCol]);
    const weight = Number(row[testWeightCol] || 0);

    if (active && stat && isFinite(weight) && weight > 0) {
      settings.push({
        stat,
        weight
      });
    }
  });

  return settings;
}


function findShadowSettingsHeaderRow_(values) {
  const required = ["Stat", "Active", "Weight"];

  for (let rowIndex = 0; rowIndex < values.length; rowIndex++) {
    const headers = mapShadowHeaders_(values[rowIndex]);
    const foundAll = required.every(header => headers[header] !== undefined);

    if (foundAll) {
      return {
        rowIndex,
        headers
      };
    }
  }

  throw new Error("Could not find Settings header row for shadow model.");
}


function mapShadowHeaders_(headerRow) {
  const map = {};

  headerRow.forEach((header, index) => {
    if (header !== "" && header !== null && header !== undefined) {
      map[String(header).trim()] = index;
    }
  });

  return map;
}


function parseShadowWeight_(value) {
  if (value === "" || value === undefined || value === null) return "";

  if (typeof value === "number") return isFinite(value) ? value : "";

  const cleaned = String(value)
    .replace("WATCH:", "")
    .replace("watch:", "")
    .trim();

  const n = Number(cleaned);
  return isFinite(n) ? n : "";
}


function parseShadowBool_(value) {
  if (value === true) return true;
  const normalized = String(value || "").trim().toUpperCase();
  return normalized === "TRUE" || normalized === "YES" || normalized === "1";
}


function cleanShadowText_(value) {
  return String(value || "").trim();
}
