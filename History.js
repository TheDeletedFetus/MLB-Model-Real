function appendPregameHistorySnapshot() {
  const ss = SpreadsheetApp.getActive();
  const historySheet = getOrCreateSheet("HISTORY");
  const archiveSheet = getOrCreateSheet("HISTORY_ARCHIVE");

  const matrixSheet = getOrCreateSheet("MODEL_MATRIX");
  const matrixValues = matrixSheet.getDataRange().getValues();

  if (matrixValues.length < 2) return;

  const matrixHeaders = matrixValues[0];
  const matrixRows = matrixValues.slice(1);

  const shadowSheet = ss.getSheetByName("MODEL_MATRIX_SHADOW");
  const shadowData = getShadowHistoryData_(shadowSheet);

  const expectedHeaders = getHistoryBaseHeaders()
    .concat(matrixHeaders.map(header => "MM_" + header))
    .concat(getShadowHistorySummaryHeaders())
    .concat(shadowData.headers.map(header => "SM_" + header));

  const historyHeaders = ensureHistoryHeadersAllowAdditions(historySheet, expectedHeaders);
  const archiveHeaders = ensureHistoryHeadersAllowAdditions(archiveSheet, expectedHeaders);

  const existingGameIds = getExistingHistoryGameIds(historySheet);

  const snapshotTime = Utilities.formatDate(
    new Date(),
    "America/New_York",
    "yyyy-MM-dd HH:mm:ss"
  );

  const historyRowsToAppend = [];
  const archiveRowsToAppend = [];

  matrixRows.forEach(matrixRow => {
    const row = rowArrayToObject(matrixHeaders, matrixRow);

    const gameId = row["Game ID"];
    if (!gameId) return;

    if (existingGameIds.has(String(gameId))) return;

    const awayTeam = row["Away Team"];
    const homeTeam = row["Home Team"];
    const modelPick = row["Model Pick"];

    const modelPickSide =
      modelPick === awayTeam ? "Away" :
      modelPick === homeTeam ? "Home" :
      "Coin Flip";

    const awayML = parseHistoryNumber(row["Away ML"]);
    const homeML = parseHistoryNumber(row["Home ML"]);

    const awayMarketProbability = americanOddsToHistoryProbability(awayML);
    const homeMarketProbability = americanOddsToHistoryProbability(homeML);

    const marketFavorite = getMarketFavorite(awayTeam, homeTeam, awayML, homeML);
    const marketUnderdog = getMarketUnderdog(awayTeam, homeTeam, awayML, homeML);

    const modelPickML =
      modelPick === awayTeam ? awayML :
      modelPick === homeTeam ? homeML :
      "";

    const modelPickMarketProbability =
      modelPick === awayTeam ? awayMarketProbability :
      modelPick === homeTeam ? homeMarketProbability :
      "";

    const modelPickType =
      modelPick === marketFavorite ? "Favorite" :
      modelPick === marketUnderdog ? "Underdog" :
      "None";

    const bestEdge = findBestEdge(row);
    const shadowRow = shadowData.byGameId[String(gameId)] || null;
    const shadow = buildShadowHistorySummary_(shadowRow, shadowData.headers, awayTeam, homeTeam, awayML, homeML, marketFavorite, marketUnderdog);

    const rowObject = {};

    rowObject["Snapshot Time"] = snapshotTime;
    rowObject["Game ID"] = gameId;
    rowObject["Game Date"] = row["Date"];

    rowObject["Away Team"] = awayTeam;
    rowObject["Home Team"] = homeTeam;
    rowObject["Away Score"] = "";
    rowObject["Home Score"] = "";
    rowObject["Winner"] = "";

    rowObject["Model Pick"] = modelPick;
    rowObject["Model Pick Side"] = modelPickSide;
    rowObject["Game Winner Correct?"] = "";

    rowObject["Away Model Score"] = row["Away Model Score"];
    rowObject["Home Model Score"] = row["Home Model Score"];
    rowObject["Model Strength"] = row["Confidence"];

    rowObject["Away Model Implied Probability"] = "";
    rowObject["Home Model Implied Probability"] = "";
    rowObject["Model Pick Implied Probability"] = "";

    rowObject["Away ML"] = row["Away ML"];
    rowObject["Home ML"] = row["Home ML"];
    rowObject["Away Market Implied Probability"] = awayMarketProbability;
    rowObject["Home Market Implied Probability"] = homeMarketProbability;
    rowObject["Model Pick Market Implied Probability"] = modelPickMarketProbability;

    rowObject["Market Favorite"] = marketFavorite;
    rowObject["Market Underdog"] = marketUnderdog;
    rowObject["Model Pick ML"] = modelPickML;
    rowObject["Model Pick Type"] = modelPickType;

    rowObject["Underdog Pick?"] = modelPickType === "Underdog";
    rowObject["Underdog Won?"] = "";
    rowObject["Favorite Pick?"] = modelPickType === "Favorite";
    rowObject["Favorite Won?"] = "";

    rowObject["Projected Edge %"] = "";
    rowObject["Beat Market?"] = "";

    rowObject["Largest Factor"] = bestEdge.name;
    rowObject["Largest Factor Value"] = bestEdge.value;
    rowObject["Flat Bet Profit"] = "";

    rowObject["Final?"] = false;
    rowObject["Source"] = "PREGAME_MODEL_MATRIX_SNAPSHOT";
    rowObject["Model Version"] = getModelVersion();

    Object.keys(shadow).forEach(key => {
      rowObject[key] = shadow[key];
    });

    matrixHeaders.forEach((header, index) => {
      rowObject["MM_" + header] = matrixRow[index];
    });

    if (shadowRow) {
      shadowData.headers.forEach((header, index) => {
        rowObject["SM_" + header] = shadowRow[index];
      });
    }

    historyRowsToAppend.push(headersToRow_(historyHeaders, rowObject));
    archiveRowsToAppend.push(headersToRow_(archiveHeaders, rowObject));
  });

  if (historyRowsToAppend.length > 0) {
    appendHistoryRows(historySheet, historyRowsToAppend, historyHeaders.length);
    appendHistoryRows(archiveSheet, archiveRowsToAppend, archiveHeaders.length);
  }
}

function getShadowHistoryData_(shadowSheet) {
  if (!shadowSheet) {
    return {
      headers: [],
      rows: [],
      byGameId: {}
    };
  }

  const values = shadowSheet.getDataRange().getValues();
  if (values.length < 2) {
    return {
      headers: values.length ? values[0] : [],
      rows: [],
      byGameId: {}
    };
  }

  const headers = values[0];
  const rows = values.slice(1);
  const gameIdCol = headers.indexOf("Game ID");
  const byGameId = {};

  if (gameIdCol !== -1) {
    rows.forEach(row => {
      const gameId = row[gameIdCol];
      if (gameId) byGameId[String(gameId)] = row;
    });
  }

  return {
    headers,
    rows,
    byGameId
  };
}

function buildShadowHistorySummary_(shadowRow, shadowHeaders, awayTeam, homeTeam, awayML, homeML, marketFavorite, marketUnderdog) {
  const summary = {};

  const shadowObj = shadowRow ? rowArrayToObject(shadowHeaders, shadowRow) : {};
  const shadowPick = shadowObj["Model Pick"] || "";

  const shadowPickSide =
    shadowPick === awayTeam ? "Away" :
    shadowPick === homeTeam ? "Home" :
    shadowPick ? "Coin Flip" : "";

  const shadowPickML =
    shadowPick === awayTeam ? awayML :
    shadowPick === homeTeam ? homeML :
    "";

  const shadowPickType =
    shadowPick === marketFavorite ? "Favorite" :
    shadowPick === marketUnderdog ? "Underdog" :
    shadowPick ? "None" : "";

  summary["Shadow Model Pick"] = shadowPick;
  summary["Shadow Model Pick Side"] = shadowPickSide;
  summary["Shadow Game Winner Correct?"] = "";
  summary["Shadow Away Model Score"] = shadowObj["Away Model Score"] || "";
  summary["Shadow Home Model Score"] = shadowObj["Home Model Score"] || "";
  summary["Shadow Model Strength"] = shadowObj["Confidence"] || "";
  summary["Shadow Pick ML"] = shadowPickML;
  summary["Shadow Pick Type"] = shadowPickType;
  summary["Shadow Flat Bet Profit"] = "";

  return summary;
}

function getShadowHistorySummaryHeaders() {
  return [
    "Shadow Model Pick",
    "Shadow Model Pick Side",
    "Shadow Game Winner Correct?",
    "Shadow Away Model Score",
    "Shadow Home Model Score",
    "Shadow Model Strength",
    "Shadow Pick ML",
    "Shadow Pick Type",
    "Shadow Flat Bet Profit"
  ];
}

function updateHistoryResults() {
  const historySheet = getOrCreateSheet("HISTORY");
  const archiveSheet = getOrCreateSheet("HISTORY_ARCHIVE");

  updateHistoryResultsForSheet(historySheet);
  updateHistoryResultsForSheet(archiveSheet);
}


function updateHistoryResultsForSheet(sheet) {
  const values = sheet.getDataRange().getValues();
  if (values.length < 2) return;

  const headers = values[0];
  const resultsByGameId = sheetToLookup("RAW_Results", "Game ID");

  const col = getHistoryColumnMap(headers);

  const updates = [];

  for (let i = 1; i < values.length; i++) {
    const row = values[i];
    const gameId = row[col.gameId];

    if (!gameId) continue;

    const result = resultsByGameId[gameId] || {};
    if (!isFinalResult(result)) continue;

    const alreadyFinal = normalizeBooleanHistory(row[col.final]);
    if (alreadyFinal === true) continue;

    const modelPick = row[col.modelPick];
    const winner = result["Winner"];

    const correct = modelPick && winner ? modelPick === winner : "";

    const modelPickType = row[col.modelPickType];
    const modelPickML = row[col.modelPickML];

    const underdogPick = modelPickType === "Underdog";
    const favoritePick = modelPickType === "Favorite";

    const underdogWon = underdogPick && correct === true;
    const favoriteWon = favoritePick && correct === true;

    const shadowPick = col.shadowModelPick !== -1 ? row[col.shadowModelPick] : "";
    const shadowCorrect = shadowPick && winner ? shadowPick === winner : "";
    const shadowPickML = col.shadowPickML !== -1 ? row[col.shadowPickML] : "";

    updates.push({
      rowNumber: i + 1,
      awayScore: result["Away Score"],
      homeScore: result["Home Score"],
      winner: winner,
      correct: correct,
      underdogPick: underdogPick,
      underdogWon: underdogWon,
      favoritePick: favoritePick,
      favoriteWon: favoriteWon,
      flatBetProfit: calculateFlatBetProfit(modelPickML, correct),
      shadowCorrect: shadowCorrect,
      shadowFlatBetProfit: calculateFlatBetProfit(shadowPickML, shadowCorrect),
      final: true
    });
  }

  updates.forEach(update => {
    sheet.getRange(update.rowNumber, col.awayScore + 1).setValue(update.awayScore);
    sheet.getRange(update.rowNumber, col.homeScore + 1).setValue(update.homeScore);
    sheet.getRange(update.rowNumber, col.winner + 1).setValue(update.winner);
    sheet.getRange(update.rowNumber, col.correct + 1).setValue(update.correct);

    sheet.getRange(update.rowNumber, col.underdogPick + 1).setValue(update.underdogPick);
    sheet.getRange(update.rowNumber, col.underdogWon + 1).setValue(update.underdogWon);
    sheet.getRange(update.rowNumber, col.favoritePick + 1).setValue(update.favoritePick);
    sheet.getRange(update.rowNumber, col.favoriteWon + 1).setValue(update.favoriteWon);

    sheet.getRange(update.rowNumber, col.flatBetProfit + 1).setValue(update.flatBetProfit);

    if (col.shadowCorrect !== -1) {
      sheet.getRange(update.rowNumber, col.shadowCorrect + 1).setValue(update.shadowCorrect);
    }

    if (col.shadowFlatBetProfit !== -1) {
      sheet.getRange(update.rowNumber, col.shadowFlatBetProfit + 1).setValue(update.shadowFlatBetProfit);
    }

    sheet.getRange(update.rowNumber, col.final + 1).setValue(update.final);
  });
}


function getHistoryColumnMap(headers) {
  return {
    snapshotTime: headers.indexOf("Snapshot Time"),
    gameId: headers.indexOf("Game ID"),
    gameDate: headers.indexOf("Game Date"),

    awayTeam: headers.indexOf("Away Team"),
    homeTeam: headers.indexOf("Home Team"),
    awayScore: headers.indexOf("Away Score"),
    homeScore: headers.indexOf("Home Score"),
    winner: headers.indexOf("Winner"),

    modelPick: headers.indexOf("Model Pick"),
    modelPickSide: headers.indexOf("Model Pick Side"),
    correct: headers.indexOf("Game Winner Correct?"),

    awayModelScore: headers.indexOf("Away Model Score"),
    homeModelScore: headers.indexOf("Home Model Score"),
    modelStrength: headers.indexOf("Model Strength"),

    awayML: headers.indexOf("Away ML"),
    homeML: headers.indexOf("Home ML"),
    modelPickML: headers.indexOf("Model Pick ML"),
    modelPickType: headers.indexOf("Model Pick Type"),

    underdogPick: headers.indexOf("Underdog Pick?"),
    underdogWon: headers.indexOf("Underdog Won?"),
    favoritePick: headers.indexOf("Favorite Pick?"),
    favoriteWon: headers.indexOf("Favorite Won?"),

    flatBetProfit: headers.indexOf("Flat Bet Profit"),
    final: headers.indexOf("Final?"),

    shadowModelPick: headers.indexOf("Shadow Model Pick"),
    shadowCorrect: headers.indexOf("Shadow Game Winner Correct?"),
    shadowPickML: headers.indexOf("Shadow Pick ML"),
    shadowFlatBetProfit: headers.indexOf("Shadow Flat Bet Profit")
  };
}


function getHistoryBaseHeaders() {
  return [
    "Snapshot Time",
    "Game ID",
    "Game Date",

    "Away Team",
    "Home Team",
    "Away Score",
    "Home Score",
    "Winner",

    "Model Pick",
    "Model Pick Side",
    "Game Winner Correct?",

    "Away Model Score",
    "Home Model Score",
    "Model Strength",

    "Away Model Implied Probability",
    "Home Model Implied Probability",
    "Model Pick Implied Probability",

    "Away ML",
    "Home ML",
    "Away Market Implied Probability",
    "Home Market Implied Probability",
    "Model Pick Market Implied Probability",

    "Market Favorite",
    "Market Underdog",
    "Model Pick ML",
    "Model Pick Type",

    "Underdog Pick?",
    "Underdog Won?",
    "Favorite Pick?",
    "Favorite Won?",

    "Projected Edge %",
    "Beat Market?",

    "Largest Factor",
    "Largest Factor Value",
    "Flat Bet Profit",

    "Final?",
    "Source",
    "Model Version"
  ];
}


function resetHistoryTables() {
  const historySheet = getOrCreateSheet("HISTORY");
  const archiveSheet = getOrCreateSheet("HISTORY_ARCHIVE");

  const matrixSheet = getOrCreateSheet("MODEL_MATRIX");
  const matrixValues = matrixSheet.getDataRange().getValues();

  const matrixHeaders = matrixValues.length > 0 ? matrixValues[0] : [];

  const headers = getHistoryBaseHeaders().concat(
    matrixHeaders.map(header => "MM_" + header)
  );

  historySheet.clearContents();
  archiveSheet.clearContents();

  historySheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  archiveSheet.getRange(1, 1, 1, headers.length).setValues([headers]);
}


function ensureExactHistoryHeaders(sheet, headers) {
  return ensureHistoryHeadersAllowAdditions(sheet, headers);
}


function ensureHistoryHeadersAllowAdditions(sheet, expectedHeaders) {
  const lastRow = sheet.getLastRow();

  if (lastRow === 0) {
    sheet.getRange(1, 1, 1, expectedHeaders.length).setValues([expectedHeaders]);
    return expectedHeaders.slice();
  }

  const currentHeaders = sheet
    .getRange(1, 1, 1, Math.max(sheet.getLastColumn(), expectedHeaders.length))
    .getValues()[0]
    .map(header => String(header || "").trim())
    .filter(header => header !== "");

  const finalHeaders = currentHeaders.slice();

  expectedHeaders.forEach(header => {
    if (finalHeaders.indexOf(header) === -1) {
      finalHeaders.push(header);
    }
  });

  if (finalHeaders.length !== currentHeaders.length) {
    sheet.getRange(1, 1, 1, finalHeaders.length).setValues([finalHeaders]);
  }

  return finalHeaders;
}


function rowArrayToObject(headers, row) {
  const obj = {};

  headers.forEach((header, index) => {
    obj[header] = row[index];
  });

  return obj;
}

function headersToRow_(headers, rowObject) {
  return headers.map(header => rowObject[header] !== undefined ? rowObject[header] : "");
}


function isFinalResult(result) {
  const finalFlag = result["Final?"];
  const status = result["Status"];
  const abstractState = result["Abstract State"];

  return (
    finalFlag === true ||
    finalFlag === "TRUE" ||
    finalFlag === "Yes" ||
    finalFlag === "Final" ||
    abstractState === "Final" ||
    status === "Final" ||
    status === "Game Over"
  );
}


function getExistingHistoryGameIds(sheet) {
  const values = sheet.getDataRange().getValues();
  const gameIds = new Set();

  if (values.length < 2) return gameIds;

  const headers = values[0];
  const gameIdCol = headers.indexOf("Game ID");

  for (let i = 1; i < values.length; i++) {
    const gameId = values[i][gameIdCol];
    if (gameId) gameIds.add(String(gameId));
  }

  return gameIds;
}


function appendHistoryRows(sheet, rows, columnCount) {
  sheet
    .getRange(sheet.getLastRow() + 1, 1, rows.length, columnCount)
    .setValues(rows);
}


function getMarketFavorite(awayTeam, homeTeam, awayML, homeML) {
  if (awayML === "" || homeML === "") return "";

  if (awayML < homeML) return awayTeam;
  if (homeML < awayML) return homeTeam;

  return "Even";
}


function getMarketUnderdog(awayTeam, homeTeam, awayML, homeML) {
  if (awayML === "" || homeML === "") return "";

  if (awayML > homeML) return awayTeam;
  if (homeML > awayML) return homeTeam;

  return "Even";
}
