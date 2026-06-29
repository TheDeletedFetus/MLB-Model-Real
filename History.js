function appendPregameHistorySnapshot() {
  const historySheet = getOrCreateSheet("HISTORY");
  const archiveSheet = getOrCreateSheet("HISTORY_ARCHIVE");

  const matrixSheet = getOrCreateSheet("MODEL_MATRIX");
  const matrixValues = matrixSheet.getDataRange().getValues();

  if (matrixValues.length < 2) return;

  const matrixHeaders = matrixValues[0];
  const matrixRows = matrixValues.slice(1);

  const historyHeaders = getHistoryBaseHeaders().concat(
    matrixHeaders.map(header => "MM_" + header)
  );

  ensureHistoryHeaders(historySheet, historyHeaders);
  ensureHistoryHeaders(archiveSheet, historyHeaders);

  const finalHistoryHeaders = historySheet
    .getRange(1, 1, 1, historySheet.getLastColumn())
    .getValues()[0];

  const finalArchiveHeaders = archiveSheet
    .getRange(1, 1, 1, archiveSheet.getLastColumn())
    .getValues()[0];

  const existingGameIds = getExistingHistoryGameIds(historySheet);

  const now = new Date();
  const snapshotTime = Utilities.formatDate(
    now,
    "America/New_York",
    "yyyy-MM-dd HH:mm:ss"
  );

  const rowsToAppendHistory = [];
  const rowsToAppendArchive = [];

  matrixRows.forEach(matrixRow => {
    const row = rowArrayToObject(matrixHeaders, matrixRow);

    const gameId = row["Game ID"];
    if (!gameId) return;

    if (existingGameIds.has(String(gameId))) return;

    const gameTime = getPregameSnapshotGameTime_(row);

    if (!gameTime) {
      Logger.log("Skipping game with missing/invalid game time: " + gameId);
      return;
    }

    if (now >= gameTime) {
      Logger.log("Skipping already-started game: " + gameId + " | " + gameTime);
      return;
    }

    const historyObject = buildPregameHistoryObject_(row, matrixHeaders, matrixRow, snapshotTime);

    rowsToAppendHistory.push(objectToRow_(historyObject, finalHistoryHeaders));
    rowsToAppendArchive.push(objectToRow_(historyObject, finalArchiveHeaders));
  });

  if (rowsToAppendHistory.length > 0) {
    appendHistoryRows(historySheet, rowsToAppendHistory, finalHistoryHeaders.length);
  }

  if (rowsToAppendArchive.length > 0) {
    appendHistoryRows(archiveSheet, rowsToAppendArchive, finalArchiveHeaders.length);
  }
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

    const awayTeam = row[col.awayTeam];
    const homeTeam = row[col.homeTeam];
    const modelPick = row[col.modelPick];
    const winner = result["Winner"];

    const correct = modelPick && winner ? modelPick === winner : "";

    const modelPickType = row[col.modelPickType];
    const modelPickML = row[col.modelPickML];

    const underdogPick = modelPickType === "Underdog";
    const favoritePick = modelPickType === "Favorite";

    const underdogWon = underdogPick && correct === true;
    const favoriteWon = favoritePick && correct === true;

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
    final: headers.indexOf("Final?")
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
  ensureHistoryHeaders(sheet, headers);
}

function ensureHistoryHeaders(sheet, requiredHeaders) {
  if (sheet.getLastRow() === 0) {
    sheet.getRange(1, 1, 1, requiredHeaders.length).setValues([requiredHeaders]);
    return;
  }

  const lastCol = Math.max(sheet.getLastColumn(), 1);
  const currentHeaders = sheet
    .getRange(1, 1, 1, lastCol)
    .getValues()[0]
    .map(header => String(header).trim());

  const existing = new Set(currentHeaders.filter(Boolean));
  const missing = requiredHeaders.filter(header => !existing.has(String(header).trim()));

  if (missing.length === 0) {
    return;
  }

  sheet
    .getRange(1, sheet.getLastColumn() + 1, 1, missing.length)
    .setValues([missing]);

  Logger.log(
    "Added missing HISTORY headers to " +
    sheet.getName() +
    ": " +
    missing.join(", ")
  );
}

function rowArrayToObject(headers, row) {
  const obj = {};

  headers.forEach((header, index) => {
    obj[header] = row[index];
  });

  return obj;
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


function americanOddsToHistoryProbability(odds) {
  odds = parseHistoryNumber(odds);

  if (odds === "") return "";

  if (odds < 0) {
    return Number((Math.abs(odds) / (Math.abs(odds) + 100)).toFixed(4));
  }

  return Number((100 / (odds + 100)).toFixed(4));
}


function calculateFlatBetProfit(americanOdds, won) {
  const odds = parseHistoryNumber(americanOdds);

  if (odds === "") return "";
  if (won === "") return "";

  if (!won) return -100;

  if (odds > 0) {
    return odds;
  }

  return Number((10000 / Math.abs(odds)).toFixed(2));
}


function parseHistoryNumber(value) {
  if (value === "" || value === undefined || value === null) return "";

  const num = Number(value);

  return isNaN(num) ? "" : num;
}


function normalizeBooleanHistory(value) {
  if (value === true || value === "TRUE" || value === "Yes") return true;
  if (value === false || value === "FALSE" || value === "No") return false;
  return null;
}


function debugHistoryGameId() {
  const sh = SpreadsheetApp.getActive().getSheetByName("HISTORY");
  const values = sh.getDataRange().getValues();
  const headers = values[0];

  const gameIdCol = headers.indexOf("Game ID");
  Logger.log("Game ID col = " + gameIdCol);

  for (let r = 1; r < values.length; r++) {
    if (String(values[r][gameIdCol]) === "822799") {
      Logger.log("FOUND ROW " + (r + 1));
      Logger.log(values[r]);
      return;
    }
  }

  Logger.log("NOT FOUND in HISTORY");
}


function validatePregameSnapshotReady() {
  const matrixSheet = getOrCreateSheet("MODEL_MATRIX");
  const values = matrixSheet.getDataRange().getValues();

  if (values.length < 2) {
    throw new Error("Pregame snapshot aborted: MODEL_MATRIX has no game rows.");
  }

  const headers = values[0];
  const rows = values.slice(1);

  const requiredHeaders = [
    "Date",
    "Game ID",
    "Away Team",
    "Home Team",
    "Away Model Score",
    "Home Model Score",
    "Model Pick",
    "Confidence"
  ];

  requiredHeaders.forEach(header => {
    if (headers.indexOf(header) === -1) {
      throw new Error("Pregame snapshot aborted: missing MODEL_MATRIX column: " + header);
    }
  });

  const gameIdCol = headers.indexOf("Game ID");
  const awayTeamCol = headers.indexOf("Away Team");
  const homeTeamCol = headers.indexOf("Home Team");
  const modelPickCol = headers.indexOf("Model Pick");
  const confidenceCol = headers.indexOf("Confidence");

  const badRows = [];

  rows.forEach((row, index) => {
    const rowNumber = index + 2;

    const gameId = row[gameIdCol];
    const awayTeam = row[awayTeamCol];
    const homeTeam = row[homeTeamCol];
    const modelPick = row[modelPickCol];
    const confidence = row[confidenceCol];

    if (!gameId || !awayTeam || !homeTeam || !modelPick || confidence === "") {
      badRows.push(rowNumber);
    }
  });

  if (badRows.length > 0) {
    throw new Error(
      "Pregame snapshot aborted: incomplete MODEL_MATRIX rows: " +
      badRows.join(", ")
    );
  }
}

function ensureHistoryOddsColumns() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("HISTORY");

  if (!sheet) {
    throw new Error("HISTORY sheet not found.");
  }

  const requiredHeaders = [
    "Open Away ML",
    "Open Home ML",
    "Latest Away ML",
    "Latest Home ML",
    "Pregame Away ML",
    "Pregame Home ML",
    "Away ML Move",
    "Home ML Move",
    "Model Pick Open ML",
    "Model Pick Pregame ML",
    "CLV",
    "Market Moved Toward Pick",
    "Steam Move Flag",
    "Odds Snapshot Count",
    "First Odds Snapshot Time",
    "Last Odds Snapshot Time"
  ];

  const lastCol = sheet.getLastColumn();
  const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  const existing = new Set(headers.map(h => String(h).trim()));

  const missing = requiredHeaders.filter(h => !existing.has(h));

  if (missing.length === 0) {
    Logger.log("No missing HISTORY odds columns.");
    return;
  }

  sheet
    .getRange(1, lastCol + 1, 1, missing.length)
    .setValues([missing]);

  Logger.log("Added HISTORY odds columns: " + missing.join(", "));
}

function buildPregameHistoryObject_(row, matrixHeaders, matrixRow, snapshotTime) {
  const gameId = row["Game ID"];
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

  const historyObject = {
    "Snapshot Time": snapshotTime,
    "Game ID": gameId,
    "Game Date": row["Date"],

    "Away Team": awayTeam,
    "Home Team": homeTeam,
    "Away Score": "",
    "Home Score": "",
    "Winner": "",

    "Model Pick": modelPick,
    "Model Pick Side": modelPickSide,
    "Game Winner Correct?": "",

    "Away Model Score": row["Away Model Score"],
    "Home Model Score": row["Home Model Score"],
    "Model Strength": row["Confidence"],

    "Away Model Implied Probability": "",
    "Home Model Implied Probability": "",
    "Model Pick Implied Probability": "",

    "Away ML": row["Away ML"],
    "Home ML": row["Home ML"],
    "Away Market Implied Probability": awayMarketProbability,
    "Home Market Implied Probability": homeMarketProbability,
    "Model Pick Market Implied Probability": modelPickMarketProbability,

    "Market Favorite": marketFavorite,
    "Market Underdog": marketUnderdog,
    "Model Pick ML": modelPickML,
    "Model Pick Type": modelPickType,

    "Underdog Pick?": modelPickType === "Underdog",
    "Underdog Won?": "",
    "Favorite Pick?": modelPickType === "Favorite",
    "Favorite Won?": "",

    "Projected Edge %": "",
    "Beat Market?": "",

    "Largest Factor": bestEdge.name,
    "Largest Factor Value": bestEdge.value,
    "Flat Bet Profit": "",

    "Final?": false,
    "Source": "PREGAME_MODEL_MATRIX_SNAPSHOT",
    "Model Version": getModelVersion()
  };

  matrixHeaders.forEach((header, index) => {
    historyObject["MM_" + header] = matrixRow[index];
  });

  return historyObject;
}

function objectToRow_(object, headers) {
  return headers.map(header => {
    const key = String(header).trim();
    return Object.prototype.hasOwnProperty.call(object, key) ? object[key] : "";
  });
}

function getPregameSnapshotGameTime_(row) {
  const possibleHeaders = [
    "Game Time",
    "Start Time",
    "Game DateTime",
    "DateTime",
    "Scheduled Time",
    "First Pitch"
  ];

  for (const header of possibleHeaders) {
    if (row[header]) {
      const date = new Date(row[header]);
      if (!isNaN(date.getTime())) return date;
    }
  }

  return null;
}