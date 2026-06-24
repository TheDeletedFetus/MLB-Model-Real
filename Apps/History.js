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

  ensureExactHistoryHeaders(historySheet, historyHeaders);
  ensureExactHistoryHeaders(archiveSheet, historyHeaders);

  const existingGameIds = getExistingHistoryGameIds(historySheet);

  const snapshotTime = Utilities.formatDate(
    new Date(),
    "America/New_York",
    "yyyy-MM-dd HH:mm:ss"
  );

  const rowsToAppend = [];

  matrixRows.forEach(matrixRow => {
    const row = rowArrayToObject(matrixHeaders, matrixRow);

    const gameId = row["Game ID"];
    if (!gameId) return;

    // One permanent pregame snapshot per game.
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

    const baseRow = [
      snapshotTime,
      gameId,
      row["Date"],

      awayTeam,
      homeTeam,
      "", // Away Score - filled postgame
      "", // Home Score - filled postgame
      "", // Winner - filled postgame

      modelPick,
      modelPickSide,
      "", // Game Winner Correct? - filled postgame

      row["Away Model Score"],
      row["Home Model Score"],
      row["Confidence"],

      "", // Away Model Implied Probability - future calibrated value
      "", // Home Model Implied Probability - future calibrated value
      "", // Model Pick Implied Probability - future calibrated value

      row["Away ML"],
      row["Home ML"],
      awayMarketProbability,
      homeMarketProbability,
      modelPickMarketProbability,

      marketFavorite,
      marketUnderdog,
      modelPickML,
      modelPickType,

      modelPickType === "Underdog",
      "", // Underdog Won? - filled postgame
      modelPickType === "Favorite",
      "", // Favorite Won? - filled postgame

      "", // Projected Edge % - future calibrated value
      "", // Beat Market? - future calibrated value

      bestEdge.name,
      bestEdge.value,
      "", // Flat Bet Profit - filled postgame

      false,
      "PREGAME_MODEL_MATRIX_SNAPSHOT_v0.1.1"
    ];

    rowsToAppend.push(baseRow.concat(matrixRow));
  });

  if (rowsToAppend.length > 0) {
    appendHistoryRows(historySheet, rowsToAppend, historyHeaders.length);
    appendHistoryRows(archiveSheet, rowsToAppend, historyHeaders.length);
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
    "Source"
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
  const lastRow = sheet.getLastRow();

  if (lastRow === 0) {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    return;
  }

  const currentHeaders = sheet
    .getRange(1, 1, 1, Math.max(sheet.getLastColumn(), headers.length))
    .getValues()[0];

  const currentText = currentHeaders.join("|");
  const expectedText = headers.join("|");

  if (currentText !== expectedText) {
    throw new Error(
      sheet.getName() +
      " headers do not match current HISTORY schema. Do not reset unless you intentionally want to rebuild HISTORY."
    );
  }
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

  return (
    finalFlag === true ||
    finalFlag === "TRUE" ||
    finalFlag === "Yes" ||
    finalFlag === "Final" ||
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