function runPregameSnapshotGatekeeper() {
  const today = Utilities.formatDate(new Date(), "America/New_York", "yyyy-MM-dd");

  if (hasPregameSnapshotRunToday(today)) {
    Logger.log("Pregame snapshot already completed for " + today);
    return;
  }

  const firstPitch = getTodayFirstPitchTime();

  if (!firstPitch) {
    Logger.log("No first pitch found for today.");
    return;
  }

  const now = new Date();
  const snapshotWindowStart = new Date(firstPitch.getTime() - 15 * 60 * 1000);

  if (now < snapshotWindowStart) {
    Logger.log("Too early for pregame snapshot. First pitch: " + firstPitch);
    return;
  }

  if (now >= firstPitch) {
    Logger.log("Pregame snapshot skipped. First pitch has already passed: " + firstPitch);
    return;
  }

  runPregameSnapshot();

  markPregameSnapshotRunToday(today);
}

function getTodayFirstPitchTime() {
  const rows = getSheetRows("RAW_Schedule");

  const gameTimes = rows
    .map(row => row["Game Time"])
    .filter(value => value)
    .map(value => new Date(value))
    .filter(date => !isNaN(date.getTime()))
    .sort((a, b) => a - b);

  if (gameTimes.length === 0) return null;

  return gameTimes[0];
}


function hasPregameSnapshotRunToday(today) {
  const sheet = getOrCreateSheet("RUN_LOG");
  const values = sheet.getDataRange().getValues();

  if (values.length < 2) return false;

  const headers = values[0];
  const dateCol = headers.indexOf("Date");
  const eventCol = headers.indexOf("Event");

  if (dateCol === -1 || eventCol === -1) return false;

  for (let i = 1; i < values.length; i++) {
    if (
      values[i][dateCol] === today &&
      values[i][eventCol] === "PREGAME_SNAPSHOT_COMPLETED"
    ) {
      return true;
    }
  }

  return false;
}


function markPregameSnapshotRunToday(today) {
  const sheet = getOrCreateSheet("RUN_LOG");

  if (sheet.getLastRow() === 0) {
    sheet.getRange(1, 1, 1, 4).setValues([[
      "Timestamp",
      "Date",
      "Event",
      "Source"
    ]]);
  }

  sheet.appendRow([
    Utilities.formatDate(new Date(), "America/New_York", "yyyy-MM-dd HH:mm:ss"),
    today,
    "PREGAME_SNAPSHOT_COMPLETED",
    "runPregameSnapshotGatekeeper"
  ]);
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