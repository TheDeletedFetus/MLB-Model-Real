/***************************************
 * OPTIMIZER VALIDATION v0.3.0
 *
 * Purpose:
 * Validate optimizer scoring against live Model_Matrix output using the
 * production scoring helper from Scoring.js:
 *   scoreModelRowWithSettings()
 *
 * Run manually after the normal model has been scored:
 *   runOptimizerValidation()
 *
 * Output:
 *   OPTIMIZER_VALIDATION
 ***************************************/

const OPT_VALIDATION = {
  MODEL_MATRIX_SHEET: "MODEL_MATRIX",
  OUTPUT_SHEET: "OPTIMIZER_VALIDATION",
  SCORE_TOLERANCE: 0.0001,

  REQUIRED_MODEL_HEADERS: {
    awayTeam: "Away Team",
    homeTeam: "Home Team",
    liveAwayScore: "Away Model Score",
    liveHomeScore: "Home Model Score",
    livePick: "Model Pick"
  }
};


function runOptimizerValidation() {
  const ss = SpreadsheetApp.getActive();
  const modelSheet = ss.getSheetByName(OPT_VALIDATION.MODEL_MATRIX_SHEET) || ss.getSheetByName("Model_Matrix");

  if (!modelSheet) throw new Error("Missing MODEL_MATRIX / Model_Matrix sheet.");

  const modelData = readOptimizerValidationModelMatrix_(modelSheet);
  const settings = getActiveModelSettings();
  const edgeStats = calculateEdgeStats(modelData.rows, modelData.headersArray, settings);

  const output = [];
  output.push([
    "Game",
    "Away Team",
    "Home Team",
    "Live Away Score",
    "Production Helper Away Score",
    "Away Score Diff",
    "Live Home Score",
    "Production Helper Home Score",
    "Home Score Diff",
    "Live Pick",
    "Production Helper Pick",
    "Used Features",
    "Score Match",
    "Pick Match"
  ]);

  let tested = 0;
  let scoreMatches = 0;
  let pickMatches = 0;
  let mismatches = 0;
  let firstMismatch = null;

  modelData.games.forEach(game => {
    const helperScore = scoreModelRowWithSettings(
      game.row,
      modelData.headersArray,
      settings,
      edgeStats
    );

    tested++;

    const awayDiff = helperScore.awayScore - game.liveAwayScore;
    const homeDiff = helperScore.homeScore - game.liveHomeScore;

    const scoreMatch = Math.abs(awayDiff) <= OPT_VALIDATION.SCORE_TOLERANCE &&
      Math.abs(homeDiff) <= OPT_VALIDATION.SCORE_TOLERANCE;

    const pickMatch = sameTeam_(game.livePick, helperScore.pick);

    if (scoreMatch) scoreMatches++;
    if (pickMatch) pickMatches++;
    if (!scoreMatch || !pickMatch) {
      mismatches++;
      if (!firstMismatch) {
        firstMismatch = {
          game,
          helperScore,
          awayDiff,
          homeDiff,
          scoreMatch,
          pickMatch
        };
      }
    }

    output.push([
      game.awayTeam + " @ " + game.homeTeam,
      game.awayTeam,
      game.homeTeam,
      game.liveAwayScore,
      helperScore.awayScore,
      awayDiff,
      game.liveHomeScore,
      helperScore.homeScore,
      homeDiff,
      game.livePick,
      helperScore.pick,
      helperScore.usedFeatures,
      scoreMatch ? "TRUE" : "FALSE",
      pickMatch ? "TRUE" : "FALSE"
    ]);
  });

  output.push(["", "", "", "", "", "", "", "", "", "", "", "", "", ""]);
  output.push(["SUMMARY", "", "", "", "", "", "", "", "", "", "", "", "", ""]);
  output.push(["Games Tested", tested, "", "", "", "", "", "", "", "", "", "", "", ""]);
  output.push(["Score Matches", scoreMatches, "", "", "", "", "", "", "", "", "", "", "", ""]);
  output.push(["Pick Matches", pickMatches, "", "", "", "", "", "", "", "", "", "", "", ""]);
  output.push(["Mismatches", mismatches, "", "", "", "", "", "", "", "", "", "", "", ""]);

  if (firstMismatch) {
    appendFirstMismatchBreakdown_(output, firstMismatch);
  }

  writeOptimizerValidationOutput_(ss, output);
}


function readOptimizerValidationModelMatrix_(sheet) {
  const values = sheet.getDataRange().getValues();
  if (values.length < 2) throw new Error("MODEL_MATRIX has no data rows.");

  const headersArray = values[0];
  const headers = mapHeaders_(headersArray);
  const rows = values.slice(1);

  const indexes = {
    awayTeam: getHeaderIndex_(headers, OPT_VALIDATION.REQUIRED_MODEL_HEADERS.awayTeam),
    homeTeam: getHeaderIndex_(headers, OPT_VALIDATION.REQUIRED_MODEL_HEADERS.homeTeam),
    liveAwayScore: getHeaderIndex_(headers, OPT_VALIDATION.REQUIRED_MODEL_HEADERS.liveAwayScore),
    liveHomeScore: getHeaderIndex_(headers, OPT_VALIDATION.REQUIRED_MODEL_HEADERS.liveHomeScore),
    livePick: getHeaderIndex_(headers, OPT_VALIDATION.REQUIRED_MODEL_HEADERS.livePick)
  };

  Object.keys(indexes).forEach(key => {
    if (indexes[key] === undefined) {
      throw new Error("MODEL_MATRIX missing required validation header: " + key);
    }
  });

  const games = [];

  rows.forEach((row, idx) => {
    const awayTeam = cleanText_(row[indexes.awayTeam]);
    const homeTeam = cleanText_(row[indexes.homeTeam]);
    const liveAwayScore = Number(row[indexes.liveAwayScore]);
    const liveHomeScore = Number(row[indexes.liveHomeScore]);
    const livePick = cleanText_(row[indexes.livePick]);

    if (!awayTeam || !homeTeam || !livePick) return;
    if (!isFinite(liveAwayScore) || !isFinite(liveHomeScore)) return;

    games.push({
      rowNumber: idx + 2,
      row,
      awayTeam,
      homeTeam,
      liveAwayScore,
      liveHomeScore,
      livePick
    });
  });

  return {
    headersArray,
    headers,
    rows,
    indexes,
    games
  };
}


function appendFirstMismatchBreakdown_(output, mismatch) {
  const game = mismatch.game;
  const helperScore = mismatch.helperScore;

  output.push(["", "", "", "", "", "", "", "", "", "", "", "", "", ""]);
  output.push(["FIRST MISMATCH DETAIL", "", "", "", "", "", "", "", "", "", "", "", "", ""]);
  output.push(["Game", game.awayTeam + " @ " + game.homeTeam, "", "", "", "", "", "", "", "", "", "", "", ""]);
  output.push(["Live Away Score", game.liveAwayScore, "Helper Away Score", helperScore.awayScore, "Diff", mismatch.awayDiff, "", "", "", "", "", "", "", ""]);
  output.push(["Live Home Score", game.liveHomeScore, "Helper Home Score", helperScore.homeScore, "Diff", mismatch.homeDiff, "", "", "", "", "", "", "", ""]);
  output.push(["Live Pick", game.livePick, "Helper Pick", helperScore.pick, "Used Features", helperScore.usedFeatures, "", "", "", "", "", "", "", ""]);
  output.push(["Final Score", helperScore.finalScore, "Total Score", helperScore.totalScore, "Total Weight", helperScore.totalWeight, "", "", "", "", "", "", "", ""]);
  output.push(["", "", "", "", "", "", "", "", "", "", "", "", "", ""]);
  output.push([
    "Feature",
    "Weight",
    "Raw Edge",
    "Mean",
    "StdDev",
    "Z-Score",
    "Contribution",
    "", "", "", "", "", "", ""
  ]);

  helperScore.contributions.forEach(item => {
    output.push([
      item.stat,
      item.weight,
      item.rawEdge,
      item.mean,
      item.stdDev,
      item.zScore,
      item.contribution,
      "", "", "", "", "", "", ""
    ]);
  });
}


function writeOptimizerValidationOutput_(ss, output) {
  let sheet = ss.getSheetByName(OPT_VALIDATION.OUTPUT_SHEET);
  if (!sheet) sheet = ss.insertSheet(OPT_VALIDATION.OUTPUT_SHEET);

  sheet.clearContents();

  if (!output.length) return;

  const width = Math.max.apply(null, output.map(row => row.length));
  const rectangular = output.map(row => {
    const copy = row.slice();
    while (copy.length < width) copy.push("");
    return copy;
  });

  sheet.getRange(1, 1, rectangular.length, width).setValues(rectangular);
  sheet.getRange(1, 1, 1, width).setFontWeight("bold");
  sheet.autoResizeColumns(1, Math.min(width, 14));
}
