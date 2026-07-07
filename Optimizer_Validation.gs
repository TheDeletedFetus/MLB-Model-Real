/***************************************
 * OPTIMIZER VALIDATION v0.2.0
 *
 * Purpose:
 * Validate shared optimizer scoring against live Model_Matrix output.
 *
 * This validator uses the same weighted-edge scorer intended for optimizer use:
 *   sharedScoreGameFromEdges_()
 *
 * Run manually after the normal model has been scored:
 *   runOptimizerValidation()
 *
 * Output:
 *   OPTIMIZER_VALIDATION
 ***************************************/

const OPT_VALIDATION = {
  MODEL_MATRIX_SHEET: "Model_Matrix",
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
  const settingsSheet = ss.getSheetByName(ITER_OPT.SETTINGS_SHEET);
  const modelSheet = ss.getSheetByName(OPT_VALIDATION.MODEL_MATRIX_SHEET);

  if (!settingsSheet) throw new Error("Missing Settings sheet.");
  if (!modelSheet) throw new Error("Missing Model_Matrix sheet.");

  const settings = readOptimizerSettings_(settingsSheet);
  const modelData = readOptimizerValidationModelMatrix_(modelSheet, settings.features);
  const weights = buildWeightMap_(settings.features, "currentWeight");

  const output = [];
  output.push([
    "Game",
    "Away Team",
    "Home Team",
    "Live Away Score",
    "Shared Away Score",
    "Away Score Diff",
    "Live Home Score",
    "Shared Home Score",
    "Home Score Diff",
    "Live Pick",
    "Shared Pick",
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
    const sharedScore = sharedScoreGameFromEdges_(
      game.row,
      modelData.edgeColumnMap,
      settings.features,
      weights,
      game.awayTeam,
      game.homeTeam
    );

    if (!sharedScore.hasScore) return;

    tested++;

    const awayDiff = sharedScore.awayScore - game.liveAwayScore;
    const homeDiff = sharedScore.homeScore - game.liveHomeScore;

    const scoreMatch = Math.abs(awayDiff) <= OPT_VALIDATION.SCORE_TOLERANCE &&
      Math.abs(homeDiff) <= OPT_VALIDATION.SCORE_TOLERANCE;

    const pickMatch = sameTeam_(game.livePick, sharedScore.pick);

    if (scoreMatch) scoreMatches++;
    if (pickMatch) pickMatches++;
    if (!scoreMatch || !pickMatch) {
      mismatches++;
      if (!firstMismatch) {
        firstMismatch = {
          game,
          sharedScore,
          awayDiff,
          homeDiff,
          scoreMatch,
          pickMatch,
          modelData,
          weights
        };
      }
    }

    output.push([
      game.awayTeam + " @ " + game.homeTeam,
      game.awayTeam,
      game.homeTeam,
      game.liveAwayScore,
      sharedScore.awayScore,
      awayDiff,
      game.liveHomeScore,
      sharedScore.homeScore,
      homeDiff,
      game.livePick,
      sharedScore.pick,
      sharedScore.usedFeatures,
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
  output.push(["Mapped Edge Columns", Object.keys(modelData.edgeColumnMap).length, "", "", "", "", "", "", "", "", "", "", "", ""]);

  if (firstMismatch) {
    appendFirstMismatchBreakdown_(output, firstMismatch, settings.features);
  }

  writeOptimizerValidationOutput_(ss, output);
}


function readOptimizerValidationModelMatrix_(sheet, features) {
  const values = sheet.getDataRange().getValues();
  const headerInfo = findOptimizerValidationHeaderRow_(values);
  const headers = headerInfo.headers;
  const headerRowIndex = headerInfo.rowIndex;

  const indexes = {
    awayTeam: getHeaderIndex_(headers, OPT_VALIDATION.REQUIRED_MODEL_HEADERS.awayTeam),
    homeTeam: getHeaderIndex_(headers, OPT_VALIDATION.REQUIRED_MODEL_HEADERS.homeTeam),
    liveAwayScore: getHeaderIndex_(headers, OPT_VALIDATION.REQUIRED_MODEL_HEADERS.liveAwayScore),
    liveHomeScore: getHeaderIndex_(headers, OPT_VALIDATION.REQUIRED_MODEL_HEADERS.liveHomeScore),
    livePick: getHeaderIndex_(headers, OPT_VALIDATION.REQUIRED_MODEL_HEADERS.livePick)
  };

  const edgeColumnMap = sharedBuildEdgeColumnMap_(headers, features, "");
  const games = [];

  for (let rowIndex = headerRowIndex + 1; rowIndex < values.length; rowIndex++) {
    const row = values[rowIndex];
    const awayTeam = cleanText_(row[indexes.awayTeam]);
    const homeTeam = cleanText_(row[indexes.homeTeam]);
    const liveAwayScore = Number(row[indexes.liveAwayScore]);
    const liveHomeScore = Number(row[indexes.liveHomeScore]);
    const livePick = cleanText_(row[indexes.livePick]);

    if (!awayTeam || !homeTeam || !livePick) continue;
    if (!isFinite(liveAwayScore) || !isFinite(liveHomeScore)) continue;

    games.push({
      rowNumber: rowIndex + 1,
      row,
      awayTeam,
      homeTeam,
      liveAwayScore,
      liveHomeScore,
      livePick
    });
  }

  return {
    headers,
    indexes,
    edgeColumnMap,
    games
  };
}


function appendFirstMismatchBreakdown_(output, mismatch, features) {
  const game = mismatch.game;
  const sharedScore = mismatch.sharedScore;
  const contributionRows = sharedContributionRowsFromEdges_(
    game.row,
    mismatch.modelData.edgeColumnMap,
    features,
    mismatch.weights
  );

  output.push(["", "", "", "", "", "", "", "", "", "", "", "", "", ""]);
  output.push(["FIRST MISMATCH DETAIL", "", "", "", "", "", "", "", "", "", "", "", "", ""]);
  output.push(["Game", game.awayTeam + " @ " + game.homeTeam, "", "", "", "", "", "", "", "", "", "", "", ""]);
  output.push(["Live Away Score", game.liveAwayScore, "Shared Away Score", sharedScore.awayScore, "Diff", mismatch.awayDiff, "", "", "", "", "", "", "", ""]);
  output.push(["Live Home Score", game.liveHomeScore, "Shared Home Score", sharedScore.homeScore, "Diff", mismatch.homeDiff, "", "", "", "", "", "", "", ""]);
  output.push(["Live Pick", game.livePick, "Shared Pick", sharedScore.pick, "Used Features", sharedScore.usedFeatures, "", "", "", "", "", "", "", ""]);
  output.push(["Weighted Edge Sum", sharedScore.weightedEdgeSum, "Scale", SHARED_SCORING.SCORE_SCALE, "", "", "", "", "", "", "", "", "", ""]);
  output.push(["", "", "", "", "", "", "", "", "", "", "", "", "", ""]);
  output.push([
    "Feature",
    "Edge",
    "Weight",
    "Raw Contribution",
    "Scaled Contribution",
    "", "", "", "", "", "", "", "", ""
  ]);

  contributionRows.forEach(item => {
    output.push([
      item.feature,
      item.edge,
      item.weight,
      item.rawContribution,
      item.scaledContribution,
      "", "", "", "", "", "", "", "", ""
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


function findOptimizerValidationHeaderRow_(values) {
  const required = [
    OPT_VALIDATION.REQUIRED_MODEL_HEADERS.awayTeam,
    OPT_VALIDATION.REQUIRED_MODEL_HEADERS.homeTeam,
    OPT_VALIDATION.REQUIRED_MODEL_HEADERS.liveAwayScore,
    OPT_VALIDATION.REQUIRED_MODEL_HEADERS.liveHomeScore,
    OPT_VALIDATION.REQUIRED_MODEL_HEADERS.livePick
  ];

  return findHeaderRow_(values, required, OPT_VALIDATION.MODEL_MATRIX_SHEET);
}
