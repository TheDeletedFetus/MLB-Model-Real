/***************************************
 * OPTIMIZER VALIDATION v0.1.0
 *
 * Purpose:
 * Validate optimizer scoring against the live Model_Matrix output.
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
    "Optimizer Away Score",
    "Away Score Diff",
    "Live Home Score",
    "Optimizer Home Score",
    "Home Score Diff",
    "Live Pick",
    "Optimizer Pick",
    "Score Match",
    "Pick Match"
  ]);

  let tested = 0;
  let scoreMatches = 0;
  let pickMatches = 0;
  let mismatches = 0;
  let firstMismatch = null;

  modelData.games.forEach(game => {
    const optScore = calculateValidationScore_(game, settings.features, weights);
    if (!optScore.hasScore) return;

    tested++;

    const awayDiff = optScore.awayScore - game.liveAwayScore;
    const homeDiff = optScore.homeScore - game.liveHomeScore;

    const scoreMatch = Math.abs(awayDiff) <= OPT_VALIDATION.SCORE_TOLERANCE &&
      Math.abs(homeDiff) <= OPT_VALIDATION.SCORE_TOLERANCE;

    const pickMatch = sameTeam_(game.livePick, optScore.pick);

    if (scoreMatch) scoreMatches++;
    if (pickMatch) pickMatches++;
    if (!scoreMatch || !pickMatch) {
      mismatches++;
      if (!firstMismatch) {
        firstMismatch = {
          game,
          optScore,
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
      optScore.awayScore,
      awayDiff,
      game.liveHomeScore,
      optScore.homeScore,
      homeDiff,
      game.livePick,
      optScore.pick,
      scoreMatch ? "TRUE" : "FALSE",
      pickMatch ? "TRUE" : "FALSE"
    ]);
  });

  output.push(["", "", "", "", "", "", "", "", "", "", "", "", ""]);
  output.push(["SUMMARY", "", "", "", "", "", "", "", "", "", "", "", ""]);
  output.push(["Games Tested", tested, "", "", "", "", "", "", "", "", "", "", ""]);
  output.push(["Score Matches", scoreMatches, "", "", "", "", "", "", "", "", "", "", ""]);
  output.push(["Pick Matches", pickMatches, "", "", "", "", "", "", "", "", "", "", ""]);
  output.push(["Mismatches", mismatches, "", "", "", "", "", "", "", "", "", "", ""]);

  if (firstMismatch) {
    appendFirstMismatchBreakdown_(output, firstMismatch, settings.features, weights);
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

  const featureColumns = resolveFeatureColumnsForSettings_(headers, features);
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

    const featureValues = {};

    features.forEach(feature => {
      const columns = featureColumns[feature.name];
      if (!columns) return;

      const awayVal = Number(row[columns.away]);
      const homeVal = Number(row[columns.home]);

      if (!isFinite(awayVal) || !isFinite(homeVal)) return;

      featureValues[feature.name] = {
        away: awayVal,
        home: homeVal
      };
    });

    games.push({
      rowNumber: rowIndex + 1,
      awayTeam,
      homeTeam,
      liveAwayScore,
      liveHomeScore,
      livePick,
      featureValues
    });
  }

  return {
    headers,
    indexes,
    featureColumns,
    games
  };
}


function calculateValidationScore_(game, features, weightsByFeature) {
  let awayScore = 0;
  let homeScore = 0;
  let used = 0;

  features.forEach(feature => {
    const weight = Number(weightsByFeature[feature.name] || 0);
    if (weight === 0) return;

    const values = game.featureValues[feature.name];
    if (!values) return;

    awayScore += values.away * weight * feature.direction;
    homeScore += values.home * weight * feature.direction;
    used++;
  });

  return {
    awayScore,
    homeScore,
    pick: awayScore > homeScore ? game.awayTeam : homeScore > awayScore ? game.homeTeam : "Tie",
    usedFeatures: used,
    hasScore: used > 0
  };
}


function appendFirstMismatchBreakdown_(output, mismatch, features, weightsByFeature) {
  const game = mismatch.game;
  const optScore = mismatch.optScore;

  output.push(["", "", "", "", "", "", "", "", "", "", "", "", ""]);
  output.push(["FIRST MISMATCH DETAIL", "", "", "", "", "", "", "", "", "", "", "", ""]);
  output.push(["Game", game.awayTeam + " @ " + game.homeTeam, "", "", "", "", "", "", "", "", "", "", ""]);
  output.push(["Live Away Score", game.liveAwayScore, "Optimizer Away Score", optScore.awayScore, "Diff", mismatch.awayDiff, "", "", "", "", "", "", ""]);
  output.push(["Live Home Score", game.liveHomeScore, "Optimizer Home Score", optScore.homeScore, "Diff", mismatch.homeDiff, "", "", "", "", "", "", ""]);
  output.push(["Live Pick", game.livePick, "Optimizer Pick", optScore.pick, "", "", "", "", "", "", "", "", ""]);
  output.push(["", "", "", "", "", "", "", "", "", "", "", "", ""]);
  output.push([
    "Feature",
    "Away Value",
    "Home Value",
    "Weight",
    "Direction",
    "Away Contribution",
    "Home Contribution",
    "Contribution Edge",
    "", "", "", "", ""
  ]);

  features.forEach(feature => {
    const values = game.featureValues[feature.name];
    if (!values) return;

    const weight = Number(weightsByFeature[feature.name] || 0);
    if (weight === 0) return;

    const awayContribution = values.away * weight * feature.direction;
    const homeContribution = values.home * weight * feature.direction;

    output.push([
      feature.name,
      values.away,
      values.home,
      weight,
      feature.direction,
      awayContribution,
      homeContribution,
      awayContribution - homeContribution,
      "", "", "", "", ""
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
  sheet.autoResizeColumns(1, Math.min(width, 13));
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
