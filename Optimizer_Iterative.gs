/***************************************
 * ITERATIVE MODEL OPTIMIZER v0.2.3
 *
 * Purpose:
 * - Read Settings once
 * - Read HISTORY once
 * - Build an in-memory game cache once
 * - Run Stage 1 individual stat tests in memory
 * - Write OPTIMIZER_ITERATIVE once
 * - Write Suggested Weight only
 * - Write OPTIMIZER_DEBUG once
 *
 * Important HISTORY format note:
 * Current HISTORY stores pregame Model_Matrix fields with MM_ prefixes,
 * e.g. MM_Away Runs/Game, MM_Home Runs/Game, MM_Away ML, MM_Home ML.
 * This optimizer explicitly supports that format.
 *
 * Run manually:
 *   runIterativeOptimizer()
 ***************************************/

const ITER_OPT = {
  SETTINGS_SHEET: "Settings",
  HISTORY_SHEET: "HISTORY",
  OUTPUT_SHEET: "OPTIMIZER_ITERATIVE",
  DEBUG_SHEET: "OPTIMIZER_DEBUG",

  DEFAULT_MIN_WEIGHT: 0,
  DEFAULT_MAX_WEIGHT: 10,

  MIN_GAMES: 50,
  MIN_WIN_RATE: 0.52,
  MIN_ROI_IMPROVEMENT: 0.005,

  SETTINGS_HEADERS: {
    stat: "Stat",
    category: "Category",
    active: "Active",
    weight: "Weight",
    direction: "Direction",
    suggestedWeight: "Suggested Weight",
    minWeight: "Min Weight",
    maxWeight: "Max Weight",
    optimize: "Optimize?"
  },

  HISTORY_ALIASES: {
    awayTeam: ["Away Team", "MM_Away Team", "Away", "AwayTeam", "Visitor", "Visitor Team", "Road Team"],
    homeTeam: ["Home Team", "MM_Home Team", "Home", "HomeTeam"],
    awayML: ["MM_Away ML", "Pregame Away ML", "Latest Away ML", "Away ML", "Away Moneyline", "Away Odds", "Away Money Line", "Away Line"],
    homeML: ["MM_Home ML", "Pregame Home ML", "Latest Home ML", "Home ML", "Home Moneyline", "Home Odds", "Home Money Line", "Home Line"],
    winner: ["Winner", "Game Winner", "Actual Winner", "Winning Team", "Final Winner"],
    awayFinal: ["Away Score", "Away Final", "Away Runs", "Away Final Score"],
    homeFinal: ["Home Score", "Home Final", "Home Runs", "Home Final Score"]
  },

  STAT_ALIASES: {
    "OPS vs Hand": ["OPS vs SP Hand"],
    "Home OPS": ["Split OPS"],
    "Road OPS": ["Split OPS"],
    "Home ERA": ["Split ERA"],
    "Road ERA": ["Split ERA"],
    "Home WHIP": ["Split WHIP"],
    "Road WHIP": ["Split WHIP"],
    "Home K/9": ["Split K/9"],
    "Road K/9": ["Split K/9"],
    "Home K/BB": ["Split K/BB"],
    "Road K/BB": ["Split K/BB"]
  }
};


function runIterativeOptimizer() {
  const ss = SpreadsheetApp.getActive();
  const settingsSheet = ss.getSheetByName(ITER_OPT.SETTINGS_SHEET);
  const historySheet = ss.getSheetByName(ITER_OPT.HISTORY_SHEET);

  if (!settingsSheet) throw new Error("Missing Settings sheet.");
  if (!historySheet) throw new Error("Missing HISTORY sheet.");

  ensureSuggestedWeightColumn_(settingsSheet);

  const settings = readOptimizerSettings_(settingsSheet);
  const historyRaw = readHistoryRaw_(historySheet, settings.features);
  const games = buildOptimizerGameCache_(historyRaw, settings.features);

  if (games.length < ITER_OPT.MIN_GAMES) {
    throw new Error("Not enough usable HISTORY rows for optimizer testing. Found " + games.length + ". Minimum required: " + ITER_OPT.MIN_GAMES + ". Check odds availability and MM_ feature mapping.");
  }

  const baselineWeights = buildWeightMap_(settings.features, "currentWeight");
  const baseline = backtestCachedGames_(games, settings.features, baselineWeights);

  const results = [];
  results.push(makeResultRow_("BASELINE", "Current live weights", baselineWeights, baseline, baseline));

  const stage1Best = [];

  settings.features
    .filter(feature => feature.active && feature.optimize)
    .forEach(feature => {
      let best = null;

      for (let weight = feature.minWeight; weight <= feature.maxWeight; weight++) {
        const testWeights = Object.assign({}, baselineWeights);
        testWeights[feature.name] = weight;

        const perf = backtestCachedGames_(games, settings.features, testWeights);
        results.push(makeResultRow_("STAGE 1 TEST", feature.name + " = " + weight, testWeights, perf, baseline));

        if (!best || isBetterResult_(perf, best.perf, baseline)) {
          best = {
            feature,
            weight,
            weights: testWeights,
            perf
          };
        }
      }

      if (best) {
        stage1Best.push(best);
        results.push(makeResultRow_("STAGE 1 BEST", best.feature.name + " best = " + best.weight, best.weights, best.perf, baseline));
      }
    });

  stage1Best.sort((a, b) => b.perf.roi - a.perf.roi);

  const accepted = stage1Best.find(candidate => passesAcceptanceRules_(candidate.perf, baseline));
  const finalCandidate = accepted || stage1Best[0];

  if (finalCandidate) {
    results.push(makeResultRow_(
      accepted ? "FINAL RECOMMENDATION" : "FINAL WATCHLIST",
      finalCandidate.feature.name,
      finalCandidate.weights,
      finalCandidate.perf,
      baseline
    ));

    writeSuggestedWeights_(settingsSheet, settings, finalCandidate.weights, Boolean(accepted));
  }

  writeOptimizerResults_(ss, results, settings.features);
  writeOptimizerDebug_(ss, games, settings.features, baselineWeights, stage1Best[0]);
  SpreadsheetApp.flush();
}


function buildOptimizerGameCache_(historyRaw, features) {
  const games = [];

  historyRaw.rows.forEach(row => {
    const awayTeam = cleanText_(row[historyRaw.indexes.awayTeam]);
    const homeTeam = cleanText_(row[historyRaw.indexes.homeTeam]);
    const winner = resolveWinner_(row, historyRaw.indexes, awayTeam, homeTeam);
    const awayML = Number(row[historyRaw.indexes.awayML]);
    const homeML = Number(row[historyRaw.indexes.homeML]);

    if (!awayTeam || !homeTeam || !winner) return;
    if (!isValidMoneyline_(awayML) || !isValidMoneyline_(homeML)) return;

    const featureValues = {};
    let usableFeatureCount = 0;

    features.forEach(feature => {
      const columns = historyRaw.featureColumns[feature.name];
      if (!columns) return;

      const awayVal = Number(row[columns.away]);
      const homeVal = Number(row[columns.home]);

      if (!isFinite(awayVal) || !isFinite(homeVal)) return;

      featureValues[feature.name] = {
        away: awayVal,
        home: homeVal
      };
      usableFeatureCount++;
    });

    if (usableFeatureCount === 0) return;

    games.push({
      awayTeam,
      homeTeam,
      winner: cleanText_(winner),
      awayML,
      homeML,
      featureValues
    });
  });

  return games;
}


function backtestCachedGames_(games, features, weightsByFeature) {
  let gamesCount = 0;
  let wins = 0;
  let losses = 0;
  let profit = 0;
  let totalRisk = 0;

  games.forEach(game => {
    let awayScore = 0;
    let homeScore = 0;

    features.forEach(feature => {
      const weight = Number(weightsByFeature[feature.name] || 0);
      if (weight === 0) return;

      const values = game.featureValues[feature.name];
      if (!values) return;

      awayScore += values.away * weight * feature.direction;
      homeScore += values.home * weight * feature.direction;
    });

    if (awayScore === homeScore) return;

    const pick = awayScore > homeScore ? game.awayTeam : game.homeTeam;
    const moneyline = pick === game.awayTeam ? game.awayML : game.homeML;

    if (!isValidMoneyline_(moneyline)) return;

    const payout = americanOddsProfit_(moneyline, 100);
    if (!isFinite(payout)) return;

    gamesCount++;
    totalRisk += 100;

    if (sameTeam_(pick, game.winner)) {
      wins++;
      profit += payout;
    } else {
      losses++;
      profit -= 100;
    }
  });

  profit = finiteOrZero_(profit);
  totalRisk = finiteOrZero_(totalRisk);

  return {
    games: gamesCount,
    wins,
    losses,
    winRate: gamesCount ? wins / gamesCount : 0,
    profit,
    roi: totalRisk ? profit / totalRisk : 0
  };
}


function writeOptimizerDebug_(ss, games, features, baselineWeights, topCandidate) {
  let sheet = ss.getSheetByName(ITER_OPT.DEBUG_SHEET);
  if (!sheet) sheet = ss.insertSheet(ITER_OPT.DEBUG_SHEET);
  sheet.clearContents();

  const output = [];
  output.push(["Debug Item", "Value"]);
  output.push(["Usable Games", games.length]);

  if (!games.length) {
    sheet.getRange(1, 1, output.length, output[0].length).setValues(output);
    return;
  }

  const game = games[0];
  const baselineScore = calculateGameScore_(game, features, baselineWeights);
  const topWeights = topCandidate ? topCandidate.weights : baselineWeights;
  const topScore = calculateGameScore_(game, features, topWeights);

  output.push(["Sample Game", game.awayTeam + " @ " + game.homeTeam]);
  output.push(["Winner", game.winner]);
  output.push(["Away ML", game.awayML]);
  output.push(["Home ML", game.homeML]);
  output.push(["Baseline Away Score", baselineScore.awayScore]);
  output.push(["Baseline Home Score", baselineScore.homeScore]);
  output.push(["Baseline Pick", baselineScore.pick]);
  output.push(["Top Test", topCandidate ? topCandidate.feature.name + " = " + topCandidate.weight : "N/A"]);
  output.push(["Top Away Score", topScore.awayScore]);
  output.push(["Top Home Score", topScore.homeScore]);
  output.push(["Top Pick", topScore.pick]);
  output.push(["", ""]);
  output.push(["Feature", "Away Value", "Home Value", "Baseline Weight", "Direction", "Baseline Away Contribution", "Baseline Home Contribution", "Top Weight", "Top Away Contribution", "Top Home Contribution"]);

  features.forEach(feature => {
    const values = game.featureValues[feature.name];
    if (!values) return;

    const baseWeight = Number(baselineWeights[feature.name] || 0);
    const candidateWeight = Number(topWeights[feature.name] || 0);

    output.push([
      feature.name,
      values.away,
      values.home,
      baseWeight,
      feature.direction,
      values.away * baseWeight * feature.direction,
      values.home * baseWeight * feature.direction,
      candidateWeight,
      values.away * candidateWeight * feature.direction,
      values.home * candidateWeight * feature.direction
    ]);
  });

  const width = Math.max.apply(null, output.map(row => row.length));
  const rectangular = output.map(row => {
    const copy = row.slice();
    while (copy.length < width) copy.push("");
    return copy;
  });

  sheet.getRange(1, 1, rectangular.length, width).setValues(rectangular);
  sheet.getRange(1, 1, 1, width).setFontWeight("bold");
  sheet.autoResizeColumns(1, Math.min(width, 10));
}


function calculateGameScore_(game, features, weightsByFeature) {
  let awayScore = 0;
  let homeScore = 0;

  features.forEach(feature => {
    const weight = Number(weightsByFeature[feature.name] || 0);
    if (weight === 0) return;

    const values = game.featureValues[feature.name];
    if (!values) return;

    awayScore += values.away * weight * feature.direction;
    homeScore += values.home * weight * feature.direction;
  });

  return {
    awayScore,
    homeScore,
    pick: awayScore > homeScore ? game.awayTeam : homeScore > awayScore ? game.homeTeam : "Tie"
  };
}


function readOptimizerSettings_(sheet) {
  const values = sheet.getDataRange().getValues();
  const headerInfo = findSettingsHeaderRow_(values);
  const headers = headerInfo.headers;
  const headerRowIndex = headerInfo.rowIndex;
  const h = ITER_OPT.SETTINGS_HEADERS;

  const features = [];

  for (let rowIndex = headerRowIndex + 1; rowIndex < values.length; rowIndex++) {
    const row = values[rowIndex];
    const name = cleanText_(row[headers[h.stat]]);
    if (!name) continue;

    const active = parseBool_(row[headers[h.active]]);
    const optimize = headers[h.optimize] === undefined ? active : parseBool_(row[headers[h.optimize]]);

    const minWeight = headers[h.minWeight] === undefined
      ? ITER_OPT.DEFAULT_MIN_WEIGHT
      : parseNumberOrDefault_(row[headers[h.minWeight]], ITER_OPT.DEFAULT_MIN_WEIGHT);

    const maxWeight = headers[h.maxWeight] === undefined
      ? ITER_OPT.DEFAULT_MAX_WEIGHT
      : parseNumberOrDefault_(row[headers[h.maxWeight]], ITER_OPT.DEFAULT_MAX_WEIGHT);

    features.push({
      rowNumber: rowIndex + 1,
      name,
      category: headers[h.category] === undefined ? "" : cleanText_(row[headers[h.category]]),
      active,
      optimize,
      currentWeight: parseNumberOrDefault_(row[headers[h.weight]], 0),
      minWeight,
      maxWeight,
      direction: parseDirection_(row[headers[h.direction]])
    });
  }

  return {
    headers,
    headerRowNumber: headerRowIndex + 1,
    features
  };
}


function readHistoryRaw_(sheet, features) {
  const values = sheet.getDataRange().getValues();
  const headerInfo = findHistoryHeaderRow_(values);
  const indexes = headerInfo.indexes;

  if (indexes.winner === undefined && (indexes.awayFinal === undefined || indexes.homeFinal === undefined)) {
    throw new Error("HISTORY must contain either a winner column or both final score columns. Found headers: " + headerInfo.foundHeaders.join(" | "));
  }

  const featureColumns = resolveFeatureColumnsForSettings_(headerInfo.headers, features);
  Logger.log("Optimizer mapped " + Object.keys(featureColumns).length + " feature columns out of " + features.length + " Settings rows.");

  return {
    headers: headerInfo.headers,
    indexes,
    rows: values.slice(headerInfo.rowIndex + 1),
    featureColumns
  };
}


function resolveFeatureColumnsForSettings_(headers, features) {
  const result = {};

  features.forEach(feature => {
    const pair = resolveFeaturePair_(headers, feature.name);

    if (pair) {
      result[feature.name] = pair;
    }
  });

  return result;
}


function resolveFeaturePair_(headers, statName) {
  const statCandidates = buildStatCandidates_(statName);

  for (let i = 0; i < statCandidates.length; i++) {
    const stat = statCandidates[i];

    const awayCandidates = [
      "MM_Away " + stat,
      "Away " + stat,
      "Away_" + stat,
      "Away" + stat
    ];

    const homeCandidates = [
      "MM_Home " + stat,
      "Home " + stat,
      "Home_" + stat,
      "Home" + stat
    ];

    const awayIdx = firstExistingHeaderIndex_(headers, awayCandidates);
    const homeIdx = firstExistingHeaderIndex_(headers, homeCandidates);

    if (awayIdx !== undefined && homeIdx !== undefined) {
      return {
        away: awayIdx,
        home: homeIdx
      };
    }
  }

  const splitPair = resolveSplitFeaturePair_(headers, statName);
  if (splitPair) return splitPair;

  return null;
}


function buildStatCandidates_(statName) {
  const candidates = [statName];
  const aliases = ITER_OPT.STAT_ALIASES[statName] || [];

  aliases.forEach(alias => candidates.push(alias));

  return candidates;
}


function resolveSplitFeaturePair_(headers, statName) {
  const splitMap = {
    "Home OPS": ["MM_Away Road OPS", "MM_Home Home OPS"],
    "Road OPS": ["MM_Away Road OPS", "MM_Home Home OPS"],
    "Home ERA": ["MM_Away Road ERA", "MM_Home Home ERA"],
    "Road ERA": ["MM_Away Road ERA", "MM_Home Home ERA"],
    "Home WHIP": ["MM_Away Road WHIP", "MM_Home Home WHIP"],
    "Road WHIP": ["MM_Away Road WHIP", "MM_Home Home WHIP"],
    "Starter Split ERA": ["MM_Away Road Starter ERA", "MM_Home Home Starter ERA"],
    "Starter Split WHIP": ["MM_Away Road Starter WHIP", "MM_Home Home Starter WHIP"],
    "Starter Split K/BB": ["MM_Away Road Starter K/BB", "MM_Home Home Starter K/BB"]
  };

  if (!splitMap[statName]) return null;

  const awayIdx = getHeaderIndex_(headers, splitMap[statName][0]);
  const homeIdx = getHeaderIndex_(headers, splitMap[statName][1]);

  if (awayIdx !== undefined && homeIdx !== undefined) {
    return {
      away: awayIdx,
      home: homeIdx
    };
  }

  return null;
}


function firstExistingHeaderIndex_(headers, candidates) {
  for (let i = 0; i < candidates.length; i++) {
    const idx = getHeaderIndex_(headers, candidates[i]);
    if (idx !== undefined) return idx;
  }

  return undefined;
}


function ensureSuggestedWeightColumn_(sheet) {
  const values = sheet.getDataRange().getValues();
  const headerInfo = findSettingsHeaderRow_(values);
  const headerRowNumber = headerInfo.rowIndex + 1;
  const headers = headerInfo.headers;
  const suggestedHeader = ITER_OPT.SETTINGS_HEADERS.suggestedWeight;

  if (headers[suggestedHeader] === undefined) {
    sheet.getRange(headerRowNumber, sheet.getLastColumn() + 1).setValue(suggestedHeader);
  }
}


function writeSuggestedWeights_(sheet, settings, weightsByFeature, accepted) {
  const suggestedCol = settings.headers[ITER_OPT.SETTINGS_HEADERS.suggestedWeight] + 1;

  settings.features.forEach(feature => {
    const suggested = weightsByFeature[feature.name];
    if (suggested === undefined) return;

    sheet.getRange(feature.rowNumber, suggestedCol).setValue(
      accepted ? suggested : "WATCH: " + suggested
    );
  });
}


function writeOptimizerResults_(ss, results, features) {
  let sheet = ss.getSheetByName(ITER_OPT.OUTPUT_SHEET);
  if (!sheet) sheet = ss.insertSheet(ITER_OPT.OUTPUT_SHEET);

  sheet.clearContents();

  const headers = [
    "Stage",
    "Test",
    "Games",
    "Wins",
    "Losses",
    "Win %",
    "Profit",
    "ROI",
    "ROI Change",
    "Decision"
  ];

  features.forEach(feature => headers.push(feature.name));

  const output = [headers];

  results.forEach(result => {
    const row = [
      result.stage,
      result.test,
      safeSheetValue_(result.games),
      safeSheetValue_(result.wins),
      safeSheetValue_(result.losses),
      safeSheetValue_(result.winRate),
      safeSheetValue_(result.profit),
      safeSheetValue_(result.roi),
      safeSheetValue_(result.roiChange),
      result.decision
    ];

    features.forEach(feature => {
      row.push(safeSheetValue_(result.weights[feature.name] || 0));
    });

    output.push(row);
  });

  sheet.getRange(1, 1, output.length, output[0].length).setValues(output);
  sheet.getRange(1, 1, 1, output[0].length).setFontWeight("bold");

  if (output.length > 1) {
    sheet.getRange(2, 6, output.length - 1, 1).setNumberFormat("0.00%");
    sheet.getRange(2, 8, output.length - 1, 2).setNumberFormat("0.00%");
    sheet.getRange(2, 7, output.length - 1, 1).setNumberFormat("$0.00");
  }

  sheet.autoResizeColumns(1, Math.min(output[0].length, 20));
}


function makeResultRow_(stage, test, weights, perf, baseline) {
  const roiChange = finiteOrZero_(perf.roi) - finiteOrZero_(baseline.roi);

  let decision = "REJECT";

  if (stage === "BASELINE") {
    decision = "CONTROL";
  } else if (passesAcceptanceRules_(perf, baseline)) {
    decision = "ACCEPT";
  } else if (finiteOrZero_(perf.profit) > finiteOrZero_(baseline.profit) || finiteOrZero_(perf.roi) > finiteOrZero_(baseline.roi)) {
    decision = "WATCHLIST";
  }

  return {
    stage,
    test,
    games: finiteOrZero_(perf.games),
    wins: finiteOrZero_(perf.wins),
    losses: finiteOrZero_(perf.losses),
    winRate: finiteOrZero_(perf.winRate),
    profit: finiteOrZero_(perf.profit),
    roi: finiteOrZero_(perf.roi),
    roiChange: finiteOrZero_(roiChange),
    decision,
    weights
  };
}


function buildWeightMap_(features, key) {
  const map = {};
  features.forEach(feature => {
    map[feature.name] = Number(feature[key] || 0);
  });
  return map;
}


function resolveWinner_(row, indexes, awayTeam, homeTeam) {
  if (indexes.winner !== undefined) {
    return cleanText_(row[indexes.winner]);
  }

  if (indexes.awayFinal !== undefined && indexes.homeFinal !== undefined) {
    const awayFinal = Number(row[indexes.awayFinal]);
    const homeFinal = Number(row[indexes.homeFinal]);

    if (isFinite(awayFinal) && isFinite(homeFinal) && awayFinal !== homeFinal) {
      return awayFinal > homeFinal ? awayTeam : homeTeam;
    }
  }

  return "";
}


function americanOddsProfit_(odds, risk) {
  const cleanOdds = Number(odds);
  const cleanRisk = Number(risk);

  if (!isValidMoneyline_(cleanOdds) || !isFinite(cleanRisk) || cleanRisk <= 0) return 0;
  if (cleanOdds > 0) return cleanRisk * (cleanOdds / 100);
  return cleanRisk * (100 / Math.abs(cleanOdds));
}


function passesAcceptanceRules_(perf, baseline) {
  if (finiteOrZero_(perf.games) < ITER_OPT.MIN_GAMES) return false;
  if (finiteOrZero_(perf.winRate) < ITER_OPT.MIN_WIN_RATE) return false;
  if ((finiteOrZero_(perf.roi) - finiteOrZero_(baseline.roi)) < ITER_OPT.MIN_ROI_IMPROVEMENT) return false;
  if (finiteOrZero_(perf.profit) <= finiteOrZero_(baseline.profit)) return false;
  return true;
}


function isBetterResult_(test, currentBest, baseline) {
  if (!currentBest) return true;

  const testPass = passesAcceptanceRules_(test, baseline);
  const bestPass = passesAcceptanceRules_(currentBest, baseline);

  if (testPass && !bestPass) return true;
  if (!testPass && bestPass) return false;

  return finiteOrZero_(test.roi) > finiteOrZero_(currentBest.roi);
}


function findSettingsHeaderRow_(values) {
  const h = ITER_OPT.SETTINGS_HEADERS;
  return findHeaderRow_(values, [h.stat, h.active, h.weight, h.direction], "Settings");
}


function findHistoryHeaderRow_(values) {
  for (let rowIndex = 0; rowIndex < values.length; rowIndex++) {
    const headers = mapHeaders_(values[rowIndex]);
    const foundHeaders = Object.keys(headers);

    const indexes = {
      awayTeam: findAliasIndex_(headers, ITER_OPT.HISTORY_ALIASES.awayTeam),
      homeTeam: findAliasIndex_(headers, ITER_OPT.HISTORY_ALIASES.homeTeam),
      awayML: findAliasIndex_(headers, ITER_OPT.HISTORY_ALIASES.awayML),
      homeML: findAliasIndex_(headers, ITER_OPT.HISTORY_ALIASES.homeML),
      winner: findAliasIndex_(headers, ITER_OPT.HISTORY_ALIASES.winner),
      awayFinal: findAliasIndex_(headers, ITER_OPT.HISTORY_ALIASES.awayFinal),
      homeFinal: findAliasIndex_(headers, ITER_OPT.HISTORY_ALIASES.homeFinal)
    };

    const hasTeams = indexes.awayTeam !== undefined && indexes.homeTeam !== undefined;
    const hasOdds = indexes.awayML !== undefined && indexes.homeML !== undefined;
    const hasResult = indexes.winner !== undefined || (indexes.awayFinal !== undefined && indexes.homeFinal !== undefined);

    if (hasTeams && hasOdds && hasResult) {
      return {
        rowIndex,
        headers,
        indexes,
        foundHeaders
      };
    }
  }

  const sampledHeaders = values
    .slice(0, 10)
    .map((row, i) => "Row " + (i + 1) + ": " + row.filter(Boolean).join(" | "))
    .join("\n");

  throw new Error(
    "Could not find usable HISTORY header row. Need away team, home team, away ML, home ML, and either winner or final scores. Sampled rows:\n" + sampledHeaders
  );
}


function findHeaderRow_(values, requiredHeaders, sheetNameForError) {
  for (let rowIndex = 0; rowIndex < values.length; rowIndex++) {
    const headers = mapHeaders_(values[rowIndex]);
    const foundAll = requiredHeaders.every(header => headers[header] !== undefined);

    if (foundAll) {
      return {
        rowIndex,
        headers
      };
    }
  }

  throw new Error(
    "Could not find header row on " + sheetNameForError +
    ". Required headers: " + requiredHeaders.join(", ")
  );
}


function findAliasIndex_(headers, aliases) {
  for (let i = 0; i < aliases.length; i++) {
    const idx = getHeaderIndex_(headers, aliases[i]);
    if (idx !== undefined) return idx;
  }

  return undefined;
}


function getHeaderIndex_(headers, headerName) {
  if (headers[headerName] !== undefined) return headers[headerName];

  const target = normalizeHeader_(headerName);
  const keys = Object.keys(headers);

  for (let i = 0; i < keys.length; i++) {
    if (normalizeHeader_(keys[i]) === target) return headers[keys[i]];
  }

  return undefined;
}


function mapHeaders_(headerRow) {
  const map = {};

  headerRow.forEach((header, index) => {
    if (header !== "" && header !== null && header !== undefined) {
      map[String(header).trim()] = index;
    }
  });

  return map;
}


function normalizeHeader_(value) {
  return String(value || "")
    .toLowerCase()
    .replace(/[^a-z0-9]/g, "");
}


function parseBool_(value) {
  if (value === true) return true;
  const normalized = String(value).trim().toUpperCase();
  return normalized === "TRUE" || normalized === "YES" || normalized === "1";
}


function parseDirection_(value) {
  const normalized = cleanText_(value).toLowerCase();

  if (["lower", "low", "less", "negative", "-1", "down"].indexOf(normalized) !== -1) {
    return -1;
  }

  if (["higher", "high", "more", "positive", "+1", "1", "up"].indexOf(normalized) !== -1) {
    return 1;
  }

  const numeric = Number(value);
  if (isFinite(numeric) && numeric < 0) return -1;

  return 1;
}


function parseNumberOrDefault_(value, fallback) {
  const n = Number(value);
  return isFinite(n) ? n : fallback;
}


function cleanText_(value) {
  return String(value || "").trim();
}


function sameTeam_(a, b) {
  return normalizeHeader_(a) === normalizeHeader_(b);
}


function isValidMoneyline_(value) {
  const n = Number(value);
  return isFinite(n) && n !== 0;
}


function finiteOrZero_(value) {
  const n = Number(value);
  return isFinite(n) ? n : 0;
}


function safeSheetValue_(value) {
  if (typeof value === "number") return isFinite(value) ? value : 0;
  return value === undefined || value === null ? "" : value;
}
