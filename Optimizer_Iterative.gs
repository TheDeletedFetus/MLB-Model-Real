/***************************************
 * ITERATIVE MODEL OPTIMIZER v0.5.0
 *
 * Purpose:
 * - Backtest current weights against HISTORY
 * - Run true iterative hill-climbing optimization
 * - Accept the single best improvement each round
 * - Stop when no improvement remains or max rounds is reached
 * - Write OPTIMIZER_ITERATIVE
 * - Write Suggested Weight only
 * - Write OPTIMIZER_DEBUG
 *
 * Scoring:
 * Uses production scoring helpers from Scoring.js:
 * - calculateEdgeStats()
 * - scoreModelRowWithSettings()
 *
 * HISTORY stores Model_Matrix snapshot columns with MM_ prefixes.
 * This optimizer strips the MM_ prefix in-memory so Scoring.js can use
 * normal Model_Matrix-style headers like "Run Differential Edge".
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
  MAX_ROUNDS: 8,

  MIN_GAMES: 50,
  MIN_WIN_RATE: 0.52,
  MIN_ROI_IMPROVEMENT: 0.005,
  MIN_PROFIT_IMPROVEMENT: 0.01,

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
    awayTeam: ["MM_Away Team", "Away Team", "Away", "AwayTeam", "Visitor", "Visitor Team", "Road Team"],
    homeTeam: ["MM_Home Team", "Home Team", "Home", "HomeTeam"],
    awayML: ["MM_Away ML", "Pregame Away ML", "Latest Away ML", "Away ML", "Away Moneyline", "Away Odds", "Away Money Line", "Away Line"],
    homeML: ["MM_Home ML", "Pregame Home ML", "Latest Home ML", "Home ML", "Home Moneyline", "Home Odds", "Home Money Line", "Home Line"],
    winner: ["Winner", "Game Winner", "Actual Winner", "Winning Team", "Final Winner"],
    awayFinal: ["Away Score", "Away Final", "Away Runs", "Away Final Score"],
    homeFinal: ["Home Score", "Home Final", "Home Runs", "Home Final Score"]
  }
};


function runIterativeOptimizer() {
  const ss = SpreadsheetApp.getActive();
  const settingsSheet = ss.getSheetByName(ITER_OPT.SETTINGS_SHEET);
  const historySheet = ss.getSheetByName(ITER_OPT.HISTORY_SHEET);

  if (!settingsSheet) throw new Error("Missing Settings sheet.");
  if (!historySheet) throw new Error("Missing HISTORY sheet.");

  ensureSuggestedWeightColumn_(settingsSheet);

  const optimizerSettings = readOptimizerSettings_(settingsSheet);
  const history = readOptimizerHistory_(historySheet);

  if (history.games.length < ITER_OPT.MIN_GAMES) {
    throw new Error("Not enough usable HISTORY rows for optimizer testing. Found " + history.games.length + ". Minimum required: " + ITER_OPT.MIN_GAMES + ".");
  }

  const startingWeights = buildWeightMap_(optimizerSettings.features, "currentWeight");
  const baseline = backtestWithProductionScoring_(history, optimizerSettings.features, startingWeights);

  let currentWeights = Object.assign({}, startingWeights);
  let currentPerf = baseline;

  const results = [];
  results.push(makeResultRow_("BASELINE", "Current live weights", currentWeights, currentPerf, baseline, "CONTROL"));

  const acceptedSteps = [];

  for (let round = 1; round <= ITER_OPT.MAX_ROUNDS; round++) {
    const candidates = [];

    optimizerSettings.features
      .filter(feature => feature.active && feature.optimize)
      .forEach(feature => {
        for (let weight = feature.minWeight; weight <= feature.maxWeight; weight++) {
          if (Number(currentWeights[feature.name] || 0) === weight) continue;

          const testWeights = Object.assign({}, currentWeights);
          testWeights[feature.name] = weight;

          const perf = backtestWithProductionScoring_(history, optimizerSettings.features, testWeights);
          const candidate = {
            round,
            feature,
            weight,
            weights: testWeights,
            perf
          };

          candidates.push(candidate);
          results.push(makeResultRow_(
            "ROUND " + round + " TEST",
            feature.name + " = " + weight,
            testWeights,
            perf,
            baseline,
            decideResult_(perf, currentPerf, baseline)
          ));
        }
      });

    candidates.sort((a, b) => compareCandidate_(a, b, currentPerf, baseline));
    const best = candidates[0];

    if (!best || !isMeaningfulImprovement_(best.perf, currentPerf, baseline)) {
      results.push(makeResultRow_(
        "STOP",
        "No meaningful improvement found after round " + round,
        currentWeights,
        currentPerf,
        baseline,
        "STOP"
      ));
      break;
    }

    currentWeights = Object.assign({}, best.weights);
    currentPerf = best.perf;
    acceptedSteps.push(best);

    results.push(makeResultRow_(
      "ROUND " + round + " ACCEPTED",
      best.feature.name + " -> " + best.weight,
      currentWeights,
      currentPerf,
      baseline,
      "ACCEPT"
    ));
  }

  const finalDecision = passesAcceptanceRules_(currentPerf, baseline) ? "FINAL RECOMMENDATION" : "FINAL WATCHLIST";
  results.push(makeResultRow_(
    finalDecision,
    acceptedSteps.length ? acceptedSteps.map(step => step.feature.name + "=" + step.weight).join(", ") : "No accepted changes",
    currentWeights,
    currentPerf,
    baseline,
    finalDecision === "FINAL RECOMMENDATION" ? "ACCEPT" : "WATCHLIST"
  ));

  writeSuggestedWeights_(settingsSheet, optimizerSettings, currentWeights, finalDecision === "FINAL RECOMMENDATION");
  writeOptimizerResults_(ss, results, optimizerSettings.features);
  writeOptimizerDebug_(ss, history, optimizerSettings.features, startingWeights, currentWeights, acceptedSteps, baseline, currentPerf);
  SpreadsheetApp.flush();
}


function backtestWithProductionScoring_(history, features, weightsByFeature) {
  const scoringSettings = buildProductionSettingsFromWeights_(features, weightsByFeature);
  const edgeStats = calculateEdgeStats(history.rows, history.scoringHeadersArray, scoringSettings);

  let gamesCount = 0;
  let wins = 0;
  let losses = 0;
  let profit = 0;
  let totalRisk = 0;

  history.games.forEach(game => {
    const score = scoreModelRowWithSettings(
      game.row,
      history.scoringHeadersArray,
      scoringSettings,
      edgeStats
    );

    if (!score || score.pick === "Coin Flip") return;

    const moneyline = sameTeam_(score.pick, game.awayTeam) ? game.awayML : game.homeML;
    if (!isValidMoneyline_(moneyline)) return;

    const payout = americanOddsProfit_(moneyline, 100);
    if (!isFinite(payout)) return;

    gamesCount++;
    totalRisk += 100;

    if (sameTeam_(score.pick, game.winner)) {
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


function buildProductionSettingsFromWeights_(features, weightsByFeature) {
  const settings = [];

  features.forEach(feature => {
    if (!feature.active) return;

    const weight = Number(weightsByFeature[feature.name] || 0);
    if (!isFinite(weight) || weight <= 0) return;

    settings.push({
      stat: feature.name,
      weight
    });
  });

  return settings;
}


function compareCandidate_(a, b, currentPerf, baseline) {
  const aImproves = isMeaningfulImprovement_(a.perf, currentPerf, baseline);
  const bImproves = isMeaningfulImprovement_(b.perf, currentPerf, baseline);

  if (aImproves && !bImproves) return -1;
  if (!aImproves && bImproves) return 1;

  const roiDiff = finiteOrZero_(b.perf.roi) - finiteOrZero_(a.perf.roi);
  if (Math.abs(roiDiff) > 0.000001) return roiDiff;

  const profitDiff = finiteOrZero_(b.perf.profit) - finiteOrZero_(a.perf.profit);
  if (Math.abs(profitDiff) > 0.000001) return profitDiff;

  return finiteOrZero_(b.perf.winRate) - finiteOrZero_(a.perf.winRate);
}


function isMeaningfulImprovement_(test, current, baseline) {
  if (finiteOrZero_(test.games) < ITER_OPT.MIN_GAMES) return false;
  if (finiteOrZero_(test.winRate) < ITER_OPT.MIN_WIN_RATE) return false;

  const roiImprovement = finiteOrZero_(test.roi) - finiteOrZero_(current.roi);
  const profitImprovement = finiteOrZero_(test.profit) - finiteOrZero_(current.profit);

  if (roiImprovement < ITER_OPT.MIN_ROI_IMPROVEMENT) return false;
  if (profitImprovement < ITER_OPT.MIN_PROFIT_IMPROVEMENT) return false;

  return finiteOrZero_(test.profit) > finiteOrZero_(baseline.profit);
}


function decideResult_(test, current, baseline) {
  if (isMeaningfulImprovement_(test, current, baseline)) return "CANDIDATE";
  if (finiteOrZero_(test.profit) > finiteOrZero_(current.profit) || finiteOrZero_(test.roi) > finiteOrZero_(current.roi)) return "WATCHLIST";
  return "REJECT";
}


function readOptimizerHistory_(sheet) {
  const values = sheet.getDataRange().getValues();
  const headerInfo = findHistoryHeaderRow_(values);
  const originalHeadersArray = values[headerInfo.rowIndex];
  const scoringHeadersArray = buildScoringHeadersFromHistoryHeaders_(originalHeadersArray);
  const rows = values.slice(headerInfo.rowIndex + 1);

  const indexes = headerInfo.indexes;
  const games = [];

  rows.forEach(row => {
    const awayTeam = cleanText_(row[indexes.awayTeam]);
    const homeTeam = cleanText_(row[indexes.homeTeam]);
    const winner = resolveWinner_(row, indexes, awayTeam, homeTeam);
    const awayML = Number(row[indexes.awayML]);
    const homeML = Number(row[indexes.homeML]);

    if (!awayTeam || !homeTeam || !winner) return;
    if (!isValidMoneyline_(awayML) || !isValidMoneyline_(homeML)) return;

    games.push({
      row,
      awayTeam,
      homeTeam,
      winner: cleanText_(winner),
      awayML,
      homeML
    });
  });

  return {
    originalHeadersArray,
    scoringHeadersArray,
    rows,
    indexes,
    games
  };
}


function buildScoringHeadersFromHistoryHeaders_(headers) {
  return headers.map(header => {
    const text = String(header || "").trim();
    if (text.indexOf("MM_") === 0) {
      return text.substring(3);
    }
    return text;
  });
}


function writeOptimizerDebug_(ss, history, features, startingWeights, finalWeights, acceptedSteps, baseline, finalPerf) {
  let sheet = ss.getSheetByName(ITER_OPT.DEBUG_SHEET);
  if (!sheet) sheet = ss.insertSheet(ITER_OPT.DEBUG_SHEET);
  sheet.clearContents();

  const output = [];
  output.push(["Debug Item", "Value"]);
  output.push(["Usable Games", history.games.length]);
  output.push(["Accepted Steps", acceptedSteps.length]);
  output.push(["Baseline Games", baseline.games]);
  output.push(["Baseline Wins", baseline.wins]);
  output.push(["Baseline Losses", baseline.losses]);
  output.push(["Baseline Win %", baseline.winRate]);
  output.push(["Baseline Profit", baseline.profit]);
  output.push(["Baseline ROI", baseline.roi]);
  output.push(["Final Games", finalPerf.games]);
  output.push(["Final Wins", finalPerf.wins]);
  output.push(["Final Losses", finalPerf.losses]);
  output.push(["Final Win %", finalPerf.winRate]);
  output.push(["Final Profit", finalPerf.profit]);
  output.push(["Final ROI", finalPerf.roi]);
  output.push(["", ""]);
  output.push(["Accepted Step", "Change", "Games", "Wins", "Losses", "Win %", "Profit", "ROI"]);

  acceptedSteps.forEach((step, idx) => {
    output.push([
      idx + 1,
      step.feature.name + " -> " + step.weight,
      step.perf.games,
      step.perf.wins,
      step.perf.losses,
      step.perf.winRate,
      step.perf.profit,
      step.perf.roi
    ]);
  });

  output.push(["", ""]);
  output.push(["Sample Game Breakdown", ""]);

  if (history.games.length) {
    const game = history.games[0];
    const finalSettings = buildProductionSettingsFromWeights_(features, finalWeights);
    const edgeStats = calculateEdgeStats(history.rows, history.scoringHeadersArray, finalSettings);
    const score = scoreModelRowWithSettings(game.row, history.scoringHeadersArray, finalSettings, edgeStats);

    output.push(["Sample Game", game.awayTeam + " @ " + game.homeTeam]);
    output.push(["Winner", game.winner]);
    output.push(["Final Away Score", score.awayScore]);
    output.push(["Final Home Score", score.homeScore]);
    output.push(["Final Pick", score.pick]);
    output.push(["", ""]);
    output.push(["Feature", "Weight", "Raw Edge", "Mean", "StdDev", "Z-Score", "Contribution"]);

    score.contributions.forEach(item => {
      output.push([
        item.stat,
        item.weight,
        item.rawEdge,
        item.mean,
        item.stdDev,
        item.zScore,
        item.contribution
      ]);
    });
  }

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
      maxWeight
    });
  }

  return {
    headers,
    headerRowNumber: headerRowIndex + 1,
    features
  };
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


function makeResultRow_(stage, test, weights, perf, baseline, overrideDecision) {
  const roiChange = finiteOrZero_(perf.roi) - finiteOrZero_(baseline.roi);

  let decision = overrideDecision || "REJECT";

  if (!overrideDecision) {
    if (stage === "BASELINE") {
      decision = "CONTROL";
    } else if (passesAcceptanceRules_(perf, baseline)) {
      decision = "ACCEPT";
    } else if (finiteOrZero_(perf.profit) > finiteOrZero_(baseline.profit) || finiteOrZero_(perf.roi) > finiteOrZero_(baseline.roi)) {
      decision = "WATCHLIST";
    }
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


function findSettingsHeaderRow_(values) {
  const h = ITER_OPT.SETTINGS_HEADERS;
  return findHeaderRow_(values, [h.stat, h.active, h.weight], "Settings");
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
