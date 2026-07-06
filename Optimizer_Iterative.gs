/***************************************
 * ITERATIVE MODEL OPTIMIZER v0.1.2
 *
 * Safe optimizer:
 * - Does NOT overwrite live Weight
 * - Writes Suggested Weight only
 * - Writes full results to OPTIMIZER_ITERATIVE
 * - Supports current Settings headers:
 *   Stat, Category, Active, Weight, Direction, Suggested Weight
 *
 * Run manually:
 *   runIterativeOptimizer()
 ***************************************/

const ITER_OPT = {
  SETTINGS_SHEET: "Settings",
  HISTORY_SHEET: "HISTORY",
  OUTPUT_SHEET: "OPTIMIZER_ITERATIVE",

  DEFAULT_MIN_WEIGHT: 0,
  DEFAULT_MAX_WEIGHT: 20,

  MAX_STAGE1_FEATURES_FOR_COMBOS: 8,
  MAX_PAIRS_TO_KEEP: 8,
  MIN_GAMES: 50,
  MIN_WIN_RATE: 0.52,
  MIN_ROI_IMPROVEMENT: 0.005,

  HEADERS: {
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

  HISTORY_HEADERS: {
    awayTeam: "Away Team",
    homeTeam: "Home Team",
    awayML: "Away Moneyline",
    homeML: "Home Moneyline",
    winner: "Winner"
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
  const history = readHistory_(historySheet);

  if (history.rows.length < ITER_OPT.MIN_GAMES) {
    throw new Error("Not enough HISTORY rows for optimizer testing. Found " + history.rows.length + ". Minimum required: " + ITER_OPT.MIN_GAMES);
  }

  const baselineWeights = {};
  settings.features.forEach(feature => {
    baselineWeights[feature.name] = feature.currentWeight;
  });

  const baseline = backtestWeights_(history, settings, baselineWeights);
  const results = [];

  results.push(makeResultRow_(
    "BASELINE",
    "Current live weights",
    baselineWeights,
    baseline,
    baseline
  ));

  const stage1 = [];

  settings.features
    .filter(feature => feature.active && feature.optimize)
    .forEach(feature => {
      let best = null;

      for (let weight = feature.minWeight; weight <= feature.maxWeight; weight++) {
        const testWeights = Object.assign({}, baselineWeights);
        testWeights[feature.name] = weight;

        const perf = backtestWeights_(history, settings, testWeights);

        if (!best || isBetterResult_(perf, best.perf, baseline)) {
          best = {
            feature,
            weight,
            perf,
            weights: testWeights
          };
        }
      }

      if (best) {
        stage1.push(best);
        results.push(makeResultRow_(
          "STAGE 1 BEST",
          best.feature.name + " best = " + best.weight,
          best.weights,
          best.perf,
          baseline
        ));
      }
    });

  stage1.sort((a, b) => b.perf.roi - a.perf.roi);

  const topSingles = stage1.slice(0, ITER_OPT.MAX_STAGE1_FEATURES_FOR_COMBOS);
  const stage2 = [];

  for (let i = 0; i < topSingles.length; i++) {
    for (let j = i + 1; j < topSingles.length; j++) {
      const a = topSingles[i];
      const b = topSingles[j];

      const testWeights = Object.assign({}, baselineWeights);
      testWeights[a.feature.name] = a.weight;
      testWeights[b.feature.name] = b.weight;

      const perf = backtestWeights_(history, settings, testWeights);

      stage2.push({
        label: a.feature.name + " + " + b.feature.name,
        weights: testWeights,
        perf
      });

      results.push(makeResultRow_(
        "STAGE 2 - PAIR",
        a.feature.name + " = " + a.weight + ", " + b.feature.name + " = " + b.weight,
        testWeights,
        perf,
        baseline
      ));
    }
  }

  stage2.sort((a, b) => b.perf.roi - a.perf.roi);

  const candidates = stage1
    .map(item => ({
      label: item.feature.name,
      weights: item.weights,
      perf: item.perf
    }))
    .concat(stage2.slice(0, ITER_OPT.MAX_PAIRS_TO_KEEP));

  candidates.sort((a, b) => b.perf.roi - a.perf.roi);

  const accepted = candidates.find(candidate => passesAcceptanceRules_(candidate.perf, baseline));
  const finalCandidate = accepted || candidates[0];

  if (finalCandidate) {
    results.push(makeResultRow_(
      accepted ? "FINAL RECOMMENDATION" : "FINAL WATCHLIST",
      finalCandidate.label,
      finalCandidate.weights,
      finalCandidate.perf,
      baseline
    ));

    writeSuggestedWeights_(settingsSheet, settings, finalCandidate.weights, Boolean(accepted));
  }

  writeOptimizerResults_(ss, results, settings.features);
  SpreadsheetApp.flush();
}


function backtestWeights_(history, settings, weightsByFeature) {
  let games = 0;
  let wins = 0;
  let losses = 0;
  let profit = 0;
  let totalRisk = 0;

  history.rows.forEach(row => {
    const awayTeam = row[history.headers[ITER_OPT.HISTORY_HEADERS.awayTeam]];
    const homeTeam = row[history.headers[ITER_OPT.HISTORY_HEADERS.homeTeam]];
    const winner = row[history.headers[ITER_OPT.HISTORY_HEADERS.winner]];

    if (!awayTeam || !homeTeam || !winner) return;

    let awayScore = 0;
    let homeScore = 0;

    settings.features.forEach(feature => {
      const weight = Number(weightsByFeature[feature.name] || 0);
      if (weight === 0) return;

      const awayHeader = "Away " + feature.name;
      const homeHeader = "Home " + feature.name;

      const awayIdx = history.headers[awayHeader];
      const homeIdx = history.headers[homeHeader];

      if (awayIdx === undefined || homeIdx === undefined) return;

      const awayVal = Number(row[awayIdx]);
      const homeVal = Number(row[homeIdx]);

      if (!isFinite(awayVal) || !isFinite(homeVal)) return;

      awayScore += awayVal * weight * feature.direction;
      homeScore += homeVal * weight * feature.direction;
    });

    if (awayScore === homeScore) return;

    const pick = awayScore > homeScore ? awayTeam : homeTeam;
    const pickedAway = pick === awayTeam;

    const moneylineHeader = pickedAway
      ? ITER_OPT.HISTORY_HEADERS.awayML
      : ITER_OPT.HISTORY_HEADERS.homeML;

    const moneyline = Number(row[history.headers[moneylineHeader]]);
    if (!isFinite(moneyline)) return;

    games++;
    totalRisk += 100;

    if (String(pick).trim() === String(winner).trim()) {
      wins++;
      profit += americanOddsProfit_(moneyline, 100);
    } else {
      losses++;
      profit -= 100;
    }
  });

  return {
    games,
    wins,
    losses,
    winRate: games ? wins / games : 0,
    profit,
    roi: totalRisk ? profit / totalRisk : 0
  };
}


function americanOddsProfit_(odds, risk) {
  if (odds > 0) return risk * (odds / 100);
  return risk * (100 / Math.abs(odds));
}


function passesAcceptanceRules_(perf, baseline) {
  if (perf.games < ITER_OPT.MIN_GAMES) return false;
  if (perf.winRate < ITER_OPT.MIN_WIN_RATE) return false;
  if ((perf.roi - baseline.roi) < ITER_OPT.MIN_ROI_IMPROVEMENT) return false;
  if (perf.profit <= baseline.profit) return false;
  return true;
}


function isBetterResult_(test, currentBest, baseline) {
  if (!currentBest) return true;

  const testPass = passesAcceptanceRules_(test, baseline);
  const bestPass = passesAcceptanceRules_(currentBest, baseline);

  if (testPass && !bestPass) return true;
  if (!testPass && bestPass) return false;

  return test.roi > currentBest.roi;
}


function readOptimizerSettings_(sheet) {
  const values = sheet.getDataRange().getValues();
  const headerInfo = findSettingsHeaderRow_(values);
  const headers = headerInfo.headers;
  const headerRowIndex = headerInfo.rowIndex;

  const features = [];

  for (let rowIndex = headerRowIndex + 1; rowIndex < values.length; rowIndex++) {
    const row = values[rowIndex];
    const name = row[headers[ITER_OPT.HEADERS.stat]];
    if (!name) continue;

    const active = parseBool_(row[headers[ITER_OPT.HEADERS.active]]);
    const optimize = headers[ITER_OPT.HEADERS.optimize] === undefined
      ? active
      : parseBool_(row[headers[ITER_OPT.HEADERS.optimize]]);

    const minWeight = headers[ITER_OPT.HEADERS.minWeight] === undefined
      ? ITER_OPT.DEFAULT_MIN_WEIGHT
      : Number(row[headers[ITER_OPT.HEADERS.minWeight]] || ITER_OPT.DEFAULT_MIN_WEIGHT);

    const maxWeight = headers[ITER_OPT.HEADERS.maxWeight] === undefined
      ? ITER_OPT.DEFAULT_MAX_WEIGHT
      : Number(row[headers[ITER_OPT.HEADERS.maxWeight]] || ITER_OPT.DEFAULT_MAX_WEIGHT);

    features.push({
      rowNumber: rowIndex + 1,
      name: String(name).trim(),
      category: headers[ITER_OPT.HEADERS.category] === undefined ? "" : String(row[headers[ITER_OPT.HEADERS.category]] || "").trim(),
      active,
      optimize,
      currentWeight: Number(row[headers[ITER_OPT.HEADERS.weight]] || 0),
      minWeight,
      maxWeight,
      direction: Number(row[headers[ITER_OPT.HEADERS.direction]] || 1)
    });
  }

  return {
    headers,
    headerRowNumber: headerRowIndex + 1,
    features
  };
}


function readHistory_(sheet) {
  const values = sheet.getDataRange().getValues();
  const headerInfo = findHistoryHeaderRow_(values);

  return {
    headers: headerInfo.headers,
    headerRowNumber: headerInfo.rowIndex + 1,
    rows: values.slice(headerInfo.rowIndex + 1)
  };
}


function ensureSuggestedWeightColumn_(sheet) {
  const values = sheet.getDataRange().getValues();
  const headerInfo = findSettingsHeaderRow_(values);
  const headerRowNumber = headerInfo.rowIndex + 1;
  const headers = headerInfo.headers;

  if (headers[ITER_OPT.HEADERS.suggestedWeight] === undefined) {
    sheet.getRange(headerRowNumber, sheet.getLastColumn() + 1).setValue(ITER_OPT.HEADERS.suggestedWeight);
  }
}


function writeSuggestedWeights_(sheet, settings, weightsByFeature, accepted) {
  const suggestedCol = settings.headers[ITER_OPT.HEADERS.suggestedWeight] + 1;

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
      result.games,
      result.wins,
      result.losses,
      result.winRate,
      result.profit,
      result.roi,
      result.roiChange,
      result.decision
    ];

    features.forEach(feature => {
      row.push(result.weights[feature.name] || 0);
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
  const roiChange = perf.roi - baseline.roi;

  let decision = "REJECT";

  if (stage === "BASELINE") {
    decision = "CONTROL";
  } else if (passesAcceptanceRules_(perf, baseline)) {
    decision = "ACCEPT";
  } else if (perf.profit > baseline.profit || perf.roi > baseline.roi) {
    decision = "WATCHLIST";
  }

  return {
    stage,
    test,
    games: perf.games,
    wins: perf.wins,
    losses: perf.losses,
    winRate: perf.winRate,
    profit: perf.profit,
    roi: perf.roi,
    roiChange,
    decision,
    weights
  };
}


function findSettingsHeaderRow_(values) {
  const required = [
    ITER_OPT.HEADERS.stat,
    ITER_OPT.HEADERS.active,
    ITER_OPT.HEADERS.weight,
    ITER_OPT.HEADERS.direction
  ];

  return findHeaderRow_(values, required, "Settings");
}


function findHistoryHeaderRow_(values) {
  const required = [
    ITER_OPT.HISTORY_HEADERS.awayTeam,
    ITER_OPT.HISTORY_HEADERS.homeTeam,
    ITER_OPT.HISTORY_HEADERS.awayML,
    ITER_OPT.HISTORY_HEADERS.homeML,
    ITER_OPT.HISTORY_HEADERS.winner
  ];

  return findHeaderRow_(values, required, "HISTORY");
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


function mapHeaders_(headerRow) {
  const map = {};

  headerRow.forEach((header, index) => {
    if (header !== "" && header !== null && header !== undefined) {
      map[String(header).trim()] = index;
    }
  });

  return map;
}


function parseBool_(value) {
  if (value === true) return true;
  const normalized = String(value).trim().toUpperCase();
  return normalized === "TRUE" || normalized === "YES" || normalized === "1";
}
