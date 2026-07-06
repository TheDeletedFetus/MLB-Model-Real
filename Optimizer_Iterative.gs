/***************************************
 * ITERATIVE MODEL OPTIMIZER v0.1.4
 *
 * Safe optimizer:
 * - Does NOT overwrite live Weight
 * - Writes Suggested Weight only
 * - Writes full results to OPTIMIZER_ITERATIVE
 * - Stage 1 only: tests each active stat individually
 *
 * Run manually:
 *   runIterativeOptimizer()
 ***************************************/

const ITER_OPT = {
  SETTINGS_SHEET: "Settings",
  HISTORY_SHEET: "HISTORY",
  OUTPUT_SHEET: "OPTIMIZER_ITERATIVE",

  DEFAULT_MIN_WEIGHT: 0,
  DEFAULT_MAX_WEIGHT: 10,

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

  HISTORY_ALIASES: {
    awayTeam: ["Away Team", "Away", "AwayTeam", "Visitor", "Visitor Team", "Road Team"],
    homeTeam: ["Home Team", "Home", "HomeTeam"],
    awayML: ["Away Moneyline", "Away ML", "Away Odds", "Away Money Line", "Away Line", "Away_Moneyline"],
    homeML: ["Home Moneyline", "Home ML", "Home Odds", "Home Money Line", "Home Line", "Home_Moneyline"],
    winner: ["Winner", "Game Winner", "Actual Winner", "Winning Team", "Final Winner"],
    awayFinal: ["Away Final", "Away Score", "Away Runs", "Away Final Score"],
    homeFinal: ["Home Final", "Home Score", "Home Runs", "Home Final Score"]
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

        results.push(makeResultRow_(
          "STAGE 1 TEST",
          feature.name + " = " + weight,
          testWeights,
          perf,
          baseline
        ));

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

  const accepted = stage1.find(candidate => passesAcceptanceRules_(candidate.perf, baseline));
  const finalCandidate = accepted || stage1[0];

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
  SpreadsheetApp.flush();
}


function backtestWeights_(history, settings, weightsByFeature) {
  let games = 0;
  let wins = 0;
  let losses = 0;
  let profit = 0;
  let totalRisk = 0;

  history.rows.forEach(row => {
    const awayTeam = row[history.indexes.awayTeam];
    const homeTeam = row[history.indexes.homeTeam];
    const winner = resolveWinner_(row, history.indexes, awayTeam, homeTeam);

    if (!awayTeam || !homeTeam || !winner) return;

    let awayScore = 0;
    let homeScore = 0;

    settings.features.forEach(feature => {
      const weight = Number(weightsByFeature[feature.name] || 0);
      if (weight === 0) return;

      const awayIdx = resolveFeatureHeaderIndex_(history.headers, "Away", feature.name);
      const homeIdx = resolveFeatureHeaderIndex_(history.headers, "Home", feature.name);

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
    const moneylineIndex = pickedAway ? history.indexes.awayML : history.indexes.homeML;
    const moneyline = Number(row[moneylineIndex]);

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


function resolveFeatureHeaderIndex_(headers, side, statName) {
  const candidates = [
    side + " " + statName,
    side + "_" + statName,
    side + statName,
    side + " Team " + statName,
    side + " " + statName + " Value"
  ];

  for (let i = 0; i < candidates.length; i++) {
    const idx = getHeaderIndex_(headers, candidates[i]);
    if (idx !== undefined) return idx;
  }

  return undefined;
}


function resolveWinner_(row, indexes, awayTeam, homeTeam) {
  if (indexes.winner !== undefined) {
    return row[indexes.winner];
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
  const indexes = headerInfo.indexes;

  if (indexes.winner === undefined && (indexes.awayFinal === undefined || indexes.homeFinal === undefined)) {
    throw new Error("HISTORY must contain either a winner column or both final score columns. Found headers: " + headerInfo.foundHeaders.join(" | "));
  }

  return {
    headers: headerInfo.headers,
    indexes,
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
