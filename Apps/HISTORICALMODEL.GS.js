function runHistoricalModelInputTest() {
  buildHistoricalModelInputForDate("2026-06-10");
}

function runHistoricalBacktestSmallBatch() {
  buildHistoricalBacktest("2026-06-01", "2026-06-10");
  buildHistoricalBacktestSummary();
}

function buildHistoricalModelInputForDate(targetDate) {
  const outputSheet = getOrCreateSheet("Historical_Model_Input");
  outputSheet.clearContents();

  const historicalGames = getSheetRows("RAW_Historical_Games");

  const gamesForDate = historicalGames.filter(row => {
    return normalizeHistoricalDate(row["Game Date"]) === targetDate;
  });

  const teamStatsBeforeDate = buildTeamStatsBeforeDate(historicalGames, targetDate);
  const recentFormBeforeDate = buildRecentFormBeforeDate(historicalGames, targetDate, 7);
  const homeAwayStatsBeforeDate = buildHomeAwayStatsBeforeDate(historicalGames, targetDate);

  const headers = getHistoricalModelHeaders();
  const output = [headers];

  gamesForDate.forEach(game => {
    output.push(buildHistoricalModelRow(
      game,
      teamStatsBeforeDate,
      recentFormBeforeDate,
      homeAwayStatsBeforeDate
    ));
  });

  if (output.length === 1) {
    output.push(buildNoHistoricalGamesFoundRow(targetDate, headers.length));
  }

  outputSheet
    .getRange(1, 1, output.length, headers.length)
    .setValues(output);

  formatHistoricalModelInput(outputSheet);
}

function buildHistoricalBacktest(startDate, endDate) {
  const outputSheet = getOrCreateSheet("Historical_Backtest");
  outputSheet.clearContents();

  const historicalGames = getSheetRows("RAW_Historical_Games");

  const gamesInRange = historicalGames.filter(row => {
    const gameDate = normalizeHistoricalDate(row["Game Date"]);
    return gameDate >= startDate && gameDate <= endDate;
  });

  const headers = [
    "Game Date",
    "Game ID",
    "Away Team",
    "Home Team",
    "Winner",
    "Historical Away Score",
    "Historical Home Score",
    "Historical Pick",
    "Historical Confidence",
    "Historical Correct?"
  ];

  const output = [headers];

  gamesInRange.forEach(game => {
    const gameDate = normalizeHistoricalDate(game["Game Date"]);

    const teamStatsBeforeDate = buildTeamStatsBeforeDate(historicalGames, gameDate);
    const recentFormBeforeDate = buildRecentFormBeforeDate(historicalGames, gameDate, 7);
    const homeAwayStatsBeforeDate = buildHomeAwayStatsBeforeDate(historicalGames, gameDate);

    const fullRow = buildHistoricalModelRow(
      game,
      teamStatsBeforeDate,
      recentFormBeforeDate,
      homeAwayStatsBeforeDate
    );

    const fullHeaders = getHistoricalModelHeaders();
    const rowObj = rowArrayToObjectHistorical(fullHeaders, fullRow);

    output.push([
      rowObj["Game Date"],
      rowObj["Game ID"],
      rowObj["Away Team"],
      rowObj["Home Team"],
      rowObj["Winner"],
      rowObj["Historical Away Score"],
      rowObj["Historical Home Score"],
      rowObj["Historical Pick"],
      rowObj["Historical Confidence"],
      rowObj["Historical Correct?"]
    ]);
  });

  outputSheet
    .getRange(1, 1, output.length, headers.length)
    .setValues(output);

  formatHistoricalBacktest(outputSheet);
}

function buildHistoricalBacktestSummary() {
  const sourceRows = getSheetRows("Historical_Backtest");
  const sheet = getOrCreateSheet("Historical_Backtest_Summary");
  sheet.clearContents();

  const overall = calculateBacktestSummary(sourceRows);

  const output = [
    ["Historical Backtest Summary", ""],
    ["", ""],
    ["Overall Results", ""],
    ["Total Games", overall.games],
    ["Correct Picks", overall.wins],
    ["Incorrect Picks", overall.losses],
    ["Win %", overall.winRate],
    ["", ""],
    ["Confidence Buckets", ""],
    ["Bucket", "Games", "Correct", "Incorrect", "Win %"]
  ];

  const buckets = [
    { name: "0-20", min: 0, max: 20 },
    { name: "21-40", min: 21, max: 40 },
    { name: "41-60", min: 41, max: 60 },
    { name: "61-80", min: 61, max: 80 },
    { name: "81-100", min: 81, max: 100 }
  ];

  buckets.forEach(bucket => {
    const rows = sourceRows.filter(row => {
      const confidence = Number(row["Historical Confidence"]);
      return !isNaN(confidence) &&
        confidence >= bucket.min &&
        confidence <= bucket.max;
    });

    const summary = calculateBacktestSummary(rows);

    output.push([
      bucket.name,
      summary.games,
      summary.wins,
      summary.losses,
      summary.winRate
    ]);
  });

  sheet
    .getRange(1, 1, output.length, 5)
    .setValues(padRows(output, 5));

  formatHistoricalBacktestSummary(sheet);
}

function getHistoricalModelHeaders() {
  return [
    "Game Date",
    "Game ID",
    "Away Team",
    "Home Team",
    "Away Score",
    "Home Score",
    "Winner",

    "Away SP",
    "Home SP",
    "Away SP ID",
    "Home SP ID",
    "Venue",

    "Away Games Before Date",
    "Home Games Before Date",
    "Away Win % Before Date",
    "Home Win % Before Date",
    "Win % Edge Before Date",

    "Away Runs/Game Before Date",
    "Home Runs/Game Before Date",
    "Runs/Game Edge Before Date",

    "Away Runs Allowed/Game Before Date",
    "Home Runs Allowed/Game Before Date",
    "Runs Allowed/Game Edge Before Date",

    "Away Run Differential Before Date",
    "Home Run Differential Before Date",
    "Run Differential Edge Before Date",

    "Away Road Games Before Date",
    "Home Home Games Before Date",
    "Away Road Win % Before Date",
    "Home Home Win % Before Date",
    "Home/Away Win % Edge Before Date",

    "Away Runs/Game Last 7 Before Date",
    "Home Runs/Game Last 7 Before Date",
    "Recent Runs/Game Edge",

    "Away Runs Allowed/Game Last 7 Before Date",
    "Home Runs Allowed/Game Last 7 Before Date",
    "Recent Runs Allowed/Game Edge",

    "Historical Away Score",
    "Historical Home Score",
    "Historical Pick",
    "Historical Confidence",
    "Historical Correct?",

    "Model Reconstruction Status",
    "Notes"
  ];
}

function buildHistoricalModelRow(
  game,
  teamStatsBeforeDate,
  recentFormBeforeDate,
  homeAwayStatsBeforeDate
) {
  const awayTeam = game["Away Team"];
  const homeTeam = game["Home Team"];

  const awaySeason = teamStatsBeforeDate[awayTeam] || blankTeamStats();
  const homeSeason = teamStatsBeforeDate[homeTeam] || blankTeamStats();

  const awayRecent = recentFormBeforeDate[awayTeam] || blankTeamStats();
  const homeRecent = recentFormBeforeDate[homeTeam] || blankTeamStats();

  const awayRoad = homeAwayStatsBeforeDate[awayTeam]?.road || blankTeamStats();
  const homeHome = homeAwayStatsBeforeDate[homeTeam]?.home || blankTeamStats();

  const awayWinPct = calculateWinPercentageHistorical(awaySeason.wins, awaySeason.games);
  const homeWinPct = calculateWinPercentageHistorical(homeSeason.wins, homeSeason.games);
  const winPctEdge = edgeHistorical(awayWinPct, homeWinPct, "higher");

  const awayRunsPerGame = calculateRunsPerGameHistorical(awaySeason.runsScored, awaySeason.games);
  const homeRunsPerGame = calculateRunsPerGameHistorical(homeSeason.runsScored, homeSeason.games);
  const runsGameEdge = edgeHistorical(awayRunsPerGame, homeRunsPerGame, "higher");

  const awayRunsAllowedPerGame = calculateRunsPerGameHistorical(awaySeason.runsAllowed, awaySeason.games);
  const homeRunsAllowedPerGame = calculateRunsPerGameHistorical(homeSeason.runsAllowed, homeSeason.games);
  const runsAllowedEdge = edgeHistorical(awayRunsAllowedPerGame, homeRunsAllowedPerGame, "lower");

  const awayRunDiff = awaySeason.runsScored - awaySeason.runsAllowed;
  const homeRunDiff = homeSeason.runsScored - homeSeason.runsAllowed;
  const runDiffEdge = edgeHistorical(awayRunDiff, homeRunDiff, "higher");

  const awayRoadWinPct = calculateWinPercentageHistorical(awayRoad.wins, awayRoad.games);
  const homeHomeWinPct = calculateWinPercentageHistorical(homeHome.wins, homeHome.games);
  const homeAwayWinPctEdge = edgeHistorical(awayRoadWinPct, homeHomeWinPct, "higher");

  const awayRecentRunsPerGame = calculateRunsPerGameHistorical(awayRecent.runsScored, awayRecent.games);
  const homeRecentRunsPerGame = calculateRunsPerGameHistorical(homeRecent.runsScored, homeRecent.games);
  const recentRunsGameEdge = edgeHistorical(awayRecentRunsPerGame, homeRecentRunsPerGame, "higher");

  const awayRecentRunsAllowedPerGame = calculateRunsPerGameHistorical(awayRecent.runsAllowed, awayRecent.games);
  const homeRecentRunsAllowedPerGame = calculateRunsPerGameHistorical(homeRecent.runsAllowed, homeRecent.games);
  const recentRunsAllowedEdge = edgeHistorical(awayRecentRunsAllowedPerGame, homeRecentRunsAllowedPerGame, "lower");

  const historicalScore = calculateSimpleHistoricalScore({
    winPctEdge,
    runsGameEdge,
    runsAllowedEdge,
    runDiffEdge,
    homeAwayWinPctEdge,
    recentRunsGameEdge,
    recentRunsAllowedEdge
  });

  const historicalPick =
    historicalScore.awayScore > historicalScore.homeScore ? awayTeam :
    historicalScore.homeScore > historicalScore.awayScore ? homeTeam :
    "Coin Flip";

  const historicalConfidence = calculateHistoricalConfidence(
    historicalScore.awayScore,
    historicalScore.homeScore
  );

  const winner = game["Winner"];
  const historicalCorrect =
    historicalPick === "Coin Flip" || !winner ? "" : historicalPick === winner;

  return [
    normalizeHistoricalDate(game["Game Date"]),
    game["Game ID"],
    awayTeam,
    homeTeam,
    game["Away Score"],
    game["Home Score"],
    winner,

    game["Away SP"],
    game["Home SP"],
    game["Away SP ID"],
    game["Home SP ID"],
    game["Venue"],

    awaySeason.games,
    homeSeason.games,
    awayWinPct,
    homeWinPct,
    winPctEdge,

    awayRunsPerGame,
    homeRunsPerGame,
    runsGameEdge,

    awayRunsAllowedPerGame,
    homeRunsAllowedPerGame,
    runsAllowedEdge,

    awayRunDiff,
    homeRunDiff,
    runDiffEdge,

    awayRoad.games,
    homeHome.games,
    awayRoadWinPct,
    homeHomeWinPct,
    homeAwayWinPctEdge,

    awayRecentRunsPerGame,
    homeRecentRunsPerGame,
    recentRunsGameEdge,

    awayRecentRunsAllowedPerGame,
    homeRecentRunsAllowedPerGame,
    recentRunsAllowedEdge,

    historicalScore.awayScore,
    historicalScore.homeScore,
    historicalPick,
    historicalConfidence,
    historicalCorrect,

    "Stage 5 Complete",
    "Historical reconstruction, scoring, batch backtest, and summary enabled"
  ];
}

function calculateSimpleHistoricalScore(edges) {
  const weights = {
    winPctEdge: 3,
    runsGameEdge: 2,
    runsAllowedEdge: 2,
    runDiffEdge: 3,
    homeAwayWinPctEdge: 2,
    recentRunsGameEdge: 2,
    recentRunsAllowedEdge: 2
  };

  let score = 0;
  let totalWeight = 0;

  Object.keys(weights).forEach(key => {
    const value = parseHistoricalNumber(edges[key]);
    if (value === "") return;

    score += value * weights[key];
    totalWeight += weights[key];
  });

  const finalScore = totalWeight > 0 ? score / totalWeight : 0;

  return {
    awayScore: roundHistoricalNumber(finalScore, 4),
    homeScore: roundHistoricalNumber(finalScore * -1, 4)
  };
}

function calculateHistoricalConfidence(awayScore, homeScore) {
  const away = parseHistoricalNumber(awayScore);
  const home = parseHistoricalNumber(homeScore);

  if (away === "" || home === "") return "";

  const scoreGap = Math.abs(away - home);

  return Math.min(100, Math.round(scoreGap * 5));
}

function calculateBacktestSummary(rows) {
  let wins = 0;
  let losses = 0;

  rows.forEach(row => {
    const correct = normalizeHistoricalBoolean(row["Historical Correct?"]);

    if (correct === true) wins++;
    if (correct === false) losses++;
  });

  const games = wins + losses;

  return {
    games,
    wins,
    losses,
    winRate: games > 0 ? Math.round((wins / games) * 1000) / 10 + "%" : ""
  };
}

function buildTeamStatsBeforeDate(historicalGames, targetDate) {
  const statsByTeam = {};

  historicalGames.forEach(game => {
    const gameDate = normalizeHistoricalDate(game["Game Date"]);

    if (!gameDate || gameDate >= targetDate) return;

    processHistoricalGameForTeamStats(statsByTeam, game);
  });

  return statsByTeam;
}

function buildRecentFormBeforeDate(historicalGames, targetDate, daysBack) {
  const statsByTeam = {};

  const target = new Date(targetDate + "T00:00:00");
  const start = new Date(target);
  start.setDate(target.getDate() - daysBack);

  const startDate = Utilities.formatDate(start, "America/New_York", "yyyy-MM-dd");

  historicalGames.forEach(game => {
    const gameDate = normalizeHistoricalDate(game["Game Date"]);

    if (!gameDate) return;
    if (gameDate < startDate) return;
    if (gameDate >= targetDate) return;

    processHistoricalGameForTeamStats(statsByTeam, game);
  });

  return statsByTeam;
}

function buildHomeAwayStatsBeforeDate(historicalGames, targetDate) {
  const statsByTeam = {};

  historicalGames.forEach(game => {
    const gameDate = normalizeHistoricalDate(game["Game Date"]);

    if (!gameDate || gameDate >= targetDate) return;

    const awayTeam = game["Away Team"];
    const homeTeam = game["Home Team"];

    const awayScore = Number(game["Away Score"]);
    const homeScore = Number(game["Home Score"]);

    if (!awayTeam || !homeTeam) return;
    if (isNaN(awayScore) || isNaN(homeScore)) return;

    if (!statsByTeam[awayTeam]) statsByTeam[awayTeam] = blankHomeAwayStats();
    if (!statsByTeam[homeTeam]) statsByTeam[homeTeam] = blankHomeAwayStats();

    statsByTeam[awayTeam].road.games++;
    statsByTeam[awayTeam].road.runsScored += awayScore;
    statsByTeam[awayTeam].road.runsAllowed += homeScore;

    statsByTeam[homeTeam].home.games++;
    statsByTeam[homeTeam].home.runsScored += homeScore;
    statsByTeam[homeTeam].home.runsAllowed += awayScore;

    if (awayScore > homeScore) {
      statsByTeam[awayTeam].road.wins++;
      statsByTeam[homeTeam].home.losses++;
    } else if (homeScore > awayScore) {
      statsByTeam[homeTeam].home.wins++;
      statsByTeam[awayTeam].road.losses++;
    }
  });

  return statsByTeam;
}

function processHistoricalGameForTeamStats(statsByTeam, game) {
  const awayTeam = game["Away Team"];
  const homeTeam = game["Home Team"];
  const awayScore = Number(game["Away Score"]);
  const homeScore = Number(game["Home Score"]);

  if (!awayTeam || !homeTeam) return;
  if (isNaN(awayScore) || isNaN(homeScore)) return;

  if (!statsByTeam[awayTeam]) statsByTeam[awayTeam] = blankTeamStats();
  if (!statsByTeam[homeTeam]) statsByTeam[homeTeam] = blankTeamStats();

  statsByTeam[awayTeam].games++;
  statsByTeam[awayTeam].runsScored += awayScore;
  statsByTeam[awayTeam].runsAllowed += homeScore;

  statsByTeam[homeTeam].games++;
  statsByTeam[homeTeam].runsScored += homeScore;
  statsByTeam[homeTeam].runsAllowed += awayScore;

  if (awayScore > homeScore) {
    statsByTeam[awayTeam].wins++;
    statsByTeam[homeTeam].losses++;
  } else if (homeScore > awayScore) {
    statsByTeam[homeTeam].wins++;
    statsByTeam[awayTeam].losses++;
  }
}

function blankTeamStats() {
  return {
    games: 0,
    wins: 0,
    losses: 0,
    runsScored: 0,
    runsAllowed: 0
  };
}

function blankHomeAwayStats() {
  return {
    home: blankTeamStats(),
    road: blankTeamStats()
  };
}

function calculateRunsPerGameHistorical(runs, games) {
  if (!games || games === 0) return "";
  return roundHistoricalNumber(runs / games, 3);
}

function calculateWinPercentageHistorical(wins, games) {
  if (!games || games === 0) return "";
  return roundHistoricalNumber(wins / games, 3);
}

function edgeHistorical(awayValue, homeValue, direction) {
  const away = parseHistoricalNumber(awayValue);
  const home = parseHistoricalNumber(homeValue);

  if (away === "" || home === "") return "";

  if (direction === "lower") {
    return roundHistoricalNumber(home - away, 4);
  }

  return roundHistoricalNumber(away - home, 4);
}

function parseHistoricalNumber(value) {
  if (value === "" || value === undefined || value === null) return "";

  const num = Number(value);
  return isNaN(num) ? "" : num;
}

function roundHistoricalNumber(value, decimals) {
  if (value === "" || isNaN(value)) return "";
  return Number(value.toFixed(decimals));
}

function normalizeHistoricalDate(value) {
  if (!value) return "";

  if (Object.prototype.toString.call(value) === "[object Date]") {
    return Utilities.formatDate(value, "America/New_York", "yyyy-MM-dd");
  }

  return value.toString().substring(0, 10);
}

function normalizeHistoricalBoolean(value) {
  if (value === true || value === "TRUE" || value === "Yes") return true;
  if (value === false || value === "FALSE" || value === "No") return false;
  return null;
}

function rowArrayToObjectHistorical(headers, row) {
  const obj = {};

  headers.forEach((header, index) => {
    obj[header] = row[index];
  });

  return obj;
}

function buildNoHistoricalGamesFoundRow(targetDate, width) {
  const row = [
    targetDate,
    "",
    "",
    "",
    "",
    "",
    "",
    "",
    "",
    "",
    "",
    "",
    "No Games Found",
    "No completed historical games found for this date"
  ];

  while (row.length < width) {
    row.splice(row.length - 2, 0, "");
  }

  return row;
}

function formatHistoricalModelInput(sheet) {
  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();

  if (lastRow < 1) return;

  sheet.getRange(1, 1, 1, lastCol)
    .setFontWeight("bold")
    .setBackground("#1f1f1f")
    .setFontColor("#ffffff");

  sheet.getRange(1, 1, lastRow, lastCol)
    .setWrap(true)
    .setVerticalAlignment("middle");

  sheet.setFrozenRows(1);
  sheet.autoResizeColumns(1, lastCol);
}

function formatHistoricalBacktest(sheet) {
  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();

  if (lastRow < 1) return;

  sheet.getRange(1, 1, 1, lastCol)
    .setFontWeight("bold")
    .setBackground("#1f1f1f")
    .setFontColor("#ffffff");

  sheet.getRange(1, 1, lastRow, lastCol)
    .setWrap(true)
    .setVerticalAlignment("middle");

  sheet.setFrozenRows(1);
  sheet.autoResizeColumns(1, lastCol);
}

function formatHistoricalBacktestSummary(sheet) {
  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();

  if (lastRow < 1) return;

  sheet.getRange(1, 1, 1, lastCol)
    .setFontWeight("bold")
    .setFontSize(14)
    .setBackground("#1f1f1f")
    .setFontColor("#ffffff");

  sheet.getRange(3, 1, 1, lastCol).setFontWeight("bold");
  sheet.getRange(9, 1, 1, lastCol).setFontWeight("bold");

  sheet.getRange(10, 1, 1, lastCol)
    .setFontWeight("bold")
    .setBackground("#1f1f1f")
    .setFontColor("#ffffff");

  sheet.getRange(1, 1, lastRow, lastCol)
    .setWrap(true)
    .setVerticalAlignment("middle");

  sheet.autoResizeColumns(1, lastCol);
}

function runHistoricalBacktestJune1To7() {
  buildHistoricalBacktest("2026-06-01", "2026-06-07");
  buildHistoricalBacktestSummary();
}

function runHistoricalBacktestJune8To14() {
  buildHistoricalBacktest("2026-06-08", "2026-06-14");
  buildHistoricalBacktestSummary();
}

function runHistoricalBacktestJune15To22() {
  buildHistoricalBacktest("2026-06-15", "2026-06-22");
  buildHistoricalBacktestSummary();
}