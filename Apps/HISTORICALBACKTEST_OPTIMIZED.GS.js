function runOptimizedHistoricalBacktest2026ToDate() {
  runOptimizedHistoricalBacktest("2026-03-26", getTodayText());
}

function runOptimizedHistoricalBacktestJuneToDate() {
  runOptimizedHistoricalBacktest("2026-06-01", getTodayText());
}

function runOptimizedHistoricalBacktest(startDate, endDate) {
  const outputSheet = getOrCreateSheet("Historical_Backtest_Optimized");
  outputSheet.clearContents();

  const historicalGames = getSheetRows("RAW_Historical_Games")
    .filter(row => {
      const gameDate = normalizeOptimizedDate(row["Game Date"]);
      return gameDate >= startDate && gameDate <= endDate;
    })
    .sort((a, b) => {
      const dateA = normalizeOptimizedDate(a["Game Date"]);
      const dateB = normalizeOptimizedDate(b["Game Date"]);
      return dateA.localeCompare(dateB);
    });

  const gamesByDate = groupHistoricalGamesByDate(historicalGames);

  const cumulativeStats = {};
  const recentGameLogs = {};

  const headers = [
    "Game Date",
    "Game ID",
    "Away Team",
    "Home Team",
    "Winner",

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
    "Historical Correct?"
  ];

  const output = [headers];

  Object.keys(gamesByDate).sort().forEach(gameDate => {
    const gamesForDate = gamesByDate[gameDate];

    gamesForDate.forEach(game => {
      const row = buildOptimizedBacktestRow(
        game,
        gameDate,
        cumulativeStats,
        recentGameLogs
      );

      output.push(row);
    });

    gamesForDate.forEach(game => {
      updateOptimizedStatsAfterGame(cumulativeStats, recentGameLogs, game, gameDate);
    });
  });

  outputSheet
    .getRange(1, 1, output.length, headers.length)
    .setValues(output);

  formatOptimizedHistoricalBacktest(outputSheet);
  buildOptimizedHistoricalBacktestSummary();
}

function buildOptimizedBacktestRow(game, gameDate, cumulativeStats, recentGameLogs) {
  const awayTeam = game["Away Team"];
  const homeTeam = game["Home Team"];

  const awaySeason = cumulativeStats[awayTeam]?.overall || blankOptimizedStats();
  const homeSeason = cumulativeStats[homeTeam]?.overall || blankOptimizedStats();

  const awayRoad = cumulativeStats[awayTeam]?.road || blankOptimizedStats();
  const homeHome = cumulativeStats[homeTeam]?.home || blankOptimizedStats();

  const awayRecent = calculateOptimizedRecentStats(recentGameLogs[awayTeam] || [], gameDate, 7);
  const homeRecent = calculateOptimizedRecentStats(recentGameLogs[homeTeam] || [], gameDate, 7);

  const awayWinPct = calculateOptimizedWinPct(awaySeason.wins, awaySeason.games);
  const homeWinPct = calculateOptimizedWinPct(homeSeason.wins, homeSeason.games);
  const winPctEdge = optimizedEdge(awayWinPct, homeWinPct, "higher");

  const awayRunsPerGame = calculateOptimizedPerGame(awaySeason.runsScored, awaySeason.games);
  const homeRunsPerGame = calculateOptimizedPerGame(homeSeason.runsScored, homeSeason.games);
  const runsGameEdge = optimizedEdge(awayRunsPerGame, homeRunsPerGame, "higher");

  const awayRunsAllowedPerGame = calculateOptimizedPerGame(awaySeason.runsAllowed, awaySeason.games);
  const homeRunsAllowedPerGame = calculateOptimizedPerGame(homeSeason.runsAllowed, homeSeason.games);
  const runsAllowedEdge = optimizedEdge(awayRunsAllowedPerGame, homeRunsAllowedPerGame, "lower");

  const awayRunDiff = awaySeason.runsScored - awaySeason.runsAllowed;
  const homeRunDiff = homeSeason.runsScored - homeSeason.runsAllowed;
  const runDiffEdge = optimizedEdge(awayRunDiff, homeRunDiff, "higher");

  const awayRoadWinPct = calculateOptimizedWinPct(awayRoad.wins, awayRoad.games);
  const homeHomeWinPct = calculateOptimizedWinPct(homeHome.wins, homeHome.games);
  const homeAwayWinPctEdge = optimizedEdge(awayRoadWinPct, homeHomeWinPct, "higher");

  const awayRecentRunsPerGame = calculateOptimizedPerGame(awayRecent.runsScored, awayRecent.games);
  const homeRecentRunsPerGame = calculateOptimizedPerGame(homeRecent.runsScored, homeRecent.games);
  const recentRunsGameEdge = optimizedEdge(awayRecentRunsPerGame, homeRecentRunsPerGame, "higher");

  const awayRecentRunsAllowedPerGame = calculateOptimizedPerGame(awayRecent.runsAllowed, awayRecent.games);
  const homeRecentRunsAllowedPerGame = calculateOptimizedPerGame(homeRecent.runsAllowed, homeRecent.games);
  const recentRunsAllowedEdge = optimizedEdge(awayRecentRunsAllowedPerGame, homeRecentRunsAllowedPerGame, "lower");

  const score = calculateOptimizedHistoricalScore({
    winPctEdge,
    runsGameEdge,
    runsAllowedEdge,
    runDiffEdge,
    homeAwayWinPctEdge,
    recentRunsGameEdge,
    recentRunsAllowedEdge
  });

  const pick =
    score.awayScore > score.homeScore ? awayTeam :
    score.homeScore > score.awayScore ? homeTeam :
    "Coin Flip";

  const confidence = calculateOptimizedConfidence(score.awayScore, score.homeScore);

  const winner = game["Winner"];
  const correct =
    pick === "Coin Flip" || !winner ? "" : pick === winner;

  return [
    gameDate,
    game["Game ID"],
    awayTeam,
    homeTeam,
    winner,

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

    score.awayScore,
    score.homeScore,
    pick,
    confidence,
    correct
  ];
}

function updateOptimizedStatsAfterGame(cumulativeStats, recentGameLogs, game, gameDate) {
  const awayTeam = game["Away Team"];
  const homeTeam = game["Home Team"];
  const awayScore = Number(game["Away Score"]);
  const homeScore = Number(game["Home Score"]);

  if (!awayTeam || !homeTeam) return;
  if (isNaN(awayScore) || isNaN(homeScore)) return;

  ensureOptimizedTeam(cumulativeStats, awayTeam);
  ensureOptimizedTeam(cumulativeStats, homeTeam);

  updateOptimizedBucket(cumulativeStats[awayTeam].overall, awayScore, homeScore);
  updateOptimizedBucket(cumulativeStats[homeTeam].overall, homeScore, awayScore);

  updateOptimizedBucket(cumulativeStats[awayTeam].road, awayScore, homeScore);
  updateOptimizedBucket(cumulativeStats[homeTeam].home, homeScore, awayScore);

  if (awayScore > homeScore) {
    cumulativeStats[awayTeam].overall.wins++;
    cumulativeStats[awayTeam].road.wins++;

    cumulativeStats[homeTeam].overall.losses++;
    cumulativeStats[homeTeam].home.losses++;
  } else if (homeScore > awayScore) {
    cumulativeStats[homeTeam].overall.wins++;
    cumulativeStats[homeTeam].home.wins++;

    cumulativeStats[awayTeam].overall.losses++;
    cumulativeStats[awayTeam].road.losses++;
  }

  if (!recentGameLogs[awayTeam]) recentGameLogs[awayTeam] = [];
  if (!recentGameLogs[homeTeam]) recentGameLogs[homeTeam] = [];

  recentGameLogs[awayTeam].push({
    gameDate,
    runsScored: awayScore,
    runsAllowed: homeScore
  });

  recentGameLogs[homeTeam].push({
    gameDate,
    runsScored: homeScore,
    runsAllowed: awayScore
  });
}

function updateOptimizedBucket(bucket, runsScored, runsAllowed) {
  bucket.games++;
  bucket.runsScored += runsScored;
  bucket.runsAllowed += runsAllowed;
}

function ensureOptimizedTeam(cumulativeStats, team) {
  if (!cumulativeStats[team]) {
    cumulativeStats[team] = {
      overall: blankOptimizedStats(),
      home: blankOptimizedStats(),
      road: blankOptimizedStats()
    };
  }
}

function calculateOptimizedRecentStats(gameLogs, targetDate, daysBack) {
  const target = new Date(targetDate + "T00:00:00");
  const start = new Date(target);
  start.setDate(target.getDate() - daysBack);

  const startDate = Utilities.formatDate(start, "America/New_York", "yyyy-MM-dd");

  const stats = blankOptimizedStats();

  gameLogs.forEach(log => {
    if (log.gameDate < startDate) return;
    if (log.gameDate >= targetDate) return;

    stats.games++;
    stats.runsScored += log.runsScored;
    stats.runsAllowed += log.runsAllowed;
  });

  return stats;
}

function groupHistoricalGamesByDate(games) {
  const grouped = {};

  games.forEach(game => {
    const gameDate = normalizeOptimizedDate(game["Game Date"]);

    if (!grouped[gameDate]) grouped[gameDate] = [];
    grouped[gameDate].push(game);
  });

  return grouped;
}

function calculateOptimizedHistoricalScore(edges) {
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
    const value = parseOptimizedNumber(edges[key]);
    if (value === "") return;

    score += value * weights[key];
    totalWeight += weights[key];
  });

  const finalScore = totalWeight > 0 ? score / totalWeight : 0;

  return {
    awayScore: roundOptimizedNumber(finalScore, 4),
    homeScore: roundOptimizedNumber(finalScore * -1, 4)
  };
}

function calculateOptimizedConfidence(awayScore, homeScore) {
  const away = parseOptimizedNumber(awayScore);
  const home = parseOptimizedNumber(homeScore);

  if (away === "" || home === "") return "";

  const scoreGap = Math.abs(away - home);
  return Math.min(100, Math.round(scoreGap * 5));
}

function buildOptimizedHistoricalBacktestSummary() {
  const sourceRows = getSheetRows("Historical_Backtest_Optimized");
  const sheet = getOrCreateSheet("Historical_Backtest_Optimized_Summary");
  sheet.clearContents();

  const overall = calculateOptimizedSummary(sourceRows);

  const output = [
    ["Optimized Historical Backtest Summary", ""],
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

    const summary = calculateOptimizedSummary(rows);

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

  formatOptimizedHistoricalBacktestSummary(sheet);
}

function calculateOptimizedSummary(rows) {
  let wins = 0;
  let losses = 0;

  rows.forEach(row => {
    const correct = normalizeOptimizedBoolean(row["Historical Correct?"]);

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

function blankOptimizedStats() {
  return {
    games: 0,
    wins: 0,
    losses: 0,
    runsScored: 0,
    runsAllowed: 0
  };
}

function calculateOptimizedPerGame(value, games) {
  if (!games || games === 0) return "";
  return roundOptimizedNumber(value / games, 3);
}

function calculateOptimizedWinPct(wins, games) {
  if (!games || games === 0) return "";
  return roundOptimizedNumber(wins / games, 3);
}

function optimizedEdge(awayValue, homeValue, direction) {
  const away = parseOptimizedNumber(awayValue);
  const home = parseOptimizedNumber(homeValue);

  if (away === "" || home === "") return "";

  if (direction === "lower") {
    return roundOptimizedNumber(home - away, 4);
  }

  return roundOptimizedNumber(away - home, 4);
}

function parseOptimizedNumber(value) {
  if (value === "" || value === undefined || value === null) return "";

  const num = Number(value);
  return isNaN(num) ? "" : num;
}

function roundOptimizedNumber(value, decimals) {
  if (value === "" || isNaN(value)) return "";
  return Number(value.toFixed(decimals));
}

function normalizeOptimizedDate(value) {
  if (!value) return "";

  if (Object.prototype.toString.call(value) === "[object Date]") {
    return Utilities.formatDate(value, "America/New_York", "yyyy-MM-dd");
  }

  return value.toString().substring(0, 10);
}

function normalizeOptimizedBoolean(value) {
  if (value === true || value === "TRUE" || value === "Yes") return true;
  if (value === false || value === "FALSE" || value === "No") return false;
  return null;
}

function formatOptimizedHistoricalBacktest(sheet) {
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

function formatOptimizedHistoricalBacktestSummary(sheet) {
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