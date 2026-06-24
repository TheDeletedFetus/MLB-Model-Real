// HISTORICALFEATURES.GS
// v0.1.0b - Historical Feature Expansion

const ADVANCED_BOXSCORE_CACHE = {};

function runAdvancedHistoricalBacktest2026ToDate() {
  runAdvancedHistoricalBacktest("2026-03-26", getTodayText());
}

function runAdvancedHistoricalBacktest(startDate, endDate) {
  const outputSheet = getOrCreateSheet("Historical_Backtest_Advanced");
  outputSheet.clearContents();

  const historicalGames = getSheetRows("RAW_Historical_Games")
    .filter(row => {
      const gameDate = normalizeOptimizedDate(row["Game Date"]);
      return gameDate >= startDate && gameDate <= endDate;
    })
    .sort((a, b) => {
      return normalizeOptimizedDate(a["Game Date"])
        .localeCompare(normalizeOptimizedDate(b["Game Date"]));
    });

  const gamesByDate = groupHistoricalGamesByDate(historicalGames);

  const teamStats = {};
  const pitcherStats = {};
  const bullpenStats = {};
  const pitcherHandCache = {};

  const headers = [
    "Game Date",
    "Game ID",
    "Away Team",
    "Home Team",
    "Winner",

    "Away SP",
    "Home SP",
    "Away SP ID",
    "Home SP ID",
    "Away SP Hand",
    "Home SP Hand",

    "Away Starter ERA Before Date",
    "Home Starter ERA Before Date",
    "Starter ERA Edge",

    "Away Starter WHIP Before Date",
    "Home Starter WHIP Before Date",
    "Starter WHIP Edge",

    "Away Bullpen ERA Before Date",
    "Home Bullpen ERA Before Date",
    "Bullpen ERA Edge",

    "Away Bullpen WHIP Before Date",
    "Home Bullpen WHIP Before Date",
    "Bullpen WHIP Edge",

    "Away OPS vs SP Hand Before Date",
    "Home OPS vs SP Hand Before Date",
    "OPS vs SP Hand Edge",

    "Away Team OPS Before Date",
    "Home Team OPS Before Date",
    "Team OPS Edge",

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
      output.push(buildAdvancedHistoricalRow(
        game,
        gameDate,
        teamStats,
        pitcherStats,
        bullpenStats,
        pitcherHandCache
      ));
    });

    gamesForDate.forEach(game => {
      updateAdvancedStatsAfterGame(
        game,
        teamStats,
        pitcherStats,
        bullpenStats,
        pitcherHandCache
      );
    });
  });

  outputSheet
    .getRange(1, 1, output.length, headers.length)
    .setValues(output);

  formatAdvancedHistoricalBacktest(outputSheet);
  buildAdvancedHistoricalBacktestSummary();
}

function buildAdvancedHistoricalRow(
  game,
  gameDate,
  teamStats,
  pitcherStats,
  bullpenStats,
  pitcherHandCache
) {
  const awayTeam = game["Away Team"];
  const homeTeam = game["Home Team"];

  const starterInfo = getAdvancedGameStarterIds(game);

  const awaySpId = starterInfo.awaySpId;
  const homeSpId = starterInfo.homeSpId;

  const awaySpHand = getHistoricalPitcherHand(awaySpId, pitcherHandCache);
  const homeSpHand = getHistoricalPitcherHand(homeSpId, pitcherHandCache);

  const awayStarter = pitcherStats[awaySpId] || blankPitcherAdvancedStats();
  const homeStarter = pitcherStats[homeSpId] || blankPitcherAdvancedStats();

  const awayBullpen = bullpenStats[awayTeam] || blankPitcherAdvancedStats();
  const homeBullpen = bullpenStats[homeTeam] || blankPitcherAdvancedStats();

  const awayTeamBucket = teamStats[awayTeam] || blankTeamAdvancedStats();
  const homeTeamBucket = teamStats[homeTeam] || blankTeamAdvancedStats();

  const awayStarterEra = calculateAdvancedERA(awayStarter);
  const homeStarterEra = calculateAdvancedERA(homeStarter);
  const starterEraEdge = edgeAdvanced(awayStarterEra, homeStarterEra, "lower");

  const awayStarterWhip = calculateAdvancedWHIP(awayStarter);
  const homeStarterWhip = calculateAdvancedWHIP(homeStarter);
  const starterWhipEdge = edgeAdvanced(awayStarterWhip, homeStarterWhip, "lower");

  const awayBullpenEra = calculateAdvancedERA(awayBullpen);
  const homeBullpenEra = calculateAdvancedERA(homeBullpen);
  const bullpenEraEdge = edgeAdvanced(awayBullpenEra, homeBullpenEra, "lower");

  const awayBullpenWhip = calculateAdvancedWHIP(awayBullpen);
  const homeBullpenWhip = calculateAdvancedWHIP(homeBullpen);
  const bullpenWhipEdge = edgeAdvanced(awayBullpenWhip, homeBullpenWhip, "lower");

  const awayOpsVsHand = calculateAdvancedOPS(
    awayTeamBucket.vsHand[homeSpHand] || blankBattingAdvancedStats()
  );

  const homeOpsVsHand = calculateAdvancedOPS(
    homeTeamBucket.vsHand[awaySpHand] || blankBattingAdvancedStats()
  );

  const opsVsHandEdge = edgeAdvanced(awayOpsVsHand, homeOpsVsHand, "higher");

  const awayOps = calculateAdvancedOPS(awayTeamBucket.overall);
  const homeOps = calculateAdvancedOPS(homeTeamBucket.overall);
  const teamOpsEdge = edgeAdvanced(awayOps, homeOps, "higher");

  const score = calculateAdvancedHistoricalScore({
    starterEraEdge,
    starterWhipEdge,
    bullpenEraEdge,
    bullpenWhipEdge,
    opsVsHandEdge,
    teamOpsEdge
  });

  const pick =
    score.awayScore > score.homeScore ? awayTeam :
    score.homeScore > score.awayScore ? homeTeam :
    "Coin Flip";

  const winner = game["Winner"];

  const correct =
    pick === "Coin Flip" || !winner ? "" : pick === winner;

  const confidence = calculateAdvancedConfidence(score.awayScore, score.homeScore);

  return [
    gameDate,
    game["Game ID"],
    awayTeam,
    homeTeam,
    winner,

    starterInfo.awaySpName,
    starterInfo.homeSpName,
    awaySpId,
    homeSpId,
    awaySpHand,
    homeSpHand,

    awayStarterEra,
    homeStarterEra,
    starterEraEdge,

    awayStarterWhip,
    homeStarterWhip,
    starterWhipEdge,

    awayBullpenEra,
    homeBullpenEra,
    bullpenEraEdge,

    awayBullpenWhip,
    homeBullpenWhip,
    bullpenWhipEdge,

    awayOpsVsHand,
    homeOpsVsHand,
    opsVsHandEdge,

    awayOps,
    homeOps,
    teamOpsEdge,

    score.awayScore,
    score.homeScore,
    pick,
    confidence,
    correct
  ];
}

function updateAdvancedStatsAfterGame(
  game,
  teamStats,
  pitcherStats,
  bullpenStats,
  pitcherHandCache
) {
  const gameId = game["Game ID"];
  const awayTeam = game["Away Team"];
  const homeTeam = game["Home Team"];

  if (!gameId || !awayTeam || !homeTeam) return;

  const boxscore = getAdvancedBoxscore(gameId);
  const starterInfo = getAdvancedGameStarterIds(game);

  const awaySpId = starterInfo.awaySpId;
  const homeSpId = starterInfo.homeSpId;

  const awaySpHand = getHistoricalPitcherHand(awaySpId, pitcherHandCache);
  const homeSpHand = getHistoricalPitcherHand(homeSpId, pitcherHandCache);

  ensureAdvancedTeam(teamStats, awayTeam);
  ensureAdvancedTeam(teamStats, homeTeam);

  processAdvancedTeamBatting(
    teamStats[awayTeam],
    boxscore.teams.away.teamStats.batting,
    homeSpHand
  );

  processAdvancedTeamBatting(
    teamStats[homeTeam],
    boxscore.teams.home.teamStats.batting,
    awaySpHand
  );

  processAdvancedPitchingStaff(
    boxscore.teams.away.players,
    boxscore.teams.away.pitchers,
    awaySpId,
    pitcherStats,
    bullpenStats,
    awayTeam
  );

  processAdvancedPitchingStaff(
    boxscore.teams.home.players,
    boxscore.teams.home.pitchers,
    homeSpId,
    pitcherStats,
    bullpenStats,
    homeTeam
  );
}

function getAdvancedBoxscore(gameId) {
  const key = String(gameId);

  if (ADVANCED_BOXSCORE_CACHE[key]) {
    return ADVANCED_BOXSCORE_CACHE[key];
  }

  const url = "https://statsapi.mlb.com/api/v1/game/" + key + "/boxscore";
  const boxscore = MLB_API(url);

  ADVANCED_BOXSCORE_CACHE[key] = boxscore;
  return boxscore;
}

function getAdvancedGameStarterIds(game) {
  const gameId = game["Game ID"];

  if (!gameId) {
    return {
      awaySpId: "",
      homeSpId: "",
      awaySpName: "",
      homeSpName: ""
    };
  }

  const boxscore = getAdvancedBoxscore(gameId);

  const awayStarterId = String(boxscore.teams.away.pitchers?.[0] || "");
  const homeStarterId = String(boxscore.teams.home.pitchers?.[0] || "");

  const awayPlayer = boxscore.teams.away.players["ID" + awayStarterId] || {};
  const homePlayer = boxscore.teams.home.players["ID" + homeStarterId] || {};

  return {
    awaySpId: awayStarterId,
    homeSpId: homeStarterId,
    awaySpName: awayPlayer.person?.fullName || "",
    homeSpName: homePlayer.person?.fullName || ""
  };
}

function processAdvancedTeamBatting(teamBucket, batting, opponentStarterHand) {
  if (!batting) return;

  addAdvancedBattingStats(teamBucket.overall, batting);

  if (opponentStarterHand === "R" || opponentStarterHand === "L") {
    addAdvancedBattingStats(teamBucket.vsHand[opponentStarterHand], batting);
  }
}

function addAdvancedBattingStats(bucket, batting) {
  bucket.atBats += Number(batting.atBats || 0);
  bucket.hits += Number(batting.hits || 0);
  bucket.walks += Number(batting.baseOnBalls || 0);
  bucket.hitByPitch += Number(batting.hitByPitch || 0);
  bucket.sacFlies += Number(batting.sacFlies || 0);
  bucket.totalBases += Number(batting.totalBases || 0);
}

function processAdvancedPitchingStaff(
  players,
  pitcherOrder,
  starterId,
  pitcherStats,
  bullpenStats,
  teamName
) {
  if (!players || !pitcherOrder) return;

  if (!bullpenStats[teamName]) {
    bullpenStats[teamName] = blankPitcherAdvancedStats();
  }

  pitcherOrder.forEach(id => {
    const personId = String(id || "");
    const player = players["ID" + personId];

    if (!player) return;

    const pitching = player.stats?.pitching;
    if (!pitching || !pitching.inningsPitched) return;

    if (!pitcherStats[personId]) {
      pitcherStats[personId] = blankPitcherAdvancedStats();
    }

    if (personId === String(starterId)) {
      addAdvancedPitchingStats(pitcherStats[personId], pitching);
    } else {
      addAdvancedPitchingStats(bullpenStats[teamName], pitching);
    }
  });
}

function addAdvancedPitchingStats(bucket, pitching) {
  bucket.outs += advancedInningsToOuts(pitching.inningsPitched);
  bucket.earnedRuns += Number(pitching.earnedRuns || 0);
  bucket.hits += Number(pitching.hits || 0);
  bucket.walks += Number(pitching.baseOnBalls || 0);
}

function calculateAdvancedHistoricalScore(edges) {
  const weights = {
    starterEraEdge: 3,
    starterWhipEdge: 3,
    bullpenEraEdge: 3,
    bullpenWhipEdge: 2,
    opsVsHandEdge: 3,
    teamOpsEdge: 2
  };

  let score = 0;
  let totalWeight = 0;

  Object.keys(weights).forEach(key => {
    const value = parseAdvancedNumber(edges[key]);
    if (value === "") return;

    score += value * weights[key];
    totalWeight += weights[key];
  });

  const finalScore = totalWeight > 0 ? score / totalWeight : 0;

  return {
    awayScore: roundAdvancedNumber(finalScore, 4),
    homeScore: roundAdvancedNumber(finalScore * -1, 4)
  };
}

function calculateAdvancedConfidence(awayScore, homeScore) {
  const away = parseAdvancedNumber(awayScore);
  const home = parseAdvancedNumber(homeScore);

  if (away === "" || home === "") return "";

  const scoreGap = Math.abs(away - home);
  return Math.min(100, Math.round(scoreGap * 10));
}

function calculateAdvancedERA(bucket) {
  if (!bucket || bucket.outs === 0) return "";

  const innings = bucket.outs / 3;
  return roundAdvancedNumber((bucket.earnedRuns * 9) / innings, 2);
}

function calculateAdvancedWHIP(bucket) {
  if (!bucket || bucket.outs === 0) return "";

  const innings = bucket.outs / 3;
  return roundAdvancedNumber((bucket.hits + bucket.walks) / innings, 2);
}

function calculateAdvancedOPS(bucket) {
  if (!bucket || bucket.atBats === 0) return "";

  const obpDenominator =
    bucket.atBats +
    bucket.walks +
    bucket.hitByPitch +
    bucket.sacFlies;

  const obp =
    obpDenominator > 0
      ? (bucket.hits + bucket.walks + bucket.hitByPitch) / obpDenominator
      : 0;

  const slg =
    bucket.atBats > 0
      ? bucket.totalBases / bucket.atBats
      : 0;

  return roundAdvancedNumber(obp + slg, 3);
}

function edgeAdvanced(awayValue, homeValue, direction) {
  const away = parseAdvancedNumber(awayValue);
  const home = parseAdvancedNumber(homeValue);

  if (away === "" || home === "") return "";

  if (direction === "lower") {
    return roundAdvancedNumber(home - away, 4);
  }

  return roundAdvancedNumber(away - home, 4);
}

function advancedInningsToOuts(inningsText) {
  const parts = inningsText.toString().split(".");
  const fullInnings = Number(parts[0] || 0);
  const extraOuts = Number(parts[1] || 0);

  return fullInnings * 3 + extraOuts;
}

function getHistoricalPitcherHand(playerId, cache) {
  if (!playerId) return "";

  const key = String(playerId);

  if (cache[key]) return cache[key];

  const url = "https://statsapi.mlb.com/api/v1/people/" + key;
  const data = MLB_API(url);

  const hand = data.people?.[0]?.pitchHand?.code || "";
  cache[key] = hand;

  return hand;
}

function blankTeamAdvancedStats() {
  return {
    overall: blankBattingAdvancedStats(),
    vsHand: {
      R: blankBattingAdvancedStats(),
      L: blankBattingAdvancedStats()
    }
  };
}

function blankBattingAdvancedStats() {
  return {
    atBats: 0,
    hits: 0,
    walks: 0,
    hitByPitch: 0,
    sacFlies: 0,
    totalBases: 0
  };
}

function blankPitcherAdvancedStats() {
  return {
    outs: 0,
    earnedRuns: 0,
    hits: 0,
    walks: 0
  };
}

function ensureAdvancedTeam(teamStats, team) {
  if (!teamStats[team]) {
    teamStats[team] = blankTeamAdvancedStats();
  }
}

function parseAdvancedNumber(value) {
  if (value === "" || value === undefined || value === null) return "";

  const num = Number(value);
  return isNaN(num) ? "" : num;
}

function roundAdvancedNumber(value, decimals) {
  if (value === "" || isNaN(value)) return "";
  return Number(value.toFixed(decimals));
}

function buildAdvancedHistoricalBacktestSummary() {
  const sourceRows = getSheetRows("Historical_Backtest_Advanced");
  const sheet = getOrCreateSheet("Historical_Backtest_Advanced_Summary");
  sheet.clearContents();

  let wins = 0;
  let losses = 0;

  sourceRows.forEach(row => {
    const correct = normalizeOptimizedBoolean(row["Historical Correct?"]);

    if (correct === true) wins++;
    if (correct === false) losses++;
  });

  const games = wins + losses;
  const winRate = games > 0 ? Math.round((wins / games) * 1000) / 10 + "%" : "";

  const output = [
    ["Advanced Historical Backtest Summary", ""],
    ["", ""],
    ["Total Games", games],
    ["Correct Picks", wins],
    ["Incorrect Picks", losses],
    ["Win %", winRate]
  ];

  sheet.getRange(1, 1, output.length, 2).setValues(output);
  sheet.autoResizeColumns(1, 2);
}

function formatAdvancedHistoricalBacktest(sheet) {
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