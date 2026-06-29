function importRecentForm() {
  const sheet = getOrCreateSheet("RAW_Recent_Form");
  sheet.clearContents();

  const today = new Date();
  const windows = [7, 14, 30];

  const teamStats = {};

  windows.forEach(daysBack => {
    const startDate = new Date(today);
    startDate.setDate(today.getDate() - daysBack);

    const start = Utilities.formatDate(startDate, "America/New_York", "yyyy-MM-dd");
    const end = Utilities.formatDate(today, "America/New_York", "yyyy-MM-dd");

    const scheduleUrl =
      "https://statsapi.mlb.com/api/v1/schedule" +
      "?sportId=1" +
      "&startDate=" + start +
      "&endDate=" + end;

    const scheduleData = MLB_API(scheduleUrl);
    const dates = scheduleData.dates || [];

    dates.forEach(dateBlock => {
      const games = dateBlock.games || [];

      games.forEach(game => {
        const status = game.status.detailedState || "";
        const isFinal = status === "Final" || status === "Game Over";

        if (!isFinal) return;

        const boxscoreUrl =
          "https://statsapi.mlb.com/api/v1/game/" +
          game.gamePk +
          "/boxscore";

        const boxscore = MLB_API(boxscoreUrl);

        processRecentFormTeam(
          teamStats,
          daysBack,
          boxscore.teams.away.team.name,
          boxscore.teams.away.teamStats
        );

        processRecentFormTeam(
          teamStats,
          daysBack,
          boxscore.teams.home.team.name,
          boxscore.teams.home.teamStats
        );
      });
    });
  });

  const rows = [[
    "Team",
    "Runs/Game Last 7",
    "OPS Last 7",
    "ERA Last 7",
    "Runs/Game Last 14",
    "OPS Last 14",
    "ERA Last 14",
    "Runs/Game Last 30",
    "OPS Last 30",
    "ERA Last 30",
    "Last Updated"
  ]];

  const todayText = Utilities.formatDate(today, "America/New_York", "yyyy-MM-dd");

  Object.keys(teamStats).sort().forEach(team => {
    rows.push([
      team,
      calculateRunsPerGame(teamStats[team][7]),
      calculateOPS(teamStats[team][7]),
      calculateERA(teamStats[team][7]),
      calculateRunsPerGame(teamStats[team][14]),
      calculateOPS(teamStats[team][14]),
      calculateERA(teamStats[team][14]),
      calculateRunsPerGame(teamStats[team][30]),
      calculateOPS(teamStats[team][30]),
      calculateERA(teamStats[team][30]),
      todayText
    ]);
  });

  sheet.getRange(1, 1, rows.length, rows[0].length).setValues(rows);
}


function processRecentFormTeam(teamStats, daysBack, teamName, stats) {
  if (!teamStats[teamName]) {
    teamStats[teamName] = {};
  }

  if (!teamStats[teamName][daysBack]) {
    teamStats[teamName][daysBack] = {
      games: 0,
      runs: 0,
      atBats: 0,
      hits: 0,
      doubles: 0,
      triples: 0,
      homeRuns: 0,
      walks: 0,
      hitByPitch: 0,
      sacrificeFlies: 0,
      totalBases: 0,
      earnedRunsAllowed: 0,
      outsPitched: 0
    };
  }

  const bucket = teamStats[teamName][daysBack];
  const batting = stats.batting || {};
  const pitching = stats.pitching || {};

  bucket.games += 1;
  bucket.runs += Number(batting.runs || 0);
  bucket.atBats += Number(batting.atBats || 0);
  bucket.hits += Number(batting.hits || 0);
  bucket.doubles += Number(batting.doubles || 0);
  bucket.triples += Number(batting.triples || 0);
  bucket.homeRuns += Number(batting.homeRuns || 0);
  bucket.walks += Number(batting.baseOnBalls || 0);
  bucket.hitByPitch += Number(batting.hitByPitch || 0);
  bucket.sacrificeFlies += Number(batting.sacFlies || 0);
  bucket.totalBases += Number(batting.totalBases || 0);

  bucket.earnedRunsAllowed += Number(pitching.earnedRuns || 0);
  bucket.outsPitched += recentInningsToOuts(pitching.inningsPitched || "0.0");
}


function calculateRunsPerGame(bucket) {
  if (!bucket || bucket.games === 0) return "";
  return roundRecentNumber(bucket.runs / bucket.games, 2);
}


function calculateOPS(bucket) {
  if (!bucket || bucket.atBats === 0) return "";

  const obpDenominator =
    bucket.atBats +
    bucket.walks +
    bucket.hitByPitch +
    bucket.sacrificeFlies;

  const obp =
    obpDenominator > 0
      ? (bucket.hits + bucket.walks + bucket.hitByPitch) / obpDenominator
      : 0;

  const slg =
    bucket.atBats > 0
      ? bucket.totalBases / bucket.atBats
      : 0;

  return roundRecentNumber(obp + slg, 3);
}


function calculateERA(bucket) {
  if (!bucket || bucket.outsPitched === 0) return "";

  const innings = bucket.outsPitched / 3;
  return roundRecentNumber((bucket.earnedRunsAllowed * 9) / innings, 2);
}


function recentInningsToOuts(inningsText) {
  const parts = inningsText.toString().split(".");
  const fullInnings = Number(parts[0] || 0);
  const extraOuts = Number(parts[1] || 0);

  return fullInnings * 3 + extraOuts;
}


function roundRecentNumber(value, decimals) {
  if (value === "" || isNaN(value)) return "";
  return Number(value.toFixed(decimals));
}