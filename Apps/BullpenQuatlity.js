function importBullpenQuality() {
  const sheet = getOrCreateSheet("RAW_Bullpen_Quality");
  sheet.clearContents();

  const today = new Date();
  const seasonStart = "2026-03-26";
  const todayText = Utilities.formatDate(today, "America/New_York", "yyyy-MM-dd");

  const scheduleUrl =
    "https://statsapi.mlb.com/api/v1/schedule" +
    "?sportId=1" +
    "&gameTypes=R" +
    "&startDate=" + seasonStart +
    "&endDate=" + todayText;

  const scheduleData = MLB_API(scheduleUrl);
  const bullpen = {};
  const validTeams = getValidMlbTeams();

  (scheduleData.dates || []).forEach(dateBlock => {
    (dateBlock.games || []).forEach(game => {
      const status = game.status.detailedState || "";
      const isFinal = status === "Final" || status === "Game Over";

      if (!isFinal) return;

      const boxscoreUrl =
        "https://statsapi.mlb.com/api/v1/game/" +
        game.gamePk +
        "/boxscore";

      const boxscore = MLB_API(boxscoreUrl);

      const awayTeam = boxscore.teams.away.team.name;
      const homeTeam = boxscore.teams.home.team.name;

      if (validTeams.has(awayTeam)) {
        processBullpenQualityTeam(
          bullpen,
          awayTeam,
          boxscore.teams.away.players
        );
      }

      if (validTeams.has(homeTeam)) {
        processBullpenQualityTeam(
          bullpen,
          homeTeam,
          boxscore.teams.home.players
        );
      }
    });
  });

  const rows = [[
    "Team",
    "Bullpen IP",
    "Bullpen ERA",
    "Bullpen WHIP",
    "Bullpen K/9",
    "Bullpen BB/9",
    "Bullpen HR/9",
    "Bullpen SO",
    "Bullpen BB",
    "Bullpen HR",
    "Last Updated",
    "Source"
  ]];

  Object.keys(bullpen).sort().forEach(team => {
    const bp = bullpen[team];
    const innings = bp.outs / 3;

    rows.push([
      team,
      outsToInningsBullpenQuality(bp.outs),
      innings > 0 ? roundBullpenQuality((bp.earnedRuns * 9) / innings, 2) : "",
      innings > 0 ? roundBullpenQuality((bp.hits + bp.walks) / innings, 2) : "",
      innings > 0 ? roundBullpenQuality((bp.strikeouts * 9) / innings, 2) : "",
      innings > 0 ? roundBullpenQuality((bp.walks * 9) / innings, 2) : "",
      innings > 0 ? roundBullpenQuality((bp.homeRuns * 9) / innings, 2) : "",
      bp.strikeouts,
      bp.walks,
      bp.homeRuns,
      todayText,
      "MLB Stats API regular-season boxscores, reliever-only"
    ]);
  });

  sheet.getRange(1, 1, rows.length, rows[0].length).setValues(rows);
}


function processBullpenQualityTeam(bullpen, teamName, players) {
  if (!bullpen[teamName]) {
    bullpen[teamName] = {
      outs: 0,
      earnedRuns: 0,
      hits: 0,
      walks: 0,
      strikeouts: 0,
      homeRuns: 0
    };
  }

  Object.keys(players).forEach(playerKey => {
    const player = players[playerKey];
    const pitching = player.stats?.pitching;

    if (!pitching) return;

    const gamesStarted = Number(pitching.gamesStarted || 0);

    if (gamesStarted > 0) return;
    if (!pitching.inningsPitched) return;

    bullpen[teamName].outs += bullpenQualityInningsToOuts(pitching.inningsPitched);
    bullpen[teamName].earnedRuns += Number(pitching.earnedRuns || 0);
    bullpen[teamName].hits += Number(pitching.hits || 0);
    bullpen[teamName].walks += Number(pitching.baseOnBalls || 0);
    bullpen[teamName].strikeouts += Number(pitching.strikeOuts || 0);
    bullpen[teamName].homeRuns += Number(pitching.homeRuns || 0);
  });
}


function bullpenQualityInningsToOuts(inningsText) {
  const parts = inningsText.toString().split(".");
  return Number(parts[0] || 0) * 3 + Number(parts[1] || 0);
}


function outsToInningsBullpenQuality(outs) {
  return Math.floor(outs / 3) + "." + (outs % 3);
}


function roundBullpenQuality(value, decimals) {
  if (value === "" || isNaN(value)) return "";
  return Number(value.toFixed(decimals));
}


function getValidMlbTeams() {
  const hittingSheet = getOrCreateSheet("RAW_Team_Hitting");
  const data = hittingSheet.getDataRange().getValues();

  const headers = data[0];
  const teamCol = headers.indexOf("Team");

  const teams = new Set();

  for (let i = 1; i < data.length; i++) {
    const team = data[i][teamCol];
    if (team) teams.add(team);
  }

  return teams;
}