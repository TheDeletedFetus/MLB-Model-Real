function importBullpenFatigue() {
  const sheet = getOrCreateSheet("RAW_Bullpen_Fatigue");
  sheet.clearContents();

  const today = new Date();
  const todayText = Utilities.formatDate(today, "America/New_York", "yyyy-MM-dd");

  const fatigueByTeam = {};

  calculateBullpenFatigueWindow(fatigueByTeam, 3);
  calculateBullpenFatigueWindow(fatigueByTeam, 7);

  const rows = [[
    "Team",
    "Bullpen IP Last 3 Days",
    "Bullpen IP Last 7 Days",
    "Bullpen Appearances Last 3 Days",
    "Bullpen Appearances Last 7 Days",
    "Bullpen Pitches Last 3 Days",
    "Bullpen Pitches Last 7 Days",
    "Last Updated"
  ]];

  Object.keys(fatigueByTeam).sort().forEach(team => {
    rows.push([
      team,
      outsToInningsFatigue(fatigueByTeam[team][3]?.outs || 0),
      outsToInningsFatigue(fatigueByTeam[team][7]?.outs || 0),
      fatigueByTeam[team][3]?.appearances || 0,
      fatigueByTeam[team][7]?.appearances || 0,
      fatigueByTeam[team][3]?.pitches || 0,
      fatigueByTeam[team][7]?.pitches || 0,
      todayText
    ]);
  });

  sheet.getRange(1, 1, rows.length, rows[0].length).setValues(rows);
}


function calculateBullpenFatigueWindow(fatigueByTeam, daysBack) {
  const today = new Date();
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

      processBullpenFatigueTeam(
        fatigueByTeam,
        daysBack,
        boxscore.teams.away.team.name,
        boxscore.teams.away.players
      );

      processBullpenFatigueTeam(
        fatigueByTeam,
        daysBack,
        boxscore.teams.home.team.name,
        boxscore.teams.home.players
      );
    });
  });
}


function processBullpenFatigueTeam(fatigueByTeam, daysBack, teamName, players) {
  if (!fatigueByTeam[teamName]) {
    fatigueByTeam[teamName] = {};
  }

  if (!fatigueByTeam[teamName][daysBack]) {
    fatigueByTeam[teamName][daysBack] = {
      outs: 0,
      appearances: 0,
      pitches: 0
    };
  }

  Object.keys(players).forEach(playerKey => {
    const player = players[playerKey];
    const pitchingStats = player.stats?.pitching;

    if (!pitchingStats) return;

    const gamesStarted = Number(pitchingStats.gamesStarted || 0);

    // Bullpen fatigue = pitchers who appeared in relief.
    // If gamesStarted is 0 and inningsPitched exists, count him as bullpen use.
    if (gamesStarted > 0) return;
    if (!pitchingStats.inningsPitched) return;

    fatigueByTeam[teamName][daysBack].outs += bullpenFatigueInningsToOuts(
      pitchingStats.inningsPitched
    );

    fatigueByTeam[teamName][daysBack].appearances += 1;
    fatigueByTeam[teamName][daysBack].pitches += Number(pitchingStats.numberOfPitches || 0);
  });
}


function bullpenFatigueInningsToOuts(inningsText) {
  const parts = inningsText.toString().split(".");
  const fullInnings = Number(parts[0] || 0);
  const extraOuts = Number(parts[1] || 0);

  return fullInnings * 3 + extraOuts;
}


function outsToInningsFatigue(outs) {
  const fullInnings = Math.floor(outs / 3);
  const remainingOuts = outs % 3;

  return fullInnings + "." + remainingOuts;
}