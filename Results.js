function importResults() {
  importResultsForDate(getYesterdayText());
}


function importResultsToday() {
  importResultsForDate(getTodayText());
}


function importResultsForDate(targetDate) {
  const sheet = getOrCreateSheet("RAW_Results");
  sheet.clearContents();

  const url =
    "https://statsapi.mlb.com/api/v1/schedule" +
    "?sportId=1" +
    "&date=" + targetDate;

  const data = MLB_API(url);

  const rows = [[
    "Date",
    "Game ID",
    "Away Team",
    "Home Team",
    "Away Score",
    "Home Score",
    "Winner",
    "Status",
    "Final?"
  ]];

  const games = data.dates[0]?.games || [];

  games.forEach(game => {
    const awayTeam = game.teams.away.team.name;
    const homeTeam = game.teams.home.team.name;

    const awayScore = game.teams.away.score ?? "";
    const homeScore = game.teams.home.score ?? "";

    const status = game.status.detailedState || "";
    const isFinal = status === "Final" || status === "Game Over";

    let winner = "";

    if (isFinal && awayScore !== "" && homeScore !== "") {
      winner = Number(awayScore) > Number(homeScore) ? awayTeam : homeTeam;
    }

    rows.push([
      targetDate,
      game.gamePk,
      awayTeam,
      homeTeam,
      awayScore,
      homeScore,
      winner,
      status,
      isFinal
    ]);
  });

  sheet.getRange(1, 1, rows.length, rows[0].length).setValues(rows);
}


function getTodayText() {
  return Utilities.formatDate(new Date(), "America/New_York", "yyyy-MM-dd");
}


function getYesterdayText() {
  const date = new Date();
  date.setDate(date.getDate() - 1);
  return Utilities.formatDate(date, "America/New_York", "yyyy-MM-dd");
}