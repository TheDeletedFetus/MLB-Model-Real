function getPitcherHand(playerId) {
  if (!playerId) return "";

  const url = "https://statsapi.mlb.com/api/v1/people/" + playerId;
  const data = MLB_API(url);

  return data.people?.[0]?.pitchHand?.code || "";
}

function importSchedule() {
  const sheet = getOrCreateSheet("RAW_Schedule");
  sheet.clearContents();

  const today = Utilities.formatDate(new Date(), "America/New_York", "yyyy-MM-dd");

  const url =
    "https://statsapi.mlb.com/api/v1/schedule" +
    "?sportId=1" +
    "&date=" + today +
    "&hydrate=probablePitcher,venue";

  const data = MLB_API(url);

  const rows = [[
    "Date",
    "Game ID",
    "Away Team",
    "Home Team",
    "Away SP",
    "Home SP",
    "Away SP ID",
    "Home SP ID",
    "Away SP Hand",
    "Home SP Hand",
    "Status",
    "Venue"
  ]];

  const games = data.dates[0]?.games || [];

  games.forEach(game => {
    const awayProbable = game.teams.away.probablePitcher || {};
    const homeProbable = game.teams.home.probablePitcher || {};

    rows.push([
      today,
      game.gamePk,
      game.teams.away.team.name,
      game.teams.home.team.name,
      awayProbable.fullName || "",
      homeProbable.fullName || "",
      awayProbable.id || "",
      homeProbable.id || "",
      getPitcherHand(awayProbable.id),
      getPitcherHand(homeProbable.id),
      game.status.detailedState || "",
      game.venue?.name || ""
    ]);
  });

  sheet.getRange(1, 1, rows.length, rows[0].length).setValues(rows);
}