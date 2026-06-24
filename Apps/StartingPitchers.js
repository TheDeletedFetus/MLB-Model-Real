function importStartingPitchers() {
  const scheduleSheet = getOrCreateSheet("RAW_Schedule");
  const pitcherSheet = getOrCreateSheet("RAW_Starting_Pitchers");

  pitcherSheet.clearContents();

  const scheduleData = scheduleSheet.getDataRange().getValues();
  const headers = scheduleData[0];

  const awaySpIdCol = headers.indexOf("Away SP ID");
  const homeSpIdCol = headers.indexOf("Home SP ID");

  const pitcherIds = [];

  for (let i = 1; i < scheduleData.length; i++) {
    const awayId = scheduleData[i][awaySpIdCol];
    const homeId = scheduleData[i][homeSpIdCol];

    if (awayId) pitcherIds.push(awayId);
    if (homeId) pitcherIds.push(homeId);
  }

  const uniquePitcherIds = [...new Set(pitcherIds)];

  const rows = [[
    "Player",
    "Player ID",
    "Team",
    "Team ID",
    "Throws",
    "Games",
    "Games Started",
    "ERA",
    "WHIP",
    "IP",
    "SO",
    "BB",
    "HR",
    "K/BB",
    "K/9",
    "BB/9",
    "HR/9"
  ]];

  uniquePitcherIds.forEach(playerId => {
    const url =
      "https://statsapi.mlb.com/api/v1/people/" +
      playerId +
      "?hydrate=stats(group=[pitching],type=[season])";

    const data = MLB_API(url);
    const player = data.people?.[0] || {};
    const stat = player.stats?.[0]?.splits?.[0]?.stat || {};
    const team = player.currentTeam || {};

    rows.push([
      player.fullName || "",
      player.id || "",
      team.name || "",
      team.id || "",
      player.pitchHand?.code || "",
      stat.gamesPlayed || "",
      stat.gamesStarted || "",
      stat.era || "",
      stat.whip || "",
      stat.inningsPitched || "",
      stat.strikeOuts || "",
      stat.baseOnBalls || "",
      stat.homeRuns || "",
      stat.strikeoutWalkRatio || "",
      stat.strikeoutsPer9Inn || "",
      stat.walksPer9Inn || "",
      stat.homeRunsPer9 || ""
    ]);
  });

  pitcherSheet.getRange(1, 1, rows.length, rows[0].length).setValues(rows);
}