function importTeamPitching() {
  const sheet = getOrCreateSheet("RAW_Team_Pitching");
  sheet.clearContents();

  const url = "https://statsapi.mlb.com/api/v1/teams/stats?group=pitching&stats=season&sportIds=1";
  const data = MLB_API(url);

  const rows = [[
    "Team",
    "Team ID",
    "Games",
    "Runs Allowed",
    "ERA",
    "WHIP",
    "K/9",
    "BB/9",
    "HR/9",
    "K-BB Ratio",
    "LOB%"
  ]];

  data.stats[0].splits.forEach(row => {
    const stat = row.stat;

    rows.push([
      row.team.name,
      row.team.id,
      stat.gamesPlayed || "",
      stat.runs || "",
      stat.era || "",
      stat.whip || "",
      stat.strikeoutsPer9Inn || "",
      stat.walksPer9Inn || "",
      stat.homeRunsPer9 || "",
      stat.strikeoutWalkRatio || "",
      stat.leftOnBasePercentage || ""
    ]);
  });

  sheet.getRange(1, 1, rows.length, rows[0].length).setValues(rows);
}