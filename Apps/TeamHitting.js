function importTeamHitting() {
  const sheet = getOrCreateSheet("RAW_Team_Hitting");
  sheet.clearContents();

  const url = "https://statsapi.mlb.com/api/v1/teams/stats?group=hitting&stats=season&sportIds=1";
  const data = MLB_API(url);

  const rows = [[
    "Team",
    "Team ID",
    "Games",
    "Runs",
    "Runs/Game",
    "AVG",
    "OBP",
    "SLG",
    "OPS",
    "HR",
    "BB",
    "SO",
    "BABIP"
  ]];

  data.stats[0].splits.forEach(row => {
    const stat = row.stat;
    const games = Number(stat.gamesPlayed || 0);
    const runs = Number(stat.runs || 0);

    rows.push([
      row.team.name,
      row.team.id,
      games,
      runs,
      games > 0 ? runs / games : "",
      stat.avg || "",
      stat.obp || "",
      stat.slg || "",
      stat.ops || "",
      stat.homeRuns || "",
      stat.baseOnBalls || "",
      stat.strikeOuts || "",
      stat.babip || ""
    ]);
  });

  sheet.getRange(1, 1, rows.length, rows[0].length).setValues(rows);
}