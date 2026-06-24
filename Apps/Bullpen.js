function importBullpen() {
  const sheet = getOrCreateSheet("RAW_Bullpen");
  sheet.clearContents();

  const url = "https://www.baseball-reference.com/leagues/majors/2026-reliever-pitching.shtml";
  const html = UrlFetchApp.fetch(url).getContentText();

  const tableMatch = html.match(/<table[\s\S]*?<\/table>/);
  if (!tableMatch) {
    throw new Error("Could not find Baseball Reference reliever table.");
  }

  const tableHtml = tableMatch[0];
  const rowMatches = [...tableHtml.matchAll(/<tr[^>]*>([\s\S]*?)<\/tr>/g)];

  const today = Utilities.formatDate(new Date(), "America/New_York", "yyyy-MM-dd");

  const rows = [[
    "Team",
    "Relief Games",
    "Games Finished",
    "Save Opportunities",
    "Saves",
    "Blown Saves",
    "Save %",
    "Holds",
    "Inherited Runners",
    "Inherited Scored",
    "Inherited Scored %",
    "Avg Leverage Index",
    "Relief Appearances >3 Outs",
    "Relief Appearances <3 Outs",
    "Multi-Inning Relief Appearances",
    "Outs Per Relief Game",
    "Pitches Per Relief Game",
    "Last Updated",
    "Source"
  ]];

  rowMatches.forEach(rowMatch => {
    const rowHtml = rowMatch[1];

    const cells = [...rowHtml.matchAll(/<(td|th)[^>]*data-stat="([^"]+)"[^>]*>([\s\S]*?)<\/(td|th)>/g)];
    const row = {};

    cells.forEach(cell => {
      const statName = cell[2];
      const value = stripHtml(cell[3]);
      row[statName] = value;
    });

    const teamName = row.team_name || "";

    if (!teamName || teamName === "Tm" || teamName === "League Average") return;

    rows.push([
      teamName,
      row.GR || "",
      row.GF || "",
      row.SvOpp || "",
      row.SV || "",
      row.BSv || "",
      row.SvOpp_perc || "",
      row.Hold || "",
      row.inherited_runners || "",
      row.inherited_score || "",
      row.inherited_score_perc || "",
      row.leverage_index_avg || "",
      row.IPouts_gt3 || "",
      row.IPouts_lt3 || "",
      row.IP_multi || "",
      row.outs_per_GR || "",
      row.pitches_per_GR || "",
      today,
      "Baseball Reference Reliever Usage"
    ]);
  });

  sheet.getRange(1, 1, rows.length, rows[0].length).setValues(rows);
}