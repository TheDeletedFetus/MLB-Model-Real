function importTeamHomeAwaySplits() {
  const sheet = getOrCreateSheet("RAW_Team_Home_Away_Splits");
  sheet.clearContents();

  const homeBatting = getBRTeamBattingSplit("Home");
  const roadBatting = getBRTeamBattingSplit("Away");

  const homePitching = getBRTeamPitchingSplit("Home");
  const roadPitching = getBRTeamPitchingSplit("Away");

  const today = Utilities.formatDate(new Date(), "America/New_York", "yyyy-MM-dd");

  const rows = [[
    "Team",
    "Home OPS",
    "Road OPS",
    "Home OBP",
    "Road OBP",
    "Home SLG",
    "Road SLG",
    "Home ERA",
    "Road ERA",
    "Home WHIP",
    "Road WHIP",
    "Home K/9",
    "Road K/9",
    "Home K/BB",
    "Road K/BB",
    "Last Updated",
    "Source"
  ]];

  Object.keys(homeBatting).sort().forEach(team => {
    const hb = homeBatting[team] || {};
    const rb = roadBatting[team] || {};
    const hp = homePitching[team] || {};
    const rp = roadPitching[team] || {};

    rows.push([
      team,
      hb.ops || "",
      rb.ops || "",
      hb.obp || "",
      rb.obp || "",
      hb.slg || "",
      rb.slg || "",
      hp.era || "",
      rp.era || "",
      hp.whip || "",
      rp.whip || "",
      hp.k9 || "",
      rp.k9 || "",
      hp.kbb || "",
      rp.kbb || "",
      today,
      "Baseball Reference home/away batting and pitching splits"
    ]);
  });

  sheet.getRange(1, 1, rows.length, rows[0].length).setValues(rows);
}


function getBRTeamBattingSplit(location) {
  const encodedLocation = encodeURIComponent(location);

  const url =
    "https://www.baseball-reference.com/tools/split_stats_lg.cgi" +
    "?full=1" +
    "&params=home%7C" + encodedLocation + "%7CML%7C2026%7Cbat%7CAB%7C";

  const html = UrlFetchApp.fetch(url).getContentText();
  const tableMatch = html.match(/<table[\s\S]*?<\/table>/);

  if (!tableMatch) {
    throw new Error("Could not find Baseball Reference batting table for " + location);
  }

  const tableHtml = tableMatch[0];
  const rowMatches = [...tableHtml.matchAll(/<tr[^>]*>([\s\S]*?)<\/tr>/g)];

  const output = {};

  rowMatches.forEach(rowMatch => {
    const rowHtml = rowMatch[1];

    const cells = [...rowHtml.matchAll(/<(td|th)[^>]*data-stat="([^"]+)"[^>]*>([\s\S]*?)<\/(td|th)>/g)];
    const row = {};

    cells.forEach(cell => {
      const statName = cell[2];
      const value = stripHtml(cell[3]);
      row[statName] = value;
    });

    const teamAbbrev = row.team || "";
    if (!teamAbbrev || teamAbbrev === "Tm" || teamAbbrev === "LgAvg") return;

    const teamName = brTeamToMlbName(teamAbbrev);
    if (!teamName) return;

    output[teamName] = {
      ops: row.onbase_plus_slugging || "",
      obp: row.onbase_perc || "",
      slg: row.slugging_perc || ""
    };
  });

  return output;
}


function getBRTeamPitchingSplit(location) {
  const encodedLocation = encodeURIComponent(location);

  const url =
    "https://www.baseball-reference.com/tools/split_stats_lg.cgi" +
    "?full=1" +
    "&params=home%7C" + encodedLocation + "%7CML%7C2026%7Cpitch%7CIP%7C";

  const html = UrlFetchApp.fetch(url).getContentText();
  const tableMatch = html.match(/<table[\s\S]*?<\/table>/);

  if (!tableMatch) {
    throw new Error("Could not find Baseball Reference pitching table for " + location);
  }

  const tableHtml = tableMatch[0];
  const rowMatches = [...tableHtml.matchAll(/<tr[^>]*>([\s\S]*?)<\/tr>/g)];

  const output = {};

  rowMatches.forEach(rowMatch => {
    const rowHtml = rowMatch[1];

    const cells = [...rowHtml.matchAll(/<(td|th)[^>]*data-stat="([^"]+)"[^>]*>([\s\S]*?)<\/(td|th)>/g)];
    const row = {};

    cells.forEach(cell => {
      const statName = cell[2];
      const value = stripHtml(cell[3]);
      row[statName] = value;
    });

    const teamAbbrev = row.team || "";
    if (!teamAbbrev || teamAbbrev === "Tm" || teamAbbrev === "LgAvg") return;

    const teamName = brTeamToMlbName(teamAbbrev);
    if (!teamName) return;

    output[teamName] = {
      era: row.earned_run_avg || "",
      whip: row.whip || "",
      k9: row.strikeouts_per_nine || "",
      kbb: row.strikeouts_per_base_on_balls || ""
    };
  });

  return output;
}