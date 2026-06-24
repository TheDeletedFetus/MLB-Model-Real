function importTeamSplits() {
  const sheet = getOrCreateSheet("RAW_Team_Splits");
  sheet.clearContents();

  const vsRhp = getBaseballReferenceSplit("vs RHP");
  const vsLhp = getBaseballReferenceSplit("vs LHP");

  const rows = [[
    "Team",
    "Team ID",
    "OPS vs RHP",
    "OPS vs LHP",
    "Last Updated",
    "Source"
  ]];

  const today = Utilities.formatDate(new Date(), "America/New_York", "yyyy-MM-dd");
  const teams = Object.keys(vsRhp);

  teams.forEach(team => {
    rows.push([
      team,
      "",
      vsRhp[team] || "",
      vsLhp[team] || "",
      today,
      "Baseball Reference"
    ]);
  });

  sheet.getRange(1, 1, rows.length, rows[0].length).setValues(rows);
}


function getBaseballReferenceSplit(splitType) {
  const encodedSplit = encodeURIComponent(splitType);

  const url =
    "https://www.baseball-reference.com/tools/split_stats_lg.cgi" +
    "?full=1" +
    "&params=plato%7C" + encodedSplit + "%7CML%7C2026%7Cbat%7CAB%7C";

  const html = UrlFetchApp.fetch(url).getContentText();

  const tableMatch = html.match(/<table[\s\S]*?<\/table>/);
  if (!tableMatch) {
    throw new Error("Could not find split table for " + splitType);
  }

  const tableHtml = tableMatch[0];
  const teamOps = {};

  const rowMatches = [...tableHtml.matchAll(/<tr[^>]*>([\s\S]*?)<\/tr>/g)];

  rowMatches.forEach(rowMatch => {
    const rowHtml = rowMatch[1];

    const cells = [...rowHtml.matchAll(/<(td|th)[^>]*data-stat="([^"]+)"[^>]*>([\s\S]*?)<\/(td|th)>/g)];

    let teamAbbrev = "";
    let ops = "";

    cells.forEach(cell => {
      const statName = cell[2];
      const rawValue = stripHtml(cell[3]);

      if (statName === "team") teamAbbrev = rawValue;
      if (statName === "onbase_plus_slugging") ops = rawValue;
    });

    const teamName = brTeamToMlbName(teamAbbrev);

    if (teamName && ops && teamName !== "LgAvg") {
      teamOps[teamName] = ops;
    }
  });

  return teamOps;
}


function stripHtml(value) {
  return value
    .replace(/<[^>]*>/g, "")
    .replace(/&nbsp;/g, " ")
    .replace(/&amp;/g, "&")
    .trim();
}


function brTeamToMlbName(abbrev) {
  const map = {
    "ARI": "Arizona Diamondbacks",
    "ATL": "Atlanta Braves",
    "BAL": "Baltimore Orioles",
    "BOS": "Boston Red Sox",
    "CHC": "Chicago Cubs",
    "CHW": "Chicago White Sox",
    "CIN": "Cincinnati Reds",
    "CLE": "Cleveland Guardians",
    "COL": "Colorado Rockies",
    "DET": "Detroit Tigers",
    "HOU": "Houston Astros",
    "KCR": "Kansas City Royals",
    "LAA": "Los Angeles Angels",
    "LAD": "Los Angeles Dodgers",
    "MIA": "Miami Marlins",
    "MIL": "Milwaukee Brewers",
    "MIN": "Minnesota Twins",
    "NYM": "New York Mets",
    "NYY": "New York Yankees",
    "ATH": "Athletics",
    "PHI": "Philadelphia Phillies",
    "PIT": "Pittsburgh Pirates",
    "SDP": "San Diego Padres",
    "SEA": "Seattle Mariners",
    "SFG": "San Francisco Giants",
    "STL": "St. Louis Cardinals",
    "TBR": "Tampa Bay Rays",
    "TEX": "Texas Rangers",
    "TOR": "Toronto Blue Jays",
    "WSN": "Washington Nationals"
  };

  return map[abbrev] || "";
}