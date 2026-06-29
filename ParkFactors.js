function importParkFactors() {
  const sheet = getOrCreateSheet("RAW_Park_Factors");
  sheet.clearContents();

  const url =
    "https://baseballsavant.mlb.com/leaderboard/statcast-park-factors";

  const html = UrlFetchApp.fetch(url).getContentText();

  const match = html.match(/var data = (\[[\s\S]*?\]);/);

  if (!match) {
    throw new Error("Could not locate park factor data.");
  }

  const parkData = JSON.parse(match[1]);

  const rows = [[
    "Venue",
    "Team",
    "Run Factor",
    "HR Factor",
    "Hit Factor",
    "PA Sample",
    "Year Range",
    "Last Updated",
    "Source"
  ]];

  const today = Utilities.formatDate(
    new Date(),
    "America/New_York",
    "yyyy-MM-dd"
  );

  parkData.forEach(park => {
    rows.push([
      park.venue_name || "",
      park.name_display_club || "",
      park.index_runs || "",
      park.index_hr || "",
      park.index_hits || "",
      park.n_pa || "",
      park.year_range || "",
      today,
      "Baseball Savant Park Factors"
    ]);
  });

  sheet
    .getRange(1, 1, rows.length, rows[0].length)
    .setValues(rows);
}