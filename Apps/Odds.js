function importOdds() {
  const rawSheet = getOrCreateSheet("RAW_Odds");
  const historySheet = getOrCreateSheet("ODDS_HISTORY");

  rawSheet.clearContents();

  const apiKey = PropertiesService
    .getScriptProperties()
    .getProperty("ODDS_API_KEY");

  if (!apiKey) {
    throw new Error("Missing ODDS_API_KEY in Script Properties.");
  }

  const url =
    "https://api.the-odds-api.com/v4/sports/baseball_mlb/odds" +
    "?apiKey=" + apiKey +
    "&regions=us" +
    "&markets=h2h" +
    "&oddsFormat=american" +
    "&bookmakers=draftkings,fanduel";

  const response = UrlFetchApp.fetch(url);
  const data = JSON.parse(response.getContentText());

  if (!Array.isArray(data)) {
    throw new Error("Odds API returned non-array response: " + response.getContentText());
  }

  const snapshotTime = Utilities.formatDate(
    new Date(),
    "America/New_York",
    "yyyy-MM-dd HH:mm:ss"
  );

  const headers = [
    "Snapshot Time",
    "Game ID",
    "Commence Time",
    "Away Team",
    "Home Team",
    "Bookmaker",
    "Away ML",
    "Home ML",
    "Source"
  ];

  const rows = [headers];

  data.forEach(game => {
    const gameId = game.id || "";
    const commenceTime = game.commence_time || "";
    const awayTeam = game.away_team || "";
    const homeTeam = game.home_team || "";

    const bookmakers = game.bookmakers || [];

    bookmakers.forEach(book => {
      const markets = book.markets || [];
      const market = markets.find(market => market.key === "h2h");
      if (!market) return;

      let awayML = "";
      let homeML = "";

      const outcomes = market.outcomes || [];

      outcomes.forEach(outcome => {
        if (outcome.name === awayTeam) awayML = outcome.price;
        if (outcome.name === homeTeam) homeML = outcome.price;
      });

      rows.push([
        snapshotTime,
        gameId,
        commenceTime,
        awayTeam,
        homeTeam,
        book.title,
        awayML,
        homeML,
        "The Odds API"
      ]);
    });
  });

  rawSheet.getRange(1, 1, rows.length, rows[0].length).setValues(rows);

  if (historySheet.getLastRow() === 0) {
    historySheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  }

  const historyRows = rows.slice(1);

  if (historyRows.length > 0) {
    historySheet
      .getRange(historySheet.getLastRow() + 1, 1, historyRows.length, headers.length)
      .setValues(historyRows);
  }
}