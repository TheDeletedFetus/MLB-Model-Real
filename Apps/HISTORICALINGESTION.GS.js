function runHistoricalIngestion2026ToDate() {
  ingestHistoricalGames("2026-03-26", getTodayText());
}

function ingestHistoricalGames(startDate, endDate) {
  const historicalSheet = getOrCreateSheet("RAW_Historical_Games");
  const logSheet = getOrCreateSheet("Historical_Ingestion_Log");

  ensureHistoricalGamesHeaders(historicalSheet);
  ensureHistoricalIngestionLogHeaders(logSheet);

  const existingGameIds = getExistingHistoricalGameIds(historicalSheet);

  const url =
    "https://statsapi.mlb.com/api/v1/schedule" +
    "?sportId=1" +
    "&gameTypes=R" +
    "&startDate=" + startDate +
    "&endDate=" + endDate +
    "&hydrate=probablePitcher,venue";

  const data = MLB_API(url);
  const dates = data.dates || [];

  const rowsToAppend = [];
  const runTime = Utilities.formatDate(
    new Date(),
    "America/New_York",
    "yyyy-MM-dd HH:mm:ss"
  );

  dates.forEach(dateBlock => {
    const gameDate = dateBlock.date;
    const games = dateBlock.games || [];

    games.forEach(game => {
      const gameId = game.gamePk;
      if (!gameId) return;
      if (existingGameIds.has(String(gameId))) return;

      const status = game.status?.detailedState || "";
      const isFinal = status === "Final" || status === "Game Over";

      if (!isFinal) return;

      const awayTeam = game.teams.away.team.name;
      const homeTeam = game.teams.home.team.name;

      const awayScore = game.teams.away.score ?? "";
      const homeScore = game.teams.home.score ?? "";

      const winner =
        awayScore > homeScore ? awayTeam :
        homeScore > awayScore ? homeTeam :
        "";

      const awayProbable = game.teams.away.probablePitcher || {};
      const homeProbable = game.teams.home.probablePitcher || {};

      rowsToAppend.push([
        gameDate,
        gameId,
        awayTeam,
        homeTeam,
        awayScore,
        homeScore,
        winner,
        status,
        true,

        awayProbable.fullName || "",
        homeProbable.fullName || "",
        awayProbable.id || "",
        homeProbable.id || "",

        game.venue?.name || "",
        runTime,
        "MLB Stats API historical schedule"
      ]);
    });
  });

  if (rowsToAppend.length > 0) {
    historicalSheet
      .getRange(
        historicalSheet.getLastRow() + 1,
        1,
        rowsToAppend.length,
        rowsToAppend[0].length
      )
      .setValues(rowsToAppend);
  }

  logSheet.appendRow([
    runTime,
    startDate,
    endDate,
    rowsToAppend.length,
    "Completed"
  ]);

  buildHistoricalIngestionSummary();
}

function ensureHistoricalGamesHeaders(sheet) {
  const headers = [
    "Game Date",
    "Game ID",
    "Away Team",
    "Home Team",
    "Away Score",
    "Home Score",
    "Winner",
    "Status",
    "Final?",

    "Away SP",
    "Home SP",
    "Away SP ID",
    "Home SP ID",

    "Venue",
    "Ingested At",
    "Source"
  ];

  if (sheet.getLastRow() === 0) {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    return;
  }

  const existingHeaders = sheet
    .getRange(1, 1, 1, headers.length)
    .getValues()[0];

  if (existingHeaders.join("|") !== headers.join("|")) {
    throw new Error(
      "RAW_Historical_Games headers do not match expected v0.0.7 schema."
    );
  }
}

function ensureHistoricalIngestionLogHeaders(sheet) {
  const headers = [
    "Run Time",
    "Start Date",
    "End Date",
    "Games Added",
    "Status"
  ];

  if (sheet.getLastRow() === 0) {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  }
}

function getExistingHistoricalGameIds(sheet) {
  const values = sheet.getDataRange().getValues();
  const ids = new Set();

  if (values.length < 2) return ids;

  const headers = values[0];
  const gameIdCol = headers.indexOf("Game ID");

  for (let i = 1; i < values.length; i++) {
    const gameId = values[i][gameIdCol];
    if (gameId) ids.add(String(gameId));
  }

  return ids;
}

function buildHistoricalIngestionSummary() {
  const sourceSheet = getOrCreateSheet("RAW_Historical_Games");
  const summarySheet = getOrCreateSheet("Historical_Ingestion_Summary");

  summarySheet.clearContents();

  const rows = getSheetRows("RAW_Historical_Games");

  const bySeason = {};
  const byTeam = {};

  rows.forEach(row => {
    const gameDate = row["Game Date"];
    if (!gameDate) return;

    const season = gameDate.toString().substring(0, 4);

    if (!bySeason[season]) {
      bySeason[season] = {
        games: 0
      };
    }

    bySeason[season].games++;

    const awayTeam = row["Away Team"];
    const homeTeam = row["Home Team"];

    if (awayTeam) {
      if (!byTeam[awayTeam]) byTeam[awayTeam] = 0;
      byTeam[awayTeam]++;
    }

    if (homeTeam) {
      if (!byTeam[homeTeam]) byTeam[homeTeam] = 0;
      byTeam[homeTeam]++;
    }
  });

  const output = [
    ["Historical Ingestion Summary", ""],
    ["", ""],
    ["Total Historical Games", rows.length],
    ["", ""],
    ["Games by Season", ""],
    ["Season", "Games"]
  ];

  Object.keys(bySeason)
    .sort()
    .forEach(season => {
      output.push([season, bySeason[season].games]);
    });

  output.push(["", ""]);
  output.push(["Games by Team", ""]);
  output.push(["Team", "Games"]);

  Object.keys(byTeam)
    .sort()
    .forEach(team => {
      output.push([team, byTeam[team]]);
    });

  summarySheet
    .getRange(1, 1, output.length, 2)
    .setValues(output);

  formatHistoricalIngestionSummary(summarySheet);
}

function formatHistoricalIngestionSummary(sheet) {
  const lastRow = sheet.getLastRow();

  sheet.getRange(1, 1, 1, 2)
    .setFontWeight("bold")
    .setFontSize(14)
    .setBackground("#1f1f1f")
    .setFontColor("#ffffff");

  sheet.getRange(5, 1, 1, 2).setFontWeight("bold");
  sheet.getRange(6, 1, 1, 2)
    .setFontWeight("bold")
    .setBackground("#1f1f1f")
    .setFontColor("#ffffff");

  sheet.autoResizeColumns(1, 2);
  sheet.getRange(1, 1, lastRow, 2).setVerticalAlignment("middle");
}

function getTodayText() {
  return Utilities.formatDate(new Date(), "America/New_York", "yyyy-MM-dd");
}