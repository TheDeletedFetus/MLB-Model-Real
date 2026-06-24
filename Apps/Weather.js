function importWeather() {
  const scheduleSheet = getOrCreateSheet("RAW_Schedule");
  const weatherSheet = getOrCreateSheet("RAW_Weather");

  weatherSheet.clearContents();

  const scheduleData = scheduleSheet.getDataRange().getValues();
  if (scheduleData.length < 2) return;

  const headers = scheduleData[0];

  const gameIdCol = headers.indexOf("Game ID");
  const awayTeamCol = headers.indexOf("Away Team");
  const homeTeamCol = headers.indexOf("Home Team");
  const venueCol = headers.indexOf("Venue");

  const today = Utilities.formatDate(new Date(), "America/New_York", "yyyy-MM-dd");

  const rows = [[
    "Date",
    "Game ID",
    "Away Team",
    "Home Team",
    "Venue",
    "Temperature",
    "Condition",
    "Wind",
    "Last Updated",
    "Source"
  ]];

  for (let i = 1; i < scheduleData.length; i++) {
    const gameId = scheduleData[i][gameIdCol];

    if (!gameId) continue;

    const url =
  "https://statsapi.mlb.com/api/v1.1/game/" +
  gameId +
  "/feed/live";

    const data = MLB_API(url);
    const weather = data.gameData?.weather || {};

    rows.push([
      today,
      gameId,
      scheduleData[i][awayTeamCol],
      scheduleData[i][homeTeamCol],
      scheduleData[i][venueCol],
      weather.temp || "",
      weather.condition || "",
      weather.wind || "",
      today,
      "MLB Stats API gameData.weather"
    ]);
  }

  weatherSheet.getRange(1, 1, rows.length, rows[0].length).setValues(rows);
}