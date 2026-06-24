function importStarterRecentForm() {
  const scheduleSheet = getOrCreateSheet("RAW_Schedule");
  const outputSheet = getOrCreateSheet("RAW_Starter_Recent_Form");

  outputSheet.clearContents();

  const scheduleData = scheduleSheet.getDataRange().getValues();
  if (scheduleData.length < 2) return;

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
    "Throws",
    "Starts Found",
    "ERA Last 3 Starts",
    "WHIP Last 3 Starts",
    "K/BB Last 3 Starts",
    "Avg IP Last 3 Starts",
    "ERA Last 5 Starts",
    "WHIP Last 5 Starts",
    "K/BB Last 5 Starts",
    "Avg IP Last 5 Starts",
    "Last Updated"
  ]];

  const todayText = Utilities.formatDate(new Date(), "America/New_York", "yyyy-MM-dd");

  uniquePitcherIds.forEach(playerId => {
    const pitcher = getStarterRecentForm(playerId);

    rows.push([
      pitcher.name,
      playerId,
      pitcher.throws,
      pitcher.startsFound,
      pitcher.last3.era,
      pitcher.last3.whip,
      pitcher.last3.kbb,
      pitcher.last3.avgIp,
      pitcher.last5.era,
      pitcher.last5.whip,
      pitcher.last5.kbb,
      pitcher.last5.avgIp,
      todayText
    ]);
  });

  outputSheet.getRange(1, 1, rows.length, rows[0].length).setValues(rows);
}


function getStarterRecentForm(playerId) {
  const personUrl = "https://statsapi.mlb.com/api/v1/people/" + playerId;
  const personData = MLB_API(personUrl);
  const person = personData.people?.[0] || {};

  const today = new Date();
  const startDate = new Date(today);
  startDate.setDate(today.getDate() - 90);

  const start = Utilities.formatDate(startDate, "America/New_York", "yyyy-MM-dd");
  const end = Utilities.formatDate(today, "America/New_York", "yyyy-MM-dd");

  const statsUrl =
    "https://statsapi.mlb.com/api/v1/people/" +
    playerId +
    "/stats" +
    "?stats=gameLog" +
    "&group=pitching" +
    "&startDate=" + start +
    "&endDate=" + end;

  const statsData = MLB_API(statsUrl);
  const splits = statsData.stats?.[0]?.splits || [];

  const starts = splits
    .filter(row => Number(row.stat?.gamesStarted || 0) > 0)
    .sort((a, b) => new Date(b.date) - new Date(a.date));

  return {
    name: person.fullName || "",
    throws: person.pitchHand?.code || "",
    startsFound: starts.length,
    last3: calculateStarterWindow(starts.slice(0, 3)),
    last5: calculateStarterWindow(starts.slice(0, 5))
  };
}


function calculateStarterWindow(starts) {
  if (!starts || starts.length === 0) {
    return {
      era: "",
      whip: "",
      kbb: "",
      avgIp: ""
    };
  }

  let outs = 0;
  let earnedRuns = 0;
  let hits = 0;
  let walks = 0;
  let strikeouts = 0;

  starts.forEach(start => {
    const stat = start.stat || {};

    outs += starterRecentInningsToOuts(stat.inningsPitched || "0.0");
    earnedRuns += Number(stat.earnedRuns || 0);
    hits += Number(stat.hits || 0);
    walks += Number(stat.baseOnBalls || 0);
    strikeouts += Number(stat.strikeOuts || 0);
  });

  const innings = outs / 3;

  return {
    era: innings > 0 ? roundStarterRecent((earnedRuns * 9) / innings, 2) : "",
    whip: innings > 0 ? roundStarterRecent((hits + walks) / innings, 2) : "",
    kbb: walks > 0 ? roundStarterRecent(strikeouts / walks, 2) : strikeouts,
    avgIp: starts.length > 0 ? outsToInningsStarterRecent(Math.round(outs / starts.length)) : ""
  };
}


function starterRecentInningsToOuts(inningsText) {
  const parts = inningsText.toString().split(".");
  const fullInnings = Number(parts[0] || 0);
  const extraOuts = Number(parts[1] || 0);

  return fullInnings * 3 + extraOuts;
}


function outsToInningsStarterRecent(outs) {
  const fullInnings = Math.floor(outs / 3);
  const remainingOuts = outs % 3;

  return fullInnings + "." + remainingOuts;
}


function roundStarterRecent(value, decimals) {
  if (value === "" || isNaN(value)) return "";
  return Number(value.toFixed(decimals));
}