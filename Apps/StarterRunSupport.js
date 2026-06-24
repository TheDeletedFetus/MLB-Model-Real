function importStarterRunSupport() {
  const scheduleSheet = getOrCreateSheet("RAW_Schedule");
  const outputSheet = getOrCreateSheet("RAW_Starter_Run_Support");

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
  const todayText = Utilities.formatDate(new Date(), "America/New_York", "yyyy-MM-dd");

  const rows = [[
    "Player",
    "Player ID",
    "Throws",
    "Starts Found",
    "Run Support Season",
    "Run Support Last 3 Starts",
    "Run Support Last 5 Starts",
    "Team Runs In Starts",
    "Last Updated",
    "Source"
  ]];

  uniquePitcherIds.forEach(playerId => {
    const data = getStarterRunSupport(playerId);

    rows.push([
      data.name,
      playerId,
      data.throws,
      data.startsFound,
      data.seasonRunSupport,
      data.last3RunSupport,
      data.last5RunSupport,
      data.totalTeamRuns,
      todayText,
      "MLB Stats API game logs + live game linescore"
    ]);
  });

  outputSheet
    .getRange(1, 1, rows.length, rows[0].length)
    .setValues(rows);
}


function getStarterRunSupport(playerId) {
  const personUrl = "https://statsapi.mlb.com/api/v1/people/" + playerId;
  const personData = MLB_API(personUrl);
  const person = personData.people?.[0] || {};

  const today = new Date();
  const seasonStart = "2026-03-26";
  const todayText = Utilities.formatDate(today, "America/New_York", "yyyy-MM-dd");

  const statsUrl =
    "https://statsapi.mlb.com/api/v1/people/" +
    playerId +
    "/stats" +
    "?stats=gameLog" +
    "&group=pitching" +
    "&startDate=" + seasonStart +
    "&endDate=" + todayText;

  const statsData = MLB_API(statsUrl);
  const splits = statsData.stats?.[0]?.splits || [];

  const starts = splits
    .filter(row => Number(row.stat?.gamesStarted || 0) > 0)
    .sort((a, b) => new Date(b.date) - new Date(a.date));

  const startRows = [];

  starts.forEach(start => {
    const gamePk = start.game?.gamePk || start.game?.pk || "";
    const teamName = start.team?.name || "";

    if (!gamePk || !teamName) return;

    const teamRuns = getStarterTeamRunsForGame(gamePk, teamName);

    if (teamRuns === "") return;

    startRows.push({
      date: start.date,
      gamePk,
      teamName,
      teamRuns
    });
  });

  return {
    name: person.fullName || "",
    throws: person.pitchHand?.code || "",
    startsFound: startRows.length,
    seasonRunSupport: calculateStarterRunSupportAverage(startRows),
    last3RunSupport: calculateStarterRunSupportAverage(startRows.slice(0, 3)),
    last5RunSupport: calculateStarterRunSupportAverage(startRows.slice(0, 5)),
    totalTeamRuns: startRows.reduce((sum, row) => sum + Number(row.teamRuns || 0), 0)
  };
}


function getStarterTeamRunsForGame(gamePk, teamName) {
  const url =
    "https://statsapi.mlb.com/api/v1.1/game/" +
    gamePk +
    "/feed/live";

  const data = MLB_API(url);

  const awayTeam = data.gameData?.teams?.away?.name || "";
  const homeTeam = data.gameData?.teams?.home?.name || "";

  const awayRuns = data.liveData?.linescore?.teams?.away?.runs;
  const homeRuns = data.liveData?.linescore?.teams?.home?.runs;

  if (teamName === awayTeam && awayRuns !== undefined && awayRuns !== null) {
    return Number(awayRuns);
  }

  if (teamName === homeTeam && homeRuns !== undefined && homeRuns !== null) {
    return Number(homeRuns);
  }

  return "";
}


function calculateStarterRunSupportAverage(starts) {
  if (!starts || starts.length === 0) return "";

  const totalRuns = starts.reduce((sum, row) => {
    return sum + Number(row.teamRuns || 0);
  }, 0);

  return roundStarterRunSupport(totalRuns / starts.length, 2);
}


function roundStarterRunSupport(value, decimals) {
  if (value === "" || isNaN(value)) return "";
  return Number(value.toFixed(decimals));
}