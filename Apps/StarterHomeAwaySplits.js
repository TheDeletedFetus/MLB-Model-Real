function importStarterHomeAwaySplits() {
  const scheduleSheet = getOrCreateSheet("RAW_Schedule");
  const outputSheet = getOrCreateSheet("RAW_Starter_Home_Away_Splits");

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
    "Home ERA",
    "Road ERA",
    "Home WHIP",
    "Road WHIP",
    "Home K/BB",
    "Road K/BB",
    "Home K/9",
    "Road K/9",
    "Home BB/9",
    "Road BB/9",
    "Last Updated",
    "Source"
  ]];

  uniquePitcherIds.forEach(playerId => {
    const data = getStarterHomeAwayData(playerId);

    rows.push([
      data.name,
      playerId,
      data.throws,
      data.home.era,
      data.road.era,
      data.home.whip,
      data.road.whip,
      data.home.kbb,
      data.road.kbb,
      data.home.k9,
      data.road.k9,
      data.home.bb9,
      data.road.bb9,
      todayText,
      "MLB Stats API home/away player pitching splits"
    ]);
  });

  outputSheet.getRange(1, 1, rows.length, rows[0].length).setValues(rows);
}


function getStarterHomeAwayData(playerId) {
  const personUrl = "https://statsapi.mlb.com/api/v1/people/" + playerId;
  const personData = MLB_API(personUrl);
  const person = personData.people?.[0] || {};

  const statsUrl =
    "https://statsapi.mlb.com/api/v1/people/" +
    playerId +
    "/stats" +
    "?stats=homeAndAway" +
    "&group=pitching";

  const statsData = MLB_API(statsUrl);
  const splits = statsData.stats?.[0]?.splits || [];

  const output = {
    name: person.fullName || "",
    throws: person.pitchHand?.code || "",
    home: blankStarterHomeAwayBucket(),
    road: blankStarterHomeAwayBucket()
  };

  splits.forEach(split => {
    const stat = split.stat || {};

    const bucket = {
      era: stat.era || "",
      whip: stat.whip || "",
      kbb: stat.strikeoutWalkRatio || "",
      k9: stat.strikeoutsPer9Inn || "",
      bb9: stat.walksPer9Inn || ""
    };

    if (split.isHome === true) {
      output.home = bucket;
    }

    if (split.isHome === false) {
      output.road = bucket;
    }
  });

  return output;
}


function blankStarterHomeAwayBucket() {
  return {
    era: "",
    whip: "",
    kbb: "",
    k9: "",
    bb9: ""
  };
}