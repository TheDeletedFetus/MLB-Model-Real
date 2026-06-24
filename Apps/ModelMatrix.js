function buildModelMatrix() {
  const sheet = getOrCreateSheet("MODEL_MATRIX");
  sheet.clearContents();

  const schedule = getSheetRows("RAW_Schedule");
  if (schedule.length === 0) return;

  const teamHitting = sheetToLookup("RAW_Team_Hitting", "Team");
  const teamPitching = sheetToLookup("RAW_Team_Pitching", "Team");
  const teamQuality = sheetToLookup("RAW_Team_Quality", "Team");
  const teamSplits = sheetToLookup("RAW_Team_Splits", "Team");
  const teamHomeAway = sheetToLookup("RAW_Team_Home_Away_Splits", "Team");
  const recentForm = sheetToLookup("RAW_Recent_Form", "Team");

  const starters = sheetToLookup("RAW_Starting_Pitchers", "Player ID");
  const starterRecent = sheetToLookup("RAW_Starter_Recent_Form", "Player ID");
  const starterRunSupport = sheetToLookup("RAW_Starter_Run_Support", "Player ID");
  const starterHomeAway = sheetToLookup("RAW_Starter_Home_Away_Splits", "Player ID");

  const bullpenQuality = sheetToLookup("RAW_Bullpen_Quality", "Team");
  const bullpenUsage = sheetToLookup("RAW_Bullpen", "Team");
  const bullpenFatigue = sheetToLookup("RAW_Bullpen_Fatigue", "Team");

  const parkFactors = sheetToLookup("RAW_Park_Factors", "Venue");
  const weather = sheetToLookup("RAW_Weather", "Game ID");
  const odds = getLatestOddsByMatchup();

  const headers = [
    "Date", "Game ID", "Away Team", "Home Team", "Away SP", "Home SP",
    "Away SP ID", "Home SP ID", "Away SP Hand", "Home SP Hand", "Venue", "Status",

    "Away Runs/Game", "Home Runs/Game", "Runs/Game Edge",
    "Away Run Differential", "Home Run Differential", "Run Differential Edge",
    "Away OPS", "Home OPS", "OPS Edge",
    "Away OBP", "Home OBP", "OBP Edge",
    "Away SLG", "Home SLG", "SLG Edge",
    "Away BABIP", "Home BABIP", "BABIP Edge",

    "Away OPS vs SP Hand", "Home OPS vs SP Hand", "OPS vs Hand Edge",
    "Away OPS vs RHP", "Home OPS vs RHP",
    "Away OPS vs LHP", "Home OPS vs LHP",

    "Away Team ERA", "Home Team ERA", "Team ERA Edge",
    "Away Team WHIP", "Home Team WHIP", "Team WHIP Edge",
    "Away Team K/9", "Home Team K/9", "Team K/9 Edge",
    "Away Team BB/9", "Home Team BB/9", "Team BB/9 Edge",
    "Away Team HR/9", "Home Team HR/9", "Team HR/9 Edge",

    "Away Road OPS", "Home Home OPS", "Split OPS Edge",
    "Away Road ERA", "Home Home ERA", "Split ERA Edge",
    "Away Road WHIP", "Home Home WHIP", "Split WHIP Edge",

    "Away Starter ERA", "Home Starter ERA", "Starter ERA Edge",
    "Away Starter WHIP", "Home Starter WHIP", "Starter WHIP Edge",
    "Away Starter K/BB", "Home Starter K/BB", "Starter K/BB Edge",
    "Away Starter K/9", "Home Starter K/9", "Starter K/9 Edge",
    "Away Starter BB/9", "Home Starter BB/9", "Starter BB/9 Edge",
    "Away Starter HR/9", "Home Starter HR/9", "Starter HR/9 Edge",

    "Away Starter ERA Last 3", "Home Starter ERA Last 3", "Starter ERA Last 3 Edge",
    "Away Starter WHIP Last 3", "Home Starter WHIP Last 3", "Starter WHIP Last 3 Edge",
    "Away Starter K/BB Last 3", "Home Starter K/BB Last 3", "Starter K/BB Last 3 Edge",
    "Away Starter Avg IP Last 3", "Home Starter Avg IP Last 3", "Starter Avg IP Last 3 Edge",

    "Away Starter ERA Last 5", "Home Starter ERA Last 5", "Starter ERA Last 5 Edge",
    "Away Starter WHIP Last 5", "Home Starter WHIP Last 5", "Starter WHIP Last 5 Edge",
    "Away Starter K/BB Last 5", "Home Starter K/BB Last 5", "Starter K/BB Last 5 Edge",
    "Away Starter Avg IP Last 5", "Home Starter Avg IP Last 5", "Starter Avg IP Last 5 Edge",

    "Away Starter Run Support", "Home Starter Run Support", "Starter Run Support Edge",
    "Away Starter Run Support Last 3", "Home Starter Run Support Last 3", "Starter Run Support Last 3 Edge",
    "Away Starter Run Support Last 5", "Home Starter Run Support Last 5", "Starter Run Support Last 5 Edge",

    "Away Road Starter ERA", "Home Home Starter ERA", "Starter Split ERA Edge",
    "Away Road Starter WHIP", "Home Home Starter WHIP", "Starter Split WHIP Edge",
    "Away Road Starter K/BB", "Home Home Starter K/BB", "Starter Split K/BB Edge",

    "Away Bullpen ERA", "Home Bullpen ERA", "Bullpen ERA Edge",
    "Away Bullpen WHIP", "Home Bullpen WHIP", "Bullpen WHIP Edge",
    "Away Bullpen K/9", "Home Bullpen K/9", "Bullpen K/9 Edge",
    "Away Bullpen BB/9", "Home Bullpen BB/9", "Bullpen BB/9 Edge",
    "Away Bullpen HR/9", "Home Bullpen HR/9", "Bullpen HR/9 Edge",

    "Away Bullpen Save %", "Home Bullpen Save %", "Bullpen Save % Edge",
    "Away Bullpen Blown Saves", "Home Bullpen Blown Saves", "Bullpen Blown Saves Edge",
    "Away Bullpen Inherited Scored %", "Home Bullpen Inherited Scored %", "Bullpen Inherited Scored % Edge",
    "Away Bullpen Avg Leverage Index", "Home Bullpen Avg Leverage Index", "Bullpen Avg Leverage Index Edge",

    "Away Bullpen IP Last 3", "Home Bullpen IP Last 3", "Bullpen IP Last 3 Edge",
    "Away Bullpen IP Last 7", "Home Bullpen IP Last 7", "Bullpen IP Last 7 Edge",
    "Away Bullpen Pitches Last 3", "Home Bullpen Pitches Last 3", "Bullpen Pitches Last 3 Edge",
    "Away Bullpen Pitches Last 7", "Home Bullpen Pitches Last 7", "Bullpen Pitches Last 7 Edge",

    "Away Team OPS Last 7", "Home Team OPS Last 7", "Team OPS Last 7 Edge",
    "Away Team OPS Last 14", "Home Team OPS Last 14", "Team OPS Last 14 Edge",
    "Away Team OPS Last 30", "Home Team OPS Last 30", "Team OPS Last 30 Edge",
    "Away Team ERA Last 7", "Home Team ERA Last 7", "Team ERA Last 7 Edge",
    "Away Team ERA Last 14", "Home Team ERA Last 14", "Team ERA Last 14 Edge",
    "Away Team ERA Last 30", "Home Team ERA Last 30", "Team ERA Last 30 Edge",

    "Park Factor Runs", "Park Factor HR", "Park Factor Hits",
    "Temperature", "Weather Condition", "Wind",

    "Away ML", "Home ML", "Away Implied Probability", "Home Implied Probability",

    "Away Model Score", "Home Model Score", "Model Pick", "Confidence"
  ];

  const output = [headers];

  schedule.forEach(game => {
    const awayTeam = game["Away Team"];
    const homeTeam = game["Home Team"];
    const awaySpId = game["Away SP ID"];
    const homeSpId = game["Home SP ID"];
    const venue = game["Venue"];
    const gameId = game["Game ID"];

    const awayHit = teamHitting[awayTeam] || {};
    const homeHit = teamHitting[homeTeam] || {};
    const awayPitch = teamPitching[awayTeam] || {};
    const homePitch = teamPitching[homeTeam] || {};
    const awayQuality = teamQuality[awayTeam] || {};
    const homeQuality = teamQuality[homeTeam] || {};
    const awaySplits = teamSplits[awayTeam] || {};
    const homeSplits = teamSplits[homeTeam] || {};
    const awayHA = teamHomeAway[awayTeam] || {};
    const homeHA = teamHomeAway[homeTeam] || {};
    const awayRecent = recentForm[awayTeam] || {};
    const homeRecent = recentForm[homeTeam] || {};

    const awayStarter = starters[awaySpId] || {};
    const homeStarter = starters[homeSpId] || {};
    const awayStarterRecent = starterRecent[awaySpId] || {};
    const homeStarterRecent = starterRecent[homeSpId] || {};
    const awayStarterRunSupport = starterRunSupport[awaySpId] || {};
    const homeStarterRunSupport = starterRunSupport[homeSpId] || {};
    const awayStarterHA = starterHomeAway[awaySpId] || {};
    const homeStarterHA = starterHomeAway[homeSpId] || {};

    const awayBPQ = bullpenQuality[awayTeam] || {};
    const homeBPQ = bullpenQuality[homeTeam] || {};
    const awayBPU = bullpenUsage[awayTeam] || {};
    const homeBPU = bullpenUsage[homeTeam] || {};
    const awayBPF = bullpenFatigue[awayTeam] || {};
    const homeBPF = bullpenFatigue[homeTeam] || {};

    const park = parkFactors[venue] || {};
    const wx = weather[gameId] || {};
    const market = odds[awayTeam + "|" + homeTeam] || {};

    const awayOpsVsHand = getOpsVsHand(awaySplits, game["Home SP Hand"]);
    const homeOpsVsHand = getOpsVsHand(homeSplits, game["Away SP Hand"]);

    output.push([
      game["Date"], gameId, awayTeam, homeTeam, game["Away SP"], game["Home SP"],
      awaySpId, homeSpId, game["Away SP Hand"], game["Home SP Hand"], venue, game["Status"],

      awayHit["Runs/Game"], homeHit["Runs/Game"], edge(awayHit["Runs/Game"], homeHit["Runs/Game"], "higher"),
      awayQuality["Run Differential"], homeQuality["Run Differential"], edge(awayQuality["Run Differential"], homeQuality["Run Differential"], "higher"),
      awayHit["OPS"], homeHit["OPS"], edge(awayHit["OPS"], homeHit["OPS"], "higher"),
      awayHit["OBP"], homeHit["OBP"], edge(awayHit["OBP"], homeHit["OBP"], "higher"),
      awayHit["SLG"], homeHit["SLG"], edge(awayHit["SLG"], homeHit["SLG"], "higher"),
      awayHit["BABIP"], homeHit["BABIP"], edge(awayHit["BABIP"], homeHit["BABIP"], "higher"),

      awayOpsVsHand, homeOpsVsHand, edge(awayOpsVsHand, homeOpsVsHand, "higher"),
      awaySplits["OPS vs RHP"], homeSplits["OPS vs RHP"],
      awaySplits["OPS vs LHP"], homeSplits["OPS vs LHP"],

      awayPitch["ERA"], homePitch["ERA"], edge(awayPitch["ERA"], homePitch["ERA"], "lower"),
      awayPitch["WHIP"], homePitch["WHIP"], edge(awayPitch["WHIP"], homePitch["WHIP"], "lower"),
      awayPitch["K/9"], homePitch["K/9"], edge(awayPitch["K/9"], homePitch["K/9"], "higher"),
      awayPitch["BB/9"], homePitch["BB/9"], edge(awayPitch["BB/9"], homePitch["BB/9"], "lower"),
      awayPitch["HR/9"], homePitch["HR/9"], edge(awayPitch["HR/9"], homePitch["HR/9"], "lower"),

      awayHA["Road OPS"], homeHA["Home OPS"], edge(awayHA["Road OPS"], homeHA["Home OPS"], "higher"),
      awayHA["Road ERA"], homeHA["Home ERA"], edge(awayHA["Road ERA"], homeHA["Home ERA"], "lower"),
      awayHA["Road WHIP"], homeHA["Home WHIP"], edge(awayHA["Road WHIP"], homeHA["Home WHIP"], "lower"),

      awayStarter["ERA"], homeStarter["ERA"], edge(awayStarter["ERA"], homeStarter["ERA"], "lower"),
      awayStarter["WHIP"], homeStarter["WHIP"], edge(awayStarter["WHIP"], homeStarter["WHIP"], "lower"),
      awayStarter["K/BB"], homeStarter["K/BB"], edge(awayStarter["K/BB"], homeStarter["K/BB"], "higher"),
      awayStarter["K/9"], homeStarter["K/9"], edge(awayStarter["K/9"], homeStarter["K/9"], "higher"),
      awayStarter["BB/9"], homeStarter["BB/9"], edge(awayStarter["BB/9"], homeStarter["BB/9"], "lower"),
      awayStarter["HR/9"], homeStarter["HR/9"], edge(awayStarter["HR/9"], homeStarter["HR/9"], "lower"),

      awayStarterRecent["ERA Last 3 Starts"], homeStarterRecent["ERA Last 3 Starts"], edge(awayStarterRecent["ERA Last 3 Starts"], homeStarterRecent["ERA Last 3 Starts"], "lower"),
      awayStarterRecent["WHIP Last 3 Starts"], homeStarterRecent["WHIP Last 3 Starts"], edge(awayStarterRecent["WHIP Last 3 Starts"], homeStarterRecent["WHIP Last 3 Starts"], "lower"),
      awayStarterRecent["K/BB Last 3 Starts"], homeStarterRecent["K/BB Last 3 Starts"], edge(awayStarterRecent["K/BB Last 3 Starts"], homeStarterRecent["K/BB Last 3 Starts"], "higher"),
      awayStarterRecent["Avg IP Last 3 Starts"], homeStarterRecent["Avg IP Last 3 Starts"], edgeInnings(awayStarterRecent["Avg IP Last 3 Starts"], homeStarterRecent["Avg IP Last 3 Starts"], "higher"),

      awayStarterRecent["ERA Last 5 Starts"], homeStarterRecent["ERA Last 5 Starts"], edge(awayStarterRecent["ERA Last 5 Starts"], homeStarterRecent["ERA Last 5 Starts"], "lower"),
      awayStarterRecent["WHIP Last 5 Starts"], homeStarterRecent["WHIP Last 5 Starts"], edge(awayStarterRecent["WHIP Last 5 Starts"], homeStarterRecent["WHIP Last 5 Starts"], "lower"),
      awayStarterRecent["K/BB Last 5 Starts"], homeStarterRecent["K/BB Last 5 Starts"], edge(awayStarterRecent["K/BB Last 5 Starts"], homeStarterRecent["K/BB Last 5 Starts"], "higher"),
      awayStarterRecent["Avg IP Last 5 Starts"], homeStarterRecent["Avg IP Last 5 Starts"], edgeInnings(awayStarterRecent["Avg IP Last 5 Starts"], homeStarterRecent["Avg IP Last 5 Starts"], "higher"),

      awayStarterRunSupport["Run Support Season"], homeStarterRunSupport["Run Support Season"], edge(awayStarterRunSupport["Run Support Season"], homeStarterRunSupport["Run Support Season"], "higher"),
      awayStarterRunSupport["Run Support Last 3 Starts"], homeStarterRunSupport["Run Support Last 3 Starts"], edge(awayStarterRunSupport["Run Support Last 3 Starts"], homeStarterRunSupport["Run Support Last 3 Starts"], "higher"),
      awayStarterRunSupport["Run Support Last 5 Starts"], homeStarterRunSupport["Run Support Last 5 Starts"], edge(awayStarterRunSupport["Run Support Last 5 Starts"], homeStarterRunSupport["Run Support Last 5 Starts"], "higher"),

      awayStarterHA["Road ERA"], homeStarterHA["Home ERA"], edge(awayStarterHA["Road ERA"], homeStarterHA["Home ERA"], "lower"),
      awayStarterHA["Road WHIP"], homeStarterHA["Home WHIP"], edge(awayStarterHA["Road WHIP"], homeStarterHA["Home WHIP"], "lower"),
      awayStarterHA["Road K/BB"], homeStarterHA["Home K/BB"], edge(awayStarterHA["Road K/BB"], homeStarterHA["Home K/BB"], "higher"),

      awayBPQ["Bullpen ERA"], homeBPQ["Bullpen ERA"], edge(awayBPQ["Bullpen ERA"], homeBPQ["Bullpen ERA"], "lower"),
      awayBPQ["Bullpen WHIP"], homeBPQ["Bullpen WHIP"], edge(awayBPQ["Bullpen WHIP"], homeBPQ["Bullpen WHIP"], "lower"),
      awayBPQ["Bullpen K/9"], homeBPQ["Bullpen K/9"], edge(awayBPQ["Bullpen K/9"], homeBPQ["Bullpen K/9"], "higher"),
      awayBPQ["Bullpen BB/9"], homeBPQ["Bullpen BB/9"], edge(awayBPQ["Bullpen BB/9"], homeBPQ["Bullpen BB/9"], "lower"),
      awayBPQ["Bullpen HR/9"], homeBPQ["Bullpen HR/9"], edge(awayBPQ["Bullpen HR/9"], homeBPQ["Bullpen HR/9"], "lower"),

      awayBPU["Save %"], homeBPU["Save %"], edgePercent(awayBPU["Save %"], homeBPU["Save %"], "higher"),
      awayBPU["Blown Saves"], homeBPU["Blown Saves"], edge(awayBPU["Blown Saves"], homeBPU["Blown Saves"], "lower"),
      awayBPU["Inherited Scored %"], homeBPU["Inherited Scored %"], edgePercent(awayBPU["Inherited Scored %"], homeBPU["Inherited Scored %"], "lower"),
      awayBPU["Avg Leverage Index"], homeBPU["Avg Leverage Index"], edge(awayBPU["Avg Leverage Index"], homeBPU["Avg Leverage Index"], "higher"),

      awayBPF["Bullpen IP Last 3 Days"], homeBPF["Bullpen IP Last 3 Days"], edgeInnings(awayBPF["Bullpen IP Last 3 Days"], homeBPF["Bullpen IP Last 3 Days"], "lower"),
      awayBPF["Bullpen IP Last 7 Days"], homeBPF["Bullpen IP Last 7 Days"], edgeInnings(awayBPF["Bullpen IP Last 7 Days"], homeBPF["Bullpen IP Last 7 Days"], "lower"),
      awayBPF["Bullpen Pitches Last 3 Days"], homeBPF["Bullpen Pitches Last 3 Days"], edge(awayBPF["Bullpen Pitches Last 3 Days"], homeBPF["Bullpen Pitches Last 3 Days"], "lower"),
      awayBPF["Bullpen Pitches Last 7 Days"], homeBPF["Bullpen Pitches Last 7 Days"], edge(awayBPF["Bullpen Pitches Last 7 Days"], homeBPF["Bullpen Pitches Last 7 Days"], "lower"),

      awayRecent["OPS Last 7"], homeRecent["OPS Last 7"], edge(awayRecent["OPS Last 7"], homeRecent["OPS Last 7"], "higher"),
      awayRecent["OPS Last 14"], homeRecent["OPS Last 14"], edge(awayRecent["OPS Last 14"], homeRecent["OPS Last 14"], "higher"),
      awayRecent["OPS Last 30"], homeRecent["OPS Last 30"], edge(awayRecent["OPS Last 30"], homeRecent["OPS Last 30"], "higher"),
      awayRecent["ERA Last 7"], homeRecent["ERA Last 7"], edge(awayRecent["ERA Last 7"], homeRecent["ERA Last 7"], "lower"),
      awayRecent["ERA Last 14"], homeRecent["ERA Last 14"], edge(awayRecent["ERA Last 14"], homeRecent["ERA Last 14"], "lower"),
      awayRecent["ERA Last 30"], homeRecent["ERA Last 30"], edge(awayRecent["ERA Last 30"], homeRecent["ERA Last 30"], "lower"),

      park["Run Factor"], park["HR Factor"], park["Hit Factor"],
      wx["Temperature"], wx["Condition"], wx["Wind"],

      market.awayML || "", market.homeML || "",
      americanOddsToProbability(market.awayML),
      americanOddsToProbability(market.homeML),

      "", "", "", ""
    ]);
  });

  sheet.getRange(1, 1, output.length, output[0].length).setValues(output);
}


function getSheetRows(sheetName) {
  const sheet = getOrCreateSheet(sheetName);
  const values = sheet.getDataRange().getValues();
  if (values.length < 2) return [];

  const headers = values[0];

  return values.slice(1).map(row => {
    const obj = {};
    headers.forEach((header, index) => {
      obj[header] = row[index];
    });
    return obj;
  });
}


function sheetToLookup(sheetName, keyColumn) {
  const rows = getSheetRows(sheetName);
  const lookup = {};

  rows.forEach(row => {
    const key = row[keyColumn];
    if (key !== "" && key !== undefined && key !== null) {
      lookup[key] = row;
    }
  });

  return lookup;
}


function getLatestOddsByMatchup() {
  const rows = getSheetRows("RAW_Odds");
  const lookup = {};

  rows.forEach(row => {
    const key = row["Away Team"] + "|" + row["Home Team"];
    lookup[key] = {
      awayML: row["Away ML"],
      homeML: row["Home ML"]
    };
  });

  return lookup;
}


function getOpsVsHand(teamSplitRow, opponentStarterHand) {
  if (opponentStarterHand === "R") return teamSplitRow["OPS vs RHP"] || "";
  if (opponentStarterHand === "L") return teamSplitRow["OPS vs LHP"] || "";
  return "";
}


function edge(awayValue, homeValue, direction) {
  const away = parseNumber(awayValue);
  const home = parseNumber(homeValue);

  if (away === "" || home === "") return "";

  if (direction === "lower") {
    return roundMatrix(home - away, 4);
  }

  return roundMatrix(away - home, 4);
}


function edgePercent(awayValue, homeValue, direction) {
  const away = parseNumber(String(awayValue).replace("%", ""));
  const home = parseNumber(String(homeValue).replace("%", ""));

  if (away === "" || home === "") return "";

  if (direction === "lower") {
    return roundMatrix(home - away, 4);
  }

  return roundMatrix(away - home, 4);
}


function edgeInnings(awayValue, homeValue, direction) {
  const away = inningsStringToDecimal(awayValue);
  const home = inningsStringToDecimal(homeValue);

  if (away === "" || home === "") return "";

  if (direction === "lower") {
    return roundMatrix(home - away, 4);
  }

  return roundMatrix(away - home, 4);
}


function inningsStringToDecimal(value) {
  if (value === "" || value === undefined || value === null) return "";

  const parts = value.toString().split(".");
  const full = Number(parts[0] || 0);
  const outs = Number(parts[1] || 0);

  return full + outs / 3;
}


function parseNumber(value) {
  if (value === "" || value === undefined || value === null) return "";

  const cleaned = value.toString().replace("%", "").trim();
  const num = Number(cleaned);

  return isNaN(num) ? "" : num;
}


function roundMatrix(value, decimals) {
  if (value === "" || isNaN(value)) return "";
  return Number(value.toFixed(decimals));
}


function americanOddsToProbability(odds) {
  const value = parseNumber(odds);
  if (value === "") return "";

  if (value < 0) {
    return roundMatrix(Math.abs(value) / (Math.abs(value) + 100), 4);
  }

  return roundMatrix(100 / (value + 100), 4);
}