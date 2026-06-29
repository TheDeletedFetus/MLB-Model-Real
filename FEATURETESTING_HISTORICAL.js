// FEATURETESTING.GS

function runFeatureTesting() {
  buildFeatureTestingReport();
}

function buildFeatureTestingReport() {
  const sheet = getOrCreateSheet("Feature_Testing");
  sheet.clearContents();

  const rows = getSheetRows("Historical_Backtest_Advanced");

  if (rows.length === 0) {
    sheet.getRange(1, 1).setValue(
      "No rows found in Historical_Backtest_Optimized. Run optimized historical backtest first."
    );
    return;
  }

const features = [
  "Win % Edge Before Date",
  "Runs/Game Edge Before Date",
  "Runs Allowed/Game Edge Before Date",
  "Run Differential Edge Before Date",
  "Home/Away Win % Edge Before Date",
  "Recent Runs/Game Edge",
  "Recent Runs Allowed/Game Edge",

  "Starter ERA Edge",
  "Starter WHIP Edge",
  "Bullpen ERA Edge",
  "Bullpen WHIP Edge",
  "OPS vs SP Hand Edge",
  "Team OPS Edge"
];
  const output = [
    ["Feature Testing Report", ""],
    ["Source", "Historical_Backtest_Optimized"],
    ["Method", "One feature at a time: positive edge picks Away, negative edge picks Home"],
    ["", ""],
    [
      "Feature",
      "Games Tested",
      "Wins",
      "Losses",
      "No Pick",
      "Win Rate",
      "Away Picks",
      "Home Picks",
      "Away Pick Win %",
      "Home Pick Win %",
      "Avg Abs Edge",
      "Interpretation"
    ]
  ];

  features.forEach(featureName => {
    output.push(calculateSingleFeaturePerformance(rows, featureName));
  });

  sheet
    .getRange(1, 1, output.length, output[4].length)
    .setValues(padRows(output, output[4].length));

  formatFeatureTestingReport(sheet);
}

function calculateSingleFeaturePerformance(rows, featureName) {
  let wins = 0;
  let losses = 0;
  let noPick = 0;

  let awayPicks = 0;
  let homePicks = 0;
  let awayWins = 0;
  let homeWins = 0;

  let edgeTotal = 0;
  let edgeCount = 0;

  rows.forEach(row => {
    const edgeValue = parseFeatureTestingNumber(row[featureName]);

    const awayTeam = row["Away Team"];
    const homeTeam = row["Home Team"];
    const winner = row["Winner"];

    if (edgeValue === "" || !awayTeam || !homeTeam || !winner) {
      noPick++;
      return;
    }

    if (edgeValue === 0) {
      noPick++;
      return;
    }

    const pick = edgeValue > 0 ? awayTeam : homeTeam;
    const correct = pick === winner;

    edgeTotal += Math.abs(edgeValue);
    edgeCount++;

    if (pick === awayTeam) {
      awayPicks++;
      if (correct) awayWins++;
    }

    if (pick === homeTeam) {
      homePicks++;
      if (correct) homeWins++;
    }

    if (correct) {
      wins++;
    } else {
      losses++;
    }
  });

  const gamesTested = wins + losses;
  const winRate = gamesTested > 0 ? formatFeatureTestingPercent(wins / gamesTested) : "";
  const awayWinRate = awayPicks > 0 ? formatFeatureTestingPercent(awayWins / awayPicks) : "";
  const homeWinRate = homePicks > 0 ? formatFeatureTestingPercent(homeWins / homePicks) : "";
  const avgAbsEdge = edgeCount > 0 ? roundFeatureTesting(edgeTotal / edgeCount, 4) : "";

  return [
    featureName,
    gamesTested,
    wins,
    losses,
    noPick,
    winRate,
    awayPicks,
    homePicks,
    awayWinRate,
    homeWinRate,
    avgAbsEdge,
    interpretFeatureTestingResult(gamesTested, wins)
  ];
}

function interpretFeatureTestingResult(games, wins) {
  if (!games || games === 0) return "No usable sample";

  const winRate = wins / games;

  if (games < 100) return "Sample too small";
  if (winRate >= 0.55) return "Strong positive signal";
  if (winRate >= 0.525) return "Possible positive signal";
  if (winRate >= 0.49 && winRate <= 0.51) return "Likely noise";
  if (winRate < 0.49) return "Potentially harmful";
  return "Weak signal";
}

function parseFeatureTestingNumber(value) {
  if (value === "" || value === undefined || value === null) return "";

  const num = Number(value);
  return isNaN(num) ? "" : num;
}

function formatFeatureTestingPercent(value) {
  return Math.round(value * 1000) / 10 + "%";
}

function roundFeatureTesting(value, decimals) {
  if (value === "" || isNaN(value)) return "";
  return Number(value.toFixed(decimals));
}

function formatFeatureTestingReport(sheet) {
  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();

  if (lastRow < 1) return;

  sheet.getRange(1, 1, 1, lastCol)
    .setFontWeight("bold")
    .setFontSize(14)
    .setBackground("#1f1f1f")
    .setFontColor("#ffffff");

  sheet.getRange(5, 1, 1, lastCol)
    .setFontWeight("bold")
    .setBackground("#1f1f1f")
    .setFontColor("#ffffff");

  sheet.getRange(1, 1, lastRow, lastCol)
    .setWrap(true)
    .setVerticalAlignment("middle");

  sheet.setFrozenRows(5);
  sheet.autoResizeColumns(1, lastCol);
}