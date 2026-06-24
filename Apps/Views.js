function buildTodayView() {
  const sheet = getOrCreateSheet("Today");
  sheet.clearContents();

  const rows = getSheetRows("Model_Matrix");

  const output = [[
    "Game",
    "Pick",
    "Model Strength",
    "Pick ML",
    "Market Implied %",
    "Best Factor",
    "Status"
  ]];

  rows.forEach(row => {
    const awayTeam = row["Away Team"];
    const homeTeam = row["Home Team"];
    const pick = row["Model Pick"];

    const awayML = row["Away ML"];
    const homeML = row["Home ML"];

    const pickML =
      pick === awayTeam ? awayML :
      pick === homeTeam ? homeML :
      "";

    const pickMarketProb =
      pick === awayTeam ? americanOddsToViewProbability(awayML) :
      pick === homeTeam ? americanOddsToViewProbability(homeML) :
      "";

    const bestEdge = findBestEdge(row);

    output.push([
      awayTeam + " @ " + homeTeam,
      pick,
      row["Confidence"],
      pickML,
      pickMarketProb === "" ? "" : Math.round(pickMarketProb * 1000) / 10 + "%",
      bestEdge.name,
      row["Status"]
    ]);
  });

  sheet.getRange(1, 1, output.length, output[0].length).setValues(output);
  formatTodayMobile(sheet, output.length, output[0].length);
}


function buildDashboard() {
  const sheet = getOrCreateSheet("DashBoard");
  sheet.clearContents();

  const historyRows = getSheetRows("HISTORY");
  const matrixRows = getSheetRows("Model_Matrix");

  const summary = buildPerformanceSummary(historyRows);

  const output = [
    ["Metric", "Value"],
    ["Total Games Tracked", summary.totalGames],
    ["Model Record", summary.wins + "-" + summary.losses],
    ["Win Rate", summary.winRate],
    ["Favorite Record", summary.favoriteWins + "-" + summary.favoriteLosses],
    ["Underdog Record", summary.underdogWins + "-" + summary.underdogLosses],
    ["Flat Bet Profit", summary.profit],
    ["ROI", summary.roi],
    ["", ""],
    ["Today's Ranked Games", ""],
    ["Rank", "Game", "Pick", "Strength", "Pick ML", "Best Factor", "Status"]
  ];

  const sorted = matrixRows
    .slice()
    .sort((a, b) => Number(b["Confidence"] || 0) - Number(a["Confidence"] || 0));

  sorted.forEach((row, index) => {
    const awayTeam = row["Away Team"];
    const homeTeam = row["Home Team"];
    const pick = row["Model Pick"];

    const pickML =
      pick === awayTeam ? row["Away ML"] :
      pick === homeTeam ? row["Home ML"] :
      "";

    const bestEdge = findBestEdge(row);

    output.push([
      index + 1,
      awayTeam + " @ " + homeTeam,
      pick,
      row["Confidence"],
      pickML,
      bestEdge.name,
      row["Status"]
    ]);
  });

  sheet.getRange(1, 1, output.length, 7).setValues(padRows(output, 7));
  formatDashboard(sheet);
}


function buildPerformanceSummary(rows) {
  let wins = 0;
  let losses = 0;
  let favoriteWins = 0;
  let favoriteLosses = 0;
  let underdogWins = 0;
  let underdogLosses = 0;
  let profit = 0;

  rows.forEach(row => {
    const correct = row["Game Winner Correct?"];
    const pickType = row["Model Pick Type"];
    const flatProfit = Number(row["Flat Bet Profit"]);

    if (correct === true || correct === "TRUE") {
      wins++;
      if (pickType === "Favorite") favoriteWins++;
      if (pickType === "Underdog") underdogWins++;
    } else if (correct === false || correct === "FALSE") {
      losses++;
      if (pickType === "Favorite") favoriteLosses++;
      if (pickType === "Underdog") underdogLosses++;
    }

    if (!isNaN(flatProfit)) {
      profit += flatProfit;
    }
  });

  const totalGames = wins + losses;
  const risked = totalGames * 100;

  return {
    totalGames,
    wins,
    losses,
    favoriteWins,
    favoriteLosses,
    underdogWins,
    underdogLosses,
    winRate: totalGames > 0 ? Math.round((wins / totalGames) * 1000) / 10 + "%" : "",
    profit: Math.round(profit * 100) / 100,
    roi: risked > 0 ? Math.round((profit / risked) * 1000) / 10 + "%" : ""
  };
}


function findBestEdge(row) {
  let best = {
    name: "",
    value: ""
  };

  Object.keys(row).forEach(key => {
    if (!key.includes("Edge")) return;

    const value = Number(row[key]);
    if (isNaN(value)) return;

    if (best.value === "" || Math.abs(value) > Math.abs(Number(best.value))) {
      best = {
        name: key,
        value: value
      };
    }
  });

  return best;
}


function americanOddsToViewProbability(odds) {
  const value = Number(odds);

  if (isNaN(value)) return "";

  if (value < 0) {
    return Math.abs(value) / (Math.abs(value) + 100);
  }

  return 100 / (value + 100);
}


function padRows(rows, width) {
  return rows.map(row => {
    const padded = row.slice();
    while (padded.length < width) {
      padded.push("");
    }
    return padded;
  });
}


function formatTodayMobile(sheet, rowCount, columnCount) {
  sheet.setFrozenRows(1);

  sheet.getRange(1, 1, 1, columnCount)
    .setFontWeight("bold")
    .setBackground("#1f1f1f")
    .setFontColor("#ffffff");

  sheet.getRange(1, 1, rowCount, columnCount)
    .setWrap(true)
    .setVerticalAlignment("middle");

  sheet.setColumnWidth(1, 160);
  sheet.setColumnWidth(2, 130);
  sheet.setColumnWidth(3, 90);
  sheet.setColumnWidth(4, 80);
  sheet.setColumnWidth(5, 100);
  sheet.setColumnWidth(6, 160);
  sheet.setColumnWidth(7, 90);

  if (rowCount > 1) {
    sheet.getRange(2, 1, rowCount - 1, columnCount).setFontSize(11);
    sheet.getRange(2, 3, rowCount - 1, 1).setHorizontalAlignment("center");
    sheet.getRange(2, 4, rowCount - 1, 2).setHorizontalAlignment("center");
  }
}


function formatDashboard(sheet) {
  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();

  if (lastRow < 1) return;

  sheet.getRange(1, 1, 1, 2)
    .setFontWeight("bold")
    .setBackground("#1f1f1f")
    .setFontColor("#ffffff");

  sheet.getRange(11, 1, 1, 7)
    .setFontWeight("bold")
    .setBackground("#1f1f1f")
    .setFontColor("#ffffff");

  sheet.getRange(1, 1, lastRow, lastCol)
    .setWrap(true)
    .setVerticalAlignment("middle");

  sheet.setFrozenRows(1);
  sheet.autoResizeColumns(1, lastCol);
}