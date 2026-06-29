function buildPerformanceDashboard() {
  const sheet = getOrCreateSheet("Performance_Dashboard");
  sheet.clearContents();

  const rows = getSheetRows("HISTORY");
  const summary = calculatePerformanceDashboardSummary(rows);
  const factorRows = calculateFactorPerformance(rows);

  const output = [
    ["MLB Model Performance Dashboard", ""],
    ["", ""],

    ["Overall Metrics", ""],
    ["Total Games", summary.totalGames],
    ["Model Record", summary.wins + "-" + summary.losses],
    ["Win Rate", summary.winRate],
    ["Flat Bet Profit", summary.profit],
    ["ROI", summary.roi],
    ["Average Pick Odds", summary.averagePickOdds],
    ["", ""],

    ["Favorite / Underdog Split", ""],
    ["Favorite Record", summary.favoriteWins + "-" + summary.favoriteLosses],
    ["Favorite Win Rate", summary.favoriteWinRate],
    ["Favorite Profit", summary.favoriteProfit],
    ["Underdog Record", summary.underdogWins + "-" + summary.underdogLosses],
    ["Underdog Win Rate", summary.underdogWinRate],
    ["Underdog Profit", summary.underdogProfit],
    ["", ""],

    ["Best Factor Performance", ""],
    ["Factor", "Games", "Wins", "Losses", "Win Rate", "Profit"]
  ];

  factorRows.forEach(row => {
    output.push([
      row.factor,
      row.games,
      row.wins,
      row.losses,
      row.winRate,
      row.profit
    ]);
  });

  sheet.getRange(1, 1, output.length, 6).setValues(padRows(output, 6));
  formatPerformanceDashboard(sheet);
}


function calculatePerformanceDashboardSummary(rows) {
  let wins = 0;
  let losses = 0;
  let profit = 0;
  let totalOdds = 0;
  let oddsCount = 0;

  let favoriteWins = 0;
  let favoriteLosses = 0;
  let favoriteProfit = 0;

  let underdogWins = 0;
  let underdogLosses = 0;
  let underdogProfit = 0;

  rows.forEach(row => {
    const correct = normalizeBoolean(row["Game Winner Correct?"]);
    const pickType = row["Model Pick Type"];
    const flatProfit = Number(row["Flat Bet Profit"]);
    const pickML = Number(row["Model Pick ML"]);

    if (correct === null) return;

    if (correct) {
      wins++;
    } else {
      losses++;
    }

    if (!isNaN(flatProfit)) {
      profit += flatProfit;
    }

    if (!isNaN(pickML)) {
      totalOdds += pickML;
      oddsCount++;
    }

    if (pickType === "Favorite") {
      if (correct) favoriteWins++;
      else favoriteLosses++;

      if (!isNaN(flatProfit)) favoriteProfit += flatProfit;
    }

    if (pickType === "Underdog") {
      if (correct) underdogWins++;
      else underdogLosses++;

      if (!isNaN(flatProfit)) underdogProfit += flatProfit;
    }
  });

  const totalGames = wins + losses;
  const favoriteGames = favoriteWins + favoriteLosses;
  const underdogGames = underdogWins + underdogLosses;

  return {
    totalGames,
    wins,
    losses,
    winRate: formatPercent(wins, totalGames),
    profit: roundMoney(profit),
    roi: totalGames > 0 ? roundPercent(profit / (totalGames * 100)) : "",

    averagePickOdds: oddsCount > 0 ? Math.round(totalOdds / oddsCount) : "",

    favoriteWins,
    favoriteLosses,
    favoriteWinRate: formatPercent(favoriteWins, favoriteGames),
    favoriteProfit: roundMoney(favoriteProfit),

    underdogWins,
    underdogLosses,
    underdogWinRate: formatPercent(underdogWins, underdogGames),
    underdogProfit: roundMoney(underdogProfit)
  };
}


function calculateFactorPerformance(rows) {
  const factorMap = {};

  rows.forEach(row => {
    const factor = row["Largest Factor"];
    const correct = normalizeBoolean(row["Game Winner Correct?"]);
    const flatProfit = Number(row["Flat Bet Profit"]);

    if (!factor || correct === null) return;

    if (!factorMap[factor]) {
      factorMap[factor] = {
        factor,
        games: 0,
        wins: 0,
        losses: 0,
        profit: 0
      };
    }

    factorMap[factor].games++;

    if (correct) {
      factorMap[factor].wins++;
    } else {
      factorMap[factor].losses++;
    }

    if (!isNaN(flatProfit)) {
      factorMap[factor].profit += flatProfit;
    }
  });

  return Object.values(factorMap)
    .map(item => {
      item.winRate = formatPercent(item.wins, item.games);
      item.profit = roundMoney(item.profit);
      return item;
    })
    .sort((a, b) => b.games - a.games);
}


function normalizeBoolean(value) {
  if (value === true || value === "TRUE" || value === "Yes") return true;
  if (value === false || value === "FALSE" || value === "No") return false;
  return null;
}


function formatPercent(numerator, denominator) {
  if (!denominator || denominator === 0) return "";
  return Math.round((numerator / denominator) * 1000) / 10 + "%";
}


function roundPercent(value) {
  return Math.round(value * 1000) / 10 + "%";
}


function roundMoney(value) {
  return Math.round(value * 100) / 100;
}


function formatPerformanceDashboard(sheet) {
  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();

  if (lastRow < 1) return;

  sheet.getRange(1, 1, 1, lastCol)
    .setFontWeight("bold")
    .setFontSize(14)
    .setBackground("#1f1f1f")
    .setFontColor("#ffffff");

  sheet.getRange(3, 1, 1, lastCol).setFontWeight("bold");
  sheet.getRange(11, 1, 1, lastCol).setFontWeight("bold");
  sheet.getRange(20, 1, 1, lastCol)
    .setFontWeight("bold")
    .setBackground("#1f1f1f")
    .setFontColor("#ffffff");

  sheet.getRange(1, 1, lastRow, lastCol)
    .setWrap(true)
    .setVerticalAlignment("middle");

  sheet.setColumnWidth(1, 220);
  sheet.setColumnWidth(2, 110);
  sheet.setColumnWidth(3, 90);
  sheet.setColumnWidth(4, 90);
  sheet.setColumnWidth(5, 110);
  sheet.setColumnWidth(6, 110);

  sheet.setFrozenRows(1);
}