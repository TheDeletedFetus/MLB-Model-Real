function buildRollingFeatureTesting() {
  const historyRows = getSheetRows("HISTORY");

  const finalizedRows = historyRows.filter(row => {
    return ftNormalizeBoolean(row["Final?"]) === true &&
      row["Winner"] &&
      row["Away Team"] &&
      row["Home Team"];
  });

  const featureColumns = getFeatureTestingColumns(historyRows);

  const featureSummary = [];
  const bucketSummary = [];

  featureColumns.forEach(featureColumn => {
    const featureRows = getValidRowsForFeature(finalizedRows, featureColumn);
    if (featureRows.length === 0) return;

    const result = testFeatureSummary(featureColumn, featureRows);
    featureSummary.push(result);

    const buckets = testFeatureBuckets(featureColumn, featureRows);
    buckets.forEach(bucket => bucketSummary.push(bucket));
  });

  featureSummary.sort((a, b) => {
    if (b.games !== a.games) return b.games - a.games;
    return b.deltaRaw - a.deltaRaw;
  });

  writeFeatureTestingSummary(featureSummary, finalizedRows.length, featureColumns.length);
  writeFeatureBucketSummary(bucketSummary, finalizedRows.length);
}


function getFeatureTestingColumns(rows) {
  if (!rows || rows.length === 0) return [];

  const sampleRow = rows[0];

  return Object.keys(sampleRow)
    .filter(header => {
      if (!header.startsWith("MM_")) return false;
      if (!header.includes("Edge")) return false;
      if (header.includes("Projected Edge")) return false;
      if (header.includes("Market")) return false;
      return true;
    })
    .sort();
}


function getValidRowsForFeature(rows, featureColumn) {
  return rows.filter(row => {
    const edgeValue = ftParseNumber(row[featureColumn]);
    if (edgeValue === "") return false;
    if (edgeValue === 0) return false;

    const winner = row["Winner"];
    const awayTeam = row["Away Team"];
    const homeTeam = row["Home Team"];

    return winner && awayTeam && homeTeam;
  });
}


function testFeatureSummary(featureColumn, rows) {
  let games = 0;
  let correct = 0;
  let incorrect = 0;

  let awayFavoredGames = 0;
  let awayFavoredCorrect = 0;

  let homeFavoredGames = 0;
  let homeFavoredCorrect = 0;

  let marketProbabilitySum = 0;
  let marketProbabilityCount = 0;

  let profit = 0;
  let profitCount = 0;

  let absEdgeCorrectSum = 0;
  let absEdgeCorrectCount = 0;

  let absEdgeIncorrectSum = 0;
  let absEdgeIncorrectCount = 0;

  rows.forEach(row => {
    const edgeValue = ftParseNumber(row[featureColumn]);
    const absEdge = Math.abs(edgeValue);

    const awayTeam = row["Away Team"];
    const homeTeam = row["Home Team"];
    const winner = row["Winner"];

    const featurePick = edgeValue > 0 ? awayTeam : homeTeam;
    const isCorrect = featurePick === winner;

    const pickMarketProbability = getFeaturePickMarketProbability(row, featurePick);
    const pickOdds = getFeaturePickOdds(row, featurePick);

    games++;

    if (isCorrect) {
      correct++;
      absEdgeCorrectSum += absEdge;
      absEdgeCorrectCount++;
    } else {
      incorrect++;
      absEdgeIncorrectSum += absEdge;
      absEdgeIncorrectCount++;
    }

    if (edgeValue > 0) {
      awayFavoredGames++;
      if (isCorrect) awayFavoredCorrect++;
    }

    if (edgeValue < 0) {
      homeFavoredGames++;
      if (isCorrect) homeFavoredCorrect++;
    }

    if (pickMarketProbability !== "") {
      marketProbabilitySum += pickMarketProbability;
      marketProbabilityCount++;
    }

    const betProfit = calculateFlatBetProfit(pickOdds, isCorrect);
    if (betProfit !== "") {
      profit += betProfit;
      profitCount++;
    }
  });

  const winRateRaw = games > 0 ? correct / games : 0;
  const marketExpectedRaw =
    marketProbabilityCount > 0 ? marketProbabilitySum / marketProbabilityCount : "";

  const deltaRaw =
    marketExpectedRaw === "" ? "" : winRateRaw - marketExpectedRaw;

  const roiRaw =
    profitCount > 0 ? profit / (profitCount * 100) : "";

  return {
    feature: cleanFeatureName(featureColumn),
    sourceColumn: featureColumn,
    games,
    correct,
    incorrect,
    winRateRaw,
    winRate: ftFormatPercent(winRateRaw),

    marketExpectedRaw,
    marketExpected: marketExpectedRaw === "" ? "" : ftFormatPercent(marketExpectedRaw),

    deltaRaw,
    delta: deltaRaw === "" ? "" : ftFormatSignedPercent(deltaRaw),

    awayFavoredGames,
    awayFavoredWinRate: awayFavoredGames > 0
      ? ftFormatPercent(awayFavoredCorrect / awayFavoredGames)
      : "",

    homeFavoredGames,
    homeFavoredWinRate: homeFavoredGames > 0
      ? ftFormatPercent(homeFavoredCorrect / homeFavoredGames)
      : "",

    avgEdgeCorrect: absEdgeCorrectCount > 0
      ? ftRound(absEdgeCorrectSum / absEdgeCorrectCount, 4)
      : "",

    avgEdgeIncorrect: absEdgeIncorrectCount > 0
      ? ftRound(absEdgeIncorrectSum / absEdgeIncorrectCount, 4)
      : "",

    flatBetProfit: ftRound(profit, 2),
    roi: roiRaw === "" ? "" : ftFormatSignedPercent(roiRaw),

    signalLabel: getFeatureSignalLabel(games, winRateRaw, deltaRaw)
  };
}


function testFeatureBuckets(featureColumn, rows) {
  const absEdges = rows
    .map(row => Math.abs(ftParseNumber(row[featureColumn])))
    .filter(value => value !== "" && !isNaN(value))
    .sort((a, b) => a - b);

  if (absEdges.length === 0) return [];

  const q1 = percentile(absEdges, 0.25);
  const q2 = percentile(absEdges, 0.50);
  const q3 = percentile(absEdges, 0.75);

  const buckets = {
    "Smallest 25%": blankFeatureBucket(cleanFeatureName(featureColumn), "Smallest 25%", 0, q1),
    "25%-50%": blankFeatureBucket(cleanFeatureName(featureColumn), "25%-50%", q1, q2),
    "50%-75%": blankFeatureBucket(cleanFeatureName(featureColumn), "50%-75%", q2, q3),
    "Largest 25%": blankFeatureBucket(cleanFeatureName(featureColumn), "Largest 25%", q3, "")
  };

  rows.forEach(row => {
    const edgeValue = ftParseNumber(row[featureColumn]);
    if (edgeValue === "" || edgeValue === 0) return;

    const absEdge = Math.abs(edgeValue);
    const bucketName = getBucketName(absEdge, q1, q2, q3);
    const bucket = buckets[bucketName];

    const awayTeam = row["Away Team"];
    const homeTeam = row["Home Team"];
    const winner = row["Winner"];

    const featurePick = edgeValue > 0 ? awayTeam : homeTeam;
    const isCorrect = featurePick === winner;

    const pickMarketProbability = getFeaturePickMarketProbability(row, featurePick);
    const pickOdds = getFeaturePickOdds(row, featurePick);
    const betProfit = calculateFlatBetProfit(pickOdds, isCorrect);

    bucket.games++;

    if (isCorrect) bucket.correct++;
    else bucket.incorrect++;

    if (pickMarketProbability !== "") {
      bucket.marketProbabilitySum += pickMarketProbability;
      bucket.marketProbabilityCount++;
    }

    if (betProfit !== "") {
      bucket.profit += betProfit;
      bucket.profitCount++;
    }
  });

  return Object.values(buckets).map(bucket => {
    const winRateRaw = bucket.games > 0 ? bucket.correct / bucket.games : "";
    const marketExpectedRaw =
      bucket.marketProbabilityCount > 0
        ? bucket.marketProbabilitySum / bucket.marketProbabilityCount
        : "";

    const deltaRaw =
      winRateRaw === "" || marketExpectedRaw === ""
        ? ""
        : winRateRaw - marketExpectedRaw;

    const roiRaw =
      bucket.profitCount > 0 ? bucket.profit / (bucket.profitCount * 100) : "";

    return {
      feature: bucket.feature,
      bucket: bucket.bucket,
      edgeMin: bucket.edgeMin,
      edgeMax: bucket.edgeMax,
      games: bucket.games,
      correct: bucket.correct,
      incorrect: bucket.incorrect,
      winRateRaw,
      winRate: winRateRaw === "" ? "" : ftFormatPercent(winRateRaw),
      marketExpected: marketExpectedRaw === "" ? "" : ftFormatPercent(marketExpectedRaw),
      delta: deltaRaw === "" ? "" : ftFormatSignedPercent(deltaRaw),
      flatBetProfit: ftRound(bucket.profit, 2),
      roi: roiRaw === "" ? "" : ftFormatSignedPercent(roiRaw)
    };
  });
}


function blankFeatureBucket(feature, bucket, edgeMin, edgeMax) {
  return {
    feature,
    bucket,
    edgeMin,
    edgeMax,
    games: 0,
    correct: 0,
    incorrect: 0,
    marketProbabilitySum: 0,
    marketProbabilityCount: 0,
    profit: 0,
    profitCount: 0
  };
}


function getBucketName(absEdge, q1, q2, q3) {
  if (absEdge <= q1) return "Smallest 25%";
  if (absEdge <= q2) return "25%-50%";
  if (absEdge <= q3) return "50%-75%";
  return "Largest 25%";
}


function percentile(values, p) {
  if (!values || values.length === 0) return "";

  const index = (values.length - 1) * p;
  const lower = Math.floor(index);
  const upper = Math.ceil(index);

  if (lower === upper) return values[lower];

  const weight = index - lower;
  return values[lower] * (1 - weight) + values[upper] * weight;
}


function getFeaturePickMarketProbability(row, featurePick) {
  const awayTeam = row["Away Team"];
  const homeTeam = row["Home Team"];

  if (featurePick === awayTeam) {
    return firstValidNumber([
      row["Away Market Implied Probability"],
      row["MM_Away Implied Probability"]
    ]);
  }

  if (featurePick === homeTeam) {
    return firstValidNumber([
      row["Home Market Implied Probability"],
      row["MM_Home Implied Probability"]
    ]);
  }

  return "";
}


function getFeaturePickOdds(row, featurePick) {
  const awayTeam = row["Away Team"];
  const homeTeam = row["Home Team"];

  if (featurePick === awayTeam) {
    return firstValidNumber([
      row["Away ML"],
      row["MM_Away ML"]
    ]);
  }

  if (featurePick === homeTeam) {
    return firstValidNumber([
      row["Home ML"],
      row["MM_Home ML"]
    ]);
  }

  return "";
}


function firstValidNumber(values) {
  for (let i = 0; i < values.length; i++) {
    const parsed = ftParseNumber(values[i]);
    if (parsed !== "") return parsed;
  }

  return "";
}


function writeFeatureTestingSummary(featureSummary, finalizedRowCount, featureCount) {
  const sheet = getOrCreateSheet("Feature_Testing");
  sheet.clearContents();

  const output = [
    ["Rolling Feature Testing", ""],
    ["Last Updated", Utilities.formatDate(new Date(), "America/New_York", "yyyy-MM-dd HH:mm:ss")],
    ["Finalized HISTORY Rows Used", finalizedRowCount],
    ["Features Tested", featureCount],
    ["", ""],
    [
      "Feature",
      "Games",
      "Correct",
      "Incorrect",
      "Win %",
      "Market Expected Win %",
      "Delta vs Market",
      "Away-Favored Games",
      "Away-Favored Win %",
      "Home-Favored Games",
      "Home-Favored Win %",
      "Avg Abs Edge When Correct",
      "Avg Abs Edge When Incorrect",
      "Flat Bet Profit",
      "ROI",
      "Signal Label"
    ]
  ];

  featureSummary.forEach(row => {
    output.push([
      row.feature,
      row.games,
      row.correct,
      row.incorrect,
      row.winRate,
      row.marketExpected,
      row.delta,
      row.awayFavoredGames,
      row.awayFavoredWinRate,
      row.homeFavoredGames,
      row.homeFavoredWinRate,
      row.avgEdgeCorrect,
      row.avgEdgeIncorrect,
      row.flatBetProfit,
      row.roi,
      row.signalLabel
    ]);
  });

  if (featureSummary.length === 0) {
    output.push([
      "No finalized HISTORY rows available yet",
      "",
      "",
      "",
      "",
      "",
      "",
      "",
      "",
      "",
      "",
      "",
      "",
      "",
      "",
      ""
    ]);
  }

  sheet
    .getRange(1, 1, output.length, 16)
    .setValues(padRows(output, 16));

  formatFeatureTestingSheet(sheet, 6);
}


function writeFeatureBucketSummary(bucketSummary, finalizedRowCount) {
  const sheet = getOrCreateSheet("Feature_Buckets");
  sheet.clearContents();

  const output = [
    ["Rolling Feature Bucket Testing", ""],
    ["Last Updated", Utilities.formatDate(new Date(), "America/New_York", "yyyy-MM-dd HH:mm:ss")],
    ["Finalized HISTORY Rows Used", finalizedRowCount],
    ["", ""],
    [
      "Feature",
      "Bucket",
      "Abs Edge Min",
      "Abs Edge Max",
      "Games",
      "Correct",
      "Incorrect",
      "Win %",
      "Market Expected Win %",
      "Delta vs Market",
      "Flat Bet Profit",
      "ROI"
    ]
  ];

  bucketSummary.forEach(row => {
    output.push([
      row.feature,
      row.bucket,
      row.edgeMin,
      row.edgeMax,
      row.games,
      row.correct,
      row.incorrect,
      row.winRate,
      row.marketExpected,
      row.delta,
      row.flatBetProfit,
      row.roi
    ]);
  });

  if (bucketSummary.length === 0) {
    output.push([
      "No finalized HISTORY rows available yet",
      "",
      "",
      "",
      "",
      "",
      "",
      "",
      "",
      "",
      "",
      ""
    ]);
  }

  sheet
    .getRange(1, 1, output.length, 12)
    .setValues(padRows(output, 12));

  formatFeatureTestingSheet(sheet, 5);
}


function cleanFeatureName(featureColumn) {
  return featureColumn.replace(/^MM_/, "");
}


function getFeatureSignalLabel(games, winRateRaw, deltaRaw) {
  if (games < 30) return "Too little data";

  if (games < 100) {
    if (deltaRaw !== "" && deltaRaw >= 0.04) return "Early positive market signal";
    if (deltaRaw !== "" && deltaRaw <= -0.04) return "Early negative market signal";
    if (winRateRaw >= 0.56) return "Early positive raw signal";
    if (winRateRaw <= 0.44) return "Early negative raw signal";
    return "Early noise";
  }

  if (deltaRaw !== "" && deltaRaw >= 0.03) return "Possible positive market signal";
  if (deltaRaw !== "" && deltaRaw >= 0.015) return "Slight positive market signal";
  if (deltaRaw !== "" && deltaRaw <= -0.03) return "Possible harmful market signal";
  if (deltaRaw !== "" && deltaRaw <= -0.015) return "Slight negative market signal";

  if (winRateRaw >= 0.55) return "Possible positive raw signal";
  if (winRateRaw <= 0.45) return "Possible harmful raw signal";

  return "Likely noise";
}


function ftParseNumber(value) {
  if (value === "" || value === undefined || value === null) return "";

  const cleaned = value.toString().replace("%", "").trim();
  const num = Number(cleaned);

  return isNaN(num) ? "" : num;
}


function ftNormalizeBoolean(value) {
  if (value === true || value === "TRUE" || value === "Yes") return true;
  if (value === false || value === "FALSE" || value === "No") return false;
  return null;
}


function ftFormatPercent(value) {
  if (value === "" || value === undefined || value === null || isNaN(value)) return "";
  return Math.round(value * 1000) / 10 + "%";
}


function ftFormatSignedPercent(value) {
  if (value === "" || value === undefined || value === null || isNaN(value)) return "";

  const pct = Math.round(value * 1000) / 10;
  return pct > 0 ? "+" + pct + "%" : pct + "%";
}


function ftRound(value, decimals) {
  if (value === "" || isNaN(value)) return "";
  return Number(value.toFixed(decimals));
}


function formatFeatureTestingSheet(sheet, headerRow) {
  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();

  if (lastRow < 1) return;

  sheet.getRange(1, 1, 1, lastCol)
    .setFontWeight("bold")
    .setFontSize(14)
    .setBackground("#1f1f1f")
    .setFontColor("#ffffff");

  sheet.getRange(headerRow, 1, 1, lastCol)
    .setFontWeight("bold")
    .setBackground("#1f1f1f")
    .setFontColor("#ffffff");

  sheet.getRange(1, 1, lastRow, lastCol)
    .setWrap(true)
    .setVerticalAlignment("middle");

  sheet.setFrozenRows(headerRow);
  sheet.autoResizeColumns(1, lastCol);
}