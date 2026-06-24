function scoreModelMatrix() {
  const sheet = getOrCreateSheet("MODEL_MATRIX");
  const values = sheet.getDataRange().getValues();

  if (values.length < 2) return;

  const headers = values[0];
  const rows = values.slice(1);
  const settings = getActiveModelSettings();

  const edgeStats = calculateEdgeStats(rows, headers, settings);

  const awayScoreCol = headers.indexOf("Away Model Score");
  const homeScoreCol = headers.indexOf("Home Model Score");
  const pickCol = headers.indexOf("Model Pick");
  const confidenceCol = headers.indexOf("Confidence");

  const updates = [];

  rows.forEach(row => {
    let totalScore = 0;
    let totalWeight = 0;

    settings.forEach(setting => {
      const edgeColName = setting.stat + " Edge";
      const edgeCol = headers.indexOf(edgeColName);
      if (edgeCol === -1) return;

      const rawEdge = parseNumber(row[edgeCol]);
      if (rawEdge === "") return;

      const stats = edgeStats[edgeColName];
      if (!stats || stats.stdDev === 0) return;

      const zScore = (rawEdge - stats.mean) / stats.stdDev;
      totalScore += zScore * setting.weight;
      totalWeight += setting.weight;
    });

    const finalScore = totalWeight > 0 ? totalScore / totalWeight : 0;

    const awayScore = roundScore(finalScore, 4);
    const homeScore = roundScore(finalScore * -1, 4);

    const awayTeam = row[headers.indexOf("Away Team")];
    const homeTeam = row[headers.indexOf("Home Team")];

    const pick =
      finalScore > 0 ? awayTeam :
      finalScore < 0 ? homeTeam :
      "Coin Flip";

    const confidence = Math.min(100, Math.round(Math.abs(finalScore) * 25));

    updates.push([awayScore, homeScore, pick, confidence]);
  });

  sheet.getRange(2, awayScoreCol + 1, updates.length, 4).setValues(updates);
}


function calculateEdgeStats(rows, headers, settings) {
  const edgeStats = {};

  settings.forEach(setting => {
    const edgeColName = setting.stat + " Edge";
    const edgeCol = headers.indexOf(edgeColName);
    if (edgeCol === -1) return;

    const values = rows
      .map(row => parseNumber(row[edgeCol]))
      .filter(value => value !== "");

    if (values.length === 0) return;

    const mean = values.reduce((sum, value) => sum + value, 0) / values.length;

    const variance =
      values.reduce((sum, value) => sum + Math.pow(value - mean, 2), 0) /
      values.length;

    edgeStats[edgeColName] = {
      mean: mean,
      stdDev: Math.sqrt(variance)
    };
  });

  return edgeStats;
}


function getActiveModelSettings() {
  const sheet = getOrCreateSheet("Settings");
  const values = sheet.getDataRange().getValues();

  if (values.length < 2) return [];

  const headers = values[0];
  const rows = values.slice(1);

  const statCol = headers.indexOf("Stat");
  const activeCol = headers.indexOf("Active");
  const weightCol = headers.indexOf("Weight");

  const settings = [];

  rows.forEach(row => {
    const stat = row[statCol];
    const active = row[activeCol];
    const weight = Number(row[weightCol] || 0);

    if ((active === true || active === "TRUE") && stat && weight > 0) {
      settings.push({
        stat: stat,
        weight: weight
      });
    }
  });

  return settings;
}


function parseNumber(value) {
  if (value === "" || value === undefined || value === null) return "";

  const cleaned = value.toString().replace("%", "").trim();
  const num = Number(cleaned);

  return isNaN(num) ? "" : num;
}


function roundScore(value, decimals) {
  if (value === "" || isNaN(value)) return "";
  return Number(value.toFixed(decimals));
}

function debugOneGameScore() {
  const sheet = getOrCreateSheet("MODEL_MATRIX");
  const values = sheet.getDataRange().getValues();

  const headers = values[0];
  const row = values[1];

  const settings = getActiveModelSettings();
  const edgeStats = calculateEdgeStats(values.slice(1), headers, settings);

  Logger.log("GAME: " + row[headers.indexOf("Away Team")] + " @ " + row[headers.indexOf("Home Team")]);

  settings.forEach(setting => {
    const edgeColName = setting.stat + " Edge";
    const edgeCol = headers.indexOf(edgeColName);

    if (edgeCol === -1) {
      Logger.log(setting.stat + " | NOT FOUND");
      return;
    }

    const rawEdge = parseNumber(row[edgeCol]);
    const stats = edgeStats[edgeColName];

    if (rawEdge === "" || !stats || stats.stdDev === 0) {
      Logger.log(setting.stat + " | skipped");
      return;
    }

    const zScore = (rawEdge - stats.mean) / stats.stdDev;
    const contribution = zScore * setting.weight;

    Logger.log(
      setting.stat +
      " | weight=" + setting.weight +
      " | rawEdge=" + rawEdge +
      " | z=" + zScore.toFixed(4) +
      " | contribution=" + contribution.toFixed(4)
    );
  });
}

