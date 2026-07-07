/***************************************
 * SHARED SCORING ENGINE v0.1.0
 *
 * Central scoring helper for validation and optimizer work.
 *
 * Current production-style assumption:
 * - Model_Matrix stores precomputed directional edge columns.
 * - Edge columns are already oriented as Away advantage.
 * - Away score = weighted edge sum / 1000
 * - Home score = -Away score
 *
 * This avoids re-scoring from raw Away/Home stat values.
 ***************************************/

const SHARED_SCORING = {
  SCORE_SCALE: 1000,
  SCORE_DECIMALS: 4,

  EDGE_ALIASES: {
    "OPS vs Hand": ["OPS vs Hand", "OPS vs SP Hand"],

    "Home OPS": ["Split OPS"],
    "Road OPS": ["Split OPS"],
    "Home ERA": ["Split ERA"],
    "Road ERA": ["Split ERA"],
    "Home WHIP": ["Split WHIP"],
    "Road WHIP": ["Split WHIP"],
    "Home K/9": ["Split K/9"],
    "Road K/9": ["Split K/9"],
    "Home K/BB": ["Split K/BB"],
    "Road K/BB": ["Split K/BB"],

    "Starter Home ERA": ["Starter Split ERA"],
    "Starter Road ERA": ["Starter Split ERA"],
    "Starter Home WHIP": ["Starter Split WHIP"],
    "Starter Road WHIP": ["Starter Split WHIP"],
    "Starter Home K/BB": ["Starter Split K/BB"],
    "Starter Road K/BB": ["Starter Split K/BB"],

    "Bullpen IP Last 3 Days": ["Bullpen IP Last 3"],
    "Bullpen IP Last 7 Days": ["Bullpen IP Last 7"],
    "Bullpen Pitches Last 3 Days": ["Bullpen Pitches Last 3"],
    "Bullpen Pitches Last 7 Days": ["Bullpen Pitches Last 7"]
  }
};


function sharedBuildEdgeColumnMap_(headers, features, prefix) {
  const map = {};
  const cleanPrefix = prefix || "";

  features.forEach(feature => {
    const idx = sharedResolveEdgeColumn_(headers, feature.name, cleanPrefix);
    if (idx !== undefined) {
      map[feature.name] = idx;
    }
  });

  return map;
}


function sharedResolveEdgeColumn_(headers, statName, prefix) {
  const candidates = sharedBuildEdgeCandidates_(statName, prefix || "");

  for (let i = 0; i < candidates.length; i++) {
    const idx = sharedGetHeaderIndex_(headers, candidates[i]);
    if (idx !== undefined) return idx;
  }

  return undefined;
}


function sharedBuildEdgeCandidates_(statName, prefix) {
  const baseNames = [statName];
  const aliases = SHARED_SCORING.EDGE_ALIASES[statName] || [];
  aliases.forEach(alias => baseNames.push(alias));

  const candidates = [];
  baseNames.forEach(name => {
    candidates.push(prefix + name + " Edge");
    candidates.push(name + " Edge");
  });

  return candidates;
}


function sharedScoreGameFromEdges_(row, edgeColumnMap, features, weightsByFeature, awayTeam, homeTeam) {
  let weightedEdgeSum = 0;
  let usedFeatures = 0;

  features.forEach(feature => {
    const weight = Number(weightsByFeature[feature.name] || 0);
    if (!isFinite(weight) || weight === 0) return;

    const col = edgeColumnMap[feature.name];
    if (col === undefined) return;

    const edge = Number(row[col]);
    if (!isFinite(edge)) return;

    weightedEdgeSum += edge * weight;
    usedFeatures++;
  });

  const awayScore = sharedRoundScore_(weightedEdgeSum / SHARED_SCORING.SCORE_SCALE);
  const homeScore = sharedRoundScore_(-awayScore);

  return {
    awayScore,
    homeScore,
    pick: awayScore > homeScore ? awayTeam : homeScore > awayScore ? homeTeam : "Tie",
    weightedEdgeSum,
    usedFeatures,
    hasScore: usedFeatures > 0
  };
}


function sharedContributionRowsFromEdges_(row, edgeColumnMap, features, weightsByFeature) {
  const rows = [];

  features.forEach(feature => {
    const weight = Number(weightsByFeature[feature.name] || 0);
    if (!isFinite(weight) || weight === 0) return;

    const col = edgeColumnMap[feature.name];
    if (col === undefined) return;

    const edge = Number(row[col]);
    if (!isFinite(edge)) return;

    const rawContribution = edge * weight;
    const scaledContribution = rawContribution / SHARED_SCORING.SCORE_SCALE;

    rows.push({
      feature: feature.name,
      edge,
      weight,
      rawContribution,
      scaledContribution
    });
  });

  return rows;
}


function sharedRoundScore_(value) {
  const n = Number(value);
  if (!isFinite(n)) return 0;

  const factor = Math.pow(10, SHARED_SCORING.SCORE_DECIMALS);
  return Math.round(n * factor) / factor;
}


function sharedGetHeaderIndex_(headers, headerName) {
  if (headers[headerName] !== undefined) return headers[headerName];

  const target = sharedNormalizeHeader_(headerName);
  const keys = Object.keys(headers);

  for (let i = 0; i < keys.length; i++) {
    if (sharedNormalizeHeader_(keys[i]) === target) return headers[keys[i]];
  }

  return undefined;
}


function sharedNormalizeHeader_(value) {
  return String(value || "")
    .toLowerCase()
    .replace(/[^a-z0-9]/g, "");
}
