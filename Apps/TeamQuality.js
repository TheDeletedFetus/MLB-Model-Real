function importTeamQuality() {
  const hittingSheet = getOrCreateSheet("RAW_Team_Hitting");
  const pitchingSheet = getOrCreateSheet("RAW_Team_Pitching");
  const qualitySheet = getOrCreateSheet("RAW_Team_Quality");

  qualitySheet.clearContents();

  const hittingData = hittingSheet.getDataRange().getValues();
  const pitchingData = pitchingSheet.getDataRange().getValues();

  const hittingHeaders = hittingData[0];
  const pitchingHeaders = pitchingData[0];

  const hitTeamCol = hittingHeaders.indexOf("Team");
  const runsScoredCol = hittingHeaders.indexOf("Runs");

  const pitchTeamCol = pitchingHeaders.indexOf("Team");
  const runsAllowedCol = pitchingHeaders.indexOf("Runs Allowed");

  const runsAllowedByTeam = {};

  for (let i = 1; i < pitchingData.length; i++) {
    const team = pitchingData[i][pitchTeamCol];
    const runsAllowed = Number(pitchingData[i][runsAllowedCol] || 0);
    runsAllowedByTeam[team] = runsAllowed;
  }

  const rows = [[
    "Team",
    "Runs Scored",
    "Runs Allowed",
    "Run Differential"
  ]];

  for (let i = 1; i < hittingData.length; i++) {
    const team = hittingData[i][hitTeamCol];
    const runsScored = Number(hittingData[i][runsScoredCol] || 0);
    const runsAllowed = runsAllowedByTeam[team] || 0;
    const runDifferential = runsScored - runsAllowed;

    rows.push([
      team,
      runsScored,
      runsAllowed,
      runDifferential
    ]);
  }

  qualitySheet.getRange(1, 1, rows.length, rows[0].length).setValues(rows);
}