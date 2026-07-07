function runMorningCoreUpdate() {
  importResults();          // yesterday's finals
  updateHistoryResults();   // grades yesterday's HISTORY rows

  importSchedule();

  importTeamHitting();
  importTeamPitching();
  importTeamQuality();
  importTeamSplits();
  importRecentForm();

  importStartingPitchers();
  importStarterRecentForm();
  importStarterHomeAwaySplits();

  importTeamHomeAwaySplits();

  importParkFactors();
  importWeather();

  buildModelMatrix();
  scoreModelMatrix();
  buildShadowModelMatrix();

  buildTodayView();
  buildDashboard();
  buildPerformanceDashboard();
  buildRollingFeatureTesting();
}

function runBullpenUpdate() {
  importBullpen();
  importBullpenQuality();
  importBullpenFatigue();

  buildModelMatrix();
  scoreModelMatrix();
  buildShadowModelMatrix();

  buildTodayView();
  buildDashboard();
  buildPerformanceDashboard();
}


function runOddsSnapshot() {
  importOdds();

  buildModelMatrix();
  scoreModelMatrix();
  buildShadowModelMatrix();

  buildTodayView();
  buildDashboard();
  buildPerformanceDashboard();
}


function runPregameSnapshot() {
  importOdds();

  buildModelMatrix();
  scoreModelMatrix();
  buildShadowModelMatrix();

  validatePregameSnapshotReady();

  buildTodayView();
  buildDashboard();

  appendPregameHistorySnapshot();

  buildPerformanceDashboard();
  buildRollingFeatureTesting();
}

function runResultsAndHistory() {
  importResults();

  updateHistoryResults();

  buildModelMatrix();
  scoreModelMatrix();
  buildShadowModelMatrix();

  buildTodayView();
  buildDashboard();
  buildPerformanceDashboard();
  buildRollingFeatureTesting();
}
