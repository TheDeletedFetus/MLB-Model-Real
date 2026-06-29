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
  buildTodayView();
  buildDashboard();
  buildPerformanceDashboard();
}


function runOddsSnapshot() {
  importOdds();

  buildModelMatrix();
  scoreModelMatrix();
  buildTodayView();
  buildDashboard();
  buildPerformanceDashboard();
}


function runPregameSnapshot() {
  importOdds();

  buildModelMatrix();
  scoreModelMatrix();

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
  buildTodayView();
  buildDashboard();
  buildPerformanceDashboard();
  buildRollingFeatureTesting();
}
