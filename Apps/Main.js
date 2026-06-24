function runMorningCoreUpdate() {
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


// NEW v0.1.1 FUNCTION
// Run this before games start.
// This stores the actual pregame model state in HISTORY.
function runPregameSnapshot() {
  buildModelMatrix();
  scoreModelMatrix();
  buildTodayView();
  buildDashboard();

  appendPregameHistorySnapshot();

  buildPerformanceDashboard();
  buildRollingFeatureTesting();
}


// UPDATED v0.1.1 FUNCTION
// This no longer appends HISTORY rows.
// It only imports results and updates existing pregame rows.
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