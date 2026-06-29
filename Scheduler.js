const ODDS_TARGET_TIMES = [
  { hour: 8, minute: 0, label: "08:00" },
  { hour: 12, minute: 0, label: "12:00" },
  { hour: 16, minute: 0, label: "16:00" },
  { hour: 18, minute: 0, label: "18:00" }
];

const ODDS_WINDOW_MINUTES = 7;

/**
 * Runs every 15 minutes.
 * Only calls odds API if current time is inside an approved odds window.
 */
function runScheduledOddsUpdate() {
  const now = new Date();
  const target = getCurrentOddsTarget_(now);

  if (!target) {
    Logger.log("Not inside odds update window.");
    return;
  }

  const runKey = buildOddsRunKey_(now, target);

  if (hasOddsWindowAlreadyRun_(runKey)) {
    Logger.log("Odds update already ran for " + runKey);
    return;
  }

  Logger.log("Running odds snapshot for " + runKey);

  runOddsSnapshot();

  markOddsWindowRun_(runKey);
}

/**
 * Dedicated odds refresh.
 * This is the only normal function that should call importOdds().
 */
function runOddsSnapshot() {
  importOdds();

  buildModelMatrix();
  scoreModelMatrix();

  buildTodayView();
  buildDashboard();
}

/**
 * Finds whether current time is close enough to one of our target odds times.
 */
function getCurrentOddsTarget_(now) {
  for (const target of ODDS_TARGET_TIMES) {
    const targetTime = new Date(now);
    targetTime.setHours(target.hour, target.minute, 0, 0);

    const diffMinutes = Math.abs(now.getTime() - targetTime.getTime()) / 60000;

    if (diffMinutes <= ODDS_WINDOW_MINUTES) {
      return target;
    }
  }

  return null;
}

/**
 * Creates unique key like:
 * ODDS_RUN_2026-06-28_12:00
 */
function buildOddsRunKey_(now, target) {
  const dateKey = Utilities.formatDate(
    now,
    Session.getScriptTimeZone(),
    "yyyy-MM-dd"
  );

  return "ODDS_RUN_" + dateKey + "_" + target.label;
}

function hasOddsWindowAlreadyRun_(runKey) {
  const props = PropertiesService.getScriptProperties();
  return props.getProperty(runKey) === "TRUE";
}

function markOddsWindowRun_(runKey) {
  const props = PropertiesService.getScriptProperties();
  props.setProperty(runKey, "TRUE");
}

/**
 * Install one master trigger.
 * Run this manually one time.
 */
function installScheduledOddsTrigger() {
  deleteTriggersForFunction_("runScheduledOddsUpdate");

  ScriptApp.newTrigger("runScheduledOddsUpdate")
    .timeBased()
    .everyMinutes(15)
    .create();
}

/**
 * Removes existing triggers for a specific function.
 */
function deleteTriggersForFunction_(functionName) {
  const triggers = ScriptApp.getProjectTriggers();

  triggers.forEach(trigger => {
    if (trigger.getHandlerFunction() === functionName) {
      ScriptApp.deleteTrigger(trigger);
    }
  });
}