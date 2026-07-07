# Project Context

## Project

MLB predictive betting model built primarily in Google Sheets and Google Apps Script.

Current versions:

- Project: `v0.2.0`
- Production model: `v0.1.0`
- Stage: Development

## Current Objective

Build a controlled, measurable, repeatable model-development platform for MLB moneyline evaluation.

The model is not expected to win every day. It is evaluated over large samples using flat-risk betting results, ROI, and out-of-sample validation.

## Core System Components

### Live Model

- Uses `MODEL_MATRIX`.
- Scored with production `Weight` values from `Settings`.
- Outputs production picks.
- Feeds the user-facing Today/dashboard views.

### Shadow Model

- Uses `MODEL_MATRIX_SHADOW`.
- Built by copying `MODEL_MATRIX` and rescoring with `Test Weight` values.
- Does not affect production picks.
- Used to forward-test optimizer suggestions.

### HISTORY

Permanent append-only game-level database.

Stores:

- Live model prediction.
- Shadow model prediction.
- Full `MM_...` snapshot of live matrix.
- Full `SM_...` snapshot of shadow matrix.
- Final result.
- Correctness.
- Flat-bet profit.

Project rule:

```text
Never delete historical feature data. Preserve columns even if features are inactive.
```

### Optimizer

- Current file: `Optimizer_Iterative.gs`.
- Uses production scoring helpers from `Scoring.js`.
- Runs hill-climbing search over feature weights.
- Writes candidate suggestions to `Suggested Weight`.
- Candidate weights should be tested through the shadow model before promotion.

### Scoring

Source of truth:

```javascript
scoreModelRowWithSettings(row, headers, settings, edgeStats)
```

Defined in `Scoring.js`.

Production scoring uses z-scored edge columns weighted by active feature weights.

## Current Daily Automation Pattern

Main runners in `Main.js`:

- `runMorningCoreUpdate()`
- `runBullpenUpdate()`
- `runOddsSnapshot()`
- `runPregameSnapshot()`
- `runResultsAndHistory()`

Current sequence after model build:

```text
buildModelMatrix()
scoreModelMatrix()
buildShadowModelMatrix()
```

This means existing triggers generally do not need to change when shadow model logic is added.

## Current Working Workflow

1. Run optimizer.
2. Review recommendations.
3. Run `setupShadowModelTestWeights()` to copy suggested weights into `Test Weight`.
4. Daily automation scores live and shadow models.
5. Pregame snapshot stores both predictions in HISTORY.
6. After enough new games, compare live vs shadow.
7. Promote shadow weights only if they outperform out of sample.

## Current Highest Priorities

1. Build `SHADOW_DASHBOARD`.
2. Add train/test validation to the iterative optimizer.
3. Compare confidence bucket performance for live vs shadow.
4. Audit HISTORY to ensure all tracked features persist.
5. Continue per-game pregame snapshot gatekeeper.
6. Add model/experiment version tracking.
7. Prepare future Python/database workflow.

## Data Integrity Rules

- Do not reset HISTORY unless explicitly intended.
- Do not delete historical columns.
- Do not evaluate shadow performance unless predictions were captured before first pitch.
- Do not promote optimizer weights solely from in-sample optimizer output.
- Do not change flat-risk ROI logic without reviewing all downstream dashboards/backtests.

## Betting Math Standard

Flat-risk betting engine:

- Risk `$100` per model pick.
- Win profit is calculated from American odds.
- Loss is `-$100`.
- ROI = profit / total risk.

## Important Known Limitations

- Optimizer still needs train/test split to reduce overfitting.
- Shadow dashboard does not exist yet.
- Pregame snapshots still need per-game gatekeeping for start-time windows.
- Lineup strength is not integrated yet.
- Python/database migration is planned but not started.

## Near-Term Coding Sequence

Recommended next work order:

1. `SHADOW_DASHBOARD`.
2. Train/test optimizer.
3. Confidence calibration.
4. HISTORY audit.
5. Pregame gatekeeper.
6. Lineup strength.
7. CLV tracking.
8. Python/database export.
