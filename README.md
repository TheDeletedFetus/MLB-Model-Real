# MLB Model

A Google Sheets and Apps Script predictive modeling platform for MLB moneyline evaluation, backtesting, optimizer research, and live-vs-shadow model testing.

This project is currently in active development. It is not designed to predict every MLB game correctly. It is designed to identify repeatable statistical edges over a large sample of games and evaluate whether those edges survive forward testing.

## Current Versions

| Component | Version | Meaning |
|---|---:|---|
| Project | `v0.2.0` | Platform now includes validated production scoring helpers, iterative optimizer, shadow model, separate shadow matrix, and shadow HISTORY persistence. |
| Production Model | `v0.1.0` | Active live scoring model and current production weights. Shadow/test weights have not been promoted. |
| Stage | Development | Not a public/commercial release. |

Version values are defined in `Version.js`.

## Core Philosophy

The model is a controlled experimentation system, not a guaranteed pick generator.

Key principles:

1. **Permanent history matters more than daily results.** Losing days and losing streaks are expected. Evaluation must happen across large samples.
2. **Never delete tracked historical data.** Inactive features should remain in HISTORY so future backtests and machine learning can evaluate them later.
3. **Do not trust optimized weights until they survive out-of-sample testing.** Backtests can overfit. Shadow model tracking exists to test optimizer recommendations on future unseen games.
4. **Production and experiments must be separated.** The live model remains stable while shadow/test models evaluate candidate improvements.
5. **Flat-risk ROI is the standard comparison metric.** Current dashboards use flat $100 risk per model pick: loss = `-$100`; win profit is based on American odds; ROI = profit / total risk.

## Current Architecture

```text
Raw Data Imports
      |
      v
MODEL_MATRIX
      |
      |-- scored with production Weight column
      v
Live Model Outputs

MODEL_MATRIX
      |
      |-- copied to MODEL_MATRIX_SHADOW
      |-- rescored with Test Weight column
      v
Shadow Model Outputs

Live + Shadow Pregame Snapshot
      |
      v
HISTORY / HISTORY_ARCHIVE
      |
      v
Results Grading / Backtesting / Dashboards / Optimizer
```

## Main Sheets

| Sheet | Purpose |
|---|---|
| `Settings` | Editable model features, active flags, production weights, optimizer suggestions, and test weights. |
| `MODEL_MATRIX` | Current production model matrix and live predictions. |
| `MODEL_MATRIX_SHADOW` | Experimental matrix using the `Test Weight` column. Generated from `MODEL_MATRIX`. |
| `Today` | User-facing current-day picks/output. |
| `HISTORY` | Permanent append-only pregame records plus final results. Now stores both live and shadow predictions. |
| `HISTORY_ARCHIVE` | Mirror/archive of permanent historical records. |
| `OPTIMIZER_ITERATIVE` | Hill-climbing optimizer output. |
| `OPTIMIZER_DEBUG` | Optimizer diagnostic output and contribution breakdowns. |
| `OPTIMIZER_VALIDATION` | Confirms validation scorer matches production scorer. |
| `Performance Dashboard` | Live model performance reporting. |
| `Backtest Dashboard` | Backtest results for current/alternate weights. |

Sheet names may vary slightly by Apps Script file, but the canonical model matrix name is `MODEL_MATRIX`.

## Main Apps Script Files

| File | Purpose |
|---|---|
| `Main.js` | Automation entry points such as morning update, odds snapshot, pregame snapshot, and results/history update. |
| `Scoring.js` | Production scoring engine. Contains `scoreModelMatrix()`, `calculateEdgeStats()`, and `scoreModelRowWithSettings()`. |
| `History.js` | Pregame HISTORY snapshot and result grading. Persists live and shadow predictions. |
| `Shadow_Model.gs` | Builds `MODEL_MATRIX_SHADOW` and scores it with `Test Weight`. |
| `Optimizer_Iterative.gs` | Iterative hill-climbing optimizer using the production scoring helpers. |
| `Optimizer_Validation.gs` | Validates helper scoring against live `MODEL_MATRIX` scores/picks. |
| `Version.js` | Project/model version metadata. |

## Scoring Model

The live model uses edge columns from `MODEL_MATRIX`, normalizes each active edge by z-score, multiplies by feature weight, and divides total weighted score by total active weight.

The production scoring source of truth is:

```javascript
scoreModelRowWithSettings(row, headers, settings, edgeStats)
```

Both the validator and optimizer are intended to use this same production scoring helper. This reduces drift between live scoring, validation, and optimization.

## Live vs Shadow Model Workflow

The shadow model is the forward-test system for optimizer recommendations.

Workflow:

1. Run the optimizer.
2. Review suggested weights.
3. Copy candidate weights into `Test Weight` using `setupShadowModelTestWeights()`.
4. Daily automations build and score both `MODEL_MATRIX` and `MODEL_MATRIX_SHADOW`.
5. Pregame snapshot writes both live and shadow predictions to HISTORY.
6. After enough new games, compare live vs shadow out-of-sample results.
7. Promote shadow weights only if they outperform live on future games.

The shadow model does not affect production picks.

## Automation Flow

The main runners in `Main.js` now call `buildShadowModelMatrix()` immediately after `scoreModelMatrix()`.

Typical flow:

```text
Import / refresh raw data
Build MODEL_MATRIX
Score live model
Build and score MODEL_MATRIX_SHADOW
Build Today / dashboards
Pregame snapshot writes live + shadow to HISTORY
Results update grades completed games
```

Existing triggers should continue to call the same runner functions. Shadow scoring is embedded inside those runners.

## HISTORY Design

HISTORY is the permanent database for model evaluation.

It stores:

- Pregame model snapshot time
- Game ID/date/teams
- Live model pick, scores, confidence, odds, pick type, and flat-bet profit
- Shadow model pick, scores, confidence, odds, pick type, and flat-bet profit
- Full `MM_...` snapshot of `MODEL_MATRIX`
- Full `SM_...` snapshot of `MODEL_MATRIX_SHADOW`
- Final results and correctness after games complete

Existing rows should not be overwritten by snapshot logic. New schema fields are appended to the right as needed.

## Optimizer Status

The current iterative optimizer:

- Uses production scoring helpers.
- Starts from current weights.
- Tests single-feature weight changes.
- Accepts the best meaningful improvement each round.
- Repeats until no meaningful improvement remains or max rounds are reached.
- Writes suggestions into `Suggested Weight`.

Current limitation:

- It can still overfit because training and evaluation currently use the same historical sample.

Next optimizer milestone:

- Add train/test validation so candidate weights are optimized on older games and evaluated on newer unseen games.

## Betting Math Standard

The project uses flat-risk ROI for model evaluation:

- Risk `$100` on every model pick.
- If the pick loses, profit is `-$100`.
- If the pick wins, profit is calculated from the American moneyline.
- ROI = total profit / total dollars risked.

Do not change this calculation without explicitly reviewing the effect across all dashboards/backtests.

## Current Development Priorities

1. Build `SHADOW_DASHBOARD` comparing live vs shadow performance.
2. Add train/test validation to the iterative optimizer.
3. Add live-vs-shadow confidence bucket comparison.
4. Audit HISTORY to confirm every tracked feature is preserved for future Python/ML work.
5. Continue per-game pregame snapshot gatekeeper work.
6. Improve model-manager/version tracking for experiments.
7. Prepare for future Python/database migration.

See `ROADMAP.md` for the longer roadmap.

## Near-Term Manual Workflow

Daily work pattern:

1. Morning automation refreshes data and generates live + shadow predictions.
2. Lunch review checks results, HISTORY integrity, and next coding priorities.
3. Evening session implements code/documentation updates in GitHub.
4. Pull updates into Apps Script with `clasp pull`.
5. Run controlled tests before relying on new production behavior.

## Important Warnings

- Do not judge the model by one day or one week.
- Do not use parlay results to evaluate model quality.
- Do not reset HISTORY unless intentionally rebuilding the database.
- Do not promote optimizer weights directly to production without shadow/out-of-sample testing.
- Do not remove historical columns for features that may be useful later.

## Long-Term Direction

The current system is evolving from a spreadsheet model into a model development platform:

- Automated data ingestion
- Versioned scoring
- Permanent historical snapshots
- Backtesting
- Iterative optimization
- Shadow model A/B testing
- Future Python/database/ML optimization

The eventual goal is a controlled, repeatable pipeline where candidate models are trained, tested, shadowed, and promoted only when they demonstrate durable out-of-sample value.
