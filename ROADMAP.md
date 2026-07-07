# MLB Model Roadmap

This roadmap tracks the project as a modeling platform, not just a pick generator.

## Current State: Project v0.2.0 / Model v0.1.0

The project now includes:

- Automated data refreshes through Google Sheets and Apps Script.
- Production scoring through `MODEL_MATRIX`.
- Shared production scoring helper in `Scoring.js`.
- Permanent pregame HISTORY snapshots.
- Result grading and flat-bet ROI calculations.
- Iterative hill-climbing optimizer.
- Optimizer validation against production scoring.
- Shadow model infrastructure.
- Separate `MODEL_MATRIX_SHADOW` sheet.
- Shadow prediction persistence in HISTORY.
- Version metadata in `Version.js`.

The production model remains `v0.1.0` because candidate shadow weights have not yet been promoted.

---

## Phase 0 — Foundation

Status: Mostly complete.

Goal: Build a working automated MLB model in Google Sheets.

Completed:

- Schedule import.
- Team offense/pitching imports.
- Starting pitcher data.
- Bullpen data.
- Recent form.
- Home/away splits.
- Park factors.
- Weather.
- Odds import.
- Model matrix.
- Scoring.
- Today view.
- Dashboard reporting.
- HISTORY database.
- Results import and grading.

Remaining:

- Continue hardening trigger reliability.
- Continue checking that all raw data refreshes are stable.

---

## Phase 1 — Permanent Model Evaluation

Status: Active / mostly complete.

Goal: Make sure every pregame prediction and every evaluated outcome is stored permanently.

Completed:

- HISTORY snapshots of full `MODEL_MATRIX` as `MM_...` columns.
- HISTORY_ARCHIVE mirror.
- Result grading.
- Flat-bet profit.
- Live model correctness tracking.
- Shadow model correctness/profit tracking added to HISTORY.
- Full shadow snapshot stored as `SM_...` columns.

Next:

- Audit HISTORY after multiple new game days to verify live and shadow rows populate correctly.
- Confirm old rows remain untouched when new schema columns are appended.
- Add experiment/model identifiers to each snapshot.
- Verify every future feature automatically persists in HISTORY.

---

## Phase 2 — Shadow Model A/B Testing

Status: Newly implemented.

Goal: Compare candidate optimizer weights against the live model on future unseen games.

Completed:

- `Test Weight` column on Settings.
- `MODEL_MATRIX_SHADOW` generation.
- Shadow scoring using production scoring helpers.
- Shadow predictions written permanently to HISTORY.
- Daily automation calls `buildShadowModelMatrix()` after live scoring.

Next:

- Build `SHADOW_DASHBOARD`.
- Compare live vs shadow:
  - Games
  - Wins/losses
  - Win %
  - Profit
  - ROI
  - Confidence buckets
  - Favorite/underdog performance
  - Home/away performance
  - Rolling performance
- Establish promotion criteria.

Promotion rule concept:

```text
Only promote shadow weights if they outperform live weights on future games that were not used by the optimizer.
```

---

## Phase 3 — Optimizer Reliability

Status: Active.

Goal: Make optimizer recommendations less prone to overfitting.

Completed:

- Iterative hill-climbing optimizer.
- Production scoring integration.
- Optimizer validation against live scoring.
- Suggested weights written to Settings.

Current limitation:

- The optimizer currently trains and evaluates on the same HISTORY sample.

Next:

1. Add train/test split:
   - Train on older games.
   - Validate on newer games.
2. Add optimizer output columns:
   - Training games
   - Training ROI
   - Validation games
   - Validation ROI
   - Validation win %
   - Delta vs baseline
3. Add time-series validation:
   - Train on games 1-N.
   - Validate on later games only.
4. Add pairwise testing after train/test validation is stable.
5. Add safeguards against low-sample confidence bucket overfitting.

Future optimizer methods:

- Pairwise search.
- Beam search.
- Random restarts.
- Genetic search.
- Bayesian optimization.
- Python-based ML feature selection.

---

## Phase 4 — Dashboard Expansion

Status: Planned.

Goal: Make performance review fast enough for daily lunch-break checks.

Next dashboard priorities:

1. `SHADOW_DASHBOARD`.
2. Live vs Shadow performance dashboard.
3. Confidence bucket comparison.
4. Rolling 7/14/30-game model comparison.
5. Favorite/underdog split comparison.
6. Model-strength/ROI calibration.
7. Optimizer recommendation summary.
8. Promotion readiness indicator.

Dashboard should answer:

- Is shadow beating live?
- Is the edge meaningful or noise?
- Is the improvement concentrated in one bucket?
- Is shadow better on favorites, underdogs, high-confidence plays, or all picks?
- Is the sample large enough to consider promotion?

---

## Phase 5 — Pregame Snapshot Gatekeeper

Status: Planned / important.

Goal: Move from broad daily snapshots to true per-game pregame snapshots.

Desired behavior:

- Check every scheduled game independently.
- Snapshot only games inside the pregame window.
- Avoid duplicate rows by Game ID.
- Preserve final pregame state before first pitch.
- Eventually support lineups and confirmed starting pitchers.

Target function:

```javascript
runPregameSnapshotGatekeeper()
```

Gatekeeper logic:

```text
For each game:
  if game has no HISTORY row
  and current time is within configured pregame window
  and required fields are present
  then append snapshot
```

This remains important because lineups and starters may become available at different times for different games.

---

## Phase 6 — Lineup Strength Integration

Status: Planned.

Goal: Add starting lineup quality before first pitch.

Approach:

1. Build lineup scraper as a separate project/module.
2. Pull confirmed lineups and starting pitchers.
3. Convert batting order to lineup strength metrics.
4. Write lineup fields into raw tabs.
5. Add lineup metrics to `MODEL_MATRIX`.
6. Preserve lineup fields in HISTORY.
7. Test lineup features through optimizer/backtesting.

Important rule:

```text
Do not reset HISTORY to add lineup data. Add new columns and preserve old rows.
```

---

## Phase 7 — Market Signal / CLV Tracking

Status: Planned.

Goal: Determine whether the model is beating the betting market, not just picking winners.

Planned metrics:

- Opening line.
- Pregame line.
- Closing line.
- Model pick line.
- Closing-line value.
- Implied probability movement.
- Favorite/underdog line movement.
- Market agreement/disagreement.

Long-term value:

- If the model consistently beats closing line, that is stronger evidence of edge than short-term win rate alone.

---

## Phase 8 — Python / Database Migration

Status: Planned.

Goal: Move historical data and optimizer work out of Google Sheets once sample size grows.

Likely path:

1. Export HISTORY to a database.
2. Preserve every pregame row exactly as captured.
3. Use Python for:
   - Larger search spaces
   - ML models
   - Feature importance
   - Cross-validation
   - Robust train/test splits
   - Model comparison
4. Keep Google Sheets as the daily operating interface initially.
5. Eventually support suggested weights being written back next to current weights.

Database requirements:

- Append-only historical records.
- Versioned model metadata.
- Feature schema tracking.
- Game-level predictions.
- Odds snapshots.
- Results and grading.
- Experiment IDs.

---

## Phase 9 — Model Manager

Status: Planned.

Goal: Move from hardcoded live/shadow models to tracked model generations.

Concept sheet: `MODELS`

| Model ID | Version | Status | Weight Column | Games | Win % | ROI |
|---|---|---|---|---:|---:|---:|
| LIVE | v0.1.0 | Active | Weight | 129 | 53.5% | -5.0% |
| SHADOW_A | v0.1.1-test | Testing | Test Weight | 20 | TBD | TBD |
| ARCHIVED_001 | v0.0.9 | Archived | N/A | 100 | 52.0% | -2.0% |

Future capabilities:

- Promote shadow to live.
- Archive failed experiments.
- Track model lineage.
- Compare multiple candidate models.
- Store experiment status.

---

## Phase 10 — Public Product / Discord Layer

Status: Deferred.

Goal: Package model output for external users only after internal validation is stronger.

Possible future features:

- Discord pick posts.
- One channel per game or per confidence tier.
- Read-this-first risk disclaimer.
- Straight-bet focus.
- No parlay framing.
- Model vs market win probability explanation.
- Public update notes.

Current stance:

- Lineup strength, optimizer validation, and shadow testing are more important than public distribution.
- Picks should not be sold until the model has stronger out-of-sample evidence.

---

## Version Milestones

### Project v0.2.x

Focus: Controlled experimentation infrastructure.

Targets:

- Shadow dashboard.
- Train/test optimizer.
- HISTORY audit.
- Confidence calibration.
- Pregame gatekeeper.

### Project v0.3.x

Focus: Data quality and feature expansion.

Targets:

- Lineup strength.
- Confirmed starting lineup integration.
- Better odds history.
- CLV tracking.

### Project v0.4.x

Focus: More robust optimization.

Targets:

- Train/test and rolling validation.
- Pairwise optimizer.
- Better model promotion workflow.
- Versioned experiment tracking.

### Project v0.5.x

Focus: Python/database bridge.

Targets:

- HISTORY export.
- Database schema.
- Python backtesting.
- ML-ready dataset.

### Project v1.0.0

Potential criteria:

- Stable data pipeline.
- No known HISTORY integrity issues.
- Model has 1,000+ historical games.
- Shadow model system has promoted at least one out-of-sample improvement.
- Backtesting and validation agree.
- Betting math is consistent across dashboards.
- Manual intervention is minimal.

---

## Current Highest Priority

The next coding milestone should be:

```text
Build SHADOW_DASHBOARD.
```

Reason:

The project now collects live and shadow predictions permanently. The missing piece is a clear reporting layer to compare them and decide whether shadow weights are outperforming live weights out of sample.
