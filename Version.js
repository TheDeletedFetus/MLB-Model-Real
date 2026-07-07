/******************************************************************************
 * VERSION.js
 * Single source of truth for project metadata.
 *
 * Versioning notes:
 * - PROJECT_VERSION tracks platform/infrastructure capability.
 * - MODEL_VERSION tracks the active production scoring/weight model.
 *
 * Current status:
 * - Project v0.2.0 reflects validated production scoring helpers,
 *   iterative optimizer, shadow model, MODEL_MATRIX_SHADOW, and shadow HISTORY persistence.
 * - Model v0.1.0 remains the active live model because shadow/test weights have not
 *   been promoted into production weights yet.
 ******************************************************************************/

const MODEL_NAME = "MLB Model";
const MODEL_STAGE = "Development";
const PROJECT_VERSION = "v0.2.0";
const MODEL_VERSION = "v0.1.0";

function getModelName() {
  return MODEL_NAME;
}

function getModelStage() {
  return MODEL_STAGE;
}

function getProjectVersion() {
  return PROJECT_VERSION;
}

function getModelVersion() {
  return MODEL_VERSION;
}

function getProjectMetadata() {
  return {
    modelName: MODEL_NAME,
    modelStage: MODEL_STAGE,
    projectVersion: PROJECT_VERSION,
    modelVersion: MODEL_VERSION
  };
}
