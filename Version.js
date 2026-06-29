/******************************************************************************
 * VERSION.gs
 * Single source of truth for project metadata.
 ******************************************************************************/

const MODEL_NAME = "MLB Model";
const MODEL_STAGE = "Development";
const PROJECT_VERSION = "v0.1.3";
const MODEL_VERSION = "v0.0.1";

function getModelName() {
  return MODEL_NAME;
}

function getModelStage() {
  return MODEL_STAGE;
}

function getModelVersion() {
  return MODEL_VERSION;
}