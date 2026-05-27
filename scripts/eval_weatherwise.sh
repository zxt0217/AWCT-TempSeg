#!/usr/bin/env bash
set -euo pipefail

DATA_ROOT=${1:-/path/to/SemanticSTF}
RESULT_DIR=${2:-exp/semanticstf/awct-tempseg}
OUT_JSON=${3:-exp/semanticstf/awct-tempseg/weatherwise_eval.json}

PYTHONPATH=$PWD python tools/analysis/eval_semanticstf_weatherwise.py \
  --data-root "$DATA_ROOT" \
  --split val \
  --model awct="$RESULT_DIR" \
  --output-json "$OUT_JSON"
