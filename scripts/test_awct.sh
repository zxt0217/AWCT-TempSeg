#!/usr/bin/env bash
set -euo pipefail

CFG=${1:-configs/semanticstf/semseg-pt-v3m1-0-tempseg-v2-awct-v11-conservative.py}
SAVE_PATH=${2:-exp/semanticstf/awct-tempseg}
DATA_ROOT=${3:-/path/to/SemanticSTF}
PRETRAIN=${4:-exp/semanticstf/awct-tempseg/model/model_best.pth}

PYTHONPATH=$PWD python tools/test.py \
  --config-file "$CFG" \
  --num-gpus 1 \
  --options \
  save_path="$SAVE_PATH" \
  data_root="$DATA_ROOT" \
  weight="$PRETRAIN"
