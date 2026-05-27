# AWCT-TempSeg
Research Release: v0.1.0

Adaptive Weather-Curriculum Temporal Segmentation for Adverse-Weather LiDAR Point Clouds.

If you use this code, please cite:

AWCT-TempSeg: Validation-Guided Adaptive Weather-Curriculum Learning for Robust Temporal LiDAR Semantic Segmentation.
DOI: https://doi.org/10.5072/zenodo.505066
Code: https://github.com/zxft0217/AWCT-TempSeg

## Introduction
AWCT-TempSeg is a temporal LiDAR semantic segmentation framework for adverse-weather scenarios. It keeps the segmentation backbone unchanged and improves robustness through a validation-guided adaptive weather curriculum during warm-start fine-tuning.

## Method Overview
- Built on a strong temporal segmentation backbone (`TempSegV2Segmentor`).
- Uses validation weather-wise mIoU to estimate weather difficulty.
- Updates weather sampling ratios conservatively with EMA-smoothed difficulty.
- Preserves a base weather distribution and applies bounded adaptive updates.
- Improves adverse-weather segmentation, especially rain and snow conditions.

## Key Features
- Temporal LiDAR semantic segmentation
- SemanticSTF support
- Validation-guided adaptive weather curriculum
- Weather-wise evaluation utilities
- Warm-start fine-tuning from TempSeg-v2 checkpoints

## Installation
```bash
conda create -n awct-tempseg python=3.10 -y
conda activate awct-tempseg
pip install -r requirements.txt
```

If you need CUDA-specific `torch-scatter` / `torch-cluster` wheels, follow the PyG installation guide matching your PyTorch/CUDA version.

## Dataset Preparation
Prepare SemanticSTF under a custom path, for example:

```text
/path/to/SemanticSTF
```

Then either:
1. Set env var `SEMANTICSTF_ROOT=/path/to/SemanticSTF`, or
2. Override in command line via `--options data_root=/path/to/SemanticSTF`.

## Training (AWCT-TempSeg)
```bash
PYTHONPATH=$PWD python tools/train.py \
  --config-file configs/semanticstf/semseg-pt-v3m1-0-tempseg-v2-awct-v11-conservative.py \
  --num-gpus 1 \
  --options \
  save_path=exp/semanticstf/awct-tempseg \
  data_root=/path/to/SemanticSTF \
  weight=/path/to/tempsegv2_model_best.pth
```

## Evaluation
```bash
PYTHONPATH=$PWD python tools/test.py \
  --config-file configs/semanticstf/semseg-pt-v3m1-0-tempseg-v2-awct-v11-conservative.py \
  --num-gpus 1 \
  --options \
  data_root=/path/to/SemanticSTF \
  save_path=exp/semanticstf/awct-tempseg \
  weight=exp/semanticstf/awct-tempseg/model/model_best.pth
```

## Weather-wise Evaluation
`eval_semanticstf_weatherwise.py` reads saved prediction files in `result/`.

```bash
PYTHONPATH=$PWD python tools/analysis/eval_semanticstf_weatherwise.py \
  --data-root /path/to/SemanticSTF \
  --split val \
  --model awct=exp/semanticstf/awct-tempseg \
  --output-json exp/semanticstf/awct-tempseg/weatherwise_eval.json
```

## Results
| Method | mIoU | mAcc | allAcc |
|---|---:|---:|---:|
| TempSeg baseline | 0.4332 | 0.5275 | 0.7531 |
| TempSeg loss-adaptive | 0.4389 | 0.5599 | 0.7724 |
| RetSeg3D AWCT | 0.5571 | 0.7134 | 0.8462 |
| AWCT-TempSeg  | 0.4889 | 0.6422 | 0.7975 |

AWCT improves mIoU and mAcc overall, with notable gains in rain and snow scenarios.

## License and Acknowledgment
This project retains the upstream open-source license in `LICENSE` and is adapted from the Pointcept framework with AWCT-specific modifications for SemanticSTF.

## Citation
```bibtex
@article{awcttempseg2026,
  title={AWCT-TempSeg: Adaptive Weather-Curriculum Temporal Segmentation for Adverse-Weather LiDAR Point Clouds},
  author={TODO},
  journal={TODO},
  year={2026}
}
```
