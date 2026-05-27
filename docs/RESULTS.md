# RESULTS

## Main Results
- Warm-ft baseline: `0.4332 / 0.5575 / 0.7731`
- WCT-v0: `0.4389 / 0.5599 / 0.7724`
- AWCT-TempSeg (seed1): `0.4393 / 0.5735 / 0.7633`
- AWCT-TempSeg (seed2): `0.4454 / 0.5690 / 0.7665`

(metric order: `mIoU / mAcc / allAcc`)

## Observation
AWCT-TempSeg improves mIoU and mAcc over warm-start baseline, and outperforms fixed WCT-v0 in the final seed-level summary.

## Final Recommended Config
- `configs/semanticstf/semseg-pt-v3m1-0-tempseg-v2-awct-v11-conservative.py`
