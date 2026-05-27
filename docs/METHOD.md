# METHOD

## AWCT-TempSeg (Final Protocol)
AWCT-TempSeg keeps the TempSeg-v2 model structure unchanged and only adapts training-time weather sampling.

### Backbone and Head
- Model: `TempSegV2Segmentor`
- Backbone: unchanged PTv3 backbone from TempSeg-v2
- Segmentation loss: unchanged from TempSeg-v2

### Training Strategy
- Warm-start fine-tuning from a trained TempSeg-v2 checkpoint
- Fixed protocol:
  - `epoch=5`
  - `eval_epoch=1`
  - `optimizer.lr=2e-4`
  - `scheduler.max_lr=[2e-4, 2e-5]`

### Adaptive Weather Curriculum
Base distribution:
- snow: 0.1852
- light_fog: 0.2222
- dense_fog: 0.2222
- rain: 0.3704

Validation-guided update:
1. Compute weather-wise validation mIoU.
2. Difficulty:
   `d_w^t = 1 - mIoU_val^t(w)`.
3. EMA smoothing:
   `D_w^t = mu * D_w^{t-1} + (1 - mu) * d_w^t`.
4. Difficulty distribution:
   `q_w^t = softmax(D_w^t / tau)`.
5. Conservative mixing with base distribution:
   `p_w^t = (1 - beta) * p_base,w + beta * q_w^t`.
6. Bounded normalization:
   `p_w^t = normalize(clamp(p_w^t, p_min, p_max))`.

Final conservative hyperparameters:
- `tau=0.5`
- `beta=0.15`
- `mu=0.8` (EMA momentum)
- `p_min=0.10`
- `p_max=0.45`
