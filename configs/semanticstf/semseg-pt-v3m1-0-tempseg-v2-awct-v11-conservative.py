_base_ = ["./semseg-pt-v3m1-0-tempseg-v2.py"]

# runtime
save_path = "exp/semanticstf/semseg-pt-v2-awct-v11-conservative-warmft-e5"
seed = 13251804
weight = None

# warm-start fine-tune schedule (same as AWCT-v1)
epoch = 5
eval_epoch = 1
optimizer = dict(type="AdamW", lr=2e-4, weight_decay=0.005)
scheduler = dict(
    type="OneCycleLR",
    max_lr=[2e-4, 2e-5],
    pct_start=0.04,
    anneal_strategy="cos",
    div_factor=10.0,
    final_div_factor=100.0,
)

# keep model unchanged
model = dict(type="TempSegV2Segmentor")

# conservative validation-guided adaptive weather curriculum trainer
train = dict(
    type="ValidationGuidedWeatherCurriculumTrainer",
    weather_sampling=dict(
        enable=True,
        # AWCT-v1.1 controls
        tau=0.5,
        beta=0.15,
        ema_momentum=0.8,
        min_ratio=0.10,
        max_ratio=0.45,
        base_distribution=dict(
            snow=0.1852,
            light_fog=0.2222,
            dense_fog=0.2222,
            rain=0.3704,
        ),
    ),
)

data = dict(
    train=dict(
        loop=5,
    )
)
