#!/usr/bin/env python3
"""
Weather-wise evaluation for SemanticSTF from saved prediction files.

This script does not run model inference. It reads *_pred.npy files produced by
Pointcept tester, loads GT labels, remaps labels to 19 train IDs, and computes
confusion-matrix metrics per weather and overall.
"""

import argparse
import json
import os
from collections import OrderedDict, defaultdict

import numpy as np


CLASS_NAMES = [
    "car",
    "bicycle",
    "motorcycle",
    "truck",
    "other-vehicle",
    "person",
    "bicyclist",
    "motorcyclist",
    "road",
    "parking",
    "sidewalk",
    "other-ground",
    "building",
    "fence",
    "vegetation",
    "trunk",
    "terrain",
    "pole",
    "traffic-sign",
]

SAFETY_CLASSES = [
    "car",
    "truck",
    "other-vehicle",
    "person",
    "bicycle",
    "bicyclist",
    "pole",
    "traffic-sign",
    "parking",
    "road",
    "vegetation",
]

WEATHER_ORDER = ["snow", "light_fog", "dense_fog", "rain"]
IGNORE_INDEX = -1
NUM_CLASSES = 19


def parse_args():
    parser = argparse.ArgumentParser("SemanticSTF weather-wise evaluator")
    parser.add_argument(
        "--data-root",
        default=os.getenv("SEMANTICSTF_ROOT", "/path/to/SemanticSTF"),
        help="SemanticSTF root path.",
    )
    parser.add_argument("--split", default="val", help="split to evaluate")
    parser.add_argument(
        "--model",
        action="append",
        required=True,
        help="model spec in form name=/abs/or/rel/result_dir",
    )
    parser.add_argument(
        "--output-json",
        default=None,
        help="optional output json path for full report",
    )
    return parser.parse_args()


def build_learning_map(ignore_index=IGNORE_INDEX):
    learning_map = {0: ignore_index, 20: ignore_index}
    for raw_id in range(1, 20):
        learning_map[raw_id] = raw_id - 1
    return learning_map


def remap_labels(raw_label, learning_map):
    mapped = np.full(raw_label.shape, IGNORE_INDEX, dtype=np.int32)
    for src, dst in learning_map.items():
        mapped[raw_label == src] = dst
    return mapped


def load_weather_map(data_root, split):
    split_file = os.path.join(data_root, split, f"{split}.txt")
    if not os.path.isfile(split_file):
        raise FileNotFoundError(f"split file not found: {split_file}")

    mapping = OrderedDict()
    with open(split_file, "r", encoding="utf-8") as f:
        for line in f:
            line = line.strip()
            if not line:
                continue
            if "," in line:
                frame_id, weather = [x.strip() for x in line.split(",", 1)]
            else:
                frame_id, weather = line, "unknown"
            mapping[frame_id] = weather or "unknown"
    return mapping


def parse_models(model_specs):
    out = OrderedDict()
    for spec in model_specs:
        if "=" not in spec:
            raise ValueError(f"invalid --model spec: {spec}")
        name, path = spec.split("=", 1)
        name, path = name.strip(), path.strip()
        if not name:
            raise ValueError(f"invalid model name in spec: {spec}")
        out[name] = path
    return out


def ensure_result_dir(path):
    # allow passing experiment root or result dir directly
    result_dir = path
    if os.path.isdir(path) and os.path.basename(path) != "result":
        probe = os.path.join(path, "result")
        if os.path.isdir(probe):
            result_dir = probe
    if not os.path.isdir(result_dir):
        raise FileNotFoundError(f"result dir not found: {result_dir}")
    return result_dir


def new_bucket():
    return {
        "cm": np.zeros((NUM_CLASSES, NUM_CLASSES), dtype=np.int64),
        "samples": 0,
        "valid_points": 0,
        "total_points": 0,
    }


def update_confusion(cm, gt, pred):
    mask = gt != IGNORE_INDEX
    gt = gt[mask]
    pred = pred[mask]
    if gt.size == 0:
        return 0
    pred = np.clip(pred.astype(np.int64), 0, NUM_CLASSES - 1)
    gt = gt.astype(np.int64)
    binc = np.bincount(gt * NUM_CLASSES + pred, minlength=NUM_CLASSES * NUM_CLASSES)
    cm += binc.reshape(NUM_CLASSES, NUM_CLASSES)
    return int(gt.size)


def compute_metrics(cm):
    cm = cm.astype(np.float64)
    intersection = np.diag(cm)
    gt_sum = cm.sum(axis=1)
    pred_sum = cm.sum(axis=0)
    union = gt_sum + pred_sum - intersection

    iou = intersection / (union + 1e-10)
    acc = intersection / (gt_sum + 1e-10)
    m_iou = float(np.mean(iou))
    m_acc = float(np.mean(acc))
    all_acc = float(intersection.sum() / (gt_sum.sum() + 1e-10))

    per_class_iou = {CLASS_NAMES[i]: float(iou[i]) for i in range(NUM_CLASSES)}
    return {
        "mIoU": m_iou,
        "mAcc": m_acc,
        "allAcc": all_acc,
        "per_class_iou": per_class_iou,
    }


def evaluate_one_model(data_root, split, weather_map, result_dir):
    learning_map = build_learning_map()
    buckets = defaultdict(new_bucket)

    for sample_name, weather in weather_map.items():
        pred_path = os.path.join(result_dir, f"{sample_name}_pred.npy")
        label_path = os.path.join(data_root, split, "labels", f"{sample_name}.label")

        if not os.path.isfile(pred_path):
            raise FileNotFoundError(f"missing prediction: {pred_path}")
        if not os.path.isfile(label_path):
            raise FileNotFoundError(f"missing label: {label_path}")

        pred = np.load(pred_path).reshape(-1)
        raw_label = np.fromfile(label_path, dtype=np.int32).reshape(-1)
        gt = remap_labels(raw_label, learning_map)

        # keep exact comparable length if minor mismatch exists
        n = min(pred.shape[0], gt.shape[0])
        pred = pred[:n]
        gt = gt[:n]

        valid_points = update_confusion(buckets[weather]["cm"], gt, pred)
        buckets[weather]["samples"] += 1
        buckets[weather]["valid_points"] += valid_points
        buckets[weather]["total_points"] += int(n)

        valid_points_o = update_confusion(buckets["overall"]["cm"], gt, pred)
        buckets["overall"]["samples"] += 1
        buckets["overall"]["valid_points"] += valid_points_o
        buckets["overall"]["total_points"] += int(n)

    # finalize metrics
    report = OrderedDict()
    for weather in WEATHER_ORDER + ["overall"]:
        if weather not in buckets:
            continue
        b = buckets[weather]
        metrics = compute_metrics(b["cm"])
        report[weather] = {
            "samples": b["samples"],
            "valid_points": b["valid_points"],
            "total_points": b["total_points"],
            **metrics,
        }

    # include unknown weathers if any
    for weather, b in buckets.items():
        if weather in report:
            continue
        metrics = compute_metrics(b["cm"])
        report[weather] = {
            "samples": b["samples"],
            "valid_points": b["valid_points"],
            "total_points": b["total_points"],
            **metrics,
        }

    return report


def print_overall_sanity(all_reports):
    print("\n=== Overall sanity (should match existing final test logs) ===")
    for name, rep in all_reports.items():
        ov = rep["overall"]
        print(
            f"{name:>16}: mIoU={ov['mIoU']:.4f} mAcc={ov['mAcc']:.4f} allAcc={ov['allAcc']:.4f} "
            f"samples={ov['samples']} valid_points={ov['valid_points']}"
        )


def print_weather_main_table(all_reports):
    print("\n=== Weather-wise main metrics ===")
    header = "model,weather,samples,valid_points,mIoU,mAcc,allAcc"
    print(header)
    for name, rep in all_reports.items():
        for weather in WEATHER_ORDER + ["overall"]:
            if weather not in rep:
                continue
            r = rep[weather]
            print(
                f"{name},{weather},{r['samples']},{r['valid_points']},"
                f"{r['mIoU']:.4f},{r['mAcc']:.4f},{r['allAcc']:.4f}"
            )


def print_weather_safety_table(all_reports):
    print("\n=== Weather-wise safety-critical class IoU ===")
    print("model,weather," + ",".join(SAFETY_CLASSES))
    for name, rep in all_reports.items():
        for weather in WEATHER_ORDER + ["overall"]:
            if weather not in rep:
                continue
            pci = rep[weather]["per_class_iou"]
            vals = [f"{pci[c]:.4f}" for c in SAFETY_CLASSES]
            print(f"{name},{weather}," + ",".join(vals))


def print_delta_wct_vs_base(all_reports, base_name, target_name):
    if base_name not in all_reports or target_name not in all_reports:
        return

    base = all_reports[base_name]
    tgt = all_reports[target_name]

    print(f"\n=== Delta: {target_name} - {base_name} (weather main metrics) ===")
    print("weather,dmIoU,dmAcc,dallAcc")
    for weather in WEATHER_ORDER + ["overall"]:
        if weather not in base or weather not in tgt:
            continue
        db = tgt[weather]
        bb = base[weather]
        print(
            f"{weather},{db['mIoU']-bb['mIoU']:+.4f},{db['mAcc']-bb['mAcc']:+.4f},{db['allAcc']-bb['allAcc']:+.4f}"
        )

    print(f"\n=== Delta: {target_name} - {base_name} (safety-critical class IoU) ===")
    print("weather," + ",".join(SAFETY_CLASSES))
    for weather in WEATHER_ORDER + ["overall"]:
        if weather not in base or weather not in tgt:
            continue
        row = [weather]
        for c in SAFETY_CLASSES:
            dv = tgt[weather]["per_class_iou"][c] - base[weather]["per_class_iou"][c]
            row.append(f"{dv:+.4f}")
        print(",".join(row))


def main():
    args = parse_args()
    weather_map = load_weather_map(args.data_root, args.split)
    model_map = parse_models(args.model)

    all_reports = OrderedDict()
    for name, model_path in model_map.items():
        result_dir = ensure_result_dir(model_path)
        report = evaluate_one_model(
            data_root=args.data_root,
            split=args.split,
            weather_map=weather_map,
            result_dir=result_dir,
        )
        all_reports[name] = report

    print_overall_sanity(all_reports)
    print_weather_main_table(all_reports)
    print_weather_safety_table(all_reports)

    # default delta focus: wct-v0 vs warmft
    if "warmft" in all_reports and "wct_v0" in all_reports:
        print_delta_wct_vs_base(all_reports, "warmft", "wct_v0")

    if args.output_json:
        payload = {
            "data_root": args.data_root,
            "split": args.split,
            "class_names": CLASS_NAMES,
            "weather_order": WEATHER_ORDER,
            "reports": all_reports,
        }
        with open(args.output_json, "w", encoding="utf-8") as f:
            json.dump(payload, f, indent=2)
        print(f"\nSaved json report to: {args.output_json}")


if __name__ == "__main__":
    main()
