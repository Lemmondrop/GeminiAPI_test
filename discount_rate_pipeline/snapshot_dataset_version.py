"""
Snapshot current experiment artifacts into test_data/<version>/.

This is intentionally copy-only. It does not delete or move root-level files.
Use it to preserve a dataset/model state before running the next experiment.
"""

import argparse
import hashlib
import json
import shutil
from datetime import datetime
from pathlib import Path

import pandas as pd


DEFAULT_GROUPS = {
    "01_raw": [
        "peer_group_discount_rates.csv",
        "discount_rate_table_pages_check.csv",
        "peer_group_discount_rates_expanded.csv",
    ],
    "02_phase_outputs": [
        "peer_group_expanded_with_corp_code.csv",
        "peer_group_expanded_with_rcpno.csv",
        "peer_group_expanded_features_raw.csv",
        "peer_group_expanded_features_enriched.csv",
    ],
    "03_training": [
        "training_features_expanded.csv",
        "training_features_expanded_segmented.csv",
        "training_features_expanded_general.csv",
        "training_features_expanded_general_excluded.csv",
    ],
    "04_models": [
        "model_general_expanded.pkl",
        "model_input_limited_expanded.pkl",
        "feature_importance_general_expanded.csv",
        "feature_importance_input_limited_expanded.csv",
        "predictions_cv_general_expanded.csv",
        "predictions_cv_input_limited_expanded.csv",
        "model_metrics_general_expanded.csv",
        "model_metrics_input_limited_expanded.csv",
    ],
}


def sha256_file(path: Path) -> str:
    h = hashlib.sha256()
    with path.open("rb") as f:
        for chunk in iter(lambda: f.read(1024 * 1024), b""):
            h.update(chunk)
    return h.hexdigest()


def csv_rows(path: Path) -> int | None:
    if path.suffix.lower() != ".csv":
        return None
    try:
        return int(len(pd.read_csv(path, encoding="utf-8-sig")))
    except Exception:
        return None


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("--version", required=True, help="folder name under test_data")
    parser.add_argument("--root", default="test_data")
    parser.add_argument("--extra-files", nargs="*", default=[])
    args = parser.parse_args()

    base_dir = Path(args.root) / args.version
    base_dir.mkdir(parents=True, exist_ok=True)

    manifest = {
        "version": args.version,
        "created_at": datetime.now().isoformat(timespec="seconds"),
        "root": base_dir.as_posix(),
        "files": [],
    }

    groups = {k: list(v) for k, v in DEFAULT_GROUPS.items()}
    if args.extra_files:
        groups.setdefault("99_extra", []).extend(args.extra_files)

    for group, files in groups.items():
        group_dir = base_dir / group
        group_dir.mkdir(parents=True, exist_ok=True)
        for file_name in files:
            src = Path(file_name)
            if not src.exists():
                continue
            dst = group_dir / src.name
            shutil.copy2(src, dst)
            manifest["files"].append(
                {
                    "group": group,
                    "source": src.as_posix(),
                    "dest": dst.as_posix(),
                    "bytes": dst.stat().st_size,
                    "sha256": sha256_file(dst),
                    "rows": csv_rows(dst),
                }
            )

    manifest_path = base_dir / "manifest.json"
    manifest_path.write_text(
        json.dumps(manifest, ensure_ascii=False, indent=2),
        encoding="utf-8",
    )

    print(f"Snapshot saved: {base_dir}")
    print(f"Files copied: {len(manifest['files'])}")
    print(f"Manifest: {manifest_path}")


if __name__ == "__main__":
    main()
