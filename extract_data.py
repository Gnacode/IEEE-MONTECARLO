#!/usr/bin/env python3
"""
extract_results.py - Compile per-detector metrics from current run + recent backups.

Reads simulation_results.json from:
  - U:/MONTECARLO/data/                      (current run)
  - U:/MONTECARLO/backups/<timestamp>/data/  (N most recent backups, default 3)

For each backup, identifies the run by reading the JSON's _simulation_parameters
(SNR, preset, etc.) and the top-level numeric key (num_nodes). Compiles all runs
into a single comparison table — one column per (num_nodes, SNR) configuration,
one row per metric × detector.

Outputs:
  - Console table (markdown-like, also pretty printed)
  - CSV file at <output_dir>/compiled_results.csv
  - JSON file at <output_dir>/compiled_results.json (raw structured form)

Usage
-----
    # Default: data/ (current) + 3 most recent backups
    python extract_results.py

    # Specify number of backups to include
    python extract_results.py --num-backups 5

    # Different backup root or current data dir
    python extract_results.py --data-dir U:/MONTECARLO/data \
                              --backup-root U:/MONTECARLO/backups \
                              --num-backups 3

    # Specify output location for CSV/JSON
    python extract_results.py --output-dir U:/MONTECARLO/compiled

Author: GNACODE INC, January 2026
"""

import argparse
import csv
import json
import os
from pathlib import Path
from datetime import datetime
from typing import Dict, List, Optional, Tuple


# ============================================================================
# Detector & metric configuration
# ============================================================================

ALL_METHODS = ['proposed', 'lipski', 'cacfar', 'oscfar', 'cusum']
METHOD_LABELS = {
    'proposed': 'TSNFA',
    'lipski':   'Lipski',
    'cacfar':   'CA-CFAR',
    'oscfar':   'OS-CFAR',
    'cusum':    'CUSUM',
}

# Metrics to extract. Each entry: (key_in_json, display_name, format_hint)
METRICS = [
    ('events_detected',              'Events Detected',         '{:.0f}'),
    ('detection_rate',               'Detection Rate (%)',      '{:.1f}%'),
    ('miss_rate',                    'Miss Rate (%)',           '{:.1f}%'),
    ('fp_clusters_outside',          'FP Clusters',             '{:.0f}'),
    ('event_precision',              'Event Precision (%)',     '{:.1f}%'),
    ('false_alarm_rate_clusters',    'FAR clusters /hr/node',   '{:.2f}'),
    ('true_positives',               'TP frames',               '{:.0f}'),
    ('false_positives',              'FP frames',               '{:.0f}'),
    ('redundancy_factor',            'Redundancy (TP/event)',   '{:.1f}'),
    ('false_alarm_rate_frames',      'FAR frames /hr/node',     '{:.2f}'),
    ('frame_precision',              'Frame Precision (%)',     '{:.1f}%'),
    ('latency_mean_ms',              'Mean Latency (ms)',       '{:.1f}'),
    ('latency_99th_ms',              '99th %ile Latency (ms)',  '{:.1f}'),
    ('network_load_bytes_per_hour',  'Network Load (bytes/hr)', '{:.0f}'),
]


# ============================================================================
# Loading & parsing
# ============================================================================

def load_run_json(json_path: Path) -> Optional[Dict]:
    """Load and parse one simulation_results.json.

    Returns a dict with:
      - num_nodes, snr_db, duration_hours, preset, source_path
      - total_events
      - methods: dict[method_name -> dict[metric -> (mean, std)]]

    Returns None if the file is missing or malformed.
    """
    if not json_path.exists():
        return None

    try:
        with open(json_path, 'r') as f:
            data = json.load(f)
    except (json.JSONDecodeError, OSError) as e:
        print(f"  WARNING: could not parse {json_path}: {e}")
        return None

    sim_params = data.get('_simulation_parameters', {})
    snr_db = sim_params.get('events', {}).get('snr_db')
    preset = sim_params.get('preset')

    # Find the per-network-size results block (top-level numeric key)
    num_nodes = None
    inner = None
    for k, v in data.items():
        if isinstance(k, str) and k.isdigit():
            num_nodes = int(k)
            inner = v
            break

    if num_nodes is None or inner is None:
        print(f"  WARNING: no per-network-size block in {json_path}")
        return None

    duration_hours = inner.get('config', {}).get('duration_hours')
    total_events = inner.get('true_events', {}).get('total')

    # Pull each detector's metrics
    methods = {}
    for method_key in ALL_METHODS:
        if method_key not in inner:
            continue
        m = inner[method_key]
        per_metric = {}
        for metric_key, _, _ in METRICS:
            v = m.get(metric_key)
            if v is None:
                per_metric[metric_key] = (None, None)
                continue
            if isinstance(v, dict):
                per_metric[metric_key] = (v.get('mean'), v.get('std'))
            else:
                per_metric[metric_key] = (float(v), 0.0)
        methods[method_key] = per_metric

    return {
        'num_nodes':       num_nodes,
        'snr_db':          snr_db,
        'duration_hours':  duration_hours,
        'preset':          preset,
        'total_events':    total_events,
        'methods':         methods,
        'source_path':     str(json_path),
        'mtime':           json_path.stat().st_mtime,
    }


def find_backup_runs(backup_root: Path, num_to_keep: int) -> List[Path]:
    """Find the N most recent backup folders containing a simulation_results.json.

    Looks for <backup_root>/<anything>/data/simulation_results.json.
    Sorts by mtime of the JSON file, descending. Returns up to num_to_keep paths.
    """
    if not backup_root.exists():
        return []

    candidates = []
    for entry in backup_root.iterdir():
        if not entry.is_dir():
            continue
        json_path = entry / 'data' / 'simulation_results.json'
        if json_path.exists():
            candidates.append((json_path.stat().st_mtime, json_path))

    candidates.sort(key=lambda x: x[0], reverse=True)
    return [p for _, p in candidates[:num_to_keep]]


def collect_runs(data_dir: Path, backup_root: Path, num_backups: int) -> List[Dict]:
    """Collect (current_run + N backups) into a list of run dicts."""
    runs = []

    # Current run from data/
    current_json = data_dir / 'simulation_results.json'
    if current_json.exists():
        run = load_run_json(current_json)
        if run is not None:
            run['origin'] = 'current'
            runs.append(run)
            print(f"  Loaded current run: "
                  f"{run['num_nodes']}n {run['snr_db']}dB "
                  f"({run['duration_hours']}h, preset={run['preset']})")
    else:
        print(f"  No current run found at {current_json}")

    # Backups
    backup_paths = find_backup_runs(backup_root, num_backups)
    for p in backup_paths:
        run = load_run_json(p)
        if run is None:
            continue
        run['origin'] = p.parent.parent.name  # the timestamp folder name
        runs.append(run)
        print(f"  Loaded backup '{run['origin']}': "
              f"{run['num_nodes']}n {run['snr_db']}dB "
              f"({run['duration_hours']}h, preset={run['preset']})")

    return runs


# ============================================================================
# Output formatting
# ============================================================================

def column_label(run: Dict) -> str:
    """Build a short label for each run, suitable for table column heading."""
    nn = run.get('num_nodes', '?')
    snr = run.get('snr_db', '?')
    if isinstance(snr, float):
        snr = f"{snr:.0f}"
    origin = run.get('origin', '?')
    if origin == 'current':
        return f"{nn}n / {snr}dB (current)"
    # Use first 13 chars of timestamp folder, e.g. 20260509_2347
    short_origin = str(origin)[:13]
    return f"{nn}n / {snr}dB ({short_origin})"


def format_value(mean: Optional[float], std: Optional[float], fmt: str) -> str:
    """Format a (mean, std) pair as a string for the table."""
    if mean is None:
        return "-"
    s = fmt.format(mean)
    if std is not None and std > 1e-9:
        # Strip trailing % from mean for ± display, then re-attach
        if s.endswith('%'):
            s = s[:-1] + f" ± {std:.1f}%"
        else:
            s = s + f" ± {std:.1f}"
    return s


def print_compiled_table(runs: List[Dict]) -> None:
    """Print a wide compiled table. One column per run, one row per metric+method."""
    if not runs:
        print("\nNo runs to display.")
        return

    headers = ["Metric", "Detector"] + [column_label(r) for r in runs]

    # Column widths
    widths = [max(20, len(h)) for h in headers]
    for r_idx, run in enumerate(runs):
        col = r_idx + 2  # offset for Metric/Detector cols
        for method in ALL_METHODS:
            for metric_key, _, fmt in METRICS:
                m = run['methods'].get(method, {})
                mean, std = m.get(metric_key, (None, None))
                cell = format_value(mean, std, fmt)
                if len(cell) > widths[col]:
                    widths[col] = len(cell)
    widths = [min(w, 30) for w in widths]  # cap at 30

    # Header
    line_total = sum(widths) + 3 * len(widths) + 1
    print()
    print("=" * line_total)
    print("Compiled Results — Current Run + Backups")
    print("=" * line_total)
    print(" | ".join(h.ljust(w) for h, w in zip(headers, widths)))
    print("-" * line_total)

    # Body — group by metric then detector
    for metric_key, metric_display, fmt in METRICS:
        for method in ALL_METHODS:
            row = [metric_display, METHOD_LABELS[method]]
            for run in runs:
                m = run['methods'].get(method, {})
                mean, std = m.get(metric_key, (None, None))
                row.append(format_value(mean, std, fmt))
            print(" | ".join(c.ljust(w) for c, w in zip(row, widths)))
        print("-" * line_total)
    print("=" * line_total)


def write_csv(runs: List[Dict], csv_path: Path) -> None:
    """Write the compiled table to a CSV file."""
    headers = ["Metric", "Detector"] + [column_label(r) for r in runs]
    with open(csv_path, 'w', newline='', encoding='utf-8') as f:
        writer = csv.writer(f)
        writer.writerow(headers)
        for metric_key, metric_display, fmt in METRICS:
            for method in ALL_METHODS:
                row = [metric_display, METHOD_LABELS[method]]
                for run in runs:
                    m = run['methods'].get(method, {})
                    mean, std = m.get(metric_key, (None, None))
                    if mean is None:
                        row.append("")
                    elif std is not None and std > 1e-9:
                        row.append(f"{mean:.4f}±{std:.4f}")
                    else:
                        row.append(f"{mean:.4f}")
                writer.writerow(row)
    print(f"\n  Compiled CSV → {csv_path}")


def write_json(runs: List[Dict], json_path: Path) -> None:
    """Write all runs as a structured JSON for downstream tooling."""
    serialisable = []
    for run in runs:
        # Convert (mean, std) tuples to dicts for JSON
        methods_serial = {}
        for method, metrics in run['methods'].items():
            methods_serial[method] = {
                k: {'mean': v[0], 'std': v[1]}
                for k, v in metrics.items()
            }
        serialisable.append({
            'origin':          run.get('origin'),
            'source_path':     run.get('source_path'),
            'num_nodes':       run.get('num_nodes'),
            'snr_db':          run.get('snr_db'),
            'duration_hours':  run.get('duration_hours'),
            'preset':          run.get('preset'),
            'total_events':    run.get('total_events'),
            'mtime':           run.get('mtime'),
            'methods':         methods_serial,
        })
    with open(json_path, 'w') as f:
        json.dump({
            'extracted_at': datetime.now().isoformat(),
            'num_runs':     len(serialisable),
            'runs':         serialisable,
        }, f, indent=2)
    print(f"  Compiled JSON → {json_path}")


# ============================================================================
# CLI
# ============================================================================

def main():
    parser = argparse.ArgumentParser(
        description='Extract & compile results from current run + N recent backups',
        formatter_class=argparse.ArgumentDefaultsHelpFormatter,
    )
    parser.add_argument('--data-dir', default='U:/MONTECARLO/data',
                        help='Directory containing the current simulation_results.json')
    parser.add_argument('--backup-root', default='U:/MONTECARLO/backups',
                        help='Root containing timestamped backup folders')
    parser.add_argument('--num-backups', type=int, default=3,
                        help='Number of most-recent backups to include')
    parser.add_argument('--output-dir', default=None,
                        help='Directory for compiled CSV/JSON output (defaults to --data-dir)')
    parser.add_argument('--no-csv', action='store_true',
                        help='Skip writing the CSV file')
    parser.add_argument('--no-json', action='store_true',
                        help='Skip writing the JSON file')

    args = parser.parse_args()

    data_dir = Path(args.data_dir).resolve()
    backup_root = Path(args.backup_root).resolve()
    output_dir = Path(args.output_dir).resolve() if args.output_dir else data_dir

    print('=' * 75)
    print('Extract & Compile Simulation Results')
    print('=' * 75)
    print(f"  Data dir:      {data_dir}")
    print(f"  Backup root:   {backup_root}")
    print(f"  Backups kept:  {args.num_backups} most recent")
    print(f"  Output dir:    {output_dir}")
    print()

    # Collect
    print("Loading runs...")
    runs = collect_runs(data_dir, backup_root, args.num_backups)
    if not runs:
        print("\nNo runs found. Check --data-dir and --backup-root.")
        return 1

    # Print compiled table
    print_compiled_table(runs)

    # Write outputs
    output_dir.mkdir(parents=True, exist_ok=True)
    if not args.no_csv:
        write_csv(runs, output_dir / 'compiled_results.csv')
    if not args.no_json:
        write_json(runs, output_dir / 'compiled_results.json')

    print('\nDone.')
    return 0


if __name__ == '__main__':
    raise SystemExit(main())