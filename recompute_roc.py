#!/usr/bin/env python3
"""
recompute_roc.py - Rebuild ROC curves from saved strength data.

Reads a *_strengths.npz file produced by simulator.py during a Monte Carlo run
and recomputes the ROC curves + plot using the corrected sweep algorithm.
This avoids re-running the simulation when only the ROC analysis needs to
be redone.

Usage
-----
    # Default - load U:/MONTECARLO/data/simulation_results_10nodes_strengths.npz
    python recompute_roc.py

    # Specify input file explicitly
    python recompute_roc.py --input U:/MONTECARLO/data/simulation_results_10nodes_strengths.npz

    # Save plot to a different location
    python recompute_roc.py --output U:/MONTECARLO/data/recomputed_roc.png

    # Adjust the FP-clustering window (default 5 sec)
    python recompute_roc.py --fp-cluster-window 10.0

    # Different number of operating points
    python recompute_roc.py --num-points 50

The .npz file must contain:
  - strengths_<method>:   (N, 3) array of (time, node_id, strength) per method
  - events_node<id>:      array of event start times for each node
  - _config_json:         single-element JSON string with simulator config

Author: GNACODE INC, January 2026
"""

import argparse
import json
import os
import sys
from collections import defaultdict
from pathlib import Path

import numpy as np
import matplotlib.pyplot as plt


# =============================================================================
# Defaults
# =============================================================================

DEFAULT_INPUT = 'U:/MONTECARLO/data/simulation_results_10nodes_strengths.npz'
FP_CLUSTER_WINDOW_SEC = 5.0
NUM_POINTS = 25
FP_FLOOR = 0.001  # for log-scale plotting when actual FP rate is zero

METHOD_LABELS = {
    'proposed': 'TSNFA',
    'lipski':   'Lipski FFT',
    'cacfar':   'CA-CFAR',
    'oscfar':   'OS-CFAR',
    'cusum':    'CUSUM',
    'cdr100':   'CDR-100*',
    'ast100':   'AST-100*',
}
COLORS = {
    'proposed': '#2ecc71', 'lipski': '#3498db', 'cacfar': '#e74c3c',
    'oscfar': '#f39c12', 'cusum': '#9b59b6',
    'cdr100': '#1abc9c', 'ast100': '#34495e',
}
MARKERS = {
    'proposed': 'o', 'lipski': 's', 'cacfar': '^', 'oscfar': 'D',
    'cusum': 'v', 'cdr100': 'P', 'ast100': 'X',
}
DRAW_ORDER = ['lipski', 'cacfar', 'oscfar', 'cusum',
              'cdr100', 'ast100', 'proposed']


# =============================================================================
# Loading
# =============================================================================

def load_strengths_npz(path):
    """Load a strengths .npz file and return (frame_strengths, true_event_times, config)."""
    if not Path(path).exists():
        raise FileNotFoundError(f"Strength file not found: {path}")

    data = np.load(path, allow_pickle=False)

    # Extract frame_strengths
    frame_strengths = {}
    for key in data.files:
        if key.startswith('strengths_'):
            method = key[len('strengths_'):]
            arr = data[key]
            frame_strengths[method] = arr

    # Extract true_event_times
    true_event_times = {}
    for key in data.files:
        if key.startswith('events_node'):
            node_id = int(key[len('events_node'):])
            true_event_times[node_id] = data[key]

    # Extract config
    config = {}
    if '_config_json' in data.files:
        config_str = str(data['_config_json'][0])
        config = json.loads(config_str)

    return frame_strengths, true_event_times, config


# =============================================================================
# ROC computation
# =============================================================================

def compute_roc(frame_strengths, true_event_times, config,
                num_points=NUM_POINTS, fp_cluster_window=FP_CLUSTER_WINDOW_SEC,
                verbose=True):
    """Compute ROC sweep for all methods.

    Returns dict: method -> {
        'thresholds':           list of threshold values
        'event_dr':             monotonic event DR (Pareto envelope)
        'event_dr_raw':         raw (non-monotonic) DR
        'fp_per_hour_per_node': FP cluster rates
        'fp_clusters_total':   total FP cluster counts
    }
    """
    frame_duration = config.get('frame_duration', 1.28)
    event_duration = config.get('event_duration', 5.0)
    gamma_d = config.get('gamma_d', 3)
    sim_duration = config.get('simulation_duration_sec', 86400.0)
    num_nodes = config.get('num_nodes', 10)
    num_sensor_nodes = max(1, num_nodes - 1)
    hours = sim_duration / 3600.0

    # Detection window
    margin = frame_duration
    window_before = margin
    window_after = event_duration + gamma_d * frame_duration + margin

    # Build per-node sorted event windows
    event_windows_by_node = {}
    for node_id, event_times in true_event_times.items():
        windows = sorted(
            (et - window_before, et + window_after) for et in event_times
        )
        event_windows_by_node[int(node_id)] = windows

    # Flat event list
    all_events = []
    for node_id, event_times in true_event_times.items():
        windows = event_windows_by_node[int(node_id)]
        for et, (ws, we) in zip(event_times, windows):
            all_events.append((int(node_id), float(ws), float(we)))
    total_true_events = len(all_events)

    if verbose:
        print(f"  Loaded {total_true_events} events across {len(event_windows_by_node)} nodes")
        print(f"  Simulation duration: {hours:.1f} hr, sensor nodes: {num_sensor_nodes}")
        print(f"  Detection window: [-{window_before:.2f}, +{window_after:.2f}] sec")

    def _trigger_in_any_event_window(node_id, t):
        windows = event_windows_by_node.get(node_id, [])
        for ws, we in windows:
            if t < ws:
                return False
            if t <= we:
                return True
        return False

    roc_data = {}

    for method, arr_2d in frame_strengths.items():
        if arr_2d.shape[0] < 10:
            continue

        times = arr_2d[:, 0]
        node_ids = arr_2d[:, 1].astype(np.int32)
        strengths = arr_2d[:, 2]

        valid_mask = strengths > 0
        if valid_mask.sum() < 10:
            continue
        valid_strengths = strengths[valid_mask]

        # Threshold range
        lo = float(np.quantile(valid_strengths, 0.001))
        hi = float(np.quantile(valid_strengths, 0.9999))
        if lo <= 0:
            lo = max(1e-6, valid_strengths.min())
        if hi <= lo:
            hi = lo * 10
        thresholds = list(np.logspace(np.log10(lo), np.log10(hi), num_points))
        if lo < 1.0 < hi:
            thresholds.append(1.0)
        thresholds.append(lo * 0.1)
        thresholds.append(hi * 10)
        thresholds = sorted(set(float(t) for t in thresholds))

        # Sort by node, then time
        sort_idx = np.lexsort((times, node_ids))
        node_ids_sorted = node_ids[sort_idx]
        times_sorted = times[sort_idx]
        strengths_sorted = strengths[sort_idx]

        # Slice indices per node
        unique_nodes, node_starts = np.unique(node_ids_sorted, return_index=True)
        node_ends = np.concatenate([node_starts[1:], [len(node_ids_sorted)]])
        node_slices = {
            int(n): (int(s), int(e))
            for n, s, e in zip(unique_nodes, node_starts, node_ends)
        }

        event_dr_list = []
        fp_per_hr_list = []
        fp_clusters_list = []

        for th in thresholds:
            # Per-event detection
            events_det = 0
            for node_id, ws, we in all_events:
                sl = node_slices.get(node_id)
                if sl is None:
                    continue
                s_idx, e_idx = sl
                sub_times = times_sorted[s_idx:e_idx]
                sub_strengths = strengths_sorted[s_idx:e_idx]
                lo_idx = np.searchsorted(sub_times, ws, side='left')
                hi_idx = np.searchsorted(sub_times, we, side='right')
                if lo_idx == hi_idx:
                    continue
                if (sub_strengths[lo_idx:hi_idx] > th).any():
                    events_det += 1

            # FP cluster counting
            fp_clusters_count = 0
            for node_id, (s_idx, e_idx) in node_slices.items():
                sub_times = times_sorted[s_idx:e_idx]
                sub_strengths = strengths_sorted[s_idx:e_idx]
                trig_mask = sub_strengths > th
                if not trig_mask.any():
                    continue
                trig_times_node = sub_times[trig_mask]

                last_fp_time = -1e18
                for t in trig_times_node:
                    if _trigger_in_any_event_window(node_id, t):
                        continue
                    if (t - last_fp_time) >= fp_cluster_window:
                        fp_clusters_count += 1
                    last_fp_time = t

            event_dr = (events_det / total_true_events * 100
                        if total_true_events > 0 else 0.0)
            fp_per_hr = (fp_clusters_count / hours / num_sensor_nodes
                         if num_sensor_nodes > 0 else 0.0)
            event_dr_list.append(event_dr)
            fp_per_hr_list.append(fp_per_hr)
            fp_clusters_list.append(fp_clusters_count)

        # Sort by FP, apply upper-envelope (Pareto frontier)
        order = np.argsort(fp_per_hr_list)
        fp_sorted = [fp_per_hr_list[i] for i in order]
        dr_sorted = [event_dr_list[i] for i in order]
        th_sorted = [thresholds[i] for i in order]
        fp_clusters_sorted = [fp_clusters_list[i] for i in order]

        dr_envelope = []
        running_max = -1.0
        for d in dr_sorted:
            if d > running_max:
                running_max = d
            dr_envelope.append(running_max)

        roc_data[method] = {
            'thresholds': th_sorted,
            'event_dr_raw': dr_sorted,
            'event_dr': dr_envelope,
            'fp_per_hour_per_node': fp_sorted,
            'fp_clusters_total': fp_clusters_sorted,
        }

        if verbose:
            print(f"  {method:8s}: {len(thresholds):3d} pts, "
                  f"DR [{min(dr_envelope):5.1f}, {max(dr_envelope):5.1f}]%, "
                  f"FP [{min(fp_sorted):.3f}, {max(fp_sorted):.2f}]/hr/node")

    return roc_data


# =============================================================================
# Plotting
# =============================================================================

def plot_roc(roc_data, save_path, snr_db=None):
    """Render the ROC plot. Same style as simulator.py's plot_roc_curves."""
    fig, ax = plt.subplots(figsize=(11, 7))

    for method in DRAW_ORDER:
        if method not in roc_data:
            continue
        data = roc_data[method]
        if not data['fp_per_hour_per_node']:
            continue

        fp_arr = np.array(data['fp_per_hour_per_node'], dtype=float)
        dr_arr = np.array(data['event_dr'], dtype=float)
        th_arr = np.array(data['thresholds'], dtype=float)

        fp_plot = np.maximum(fp_arr, FP_FLOOR)

        label = METHOD_LABELS.get(method, method)
        color = COLORS.get(method, '#777777')
        marker = MARKERS.get(method, 'o')

        is_proposed = (method == 'proposed')
        ax.plot(fp_plot, dr_arr,
                marker=marker, color=color, label=label,
                linewidth=2.5 if is_proposed else 1.8,
                markersize=8 if is_proposed else 6,
                alpha=0.95 if is_proposed else 0.85,
                zorder=10 if is_proposed else 5)

        # Canonical operating point (threshold closest to 1.0)
        canonical_idx = int(np.argmin(np.abs(th_arr - 1.0)))
        ax.plot(fp_plot[canonical_idx], dr_arr[canonical_idx],
                marker='*', color=color, markersize=18,
                markeredgecolor='black', markeredgewidth=1.5,
                zorder=15)

    ax.set_xscale('log')
    ax.set_xlim([FP_FLOOR * 0.5, 1e3])
    ax.set_ylim([-2, 105])

    ax.set_xlabel('False-alarm cluster rate (per hour per node)', fontsize=12)
    ax.set_ylabel('Event detection rate (%)', fontsize=12)
    title = 'ROC Curves — TSNFA + 6 Comparators'
    if snr_db is not None:
        title += f'  (SNR = {snr_db:.0f} dB)'
    title += '\n★ marks canonical operating point  •  * = statistical proxy'
    ax.set_title(title, fontsize=13, fontweight='bold')
    ax.grid(True, which='both', alpha=0.3)
    ax.legend(loc='lower right', fontsize=10, framealpha=0.95)

    ax.annotate('FP=0\n(floored\nto 10⁻³)',
                xy=(FP_FLOOR, 100), xytext=(FP_FLOOR * 0.7, 80),
                fontsize=8, color='gray', ha='center',
                arrowprops=dict(arrowstyle='->', color='gray', lw=0.8))
    ax.axhline(y=100, color='gray', linestyle='--', alpha=0.4, linewidth=0.8)

    plt.tight_layout()
    plt.savefig(save_path, dpi=150, bbox_inches='tight')
    print(f"\n  ROC plot saved to: {save_path}")
    plt.close(fig)


# =============================================================================
# CLI
# =============================================================================

def main():
    parser = argparse.ArgumentParser(
        description='Recompute ROC curves from saved strength data',
        formatter_class=argparse.ArgumentDefaultsHelpFormatter,
    )
    parser.add_argument('--input', default=DEFAULT_INPUT,
                        help='Path to *_strengths.npz file from a Monte Carlo run')
    parser.add_argument('--output', default=None,
                        help='Path to save ROC plot (defaults to <input>_recomputed_roc.png)')
    parser.add_argument('--num-points', type=int, default=NUM_POINTS,
                        help='Number of threshold operating points to evaluate')
    parser.add_argument('--fp-cluster-window', type=float, default=FP_CLUSTER_WINDOW_SEC,
                        help='Seconds for clustering consecutive FPs into one event')
    parser.add_argument('--save-roc-json', default=None,
                        help='Optional path to save the recomputed ROC arrays as JSON')

    args = parser.parse_args()

    input_path = Path(args.input).resolve()
    if args.output is None:
        # Default: same dir as input, with _recomputed_roc.png suffix
        out_dir = input_path.parent
        stem = input_path.stem.replace('_strengths', '_recomputed_roc')
        output_path = out_dir / f"{stem}.png"
    else:
        output_path = Path(args.output).resolve()

    print('=' * 75)
    print('Recompute ROC from saved strength data')
    print('=' * 75)
    print(f"  Input:        {input_path}")
    print(f"  Output:       {output_path}")
    print(f"  Num points:   {args.num_points}")
    print(f"  FP clust win: {args.fp_cluster_window}s")
    print()

    # Load
    print("Loading strength data...")
    frame_strengths, true_event_times, config = load_strengths_npz(input_path)

    sizes = {m: arr.shape[0] for m, arr in frame_strengths.items()}
    total_frames = sum(sizes.values())
    print(f"  Loaded {len(frame_strengths)} detectors, "
          f"{total_frames:,} total frame-strength records")
    for m, n in sorted(sizes.items()):
        print(f"    {m:8s}: {n:,} records")
    print()

    # Compute ROC
    print("Computing ROC sweep...")
    roc_data = compute_roc(
        frame_strengths, true_event_times, config,
        num_points=args.num_points,
        fp_cluster_window=args.fp_cluster_window,
        verbose=True,
    )

    # Plot
    snr_db = config.get('event_snr_db')
    plot_roc(roc_data, output_path, snr_db=snr_db)

    # Optionally save the recomputed ROC arrays as JSON
    if args.save_roc_json:
        json_path = Path(args.save_roc_json).resolve()
        # Convert any numpy types to native Python for JSON
        def _convert(o):
            if isinstance(o, dict):
                return {k: _convert(v) for k, v in o.items()}
            if isinstance(o, list):
                return [_convert(x) for x in o]
            if isinstance(o, np.integer):
                return int(o)
            if isinstance(o, np.floating):
                return float(o)
            if isinstance(o, np.ndarray):
                return _convert(o.tolist())
            return o

        with open(json_path, 'w') as f:
            json.dump(_convert(roc_data), f, indent=2)
        print(f"  ROC arrays saved to: {json_path}")

    print('=' * 75)
    print('Done.')


if __name__ == '__main__':
    main()